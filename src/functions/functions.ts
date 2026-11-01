import { getStoredAuth, ensureValidToken } from "../taskpane/services/lookerAuth";
import { runInlineQuery, QueryPayload } from "../taskpane/services/queryTaskRunner";

export type QueryRunnerFn = (
  baseUrl: string,
  token: string,
  payload: QueryPayload
) => Promise<Record<string, any>[]>;

interface CacheEntry {
  timestamp: number;
  value: any;
}

const VALUE_CACHE = new Map<string, CacheEntry>();
const QUERY_CACHE = new Map<string, CacheEntry>();
const DEFAULT_CACHE_TTL_MS = 60_000;

export function clearCustomFunctionCache(): void {
  VALUE_CACHE.clear();
  QUERY_CACHE.clear();
}

/**
 * Parses flexible filter inputs passed from Excel formulas:
 * - Alternating pairs: ["orders.region", "West", "orders.year", "2025"]
 * - 2D range pairs: [[["orders.region", "West"], ["orders.year", "2025"]]]
 * - Key=value semicolon/comma string: "orders.region=West;orders.year=2025"
 * - JSON object string: '{"orders.region":"West"}'
 */
export function parseFormulaFilters(rawArgs: any[]): Record<string, string> {
  const filters: Record<string, string> = {};
  if (!rawArgs || rawArgs.length === 0) return filters;

  // Flatten 2D range arrays if passed as a single range argument
  const flat: string[] = [];
  for (const arg of rawArgs) {
    if (arg === null || arg === undefined || arg === "") continue;
    if (Array.isArray(arg)) {
      for (const row of arg) {
        if (Array.isArray(row)) {
          if (row.length >= 2 && row[0] !== null && row[0] !== undefined && String(row[0]).trim() !== "") {
            filters[String(row[0]).trim()] = String(row[1] ?? "").trim();
          } else if (row.length === 1 && row[0] !== null && row[0] !== undefined) {
            flat.push(String(row[0]).trim());
          }
        } else if (row !== null && row !== undefined) {
          flat.push(String(row).trim());
        }
      }
    } else {
      flat.push(String(arg).trim());
    }
  }

  if (flat.length === 1) {
    const single = flat[0];
    if (!single) return filters;
    if (single.startsWith("{")) {
      try {
        const parsed = JSON.parse(single);
        for (const [k, v] of Object.entries(parsed)) {
          if (v !== undefined && v !== null) {
            filters[k.trim()] = String(v).trim();
          }
        }
        return filters;
      } catch {}
    }
    if (single.includes("=")) {
      const pairs = single.split(/[;|]/);
      for (const pair of pairs) {
        const eqIdx = pair.indexOf("=");
        if (eqIdx > 0) {
          const k = pair.slice(0, eqIdx).trim();
          const v = pair.slice(eqIdx + 1).trim();
          if (k) filters[k] = v;
        }
      }
      return filters;
    }
  }

  for (let i = 0; i + 1 < flat.length; i += 2) {
    const k = flat[i];
    const v = flat[i + 1];
    if (k) {
      filters[k] = v ?? "";
    }
  }

  return filters;
}

interface PendingValueRequest {
  primaryVal: string;
  cacheKey: string;
  resolve: (val: any) => void;
  reject: (err: any) => void;
}

interface BatchGroup {
  model: string;
  explore: string;
  measure: string;
  primaryDim: string | null;
  staticFilters: Record<string, string>;
  requests: PendingValueRequest[];
  timer: ReturnType<typeof setTimeout> | null;
}

/**
 * Coalesces concurrent `=LOOKER.VALUE(...)` formula calls made within the same calculation tick
 * into a single grouped Looker query per (model, explore, measure, primaryDim, staticFilters).
 */
export class LookerValueBatcher {
  private groups = new Map<string, BatchGroup>();
  private windowMs: number;
  private runner: QueryRunnerFn;
  private baseUrlProvider: () => Promise<{ baseUrl: string; accessToken: string }>;

  constructor(
    options?: {
      windowMs?: number;
      runner?: QueryRunnerFn;
      baseUrlProvider?: () => Promise<{ baseUrl: string; accessToken: string }>;
    }
  ) {
    this.windowMs = options?.windowMs ?? 80;
    this.runner = options?.runner ?? runInlineQuery;
    this.baseUrlProvider =
      options?.baseUrlProvider ??
      (async () => {
        const stored = getStoredAuth();
        if (!stored) {
          throw new Error("Not signed in to Looker. Open the Looker taskpane to sign in.");
        }
        const valid = await ensureValidToken(stored);
        return { baseUrl: valid.baseUrl, accessToken: valid.accessToken };
      });
  }

  public requestValue(
    model: string,
    explore: string,
    measure: string,
    filters: Record<string, string>
  ): Promise<any> {
    const filterEntries = Object.entries(filters).filter(([k, v]) => k && v !== "");
    const primaryDim = filterEntries.length > 0 ? filterEntries[0][0] : null;
    const primaryVal = filterEntries.length > 0 ? filterEntries[0][1] : "";

    const staticFilters: Record<string, string> = {};
    for (let i = 1; i < filterEntries.length; i++) {
      staticFilters[filterEntries[i][0]] = filterEntries[i][1];
    }

    const staticKey = JSON.stringify(
      Object.entries(staticFilters).sort(([a], [b]) => a.localeCompare(b))
    );
    const cacheKey = `${model}::${explore}::${measure}::${primaryDim || ""}::${primaryVal}::${staticKey}`;

    const cached = VALUE_CACHE.get(cacheKey);
    if (cached && Date.now() - cached.timestamp < DEFAULT_CACHE_TTL_MS) {
      return Promise.resolve(cached.value);
    }

    const groupKey = `${model}::${explore}::${measure}::${primaryDim || "__SCALAR__"}::${staticKey}`;
    let group = this.groups.get(groupKey);
    if (!group) {
      group = {
        model,
        explore,
        measure,
        primaryDim,
        staticFilters,
        requests: [],
        timer: null,
      };
      this.groups.set(groupKey, group);
    }

    return new Promise((resolve, reject) => {
      group!.requests.push({ primaryVal, cacheKey, resolve, reject });
      if (!group!.timer) {
        group!.timer = setTimeout(() => {
          this.flushGroup(groupKey);
        }, this.windowMs);
      }
    });
  }

  private async flushGroup(groupKey: string): Promise<void> {
    const group = this.groups.get(groupKey);
    if (!group) return;
    this.groups.delete(groupKey);

    try {
      const { baseUrl, accessToken } = await this.baseUrlProvider();

      if (!group.primaryDim) {
        // Scalar query (no dimension filter, or static filters only)
        const payload: QueryPayload = {
          model: group.model,
          view: group.explore,
          fields: [group.measure],
          filters:
            Object.keys(group.staticFilters).length > 0 ? group.staticFilters : null,
          limit: "1",
        };
        const rows = await this.runner(baseUrl, accessToken, payload);
        const val = rows?.[0]?.[group.measure] ?? 0;
        for (const req of group.requests) {
          VALUE_CACHE.set(req.cacheKey, { timestamp: Date.now(), value: val });
          req.resolve(val);
        }
        return;
      }

      // Coalesce all unique primaryDim values requested in this batch window into one query
      const uniqueVals = Array.from(
        new Set(group.requests.map((r) => r.primaryVal).filter((v) => v !== ""))
      );
      const combinedFilters: Record<string, string> = {
        ...group.staticFilters,
        [group.primaryDim]: uniqueVals.join(","),
      };

      const payload: QueryPayload = {
        model: group.model,
        view: group.explore,
        fields: [group.primaryDim, group.measure],
        filters: combinedFilters,
        limit: String(Math.max(500, uniqueVals.length * 2)),
      };

      const rows = await this.runner(baseUrl, accessToken, payload);
      const lookup = new Map<string, any>();
      for (const r of rows || []) {
        const dimKey = String(r[group.primaryDim] ?? "").toLowerCase();
        lookup.set(dimKey, r[group.measure]);
      }

      for (const req of group.requests) {
        const matched = lookup.get(req.primaryVal.toLowerCase()) ?? 0;
        VALUE_CACHE.set(req.cacheKey, { timestamp: Date.now(), value: matched });
        req.resolve(matched);
      }
    } catch (err: any) {
      for (const req of group.requests) {
        req.reject(err);
      }
    }
  }
}

export const defaultValueBatcher = new LookerValueBatcher();

/**
 * Excel Custom Function: `=LOOKER.VALUE(model, explore, measure, [filterField1, filterVal1, ...])`
 * Evaluates a single governed Looker metric (supporting non-additive measures) with automatic micro-batching.
 */
export async function lookerValue(
  model: string,
  explore: string,
  measure: string,
  ...filterArgs: any[]
): Promise<any> {
  if (!model || !explore || !measure) {
    return "#VALUE!";
  }
  const filters = parseFormulaFilters(filterArgs);
  return await defaultValueBatcher.requestValue(
    String(model).trim(),
    String(explore).trim(),
    String(measure).trim(),
    filters
  );
}

/**
 * Excel Custom Function: `=LOOKER.QUERY(model, explore, fieldsCsv, [filters], [sortsCsv], [limit])`
 * Evaluates a Looker query and spills a 2D array (header row + data rows) into the worksheet.
 */
export async function lookerQuery(
  model: string,
  explore: string,
  fieldsCsv: string,
  filtersArg?: any,
  sortsCsv?: string,
  limit?: number | string,
  customRunner?: QueryRunnerFn,
  customAuthProvider?: () => Promise<{ baseUrl: string; accessToken: string }>
): Promise<(string | number | boolean)[][]> {
  if (!model || !explore || !fieldsCsv) {
    return [["#VALUE!"]];
  }

  const fields = String(fieldsCsv)
    .split(",")
    .map((f) => f.trim())
    .filter(Boolean);
  if (fields.length === 0) {
    return [["#VALUE!"]];
  }

  const filters = filtersArg ? parseFormulaFilters([filtersArg]) : {};
  const sorts = sortsCsv
    ? String(sortsCsv)
        .split(",")
        .map((s) => s.trim())
        .filter(Boolean)
    : [];
  const limitStr = limit !== undefined && limit !== null && String(limit).trim() !== ""
    ? String(limit).trim()
    : "500";

  const cacheKey = JSON.stringify({ model, explore, fields, filters, sorts, limitStr });
  const cached = QUERY_CACHE.get(cacheKey);
  if (cached && Date.now() - cached.timestamp < DEFAULT_CACHE_TTL_MS) {
    return cached.value;
  }

  const authProvider =
    customAuthProvider ??
    (async () => {
      const stored = getStoredAuth();
      if (!stored) {
        throw new Error("Not signed in to Looker. Open the Looker taskpane to sign in.");
      }
      const valid = await ensureValidToken(stored);
      return { baseUrl: valid.baseUrl, accessToken: valid.accessToken };
    });

  const { baseUrl, accessToken } = await authProvider();
  const runner = customRunner ?? runInlineQuery;

  const rows = await runner(baseUrl, accessToken, {
    model: String(model).trim(),
    view: String(explore).trim(),
    fields,
    filters: Object.keys(filters).length > 0 ? filters : null,
    sorts,
    limit: limitStr,
  });

  const headerRow = fields;
  const dataMatrix: (string | number | boolean)[][] = [headerRow];
  for (const r of rows || []) {
    dataMatrix.push(
      fields.map((f) => {
        const val = r[f];
        if (val === null || val === undefined) return "";
        if (typeof val === "object") return JSON.stringify(val);
        return val;
      })
    );
  }

  QUERY_CACHE.set(cacheKey, { timestamp: Date.now(), value: dataMatrix });
  return dataMatrix;
}

/**
 * Registers `LOOKER.VALUE` and `LOOKER.QUERY` with the Office.js CustomFunctions runtime when present.
 */
export function registerLookerCustomFunctions(): void {
  try {
    const cf = (globalThis as any).CustomFunctions;
    if (cf && typeof cf.associate === "function") {
      cf.associate("VALUE", lookerValue);
      cf.associate("QUERY", lookerQuery);
    }
  } catch (e) {
    console.warn("CustomFunctions association skipped:", e);
  }
}
