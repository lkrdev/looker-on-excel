export interface QueryPayload {
  model: string;
  view: string;
  fields: string[];
  filters?: Record<string, string> | null;
  sorts?: string[];
  limit?: string;
  pivots?: string[] | null;
  total?: boolean;
  row_total?: "right" | "left" | null;
  subtotals?: string[] | null;
}

export interface PivotDescriptor {
  key: string;
  is_total?: boolean;
  label?: string;
  data?: Record<string, string>;
}

export interface QueryExecutionBundle {
  rows: Record<string, any>[];
  totalsData?: Record<string, any> | null;
  subtotalsData?: Record<string, Record<string, any>[]> | null;
  pivotsMeta?: PivotDescriptor[] | null;
  resultFormat: "json" | "json_detail" | "split_json_and_total";
}

export interface QueryTaskController {
  taskId: string;
  cancel: () => Promise<void>;
  waitForResults: (onStatus: (status: string, elapsedSeconds: number) => void) => Promise<any[]>;
  waitForBundle: (
    onStatus: (status: string, elapsedSeconds: number) => void
  ) => Promise<QueryExecutionBundle>;
}

export type ResultFormatStrategy = "json" | "json_detail" | "split_json_and_total";

/**
 * Determines whether to use lightweight `json`, full `json_detail`, or a split
 * `json` + 1-row `json_detail` total query when row count >= 5,000.
 */
export function resolveResultFormatStrategy(
  queryPayload: QueryPayload,
  highVolumeThreshold: number = 5000
): ResultFormatStrategy {
  const hasTotal = Boolean(queryPayload.total);
  const hasRowTotal = Boolean(queryPayload.row_total);
  const hasSubtotals = Boolean(queryPayload.subtotals && queryPayload.subtotals.length > 0);
  const hasMultiPivots = Boolean(queryPayload.pivots && queryPayload.pivots.length > 1);

  if (!hasTotal && !hasRowTotal && !hasSubtotals && !hasMultiPivots) {
    return "json";
  }

  // If only Column Grand Total (total: true) is requested on a large query (5k+ rows or -1),
  // avoid wrapping 5,000+ rows in verbose json_detail cell objects. Stream rows via "json"
  // and fetch the 1-row totals_data object via a lightweight limit="1" json_detail query.
  if (hasTotal && !hasRowTotal && !hasSubtotals && !hasMultiPivots) {
    const numericLimit = Number(queryPayload.limit || "500");
    if (queryPayload.limit === "-1" || (Number.isFinite(numericLimit) && numericLimit >= highVolumeThreshold)) {
      return "split_json_and_total";
    }
  }

  return "json_detail";
}

/**
 * Unwraps a Looker `json_detail` cell or pivoted cell map into the flat / nested value structure
 * expected by `excelWriter.ts`.
 */
export function unwrapJsonDetailCell(
  cellOrMap: any,
  pivots?: string[] | null,
  pivotsMeta?: PivotDescriptor[] | null
): any {
  if (cellOrMap === null || cellOrMap === undefined) {
    return null;
  }
  if (typeof cellOrMap !== "object" || Array.isArray(cellOrMap)) {
    return cellOrMap;
  }

  // Direct non-pivoted cell object: { value: X, rendered?: string, ... }
  if ("value" in cellOrMap && Object.keys(cellOrMap).every((k) => k === "value" || k === "rendered" || k === "html" || k === "links" || k === "filterable_value")) {
    return cellOrMap.value;
  }

  const primaryPivotField = pivots && pivots.length > 0 ? pivots[0] : null;

  // Check if already nested by pivot dimension name: { "orders.created_year": { "2024": { value: 100 } } }
  const keys = Object.keys(cellOrMap);
  if (
    primaryPivotField &&
    keys.includes(primaryPivotField) &&
    cellOrMap[primaryPivotField] &&
    typeof cellOrMap[primaryPivotField] === "object" &&
    !("value" in cellOrMap[primaryPivotField])
  ) {
    const innerObj = cellOrMap[primaryPivotField];
    const normalizedInner: Record<string, any> = {};
    for (const [pKey, pCell] of Object.entries(innerObj)) {
      normalizedInner[pKey] =
        pCell && typeof pCell === "object" && "value" in (pCell as any)
          ? (pCell as any).value
          : pCell;
    }
    const result: Record<string, any> = {
      [primaryPivotField]: normalizedInner,
    };
    if ("$$$_row_total_$$$" in cellOrMap) {
      const rt = cellOrMap["$$$_row_total_$$$"];
      result["$$$_row_total_$$$"] =
        rt && typeof rt === "object" && "value" in rt ? rt.value : rt;
    }
    return result;
  }

  // Standard Looker `json_detail` pivoted measure cell:
  // { "2024": { value: 100 }, "2025": { value: 200 }, "$$$_row_total_$$$": { value: 300 } }
  const pivotValuesMap: Record<string, any> = {};
  let rowTotalVal: any = undefined;
  let hasRowTotalKey = false;

  const orderedKeys =
    pivotsMeta && pivotsMeta.length > 0
      ? [
          ...pivotsMeta.map((pm) => pm.key).filter((k) => k in cellOrMap),
          ...keys.filter((k) => !pivotsMeta.some((pm) => pm.key === k)),
        ]
      : keys;

  for (const k of orderedKeys) {
    const entry = cellOrMap[k];
    const unwrapped =
      entry && typeof entry === "object" && "value" in entry ? entry.value : entry;
    if (k === "$$$_row_total_$$$") {
      hasRowTotalKey = true;
      rowTotalVal = unwrapped;
    } else {
      pivotValuesMap[k] = unwrapped;
    }
  }

  const effectivePivotField = primaryPivotField || "__pivot_key__";
  const wrapped: Record<string, any> = {
    [effectivePivotField]: pivotValuesMap,
  };
  if (hasRowTotalKey) {
    wrapped["$$$_row_total_$$$"] = rowTotalVal;
  }
  return wrapped;
}

export function normalizeJsonDetailRow(
  rawRow: Record<string, any>,
  pivots?: string[] | null,
  pivotsMeta?: PivotDescriptor[] | null
): Record<string, any> {
  const out: Record<string, any> = {};
  for (const [fieldKey, cell] of Object.entries(rawRow || {})) {
    out[fieldKey] = unwrapJsonDetailCell(cell, pivots, pivotsMeta);
  }
  return out;
}

/**
 * Normalizes either a raw `json` array or a `json_detail` response object into a unified QueryExecutionBundle.
 */
export function normalizeQueryResult(
  rawResult: any,
  pivots?: string[] | null,
  resultFormat: ResultFormatStrategy = "json"
): QueryExecutionBundle {
  if (Array.isArray(rawResult)) {
    return {
      rows: rawResult,
      totalsData: null,
      subtotalsData: null,
      pivotsMeta: null,
      resultFormat,
    };
  }

  if (rawResult && typeof rawResult === "object") {
    const pivotsMeta: PivotDescriptor[] | null = Array.isArray(rawResult.pivots)
      ? rawResult.pivots
      : null;

    const rawRows: any[] = Array.isArray(rawResult.data) ? rawResult.data : [];
    const rows = rawRows.map((r) => normalizeJsonDetailRow(r, pivots, pivotsMeta));

    const totalsData =
      rawResult.totals_data && typeof rawResult.totals_data === "object"
        ? normalizeJsonDetailRow(rawResult.totals_data, pivots, pivotsMeta)
        : null;

    let subtotalsData: Record<string, Record<string, any>[]> | null = null;
    if (rawResult.subtotals_data && typeof rawResult.subtotals_data === "object") {
      subtotalsData = {};
      for (const [groupKey, groupRows] of Object.entries(rawResult.subtotals_data)) {
        if (Array.isArray(groupRows)) {
          subtotalsData[groupKey] = groupRows.map((r) =>
            normalizeJsonDetailRow(r, pivots, pivotsMeta)
          );
        }
      }
    }

    return {
      rows,
      totalsData,
      subtotalsData,
      pivotsMeta,
      resultFormat: "json_detail",
    };
  }

  return {
    rows: [],
    totalsData: null,
    subtotalsData: null,
    pivotsMeta: null,
    resultFormat,
  };
}

/**
 * Executes a Looker query asynchronously via the Query Task API with live polling and cancellation.
 * Automatically selects lightweight `json` vs `json_detail` (or split total query for 5k+ rows).
 */
export async function runQueryTask(
  baseUrl: string,
  token: string,
  queryPayload: QueryPayload
): Promise<QueryTaskController> {
  const root = baseUrl.replace(/\/$/, "");
  const headers = {
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json",
  };

  const strategy = resolveResultFormatStrategy(queryPayload);
  const taskResultFormat = strategy === "json_detail" ? "json_detail" : "json";

  // 1. Create the immutable Looker Query object to obtain a query_id
  const queryRes = await fetch(`${root}/api/4.0/queries`, {
    method: "POST",
    headers,
    body: JSON.stringify(queryPayload),
  });

  if (!queryRes.ok) {
    const errText = await queryRes.text();
    throw new Error(`Failed to create query (${queryRes.status}): ${errText}`);
  }

  const queryObj = await queryRes.json();
  const queryId = String(queryObj.id);

  // 2. Launch asynchronous Query Task with query_id
  const createRes = await fetch(`${root}/api/4.0/query_tasks`, {
    method: "POST",
    headers,
    body: JSON.stringify({
      query_id: queryId,
      result_format: taskResultFormat,
    }),
  });

  if (!createRes.ok) {
    const errText = await createRes.text();
    throw new Error(`Failed to create query task (${createRes.status}): ${errText}`);
  }

  const { id: taskId } = await createRes.json();
  let isCanceled = false;
  const abortController = new AbortController();

  const cancel = async () => {
    isCanceled = true;
    abortController.abort();
  };

  const waitForBundle = async (
    onStatus: (status: string, elapsedSeconds: number) => void
  ): Promise<QueryExecutionBundle> => {
    const startTime = Date.now();

    while (!isCanceled) {
      const elapsed = Math.round((Date.now() - startTime) / 1000);

      try {
        const statusRes = await fetch(`${root}/api/4.0/query_tasks/${taskId}`, {
          headers,
          signal: abortController.signal,
        });
        if (!statusRes.ok) {
          throw new Error(`Failed to check query task status (${statusRes.status})`);
        }

        const taskInfo = await statusRes.json();
        const status = taskInfo.status || "running";
        onStatus(status, elapsed);

        if (status === "complete") {
          const resultsRes = await fetch(
            `${root}/api/4.0/query_tasks/${taskId}/results`,
            { headers, signal: abortController.signal }
          );
          if (!resultsRes.ok) {
            throw new Error(`Failed to fetch query results (${resultsRes.status})`);
          }
          const rawMain = await resultsRes.json();
          const bundle = normalizeQueryResult(rawMain, queryPayload.pivots, strategy);

          // If using split_json_and_total on a 5k+ row query, fetch the 1-row totals_data via inline json_detail
          if (strategy === "split_json_and_total" && queryPayload.total) {
            try {
              const totalOnlyPayload: QueryPayload = {
                ...queryPayload,
                limit: "1",
                total: true,
              };
              const totalRes = await fetch(`${root}/api/4.0/queries/run/json_detail`, {
                method: "POST",
                headers,
                body: JSON.stringify(totalOnlyPayload),
                signal: abortController.signal,
              });
              if (totalRes.ok) {
                const rawTotalDetail = await totalRes.json();
                const normalizedTotal = normalizeQueryResult(
                  rawTotalDetail,
                  queryPayload.pivots,
                  "json_detail"
                );
                bundle.totalsData = normalizedTotal.totalsData;
                bundle.resultFormat = "split_json_and_total";
              }
            } catch (e) {
              console.warn("Companion Grand Total query failed:", e);
            }
          }

          return bundle;
        } else if (status === "error" || status === "failed" || status === "killed") {
          throw new Error(`Looker query failed with status: ${status}`);
        }
      } catch (err: any) {
        if (isCanceled || err?.name === "AbortError") {
          throw new Error("Query was canceled by the user.");
        }
        throw err;
      }

      await new Promise((resolve) => setTimeout(resolve, 800));
    }

    throw new Error("Query was canceled by the user.");
  };

  const waitForResults = async (
    onStatus: (status: string, elapsedSeconds: number) => void
  ): Promise<any[]> => {
    const bundle = await waitForBundle(onStatus);
    return bundle.rows;
  };

  return { taskId, cancel, waitForResults, waitForBundle };
}

/**
 * Fallback direct inline query runner for quick small queries.
 */
export async function runInlineQuery(
  baseUrl: string,
  token: string,
  queryPayload: QueryPayload
): Promise<any[]> {
  const strategy = resolveResultFormatStrategy(queryPayload);
  const fmt = strategy === "json_detail" ? "json_detail" : "json";
  const url = `${baseUrl.replace(/\/$/, "")}/api/4.0/queries/run/${fmt}`;
  const res = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(queryPayload),
  });

  if (!res.ok) {
    const errText = await res.text();
    throw new Error(`Query failed (${res.status}): ${errText}`);
  }

  const raw = await res.json();
  return normalizeQueryResult(raw, queryPayload.pivots, strategy).rows;
}
