import { QueryPayload } from "./queryTaskRunner";

export interface ExploreSummary {
  name: string;
  label: string;
}

export interface ModelSummary {
  name: string;
  label: string;
  explores: ExploreSummary[];
}

export interface FieldDefinition {
  name: string;
  label: string;
  label_short?: string;
  view_label?: string;
  type?: string;
  value_format?: string | null;
  value_format_name?: string | null;
  description?: string | null;
  category?: "dimension" | "measure" | "parameter";
  field_group_label?: string | null;
  field_group_variant?: string | null;
  dimension_group?: string | null;
  group_label?: string | null;
  synonyms?: string[];
  tags?: string[];
  drill_fields?: string[];
}

export interface ParameterDefinition {
  name: string;
  label: string;
  label_short?: string;
  view_label?: string;
  type?: string;
  default_value?: string | null;
  default_filter_value?: string | null;
  has_allowed_values?: boolean;
  enumerations?: Array<{ label: string; value: string }> | null;
}

export interface AlwaysFilterItem {
  field: string;
  values: string[];
  name?: string;
  value?: string;
}

export interface ExploreDetail {
  id: string;
  name: string;
  label: string;
  always_filter?: AlwaysFilterItem[];
  conditionally_filter?: Array<{ fields: string[]; values: string[] }>;
  fields: {
    dimensions: FieldDefinition[];
    measures: FieldDefinition[];
    parameters?: ParameterDefinition[];
  };
}

export interface LookerWriteQuery {
  model: string;
  view: string;
  fields?: string[] | null;
  pivots?: string[] | null;
  fill_fields?: string[] | null;
  filters?: Record<string, string> | null;
  sorts?: string[] | null;
  limit?: string | null;
  total?: boolean | null;
  row_total?: string | null;
  subtotals?: string[] | null;
}

export interface LookSummary {
  id: string;
  title: string;
  description?: string | null;
  folder?: { id: string; name: string } | null;
  model?: { id: string; label?: string } | null;
  updated_at?: string | null;
}

export interface LookDetail {
  id: string;
  title: string;
  description?: string | null;
  query?: LookerWriteQuery | null;
}

export interface DashboardSummary {
  id: string;
  title: string;
  description?: string | null;
  folder?: { id: string; name: string } | null;
}

export interface DashboardFilterDef {
  id?: string;
  name: string;
  title: string;
  type?: string;
  default_value?: string | null;
  dimension?: string | null;
}

export interface DashboardTileListener {
  dashboard_filter_name: string;
  field: string;
}

export interface DashboardElementDef {
  id: string;
  title?: string | null;
  type?: string | null;
  query?: LookerWriteQuery | null;
  result_maker?: {
    query?: LookerWriteQuery | null;
    filterables?: Array<{ listen?: DashboardTileListener[] | null }> | null;
  } | null;
}

export interface DashboardDetail {
  id: string;
  title: string;
  description?: string | null;
  dashboard_filters?: DashboardFilterDef[];
  dashboard_elements?: DashboardElementDef[];
}

/**
 * Converts a Looker Look or Dashboard Tile `query` object into a normalized QueryPayload.
 */
export function convertWriteQueryToPayload(wq: LookerWriteQuery): QueryPayload {
  return {
    model: wq.model,
    view: wq.view,
    fields: Array.isArray(wq.fields) ? wq.fields : [],
    filters:
      wq.filters && Object.keys(wq.filters).length > 0 ? { ...wq.filters } : null,
    sorts: Array.isArray(wq.sorts) ? wq.sorts : [],
    limit: wq.limit || "500",
    pivots: Array.isArray(wq.pivots) && wq.pivots.length > 0 ? wq.pivots : null,
    fill_fields:
      Array.isArray(wq.fill_fields) && wq.fill_fields.length > 0
        ? wq.fill_fields
        : undefined,
    total: wq.total ? true : undefined,
    row_total:
      wq.row_total === "right" || wq.row_total === "left" ? wq.row_total : undefined,
    subtotals:
      Array.isArray(wq.subtotals) && wq.subtotals.length > 0 ? wq.subtotals : null,
  };
}

/**
 * Stably orders query field keys so that Dimensions always appear before Measures
 * (matching Looker's Explore data table layout: non-pivoted dimensions -> pivoted dimensions -> measures),
 * while preserving the relative order within each category.
 *
 * Looker's Saved Look and Dashboard APIs return `query.fields` in the raw chronological order
 * that the author clicked fields in the UI, which can interleave measures before dimensions.
 */
export function orderFieldsDimensionsFirst(
  fieldKeys: string[],
  availableFields: FieldDefinition[],
  pivots?: string[] | null
): string[] {
  if (!Array.isArray(fieldKeys) || fieldKeys.length <= 1) {
    return Array.isArray(fieldKeys) ? [...fieldKeys] : [];
  }

  const fieldMap = new Map<string, FieldDefinition>();
  for (const f of availableFields) {
    fieldMap.set(f.name, f);
  }
  const pivotSet = new Set<string>(Array.isArray(pivots) ? pivots : []);

  const nonPivotedDimensions: string[] = [];
  const pivotedDimensions: string[] = [];
  const unknownFields: string[] = [];
  const measures: string[] = [];

  for (const key of fieldKeys) {
    const def = fieldMap.get(key);
    if (def?.category === "dimension") {
      if (pivotSet.has(key)) {
        pivotedDimensions.push(key);
      } else {
        nonPivotedDimensions.push(key);
      }
    } else if (def?.category === "measure") {
      measures.push(key);
    } else if (pivotSet.has(key)) {
      pivotedDimensions.push(key);
    } else {
      unknownFields.push(key);
    }
  }

  return [
    ...nonPivotedDimensions,
    ...pivotedDimensions,
    ...unknownFields,
    ...measures,
  ];
}

/**
 * Extracts the query and any dashboard filter `listen` bindings from a DashboardElementDef.
 */
export function extractTileQueryAndListeners(el: DashboardElementDef): {
  query: LookerWriteQuery | null;
  listeners: DashboardTileListener[];
} {
  const query = el.query || el.result_maker?.query || null;
  const listeners: DashboardTileListener[] = [];
  const filterables = el.result_maker?.filterables || [];
  for (const f of filterables) {
    if (Array.isArray(f.listen)) {
      for (const l of f.listen) {
        if (l.dashboard_filter_name && l.field) {
          listeners.push({
            dashboard_filter_name: l.dashboard_filter_name,
            field: l.field,
          });
        }
      }
    }
  }
  return { query, listeners };
}

/**
 * Applies active dashboard-level filter values onto a tile's QueryPayload using its `listen` mappings.
 */
export function applyDashboardFiltersToTilePayload(
  basePayload: QueryPayload,
  listeners: DashboardTileListener[],
  dashboardFilterValues: Record<string, string>
): QueryPayload {
  const mergedFilters: Record<string, string> = {
    ...(basePayload.filters || {}),
  };

  for (const listener of listeners) {
    const val = dashboardFilterValues[listener.dashboard_filter_name];
    if (val !== undefined && val !== null && String(val).trim() !== "") {
      mergedFilters[listener.field] = String(val).trim();
    }
  }

  return {
    ...basePayload,
    filters: Object.keys(mergedFilters).length > 0 ? mergedFilters : null,
  };
}

/**
 * Normalizes Looker API 4.0 `always_filter` items (which return `{ name, value }`)
 * and legacy/test `{ field, values }` shapes into a consistent AlwaysFilterItem[].
 */
export function normalizeAlwaysFilters(
  raw?: Array<{ name?: string; value?: string; field?: string; values?: string[] }> | null
): AlwaysFilterItem[] {
  if (!Array.isArray(raw)) return [];
  return raw
    .map((item) => {
      const field = (item.field || item.name || "").trim();
      const values =
        Array.isArray(item.values) && item.values.length > 0
          ? item.values
          : item.value !== undefined && item.value !== null && String(item.value).trim() !== ""
          ? [String(item.value).trim()]
          : [];
      return {
        field,
        values,
        name: item.name || field,
        value: item.value ?? values.join(", "),
      };
    })
    .filter((item) => item.field !== "");
}

export interface UserSummary {
  id: number;
  first_name?: string;
  last_name?: string;
  email?: string;
}

export async function getCurrentUser(baseUrl: string, token: string): Promise<UserSummary> {
  const url = `${baseUrl.replace(/\/$/, "")}/api/4.0/user?fields=id,first_name,last_name,email`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) throw new Error(`Failed to fetch user: ${res.statusText}`);
  return await res.json();
}

/**
 * Fetches LookML models with explores, using fields filter for 70% payload pruning.
 */
export async function getModels(baseUrl: string, token: string): Promise<ModelSummary[]> {
  const fields = "name,label,explores(name,label)";
  const url = `${baseUrl.replace(/\/$/, "")}/api/4.0/lookml_models?fields=${encodeURIComponent(fields)}`;

  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!res.ok) {
    throw new Error(`Failed to fetch models: ${res.statusText}`);
  }

  return await res.json();
}

/**
 * Fetches explore schema with dimensions, measures, formats, and parameters, pruned by 90% via fields parameter.
 */
export async function getExplore(
  baseUrl: string,
  token: string,
  model: string,
  explore: string
): Promise<ExploreDetail> {
  const fields =
    "id,name,label,always_filter,conditionally_filter,fields(dimensions(name,label,label_short,type,value_format,value_format_name,view_label,description,field_group_label,field_group_variant,dimension_group,synonyms,tags,drill_fields),measures(name,label,label_short,type,value_format,value_format_name,view_label,description,field_group_label,field_group_variant,synonyms,tags,drill_fields),parameters(name,label,label_short,view_label,type,default_value,default_filter_value,has_allowed_values,enumerations,synonyms))";
  const url = `${baseUrl.replace(/\/$/, "")}/api/4.0/lookml_models/${encodeURIComponent(model)}/explores/${encodeURIComponent(explore)}?fields=${encodeURIComponent(fields)}`;

  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!res.ok) {
    throw new Error(`Failed to fetch explore details: ${res.statusText}`);
  }

  const data: ExploreDetail = await res.json();
  data.always_filter = normalizeAlwaysFilters(data.always_filter as any);

  if (data.fields?.dimensions) {
    data.fields.dimensions.forEach((d) => (d.category = "dimension"));
  }
  if (data.fields?.measures) {
    data.fields.measures.forEach((m) => (m.category = "measure"));
  }
  if (data.fields?.parameters) {
    data.fields.parameters.forEach((p) => {
      if (!p.default_value && p.default_filter_value) {
        p.default_value = p.default_filter_value;
      }
    });
  }

  return data;
}

/**
 * Retrieves field suggestions for a specific explore and field.
 */
export async function getFieldSuggestions(
  baseUrl: string,
  token: string,
  modelName: string,
  exploreName: string,
  fieldName: string,
  term?: string
): Promise<string[]> {
  try {
    const root = baseUrl.replace(/\/$/, "");
    const query = term ? `?term=${encodeURIComponent(term)}` : "";
    const url = `${root}/api/4.0/models/${encodeURIComponent(modelName)}/views/${encodeURIComponent(exploreName)}/fields/${encodeURIComponent(fieldName)}/suggestions${query}`;

    const res = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
    });

    if (!res.ok) {
      return [];
    }

    const data = await res.json();
    return Array.isArray(data?.suggestions) ? data.suggestions : [];
  } catch (e) {
    console.warn("Failed to fetch field suggestions:", e);
    return [];
  }
}

/**
 * Fetches all accessible Saved Looks from the Looker instance.
 */
export async function getLooks(baseUrl: string, token: string): Promise<LookSummary[]> {
  const fields = "id,title,description,folder(id,name),model(id,label),updated_at";
  const url = `${baseUrl.replace(/\/$/, "")}/api/4.0/looks?fields=${encodeURIComponent(fields)}`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    throw new Error(`Failed to fetch Looks: ${res.statusText}`);
  }
  const list = await res.json();
  return Array.isArray(list)
    ? list.map((l: any) => ({ ...l, id: String(l.id) }))
    : [];
}

/**
 * Fetches a specific Saved Look including its underlying semantic Query definition.
 */
export async function getLook(
  baseUrl: string,
  token: string,
  lookId: string
): Promise<LookDetail> {
  const fields =
    "id,title,description,query(model,view,fields,pivots,fill_fields,filters,sorts,limit,total,row_total,subtotals)";
  const url = `${baseUrl.replace(/\/$/, "")}/api/4.0/looks/${encodeURIComponent(lookId)}?fields=${encodeURIComponent(fields)}`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    throw new Error(`Failed to fetch Look ${lookId}: ${res.statusText}`);
  }
  const data = await res.json();
  return { ...data, id: String(data.id) };
}

/**
 * Fetches all accessible Dashboards from the Looker instance.
 */
export async function getDashboards(baseUrl: string, token: string): Promise<DashboardSummary[]> {
  const fields = "id,title,description,folder(id,name)";
  const url = `${baseUrl.replace(/\/$/, "")}/api/4.0/dashboards?fields=${encodeURIComponent(fields)}`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    throw new Error(`Failed to fetch Dashboards: ${res.statusText}`);
  }
  const list = await res.json();
  return Array.isArray(list)
    ? list.map((d: any) => ({ ...d, id: String(d.id) }))
    : [];
}

/**
 * Fetches a specific Dashboard including its dashboard_filters and dashboard_elements (tiles + queries + listeners).
 */
export async function getDashboard(
  baseUrl: string,
  token: string,
  dashboardId: string
): Promise<DashboardDetail> {
  const fields =
    "id,title,description,dashboard_filters(id,name,title,type,default_value,dimension),dashboard_elements(id,title,type,query(model,view,fields,pivots,fill_fields,filters,sorts,limit,total,row_total,subtotals),result_maker(query(model,view,fields,pivots,fill_fields,filters,sorts,limit,total,row_total,subtotals),filterables(listen)))";
  const url = `${baseUrl.replace(/\/$/, "")}/api/4.0/dashboards/${encodeURIComponent(dashboardId)}?fields=${encodeURIComponent(fields)}`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    throw new Error(`Failed to fetch Dashboard ${dashboardId}: ${res.statusText}`);
  }
  const data = await res.json();
  return { ...data, id: String(data.id) };
}
