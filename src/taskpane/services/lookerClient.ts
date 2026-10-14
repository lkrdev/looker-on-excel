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
    "id,name,label,always_filter,conditionally_filter,fields(dimensions(name,label,label_short,type,value_format,value_format_name,view_label,description,field_group_label,field_group_variant,dimension_group,synonyms),measures(name,label,label_short,type,value_format,value_format_name,view_label,description,field_group_label,field_group_variant,synonyms),parameters(name,label,label_short,view_label,type,default_value,default_filter_value,has_allowed_values,enumerations,synonyms))";
  const url = `${baseUrl.replace(/\/$/, "")}/api/4.0/lookml_models/${encodeURIComponent(model)}/explores/${encodeURIComponent(explore)}?fields=${encodeURIComponent(fields)}`;

  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!res.ok) {
    throw new Error(`Failed to fetch explore details: ${res.statusText}`);
  }

  const data: ExploreDetail = await res.json();
  data.always_filter = normalizeAlwaysFilters(data.always_filter as any);

  // Tag categories for easy UI consumption
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
 * GET /api/4.0/models/{model_name}/views/{view_name}/fields/{field_name}/suggestions
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

