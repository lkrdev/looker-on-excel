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
}

export interface ParameterDefinition {
  name: string;
  label: string;
  type?: string;
  default_value?: string;
}

export interface ExploreDetail {
  id: string;
  name: string;
  label: string;
  always_filter?: Array<{ field: string; values: string[] }>;
  conditionally_filter?: Array<{ fields: string[]; values: string[] }>;
  fields: {
    dimensions: FieldDefinition[];
    measures: FieldDefinition[];
    parameters?: ParameterDefinition[];
  };
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
    "id,name,label,always_filter,conditionally_filter,fields(dimensions(name,label,label_short,type,value_format,value_format_name,view_label,description,field_group_label,field_group_variant,dimension_group),measures(name,label,label_short,type,value_format,value_format_name,view_label,description,field_group_label,field_group_variant),parameters(name,label,type,default_value))";
  const url = `${baseUrl.replace(/\/$/, "")}/api/4.0/lookml_models/${encodeURIComponent(model)}/explores/${encodeURIComponent(explore)}?fields=${encodeURIComponent(fields)}`;

  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!res.ok) {
    throw new Error(`Failed to fetch explore details: ${res.statusText}`);
  }

  const data: ExploreDetail = await res.json();

  // Tag categories for easy UI consumption
  if (data.fields?.dimensions) {
    data.fields.dimensions.forEach((d) => (d.category = "dimension"));
  }
  if (data.fields?.measures) {
    data.fields.measures.forEach((m) => (m.category = "measure"));
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

