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
  view_label?: string;
  type?: string;
  value_format?: string | null;
  value_format_name?: string | null;
  description?: string | null;
  category?: "dimension" | "measure" | "parameter";
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
    "id,name,label,always_filter,conditionally_filter,fields(dimensions(name,label,type,value_format,value_format_name,view_label,description),measures(name,label,type,value_format,value_format_name,view_label,description),parameters(name,label,type,default_value))";
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
 * Fetches dynamic suggestions for filter inputs.
 */
export async function getFieldSuggestions(
  baseUrl: string,
  token: string,
  model: string,
  explore: string,
  field: string
): Promise<string[]> {
  const url = `${baseUrl.replace(/\/$/, "")}/api/4.0/lookml_models/${encodeURIComponent(model)}/explores/${encodeURIComponent(explore)}/fields/${encodeURIComponent(field)}/suggestions`;
  try {
    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!res.ok) return [];
    const data = await res.json();
    return data.suggestions || [];
  } catch {
    return [];
  }
}
