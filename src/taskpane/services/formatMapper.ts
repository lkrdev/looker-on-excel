export const LOOKER_EXCEL_FORMAT_MAP: Record<string, string> = {
  usd: "$#,##0.00",
  usd_0: "$#,##0",
  eur: "[$€-2] #,##0.00",
  eur_0: "[$€-2] #,##0",
  gbp: "[$£-809] #,##0.00",
  decimal_0: "#,##0",
  decimal_1: "#,##0.0",
  decimal_2: "#,##0.00",
  percent_0: "0%",
  percent_1: "0.0%",
  percent_2: "0.00%",
  id: "00000",
};

export function resolveExcelNumberFormat(field: {
  type?: string;
  value_format?: string | null;
  value_format_name?: string | null;
}): string | undefined {
  // 1. Explicit LookML format (uses standard Excel syntax)
  if (field.value_format) {
    return field.value_format;
  }

  // 2. Named Looker format translation
  if (field.value_format_name && LOOKER_EXCEL_FORMAT_MAP[field.value_format_name]) {
    return LOOKER_EXCEL_FORMAT_MAP[field.value_format_name];
  }

  // 3. Field data type heuristics
  switch (field.type) {
    case "date_date":
      return "yyyy-mm-dd";
    case "date_time":
      return "yyyy-mm-dd hh:mm:ss";
    case "int":
    case "count":
      return "#,##0";
    case "number":
    case "sum":
    case "average":
    case "min":
    case "max":
      return "#,##0.00";
    default:
      return undefined;
  }
}
