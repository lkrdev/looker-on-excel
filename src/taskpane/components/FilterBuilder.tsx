import React, { useState, useEffect, useCallback } from "react";
import {
  makeStyles,
  tokens,
  Select,
  Input,
  Combobox,
  Option,
  Button,
  Caption1,
  Badge,
} from "@fluentui/react-components";
import { AddRegular, DeleteRegular } from "@fluentui/react-icons";
import { FieldDefinition, getFieldSuggestions } from "../services/lookerClient";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    padding: "0 16px 12px 16px",
  },
  headerRow: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
  },
  filterCard: {
    display: "flex",
    flexDirection: "column",
    gap: "6px",
    padding: "8px 10px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  fieldRow: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    width: "100%",
  },
  fieldSelect: {
    flex: 1,
    minWidth: 0,
    width: "100%",
    "& select": {
      minWidth: 0,
      width: "100%",
      textOverflow: "ellipsis",
    },
  },
  deleteButton: {
    minWidth: "28px",
    flexShrink: 0,
  },
  conditionRow: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    width: "100%",
    flexWrap: "wrap",
  },
  operatorSelect: {
    width: "125px",
    minWidth: "95px",
    flexShrink: 0,
    "& select": {
      width: "100%",
    },
  },
  valueControl: {
    flex: 1,
    minWidth: 0,
    width: "100%",
    "& input": {
      minWidth: 0,
      width: "100%",
    },
  },
  compoundControl: {
    display: "flex",
    alignItems: "center",
    gap: "4px",
    flex: 1,
    minWidth: 0,
    flexWrap: "wrap",
  },
  suggestionsBar: {
    display: "flex",
    flexWrap: "wrap",
    gap: "4px",
    paddingLeft: "2px",
    marginTop: "2px",
  },
  emptyHint: {
    color: tokens.colorNeutralForeground4,
    fontStyle: "italic",
    padding: "2px 0",
  },
  rangeSeparator: {
    fontSize: "11px",
    color: tokens.colorNeutralForeground3,
    whiteSpace: "nowrap",
  },
  staticLabel: {
    color: tokens.colorNeutralForeground4,
    fontStyle: "italic",
    fontSize: "12px",
    paddingLeft: "4px",
  },
});

export interface FilterCondition {
  id: string;
  field: string;
  operator: string;
  value: string;
  startValue?: string;
  endValue?: string;
  unit?: string;
}

export function isDateField(field?: FieldDefinition): boolean {
  if (!field) return false;
  if (field.dimension_group) return true;
  if (!field.type) return false;
  const t = field.type.toLowerCase();
  return t.startsWith("date") || t.startsWith("time") || t === "timestamp";
}

export function isNumberField(field?: FieldDefinition): boolean {
  if (!field || !field.type) return false;
  const t = field.type.toLowerCase();
  return [
    "number",
    "int",
    "integer",
    "count",
    "count_distinct",
    "sum",
    "sum_distinct",
    "average",
    "average_distinct",
    "min",
    "max",
    "median",
    "percentile",
    "percent_of_total",
    "percent_of_previous",
    "running_total",
    "duration",
    "distance",
  ].includes(t);
}

export function isStringField(field?: FieldDefinition): boolean {
  return !isDateField(field) && !isNumberField(field);
}

export function compileFilterToLookerExpression(
  f: FilterCondition,
  fieldDef?: FieldDefinition
): string | null {
  if (isDateField(fieldDef)) {
    switch (f.operator) {
      case "is_in_the_last":
        return `${f.value || "7"} ${f.unit || "days"}`;
      case "is_on_the_day":
        return f.value ? f.value : null;
      case "is_in_range": {
        const start = f.startValue || f.value;
        const end = f.endValue;
        if (start && end) return `${start} to ${end}`;
        if (start) return `after ${start}`;
        if (end) return `before ${end}`;
        return null;
      }
      case "is_before":
      case "before":
        return f.value ? `before ${f.value}` : null;
      case "is_on_or_after":
      case "after":
        return f.value ? `after ${f.value}` : null;
      case "is_in_the_year":
        return f.value ? `${f.value}` : null;
      case "is_in_the_month":
        return f.value ? `${f.value.replace("-", "/")}` : null;
      case "is_this":
        return `this ${f.unit || "month"}`;
      case "is_next":
        return `next ${f.value || "1"} ${f.unit || "months"}`;
      case "is_previous":
        return `previous ${f.unit || "month"}`;
      case "is":
        return f.value || null;
      case "is_null":
        return "null";
      case "is_not_null":
        return "-null";
      case "is_any_time":
        return null;
      case "matches_advanced":
        return f.value || null;
      default:
        return f.value || null;
    }
  }

  if (isNumberField(fieldDef)) {
    switch (f.operator) {
      case "is":
        return f.value !== "" && f.value !== undefined ? `${f.value}` : null;
      case "is_not":
        return f.value !== "" && f.value !== undefined ? `NOT ${f.value}` : null;
      case "greater_than":
      case ">":
        return f.value !== "" && f.value !== undefined ? `>${f.value}` : null;
      case "greater_than_or_equal":
      case ">=":
        return f.value !== "" && f.value !== undefined ? `>=${f.value}` : null;
      case "less_than":
      case "<":
        return f.value !== "" && f.value !== undefined ? `<${f.value}` : null;
      case "less_than_or_equal":
      case "<=":
        return f.value !== "" && f.value !== undefined ? `<=${f.value}` : null;
      case "is_between": {
        const min = f.startValue || f.value;
        const max = f.endValue;
        if (min !== "" && min !== undefined && max !== "" && max !== undefined) {
          return `[${min}, ${max}]`;
        }
        if (min !== "" && min !== undefined) return `>=${min}`;
        if (max !== "" && max !== undefined) return `<=${max}`;
        return null;
      }
      case "is_not_between": {
        const min = f.startValue || f.value;
        const max = f.endValue;
        if (min !== "" && min !== undefined && max !== "" && max !== undefined) {
          return `NOT [${min}, ${max}]`;
        }
        return null;
      }
      case "is_null":
        return "null";
      case "is_not_null":
        return "-null";
      case "matches_advanced":
        return f.value || null;
      default:
        return f.value !== "" && f.value !== undefined ? `${f.value}` : null;
    }
  }

  // String field
  switch (f.operator) {
    case "is":
      return f.value ? f.value : null;
    case "is_not":
      return f.value ? `-${f.value}` : null;
    case "contains":
      return f.value ? `%${f.value}%` : null;
    case "does_not_contain":
      return f.value ? `-%${f.value}%` : null;
    case "starts_with":
      return f.value ? `${f.value}%` : null;
    case "ends_with":
      return f.value ? `%${f.value}` : null;
    case "is_null":
      return "EMPTY";
    case "is_not_null":
      return "-EMPTY";
    case "matches_advanced":
      return f.value || null;
    default:
      return f.value || null;
  }
}

interface FilterBuilderProps {
  fields: FieldDefinition[];
  filters: FilterCondition[];
  onAddFilter: () => void;
  onUpdateFilter: (id: string, updates: Partial<FilterCondition>) => void;
  onRemoveFilter: (id: string) => void;
  baseUrl?: string;
  token?: string;
  modelName?: string;
  exploreName?: string;
}

export const FilterBuilder: React.FC<FilterBuilderProps> = ({
  fields,
  filters,
  onAddFilter,
  onUpdateFilter,
  onRemoveFilter,
  baseUrl,
  token,
  modelName,
  exploreName,
}) => {
  const styles = useStyles();

  const [suggestionsCache, setSuggestionsCache] = useState<Record<string, string[]>>({});
  const [loadingFields, setLoadingFields] = useState<Record<string, boolean>>({});

  // Fetch field suggestions from Looker API - ONLY for string fields
  const fetchSuggestions = useCallback(
    async (fieldName: string) => {
      const fieldDef = fields.find((fld) => fld.name === fieldName);
      if (!isStringField(fieldDef)) return;
      if (!baseUrl || !token || !modelName || !exploreName || !fieldName) return;
      if (suggestionsCache[fieldName] || loadingFields[fieldName]) return;

      setLoadingFields((prev) => ({ ...prev, [fieldName]: true }));
      try {
        const list = await getFieldSuggestions(baseUrl, token, modelName, exploreName, fieldName);
        setSuggestionsCache((prev) => ({ ...prev, [fieldName]: list }));
      } catch (err) {
        console.warn("Error fetching filter suggestions:", err);
      } finally {
        setLoadingFields((prev) => ({ ...prev, [fieldName]: false }));
      }
    },
    [baseUrl, token, modelName, exploreName, fields, suggestionsCache, loadingFields]
  );

  // Automatically fetch suggestions only when a string filter field is configured
  useEffect(() => {
    filters.forEach((f) => {
      const fieldDef = fields.find((fld) => fld.name === f.field);
      if (f.field && isStringField(fieldDef) && !suggestionsCache[f.field] && !loadingFields[f.field]) {
        fetchSuggestions(f.field);
      }
    });
  }, [filters, fields, fetchSuggestions, suggestionsCache, loadingFields]);

  const handleFieldChange = (filterId: string, newField: string) => {
    const fieldDef = fields.find((fld) => fld.name === newField);
    let defaultOp = "is";
    let defaultValue = "";
    let defaultUnit: string | undefined = undefined;

    if (isDateField(fieldDef)) {
      defaultOp = "is_in_the_last";
      defaultValue = "7";
      defaultUnit = "days";
    }

    onUpdateFilter(filterId, {
      field: newField,
      operator: defaultOp,
      value: defaultValue,
      unit: defaultUnit,
      startValue: undefined,
      endValue: undefined,
    });

    if (newField && isStringField(fieldDef)) {
      fetchSuggestions(newField);
    }
  };

  const handleOperatorChange = (filter: FilterCondition, newOp: string, fieldDef?: FieldDefinition) => {
    const updates: Partial<FilterCondition> = { operator: newOp };

    if (isDateField(fieldDef)) {
      if (newOp === "is_in_the_last") {
        updates.value = filter.value || "7";
        updates.unit = filter.unit || "days";
      } else if (newOp === "is_in_range") {
        updates.startValue = filter.startValue || filter.value || "";
        updates.endValue = filter.endValue || "";
      } else if (newOp === "is_this" || newOp === "is_previous") {
        updates.unit = filter.unit || "month";
      } else if (newOp === "is_next") {
        updates.value = filter.value || "1";
        updates.unit = filter.unit || "months";
      }
    }

    onUpdateFilter(filter.id, updates);
  };

  return (
    <div className={styles.container}>
      <div className={styles.headerRow}>
        <Caption1 style={{ fontWeight: 600 }}>Filters ({filters.length})</Caption1>
        <Button
          size="small"
          appearance="subtle"
          icon={<AddRegular />}
          onClick={onAddFilter}
          disabled={fields.length === 0}
        >
          Add Filter
        </Button>
      </div>

      {filters.length === 0 && (
        <Caption1 className={styles.emptyHint}>
          No filters applied. Click &quot;Add Filter&quot; to narrow results.
        </Caption1>
      )}

      {filters.map((f) => {
        const selectedFieldObj = fields.find((fld) => fld.name === f.field);
        const isDate = isDateField(selectedFieldObj);
        const isNumber = isNumberField(selectedFieldObj);
        const isString = isStringField(selectedFieldObj);

        const fieldSuggestions = isString && f.field ? suggestionsCache[f.field] || [] : [];
        const isLoadingSuggestions = isString && f.field ? !!loadingFields[f.field] : false;

        return (
          <div key={f.id} className={styles.filterCard}>
            {/* Row 1: Field Selection & Delete Button */}
            <div className={styles.fieldRow}>
              <Select
                size="small"
                className={styles.fieldSelect}
                value={f.field}
                title={selectedFieldObj?.label || f.field || "Select Field"}
                onChange={(_, data) => handleFieldChange(f.id, data.value)}
              >
                <option value="">Select Field</option>
                {fields.map((fld) => (
                  <option key={fld.name} value={fld.name}>
                    {fld.label || fld.name}
                  </option>
                ))}
              </Select>
              <Button
                size="small"
                appearance="subtle"
                className={styles.deleteButton}
                icon={<DeleteRegular />}
                onClick={() => onRemoveFilter(f.id)}
                title="Remove Filter"
              />
            </div>

            {/* Row 2: Operator & Specialized Value Input */}
            <div className={styles.conditionRow}>
              {/* Operator Select based on data type */}
              <Select
                size="small"
                className={styles.operatorSelect}
                value={f.operator}
                onChange={(_, data) => handleOperatorChange(f, data.value, selectedFieldObj)}
              >
                {isDate ? (
                  <>
                    <option value="is_in_the_last">is in the last</option>
                    <option value="is_on_the_day">is on the day</option>
                    <option value="is_in_range">is in range</option>
                    <option value="is_before">is before</option>
                    <option value="is_on_or_after">is on or after</option>
                    <option value="is_in_the_year">is in the year</option>
                    <option value="is_in_the_month">is in the month</option>
                    <option value="is_this">is this</option>
                    <option value="is_next">is next</option>
                    <option value="is_previous">is previous</option>
                    <option value="is">is</option>
                    <option value="is_null">is null</option>
                    <option value="is_not_null">is not null</option>
                    <option value="is_any_time">is any time</option>
                    <option value="matches_advanced">matches (advanced)</option>
                  </>
                ) : isNumber ? (
                  <>
                    <option value="is">is</option>
                    <option value="is_not">is not</option>
                    <option value="greater_than">is &gt;</option>
                    <option value="greater_than_or_equal">is &gt;=</option>
                    <option value="less_than">is &lt;</option>
                    <option value="less_than_or_equal">is &lt;=</option>
                    <option value="is_between">is between</option>
                    <option value="is_not_between">is not between</option>
                    <option value="is_null">is null</option>
                    <option value="is_not_null">is not null</option>
                    <option value="matches_advanced">matches (advanced)</option>
                  </>
                ) : (
                  <>
                    <option value="is">is</option>
                    <option value="is_not">is not</option>
                    <option value="contains">contains</option>
                    <option value="does_not_contain">does not contain</option>
                    <option value="starts_with">starts with</option>
                    <option value="ends_with">ends with</option>
                    <option value="is_null">is null</option>
                    <option value="is_not_null">is not null</option>
                    <option value="matches_advanced">matches (advanced)</option>
                  </>
                )}
              </Select>

              {/* DATE FIELDS CONTROLS */}
              {isDate && (
                <>
                  {f.operator === "is_in_the_last" && (
                    <div className={styles.compoundControl}>
                      <Input
                        size="small"
                        type="number"
                        min={1}
                        value={f.value ?? "7"}
                        onChange={(_, data) => onUpdateFilter(f.id, { value: data.value })}
                        style={{ width: "65px", minWidth: "50px" }}
                      />
                      <Select
                        size="small"
                        value={f.unit || "days"}
                        onChange={(_, data) => onUpdateFilter(f.id, { unit: data.value })}
                        style={{ flex: 1, minWidth: "75px" }}
                      >
                        <option value="days">days</option>
                        <option value="weeks">weeks</option>
                        <option value="months">months</option>
                        <option value="quarters">quarters</option>
                        <option value="years">years</option>
                      </Select>
                    </div>
                  )}

                  {f.operator === "is_in_range" && (
                    <div className={styles.compoundControl}>
                      <Input
                        size="small"
                        type="date"
                        value={f.startValue ?? f.value ?? ""}
                        onChange={(_, data) => onUpdateFilter(f.id, { startValue: data.value })}
                        style={{ flex: 1, minWidth: "110px" }}
                      />
                      <Caption1 className={styles.rangeSeparator}>until (before)</Caption1>
                      <Input
                        size="small"
                        type="date"
                        value={f.endValue ?? ""}
                        onChange={(_, data) => onUpdateFilter(f.id, { endValue: data.value })}
                        style={{ flex: 1, minWidth: "110px" }}
                      />
                    </div>
                  )}

                  {(f.operator === "is_on_the_day" ||
                    f.operator === "is_before" ||
                    f.operator === "is_on_or_after") && (
                    <Input
                      size="small"
                      type="date"
                      className={styles.valueControl}
                      value={f.value ?? ""}
                      onChange={(_, data) => onUpdateFilter(f.id, { value: data.value })}
                    />
                  )}

                  {f.operator === "is_in_the_year" && (
                    <Input
                      size="small"
                      type="number"
                      placeholder="YYYY (e.g. 2026)"
                      className={styles.valueControl}
                      value={f.value ?? ""}
                      onChange={(_, data) => onUpdateFilter(f.id, { value: data.value })}
                    />
                  )}

                  {f.operator === "is_in_the_month" && (
                    <Input
                      size="small"
                      type="month"
                      className={styles.valueControl}
                      value={f.value ?? ""}
                      onChange={(_, data) => onUpdateFilter(f.id, { value: data.value })}
                    />
                  )}

                  {(f.operator === "is_this" || f.operator === "is_previous") && (
                    <Select
                      size="small"
                      className={styles.valueControl}
                      value={f.unit || "month"}
                      onChange={(_, data) => onUpdateFilter(f.id, { unit: data.value })}
                    >
                      <option value="day">day</option>
                      <option value="week">week</option>
                      <option value="month">month</option>
                      <option value="quarter">quarter</option>
                      <option value="year">year</option>
                    </Select>
                  )}

                  {f.operator === "is_next" && (
                    <div className={styles.compoundControl}>
                      <Input
                        size="small"
                        type="number"
                        min={1}
                        value={f.value ?? "1"}
                        onChange={(_, data) => onUpdateFilter(f.id, { value: data.value })}
                        style={{ width: "65px", minWidth: "50px" }}
                      />
                      <Select
                        size="small"
                        value={f.unit || "months"}
                        onChange={(_, data) => onUpdateFilter(f.id, { unit: data.value })}
                        style={{ flex: 1, minWidth: "75px" }}
                      >
                        <option value="days">days</option>
                        <option value="weeks">weeks</option>
                        <option value="months">months</option>
                        <option value="quarters">quarters</option>
                        <option value="years">years</option>
                      </Select>
                    </div>
                  )}

                  {(f.operator === "is_null" ||
                    f.operator === "is_not_null" ||
                    f.operator === "is_any_time") && (
                    <Caption1 className={styles.staticLabel}>
                      {f.operator === "is_null"
                        ? "is null (empty)"
                        : f.operator === "is_not_null"
                        ? "is not null"
                        : "any time"}
                    </Caption1>
                  )}

                  {(f.operator === "is" || f.operator === "matches_advanced") && (
                    <Input
                      size="small"
                      className={styles.valueControl}
                      placeholder="Looker date expression (e.g. last 30 days)"
                      value={f.value ?? ""}
                      onChange={(_, data) => onUpdateFilter(f.id, { value: data.value })}
                    />
                  )}
                </>
              )}

              {/* NUMBER FIELDS CONTROLS */}
              {isNumber && (
                <>
                  {(f.operator === "is_between" || f.operator === "is_not_between") && (
                    <div className={styles.compoundControl}>
                      <Input
                        size="small"
                        type="number"
                        placeholder="Min"
                        value={f.startValue ?? f.value ?? ""}
                        onChange={(_, data) => onUpdateFilter(f.id, { startValue: data.value })}
                        style={{ flex: 1, minWidth: "60px" }}
                      />
                      <Caption1 className={styles.rangeSeparator}>to</Caption1>
                      <Input
                        size="small"
                        type="number"
                        placeholder="Max"
                        value={f.endValue ?? ""}
                        onChange={(_, data) => onUpdateFilter(f.id, { endValue: data.value })}
                        style={{ flex: 1, minWidth: "60px" }}
                      />
                    </div>
                  )}

                  {(f.operator === "is_null" || f.operator === "is_not_null") && (
                    <Caption1 className={styles.staticLabel}>
                      {f.operator === "is_null" ? "is null" : "is not null"}
                    </Caption1>
                  )}

                  {f.operator === "matches_advanced" && (
                    <Input
                      size="small"
                      className={styles.valueControl}
                      placeholder="e.g. > 100 AND <= 500"
                      value={f.value ?? ""}
                      onChange={(_, data) => onUpdateFilter(f.id, { value: data.value })}
                    />
                  )}

                  {f.operator !== "is_between" &&
                    f.operator !== "is_not_between" &&
                    f.operator !== "is_null" &&
                    f.operator !== "is_not_null" &&
                    f.operator !== "matches_advanced" && (
                      <Input
                        size="small"
                        type="number"
                        className={styles.valueControl}
                        placeholder="any value"
                        value={f.value ?? ""}
                        onChange={(_, data) => onUpdateFilter(f.id, { value: data.value })}
                      />
                    )}
                </>
              )}

              {/* STRING FIELDS CONTROLS */}
              {isString && (
                <>
                  {(f.operator === "is_null" || f.operator === "is_not_null") && (
                    <Caption1 className={styles.staticLabel}>
                      {f.operator === "is_null" ? "is empty / null" : "is not empty"}
                    </Caption1>
                  )}

                  {f.operator !== "is_null" && f.operator !== "is_not_null" && (
                    <>
                      {fieldSuggestions.length > 0 && (f.operator === "is" || f.operator === "is_not") ? (
                        <Combobox
                          size="small"
                          freeform
                          className={styles.valueControl}
                          placeholder={isLoadingSuggestions ? "Loading..." : "Value or pick..."}
                          value={f.value}
                          onChange={(e) => onUpdateFilter(f.id, { value: e.target.value })}
                          onOptionSelect={(_, data) =>
                            onUpdateFilter(f.id, { value: data.optionValue || data.optionText || "" })
                          }
                        >
                          {fieldSuggestions.map((s) => (
                            <Option key={s} value={s}>
                              {s}
                            </Option>
                          ))}
                        </Combobox>
                      ) : (
                        <Input
                          size="small"
                          className={styles.valueControl}
                          placeholder={
                            isLoadingSuggestions
                              ? "Checking..."
                              : f.operator === "contains"
                              ? "Contains text..."
                              : f.operator === "starts_with"
                              ? "Starts with..."
                              : f.operator === "ends_with"
                              ? "Ends with..."
                              : "Filter value..."
                          }
                          value={f.value}
                          onChange={(_, data) => onUpdateFilter(f.id, { value: data.value })}
                          onFocus={() => {
                            if (f.field && !suggestionsCache[f.field]) {
                              fetchSuggestions(f.field);
                            }
                          }}
                        />
                      )}
                    </>
                  )}
                </>
              )}
            </div>

            {/* Quick Suggestion Pills for Compact Option Lists (<= 8 items) - STRING FIELDS ONLY */}
            {isString && fieldSuggestions.length > 0 && fieldSuggestions.length <= 8 && (
              <div className={styles.suggestionsBar}>
                <Caption1
                  style={{
                    color: tokens.colorNeutralForeground4,
                    fontSize: "10px",
                    alignSelf: "center",
                  }}
                >
                  Suggestions:
                </Caption1>
                {fieldSuggestions.map((s) => (
                  <Badge
                    key={s}
                    size="small"
                    appearance={f.value === s ? "filled" : "outline"}
                    color={f.value === s ? "brand" : "informative"}
                    style={{ cursor: "pointer", fontSize: "10px" }}
                    onClick={() => onUpdateFilter(f.id, { value: s })}
                  >
                    {s}
                  </Badge>
                ))}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
};
