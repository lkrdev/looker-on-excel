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

export function splitMultiValueString(val?: string): string[] {
  if (!val) return [];
  return val
    .split(",")
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
}

export function toggleMultiValueToken(currentVal: string | undefined, token: string): string {
  const tokens = splitMultiValueString(currentVal);
  const exists = tokens.includes(token);
  const next = exists ? tokens.filter((t) => t !== token) : [...tokens, token];
  return next.join(", ");
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

  // String field (supports multi-value comma-separated selections)
  const items = splitMultiValueString(f.value);
  switch (f.operator) {
    case "is":
      return items.length > 0 ? items.join(", ") : null;
    case "is_not":
      return items.length > 0
        ? items.map((t) => (t.startsWith("-") ? t : `-${t}`)).join(", ")
        : null;
    case "contains":
      return items.length > 0 ? items.map((t) => `%${t}%`).join(", ") : null;
    case "does_not_contain":
      return items.length > 0 ? items.map((t) => `-%${t}%`).join(", ") : null;
    case "starts_with":
      return items.length > 0 ? items.map((t) => `${t}%`).join(", ") : null;
    case "ends_with":
      return items.length > 0 ? items.map((t) => `%${t}`).join(", ") : null;
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

/**
 * Parses a Looker filter expression (e.g. from `always_filter` or saved metadata)
 * into a structured FilterCondition for interactive editing.
 */
export function parseLookerExpressionToFilterCondition(
  field: string,
  expr: string,
  fieldDef?: FieldDefinition,
  id: string = `filter_${field}`
): FilterCondition {
  const trimmed = (expr || "").trim();
  const upper = trimmed.toUpperCase();

  if (isDateField(fieldDef)) {
    if (!trimmed) {
      return { id, field, operator: "is_in_the_last", value: "7", unit: "days" };
    }
    if (upper === "NULL") return { id, field, operator: "is_null", value: "" };
    if (upper === "-NULL" || upper === "NOT NULL") {
      return { id, field, operator: "is_not_null", value: "" };
    }
    const lastMatch = trimmed.match(/^(\d+)\s+(days|weeks|months|quarters|years)$/i);
    if (lastMatch) {
      return {
        id,
        field,
        operator: "is_in_the_last",
        value: lastMatch[1],
        unit: lastMatch[2].toLowerCase(),
      };
    }
    const thisMatch = trimmed.match(/^this\s+(day|week|month|quarter|year)$/i);
    if (thisMatch) {
      return { id, field, operator: "is_this", value: "", unit: thisMatch[1].toLowerCase() };
    }
    const prevMatch = trimmed.match(/^previous\s+(day|week|month|quarter|year)$/i);
    if (prevMatch) {
      return { id, field, operator: "is_previous", value: "", unit: prevMatch[1].toLowerCase() };
    }
    const nextMatch = trimmed.match(/^next\s+(\d+)\s+(days|weeks|months|quarters|years)$/i);
    if (nextMatch) {
      return {
        id,
        field,
        operator: "is_next",
        value: nextMatch[1],
        unit: nextMatch[2].toLowerCase(),
      };
    }
    const rangeMatch = trimmed.match(/^(\d{4}-\d{2}-\d{2})\s+to\s+(\d{4}-\d{2}-\d{2})$/i);
    if (rangeMatch) {
      return {
        id,
        field,
        operator: "is_in_range",
        value: "",
        startValue: rangeMatch[1],
        endValue: rangeMatch[2],
      };
    }
    const beforeMatch = trimmed.match(/^before\s+(\d{4}-\d{2}-\d{2})$/i);
    if (beforeMatch) {
      return { id, field, operator: "is_before", value: beforeMatch[1] };
    }
    const afterMatch = trimmed.match(/^after\s+(\d{4}-\d{2}-\d{2})$/i);
    if (afterMatch) {
      return { id, field, operator: "is_on_or_after", value: afterMatch[1] };
    }
    if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
      return { id, field, operator: "is_on_the_day", value: trimmed };
    }
    return { id, field, operator: "is", value: trimmed };
  }

  if (isNumberField(fieldDef)) {
    if (upper === "NULL") return { id, field, operator: "is_null", value: "" };
    if (upper === "-NULL" || upper === "NOT NULL") {
      return { id, field, operator: "is_not_null", value: "" };
    }
    const notBetween = trimmed.match(/^NOT\s*\[\s*([^,\]]+)\s*,\s*([^\]]+)\s*\]$/i);
    if (notBetween) {
      return {
        id,
        field,
        operator: "is_not_between",
        value: "",
        startValue: notBetween[1].trim(),
        endValue: notBetween[2].trim(),
      };
    }
    const between = trimmed.match(/^\[\s*([^,\]]+)\s*,\s*([^\]]+)\s*\]$/);
    if (between) {
      return {
        id,
        field,
        operator: "is_between",
        value: "",
        startValue: between[1].trim(),
        endValue: between[2].trim(),
      };
    }
    if (trimmed.startsWith(">=")) {
      return { id, field, operator: "greater_than_or_equal", value: trimmed.slice(2).trim() };
    }
    if (trimmed.startsWith("<=")) {
      return { id, field, operator: "less_than_or_equal", value: trimmed.slice(2).trim() };
    }
    if (trimmed.startsWith(">")) {
      return { id, field, operator: "greater_than", value: trimmed.slice(1).trim() };
    }
    if (trimmed.startsWith("<")) {
      return { id, field, operator: "less_than", value: trimmed.slice(1).trim() };
    }
    if (/^NOT\s+/i.test(trimmed)) {
      return { id, field, operator: "is_not", value: trimmed.replace(/^NOT\s+/i, "").trim() };
    }
    return { id, field, operator: "is", value: trimmed };
  }

  // String field
  if (upper === "-NULL" || upper === "-EMPTY" || upper === "NOT NULL") {
    return { id, field, operator: "is_not_null", value: "" };
  }
  if (upper === "NULL" || upper === "EMPTY") {
    return { id, field, operator: "is_null", value: "" };
  }

  const parts = splitMultiValueString(trimmed);
  if (parts.length > 0) {
    if (parts.every((p) => p.startsWith("-%") && p.endsWith("%") && p.length > 3)) {
      return {
        id,
        field,
        operator: "does_not_contain",
        value: parts.map((p) => p.slice(2, -1)).join(", "),
      };
    }
    if (parts.every((p) => p.startsWith("%") && p.endsWith("%") && p.length > 2)) {
      return {
        id,
        field,
        operator: "contains",
        value: parts.map((p) => p.slice(1, -1)).join(", "),
      };
    }
    if (parts.every((p) => p.endsWith("%") && !p.startsWith("%") && !p.startsWith("-") && p.length > 1)) {
      return {
        id,
        field,
        operator: "starts_with",
        value: parts.map((p) => p.slice(0, -1)).join(", "),
      };
    }
    if (parts.every((p) => p.startsWith("%") && !p.endsWith("%") && p.length > 1)) {
      return {
        id,
        field,
        operator: "ends_with",
        value: parts.map((p) => p.slice(1)).join(", "),
      };
    }
    if (parts.every((p) => p.startsWith("-") && p.length > 1)) {
      return {
        id,
        field,
        operator: "is_not",
        value: parts.map((p) => p.slice(1)).join(", "),
      };
    }
  }

  return { id, field, operator: "is", value: parts.join(", ") };
}

export interface FilterConditionRowProps {
  filter: FilterCondition;
  fieldDef?: FieldDefinition;
  fieldSuggestions: string[];
  isLoadingSuggestions: boolean;
  onUpdate: (updates: Partial<FilterCondition>) => void;
  onRequestSuggestions?: () => void;
}

/**
 * Shared filter operator + value input + multi-value suggestions control used by
 * both FilterBuilder (optional filters) and PromptsDialog (always_filter required filters).
 */
export const FilterConditionRow: React.FC<FilterConditionRowProps> = ({
  filter: f,
  fieldDef: selectedFieldObj,
  fieldSuggestions,
  isLoadingSuggestions,
  onUpdate,
  onRequestSuggestions,
}) => {
  const styles = useStyles();
  const isDate = isDateField(selectedFieldObj);
  const isNumber = isNumberField(selectedFieldObj);
  const isString = isStringField(selectedFieldObj);

  const selectedStringTokens = isString ? splitMultiValueString(f.value) : [];

  const handleOperatorChange = (newOp: string) => {
    const updates: Partial<FilterCondition> = { operator: newOp };

    if (isDate) {
      if (newOp === "is_in_the_last") {
        updates.value = f.value || "7";
        updates.unit = f.unit || "days";
      } else if (newOp === "is_in_range") {
        updates.startValue = f.startValue || f.value || "";
        updates.endValue = f.endValue || "";
      } else if (newOp === "is_this" || newOp === "is_previous") {
        updates.unit = f.unit || "month";
      } else if (newOp === "is_next") {
        updates.value = f.value || "1";
        updates.unit = f.unit || "months";
      }
    }

    onUpdate(updates);
  };

  const handleSuggestionToggle = (suggestion: string) => {
    const nextOp =
      f.operator === "is_null" || f.operator === "is_not_null" ? "is" : f.operator;
    const baseVal =
      f.operator === "is_null" || f.operator === "is_not_null" ? "" : f.value;
    const nextVal = toggleMultiValueToken(baseVal, suggestion);
    onUpdate({ operator: nextOp, value: nextVal });
  };

  return (
    <>
      <div className={styles.conditionRow}>
        {/* Operator Select based on data type */}
        <Select
          size="small"
          className={styles.operatorSelect}
          value={f.operator}
          onChange={(_, data) => handleOperatorChange(data.value)}
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
                  onChange={(_, data) => onUpdate({ value: data.value })}
                  style={{ width: "65px", minWidth: "50px" }}
                />
                <Select
                  size="small"
                  value={f.unit || "days"}
                  onChange={(_, data) => onUpdate({ unit: data.value })}
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
                  onChange={(_, data) => onUpdate({ startValue: data.value })}
                  style={{ flex: 1, minWidth: "110px" }}
                />
                <Caption1 className={styles.rangeSeparator}>until (before)</Caption1>
                <Input
                  size="small"
                  type="date"
                  value={f.endValue ?? ""}
                  onChange={(_, data) => onUpdate({ endValue: data.value })}
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
                onChange={(_, data) => onUpdate({ value: data.value })}
              />
            )}

            {f.operator === "is_in_the_year" && (
              <Input
                size="small"
                type="number"
                placeholder="YYYY (e.g. 2026)"
                className={styles.valueControl}
                value={f.value ?? ""}
                onChange={(_, data) => onUpdate({ value: data.value })}
              />
            )}

            {f.operator === "is_in_the_month" && (
              <Input
                size="small"
                type="month"
                className={styles.valueControl}
                value={f.value ?? ""}
                onChange={(_, data) => onUpdate({ value: data.value })}
              />
            )}

            {(f.operator === "is_this" || f.operator === "is_previous") && (
              <Select
                size="small"
                className={styles.valueControl}
                value={f.unit || "month"}
                onChange={(_, data) => onUpdate({ unit: data.value })}
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
                  onChange={(_, data) => onUpdate({ value: data.value })}
                  style={{ width: "65px", minWidth: "50px" }}
                />
                <Select
                  size="small"
                  value={f.unit || "months"}
                  onChange={(_, data) => onUpdate({ unit: data.value })}
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
                onChange={(_, data) => onUpdate({ value: data.value })}
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
                  onChange={(_, data) => onUpdate({ startValue: data.value })}
                  style={{ flex: 1, minWidth: "60px" }}
                />
                <Caption1 className={styles.rangeSeparator}>to</Caption1>
                <Input
                  size="small"
                  type="number"
                  placeholder="Max"
                  value={f.endValue ?? ""}
                  onChange={(_, data) => onUpdate({ endValue: data.value })}
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
                onChange={(_, data) => onUpdate({ value: data.value })}
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
                  onChange={(_, data) => onUpdate({ value: data.value })}
                />
              )}
          </>
        )}

        {/* STRING FIELDS CONTROLS (Multi-Value Capable) */}
        {isString && (
          <>
            {(f.operator === "is_null" || f.operator === "is_not_null") && (
              <Caption1 className={styles.staticLabel}>
                {f.operator === "is_null" ? "is empty / null" : "is not empty"}
              </Caption1>
            )}

            {f.operator !== "is_null" && f.operator !== "is_not_null" && (
              <>
                {fieldSuggestions.length > 0 &&
                (f.operator === "is" || f.operator === "is_not") ? (
                  <Combobox
                    size="small"
                    multiselect
                    freeform
                    className={styles.valueControl}
                    placeholder={
                      isLoadingSuggestions
                        ? "Loading..."
                        : "Select one or more values (or comma-separate)..."
                    }
                    value={f.value}
                    selectedOptions={selectedStringTokens}
                    onChange={(e) => onUpdate({ value: e.target.value })}
                    onOptionSelect={(_, data) =>
                      onUpdate({ value: data.selectedOptions.join(", ") })
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
                        ? "Contains text (comma-separate for multiple)..."
                        : f.operator === "starts_with"
                        ? "Starts with (comma-separate for multiple)..."
                        : f.operator === "ends_with"
                        ? "Ends with (comma-separate for multiple)..."
                        : "Filter value(s) (comma-separate for multiple)..."
                    }
                    value={f.value}
                    onChange={(_, data) => onUpdate({ value: data.value })}
                    onFocus={onRequestSuggestions}
                  />
                )}
              </>
            )}
          </>
        )}
      </div>

      {/* Multi-Select Quick Suggestion Pills for String Fields */}
      {isString && fieldSuggestions.length > 0 && fieldSuggestions.length <= 10 && (
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
          {fieldSuggestions.map((s) => {
            const isSelected =
              f.operator !== "is_null" &&
              f.operator !== "is_not_null" &&
              selectedStringTokens.includes(s);
            return (
              <Badge
                key={s}
                size="small"
                appearance={isSelected ? "filled" : "outline"}
                color={isSelected ? "brand" : "informative"}
                style={{ cursor: "pointer", fontSize: "10px" }}
                onClick={() => handleSuggestionToggle(s)}
              >
                {s}
              </Badge>
            );
          })}
        </div>
      )}
    </>
  );
};

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

            {/* Row 2: Operator & Specialized Value Input + Suggestions */}
            <FilterConditionRow
              filter={f}
              fieldDef={selectedFieldObj}
              fieldSuggestions={fieldSuggestions}
              isLoadingSuggestions={isLoadingSuggestions}
              onUpdate={(updates) => onUpdateFilter(f.id, updates)}
              onRequestSuggestions={() => {
                if (f.field && !suggestionsCache[f.field]) {
                  fetchSuggestions(f.field);
                }
              }}
            />
          </div>
        );
      })}
    </div>
  );
};
