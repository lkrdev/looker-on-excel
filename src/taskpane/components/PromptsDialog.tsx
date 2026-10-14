import React, { useState, useEffect, useCallback } from "react";
import {
  makeStyles,
  tokens,
  MessageBar,
  MessageBarBody,
  Label,
  Input,
  Select,
  Caption1,
} from "@fluentui/react-components";
import { WarningRegular } from "@fluentui/react-icons";
import {
  ParameterDefinition,
  FieldDefinition,
  AlwaysFilterItem,
  getFieldSuggestions,
} from "../services/lookerClient";
import {
  FilterCondition,
  FilterConditionRow,
  isStringField,
  parseLookerExpressionToFilterCondition,
  compileFilterToLookerExpression,
} from "./FilterBuilder";

const useStyles = makeStyles({
  container: {
    margin: "0 16px 12px 16px",
    display: "flex",
    flexDirection: "column",
    gap: "12px",
  },
  section: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  sectionHeader: {
    fontWeight: 600,
    color: tokens.colorNeutralForeground2,
    textTransform: "uppercase",
    letterSpacing: "0.5px",
  },
  promptCard: {
    display: "flex",
    flexDirection: "column",
    gap: "6px",
    padding: "8px 10px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  labelStack: {
    display: "flex",
    flexDirection: "column",
    gap: "1px",
  },
  viewCaption: {
    fontSize: "11px",
    color: tokens.colorNeutralForeground3,
    lineHeight: "14px",
  },
  promptRow: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
  },
});

interface PromptsDialogProps {
  parameters?: ParameterDefinition[];
  alwaysFilters?: Array<
    AlwaysFilterItem | { field?: string; values?: string[]; name?: string; value?: string }
  >;
  fields?: FieldDefinition[];
  promptValues: Record<string, string>;
  onPromptChange: (field: string, value: string) => void;
  baseUrl?: string;
  token?: string;
  modelName?: string;
  exploreName?: string;
}

export const PromptsDialog: React.FC<PromptsDialogProps> = ({
  parameters = [],
  alwaysFilters = [],
  fields = [],
  promptValues,
  onPromptChange,
  baseUrl,
  token,
  modelName,
  exploreName,
}) => {
  const styles = useStyles();

  const [suggestionsCache, setSuggestionsCache] = useState<Record<string, string[]>>({});
  const [loadingFields, setLoadingFields] = useState<Record<string, boolean>>({});
  const [conditionStates, setConditionStates] = useState<Record<string, FilterCondition>>({});

  // Normalize always_filter entries so both { name, value } and { field, values } work seamlessly
  const normalizedAlwaysFilters = alwaysFilters
    .map((f) => {
      const field = (f.field || f.name || "").trim();
      const defaultVal =
        Array.isArray(f.values) && f.values.length > 0
          ? f.values.join(", ")
          : f.value !== undefined && f.value !== null
          ? String(f.value)
          : "";
      return { field, defaultVal };
    })
    .filter((f) => f.field !== "");

  const fetchSuggestions = useCallback(
    async (fieldName: string) => {
      const fieldDef = fields.find((fld) => fld.name === fieldName);
      if (fieldDef && !isStringField(fieldDef)) return;
      if (!baseUrl || !token || !modelName || !exploreName || !fieldName) return;
      if (suggestionsCache[fieldName] || loadingFields[fieldName]) return;

      setLoadingFields((prev) => ({ ...prev, [fieldName]: true }));
      try {
        const list = await getFieldSuggestions(baseUrl, token, modelName, exploreName, fieldName);
        setSuggestionsCache((prev) => ({ ...prev, [fieldName]: list }));
      } catch (err) {
        console.warn("Error fetching required filter suggestions:", err);
      } finally {
        setLoadingFields((prev) => ({ ...prev, [fieldName]: false }));
      }
    },
    [baseUrl, token, modelName, exploreName, fields, suggestionsCache, loadingFields]
  );

  // Automatically fetch suggestions for string always_filter fields
  useEffect(() => {
    normalizedAlwaysFilters.forEach(({ field }) => {
      const fieldDef = fields.find((fld) => fld.name === field);
      if (
        field &&
        (!fieldDef || isStringField(fieldDef)) &&
        !suggestionsCache[field] &&
        !loadingFields[field]
      ) {
        fetchSuggestions(field);
      }
    });
  }, [normalizedAlwaysFilters, fields, fetchSuggestions, suggestionsCache, loadingFields]);

  // Reset local condition states when model/explore changes
  useEffect(() => {
    setConditionStates({});
  }, [modelName, exploreName]);

  if (parameters.length === 0 && normalizedAlwaysFilters.length === 0) {
    return null;
  }

  return (
    <div className={styles.container}>
      {/* 1. True Required Filters (always_filter defined in LookML) */}
      {normalizedAlwaysFilters.length > 0 && (
        <div className={styles.section}>
          <MessageBar intent="warning" icon={<WarningRegular />}>
            <MessageBarBody>
              <strong>Required Filter{normalizedAlwaysFilters.length > 1 ? "s" : ""}</strong>
              <Caption1 style={{ display: "block" }}>
                This Explore requires the following filter
                {normalizedAlwaysFilters.length > 1 ? "s" : ""} before running:
              </Caption1>
            </MessageBarBody>
          </MessageBar>

          {normalizedAlwaysFilters.map(({ field, defaultVal }) => {
            const paramDef = parameters.find((p) => p.name === field);
            const fieldDef = fields.find((fld) => fld.name === field);
            const currentExpr = promptValues[field] ?? defaultVal;

            // If always_filter references a parameter, render single-value control
            if (paramDef) {
              return (
                <div key={field} className={styles.promptCard} title={field}>
                  <div className={styles.labelStack}>
                    {paramDef.view_label && (
                      <Caption1 className={styles.viewCaption}>{paramDef.view_label}</Caption1>
                    )}
                    <Label size="small" weight="semibold" required>
                      {paramDef.label_short || paramDef.label || field}
                    </Label>
                  </div>
                  {paramDef.enumerations && paramDef.enumerations.length > 0 ? (
                    <Select
                      size="small"
                      value={currentExpr}
                      onChange={(_, data) => onPromptChange(field, data.value)}
                    >
                      {paramDef.enumerations.map((en) => (
                        <option key={en.value} value={en.value}>
                          {en.label || en.value}
                        </option>
                      ))}
                    </Select>
                  ) : (
                    <Input
                      size="small"
                      value={currentExpr}
                      onChange={(_, data) => onPromptChange(field, data.value)}
                      placeholder={
                        paramDef.default_value
                          ? `Default: ${paramDef.default_value}`
                          : `Enter ${paramDef.label_short || paramDef.label || field}`
                      }
                    />
                  )}
                </div>
              );
            }

            // Standard LookML field (dimension/measure/filter): full FilterConditionRow with multi-value string support
            const activeCondition: FilterCondition =
              conditionStates[field] ??
              parseLookerExpressionToFilterCondition(
                field,
                currentExpr,
                fieldDef,
                `always_${field}`
              );

            const isString = !fieldDef || isStringField(fieldDef);
            const fieldSuggestions = isString ? suggestionsCache[field] || [] : [];
            const isLoadingSuggestions = isString ? !!loadingFields[field] : false;

            const handleConditionUpdate = (updates: Partial<FilterCondition>) => {
              const nextCondition: FilterCondition = { ...activeCondition, ...updates };
              setConditionStates((prev) => ({ ...prev, [field]: nextCondition }));
              const compiled = compileFilterToLookerExpression(nextCondition, fieldDef);
              onPromptChange(field, compiled ?? "");
            };

            return (
              <div key={field} className={styles.promptCard} title={field}>
                <div className={styles.labelStack}>
                  {fieldDef?.view_label && (
                    <Caption1 className={styles.viewCaption}>{fieldDef.view_label}</Caption1>
                  )}
                  <Label size="small" weight="semibold" required>
                    {fieldDef?.label_short || fieldDef?.label || field}
                  </Label>
                </div>

                <FilterConditionRow
                  filter={activeCondition}
                  fieldDef={fieldDef}
                  fieldSuggestions={fieldSuggestions}
                  isLoadingSuggestions={isLoadingSuggestions}
                  onUpdate={handleConditionUpdate}
                  onRequestSuggestions={() => {
                    if (isString && !suggestionsCache[field]) {
                      fetchSuggestions(field);
                    }
                  }}
                />
              </div>
            );
          })}
        </div>
      )}

      {/* 2. Optional LookML Parameters (Single Value Only) */}
      {parameters.length > 0 && (
        <div className={styles.section}>
          <Caption1 className={styles.sectionHeader}>
            Parameters ({parameters.length})
          </Caption1>
          {parameters.map((p) => (
            <div key={p.name} className={styles.promptRow} title={p.name}>
              <div className={styles.labelStack}>
                {p.view_label && (
                  <Caption1 className={styles.viewCaption}>{p.view_label}</Caption1>
                )}
                <Label size="small" weight="regular">
                  {p.label_short || p.label || p.name}
                </Label>
              </div>
              {p.enumerations && p.enumerations.length > 0 ? (
                <Select
                  size="small"
                  value={promptValues[p.name] ?? ""}
                  onChange={(_, data) => onPromptChange(p.name, data.value)}
                >
                  <option value="">
                    {p.default_value ? `Default (${p.default_value})` : "Select value (optional)"}
                  </option>
                  {p.enumerations.map((en) => (
                    <option key={en.value} value={en.value}>
                      {en.label || en.value}
                    </option>
                  ))}
                </Select>
              ) : (
                <Input
                  size="small"
                  value={promptValues[p.name] ?? ""}
                  onChange={(_, data) => onPromptChange(p.name, data.value)}
                  placeholder={
                    p.default_value
                      ? `Default: ${p.default_value}`
                      : `Enter ${p.label_short || p.label || p.name}`
                  }
                />
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
};
