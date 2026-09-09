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
  filterRow: {
    display: "grid",
    gridTemplateColumns: "1fr 85px 1fr 30px",
    gap: "6px",
    alignItems: "center",
  },
  suggestionsBar: {
    display: "flex",
    flexWrap: "wrap",
    gap: "4px",
    gridColumn: "1 / -1",
    paddingLeft: "4px",
    marginTop: "-2px",
    marginBottom: "4px",
  },
});

export interface FilterCondition {
  id: string;
  field: string;
  operator: string;
  value: string;
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

  // Fetch field suggestions from Looker API:
  // GET /api/4.0/models/{model_name}/views/{view_name}/fields/{field_name}/suggestions
  const fetchSuggestions = useCallback(
    async (fieldName: string) => {
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
    [baseUrl, token, modelName, exploreName, suggestionsCache, loadingFields]
  );

  // Automatically fetch suggestions when a filter field is configured
  useEffect(() => {
    filters.forEach((f) => {
      if (f.field && !suggestionsCache[f.field] && !loadingFields[f.field]) {
        fetchSuggestions(f.field);
      }
    });
  }, [filters, fetchSuggestions, suggestionsCache, loadingFields]);

  const handleFieldChange = (filterId: string, newField: string) => {
    onUpdateFilter(filterId, { field: newField, value: "" });
    if (newField) {
      fetchSuggestions(newField);
    }
  };

  return (
    <div className={styles.container}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
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

      {filters.map((f) => {
        const fieldSuggestions = f.field ? suggestionsCache[f.field] || [] : [];
        const isLoadingSuggestions = f.field ? !!loadingFields[f.field] : false;

        return (
          <React.Fragment key={f.id}>
            <div className={styles.filterRow}>
              {/* Field Select */}
              <Select
                size="small"
                value={f.field}
                onChange={(_, data) => handleFieldChange(f.id, data.value)}
              >
                <option value="">Select Field</option>
                {fields.map((fld) => (
                  <option key={fld.name} value={fld.name}>
                    {fld.label || fld.name}
                  </option>
                ))}
              </Select>

              {/* Operator Select */}
              <Select
                size="small"
                value={f.operator}
                onChange={(_, data) => onUpdateFilter(f.id, { operator: data.value })}
              >
                <option value="is">is</option>
                <option value="is_not">is not</option>
                <option value="contains">contains</option>
                <option value="before">before</option>
                <option value="after">after</option>
                <option value="greater_than">&gt;</option>
                <option value="less_than">&lt;</option>
              </Select>

              {/* Filter Value: Combobox with suggestions if available, else Input */}
              {fieldSuggestions.length > 0 ? (
                <Combobox
                  size="small"
                  freeform
                  placeholder={isLoadingSuggestions ? "Loading..." : "Value or pick..."}
                  value={f.value}
                  onChange={(e) => onUpdateFilter(f.id, { value: e.target.value })}
                  onOptionSelect={(_, data) =>
                    onUpdateFilter(f.id, { value: data.optionValue || data.optionText || "" })
                  }
                  style={{ minWidth: 0 }}
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
                  placeholder={isLoadingSuggestions ? "Checking..." : "Value..."}
                  value={f.value}
                  onChange={(_, data) => onUpdateFilter(f.id, { value: data.value })}
                  onFocus={() => {
                    if (f.field && !suggestionsCache[f.field]) {
                      fetchSuggestions(f.field);
                    }
                  }}
                />
              )}

              {/* Remove Filter Button */}
              <Button
                size="small"
                appearance="subtle"
                icon={<DeleteRegular />}
                onClick={() => onRemoveFilter(f.id)}
                title="Remove Filter"
              />
            </div>

            {/* Quick Suggestion Pills for Compact Option Lists (<= 8 items) */}
            {fieldSuggestions.length > 0 && fieldSuggestions.length <= 8 && (
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
          </React.Fragment>
        );
      })}
    </div>
  );
};
