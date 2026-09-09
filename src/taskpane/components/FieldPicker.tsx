import React, { useState, useMemo } from "react";
import {
  makeStyles,
  tokens,
  Input,
  Checkbox,
  Badge,
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
  Caption1,
  Body2,
  Button,
} from "@fluentui/react-components";
import {
  SearchRegular,
  DismissRegular,
  NumberSymbolRegular,
  TextAlignLeftRegular,
} from "@fluentui/react-icons";
import { FieldDefinition } from "../services/lookerClient";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "10px",
    padding: "0 16px 12px 16px",
  },
  selectedBar: {
    display: "flex",
    flexWrap: "wrap",
    gap: "6px",
    padding: "8px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: "4px",
    maxHeight: "100px",
    overflowY: "auto",
  },
  selectedChip: {
    display: "inline-flex",
    alignItems: "center",
    gap: "4px",
    padding: "2px 8px",
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: "12px",
    fontSize: "12px",
  },
  fieldRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "2px 0",
  },
  fieldLabel: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
  },
});

interface FieldPickerProps {
  dimensions: FieldDefinition[];
  measures: FieldDefinition[];
  selectedFields: string[];
  onToggleField: (fieldName: string) => void;
  onClearSelected: () => void;
}

export const FieldPicker: React.FC<FieldPickerProps> = ({
  dimensions,
  measures,
  selectedFields,
  onToggleField,
  onClearSelected,
}) => {
  const styles = useStyles();
  const [searchTerm, setSearchTerm] = useState("");

  const allFields = useMemo(
    () => [...dimensions, ...measures],
    [dimensions, measures]
  );

  // Filter fields based on search
  const filteredFields = useMemo(() => {
    if (!searchTerm.trim()) return allFields;
    const term = searchTerm.toLowerCase();
    return allFields.filter(
      (f) =>
        f.name.toLowerCase().includes(term) ||
        (f.label && f.label.toLowerCase().includes(term)) ||
        (f.view_label && f.view_label.toLowerCase().includes(term))
    );
  }, [allFields, searchTerm]);

  // Group by view_label
  const groupedViews = useMemo(() => {
    const groups: Record<string, FieldDefinition[]> = {};
    filteredFields.forEach((f) => {
      const view = f.view_label || "Other";
      if (!groups[view]) groups[view] = [];
      groups[view].push(f);
    });
    return groups;
  }, [filteredFields]);

  return (
    <div className={styles.container}>
      <Input
        size="small"
        contentBefore={<SearchRegular />}
        placeholder="Search dimensions & measures..."
        value={searchTerm}
        onChange={(_, data) => setSearchTerm(data.value)}
      />

      {/* Selected fields chip drawer */}
      {selectedFields.length > 0 && (
        <div>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "4px" }}>
            <Caption1 style={{ fontWeight: 600 }}>
              Selected Fields ({selectedFields.length})
            </Caption1>
            <Button size="small" appearance="subtle" onClick={onClearSelected}>
              Clear All
            </Button>
          </div>
          <div className={styles.selectedBar}>
            {selectedFields.map((fieldKey) => {
              const def = allFields.find((f) => f.name === fieldKey);
              return (
                <div key={fieldKey} className={styles.selectedChip}>
                  <span>{def?.label || fieldKey}</span>
                  <DismissRegular
                    style={{ cursor: "pointer", fontSize: "12px" }}
                    onClick={() => onToggleField(fieldKey)}
                  />
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* Grouped Tree Accordion */}
      <div style={{ maxHeight: "260px", overflowY: "auto" }}>
        <Accordion collapsible multiple defaultOpenItems={Object.keys(groupedViews).slice(0, 2)}>
          {Object.entries(groupedViews).map(([viewName, fields]) => (
            <AccordionItem key={viewName} value={viewName}>
              <AccordionHeader size="small">
                {viewName} ({fields.length})
              </AccordionHeader>
              <AccordionPanel>
                <div style={{ display: "flex", flexDirection: "column", gap: "2px" }}>
                  {fields.map((f) => {
                    const isSelected = selectedFields.includes(f.name);
                    const isMeasure = f.category === "measure";

                    return (
                      <div key={f.name} className={styles.fieldRow}>
                        <div className={styles.fieldLabel}>
                          <Checkbox
                            checked={isSelected}
                            onChange={() => onToggleField(f.name)}
                            label={
                              <Body2 style={{ fontSize: "12px" }}>
                                {f.label || f.name}
                              </Body2>
                            }
                          />
                        </div>
                        <div style={{ display: "flex", alignItems: "center", gap: "4px" }}>
                          {f.value_format && (
                            <Badge size="small" appearance="outline" color="informative">
                              {f.value_format}
                            </Badge>
                          )}
                          <Badge
                            size="small"
                            appearance="tint"
                            color={isMeasure ? "warning" : "brand"}
                          >
                            {isMeasure ? "M" : "D"}
                          </Badge>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </AccordionPanel>
            </AccordionItem>
          ))}
        </Accordion>
      </div>
    </div>
  );
};
