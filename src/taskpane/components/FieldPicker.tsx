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
  Popover,
  PopoverTrigger,
  PopoverSurface,
} from "@fluentui/react-components";
import {
  SearchRegular,
  DismissRegular,
  Info16Regular,
  ChevronRight16Regular,
  ChevronDown16Regular,
  Folder16Regular,
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
    gap: "4px",
    minWidth: 0,
    flex: 1,
  },
  subgroupHeader: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    padding: "4px 6px",
    cursor: "pointer",
    borderRadius: "4px",
    userSelect: "none",
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
  },
  subgroupContent: {
    display: "flex",
    flexDirection: "column",
    gap: "2px",
    paddingLeft: "16px",
    marginLeft: "8px",
    borderLeft: `2px solid ${tokens.colorNeutralStroke2}`,
  },
  infoButton: {
    background: "transparent",
    border: "none",
    padding: "2px",
    margin: "0 2px",
    cursor: "pointer",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    color: tokens.colorNeutralForeground3,
    borderRadius: "4px",
    ":hover": {
      color: tokens.colorBrandForeground1,
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
  },
});

interface FieldPickerProps {
  dimensions: FieldDefinition[];
  measures: FieldDefinition[];
  selectedFields: string[];
  onToggleField: (fieldName: string) => void;
  onClearSelected: () => void;
}

interface SubgroupStructure {
  groupName: string;
  fields: FieldDefinition[];
}

interface ViewStructure {
  viewName: string;
  subgroups: SubgroupStructure[];
  ungrouped: FieldDefinition[];
  totalCount: number;
}

function resolveGroupLabel(f: FieldDefinition): string | null {
  if (f.field_group_label) return f.field_group_label;
  if (f.group_label) return f.group_label;
  if (f.dimension_group) {
    const parts = f.dimension_group.split(".");
    const raw = parts[parts.length - 1];
    return raw.charAt(0).toUpperCase() + raw.slice(1).replace(/_/g, " ") + " Date";
  }
  return null;
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
  const [openSubgroups, setOpenSubgroups] = useState<Record<string, boolean>>({});

  const toggleSubgroup = (key: string) => {
    setOpenSubgroups((prev) => ({ ...prev, [key]: !prev[key] }));
  };

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
        (f.label_short && f.label_short.toLowerCase().includes(term)) ||
        (f.view_label && f.view_label.toLowerCase().includes(term)) ||
        (f.field_group_label && f.field_group_label.toLowerCase().includes(term)) ||
        (f.field_group_variant && f.field_group_variant.toLowerCase().includes(term))
    );
  }, [allFields, searchTerm]);

  // Group by view_label, then sub-group by field_group_label / dimension_group
  const viewsHierarchy = useMemo(() => {
    const viewsMap: Record<string, { grouped: Record<string, FieldDefinition[]>; ungrouped: FieldDefinition[] }> = {};

    filteredFields.forEach((f) => {
      const view = f.view_label || "Other";
      if (!viewsMap[view]) {
        viewsMap[view] = { grouped: {}, ungrouped: [] };
      }

      const group = resolveGroupLabel(f);
      if (group) {
        if (!viewsMap[view].grouped[group]) {
          viewsMap[view].grouped[group] = [];
        }
        viewsMap[view].grouped[group].push(f);
      } else {
        viewsMap[view].ungrouped.push(f);
      }
    });

    const result: ViewStructure[] = Object.entries(viewsMap).map(([viewName, data]) => {
      const subgroups: SubgroupStructure[] = Object.entries(data.grouped).map(([groupName, fields]) => ({
        groupName,
        fields,
      }));

      // Sort subgroups alphabetically
      subgroups.sort((a, b) => a.groupName.localeCompare(b.groupName));

      const totalCount =
        subgroups.reduce((acc, g) => acc + g.fields.length, 0) + data.ungrouped.length;

      return {
        viewName,
        subgroups,
        ungrouped: data.ungrouped,
        totalCount,
      };
    });

    return result;
  }, [filteredFields]);

  // Render an individual field row with Checkbox, Info popover, and Badges
  const renderFieldRow = (f: FieldDefinition, isNested: boolean = false) => {
    const isSelected = selectedFields.includes(f.name);
    const isMeasure = f.category === "measure";
    const displayName = isNested && f.field_group_variant
      ? f.field_group_variant
      : f.label_short || f.label || f.name;

    return (
      <div key={f.name} className={styles.fieldRow}>
        <div className={styles.fieldLabel}>
          <Checkbox
            checked={isSelected}
            onChange={() => onToggleField(f.name)}
            label={<Body2 style={{ fontSize: "12px" }}>{displayName}</Body2>}
          />

          {/* Field Description Info Click Popover */}
          <Popover withArrow positioning="after-top">
            <PopoverTrigger disableButtonEnhancement>
              <button
                type="button"
                className={styles.infoButton}
                title="Click for field details & description"
                onClick={(e) => e.stopPropagation()}
              >
                <Info16Regular />
              </button>
            </PopoverTrigger>
            <PopoverSurface style={{ maxWidth: "260px", padding: "10px 12px" }}>
              <div style={{ display: "flex", flexDirection: "column", gap: "4px" }}>
                <Body2 style={{ fontWeight: 600 }}>{f.label || f.name}</Body2>
                <Caption1 style={{ color: tokens.colorNeutralForeground2, lineHeight: "16px" }}>
                  {f.description || "No description provided in LookML."}
                </Caption1>
                <div
                  style={{
                    marginTop: "6px",
                    paddingTop: "6px",
                    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
                    fontSize: "11px",
                    color: tokens.colorNeutralForeground4,
                  }}
                >
                  <code>{f.name}</code> {f.type ? `• ${f.type}` : ""}
                </div>
              </div>
            </PopoverSurface>
          </Popover>
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
  };

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
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              marginBottom: "4px",
            }}
          >
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
      <div style={{ maxHeight: "300px", overflowY: "auto" }}>
        <Accordion
          collapsible
          multiple
          defaultOpenItems={viewsHierarchy.slice(0, 2).map((v) => v.viewName)}
        >
          {viewsHierarchy.map((view) => (
            <AccordionItem key={view.viewName} value={view.viewName}>
              <AccordionHeader size="small">
                {view.viewName} ({view.totalCount})
              </AccordionHeader>
              <AccordionPanel>
                <div style={{ display: "flex", flexDirection: "column", gap: "4px" }}>
                  {/* Second-Level Subgroups (dimension_group / group_label) */}
                  {view.subgroups.map((subgroup) => {
                    const groupKey = `${view.viewName}::${subgroup.groupName}`;
                    const isExpanded = searchTerm.trim()
                      ? true
                      : !!openSubgroups[groupKey];

                    return (
                      <div key={subgroup.groupName}>
                        <div
                          className={styles.subgroupHeader}
                          onClick={() => toggleSubgroup(groupKey)}
                        >
                          {isExpanded ? (
                            <ChevronDown16Regular style={{ color: tokens.colorNeutralForeground3 }} />
                          ) : (
                            <ChevronRight16Regular style={{ color: tokens.colorNeutralForeground3 }} />
                          )}
                          <Folder16Regular style={{ color: tokens.colorBrandForeground1 }} />
                          <Body2 style={{ fontWeight: 600, fontSize: "12px", flex: 1 }}>
                            {subgroup.groupName} ({subgroup.fields.length})
                          </Body2>
                        </div>

                        {isExpanded && (
                          <div className={styles.subgroupContent}>
                            {subgroup.fields.map((f) => renderFieldRow(f, true))}
                          </div>
                        )}
                      </div>
                    );
                  })}

                  {/* Ungrouped Fields */}
                  {view.ungrouped.map((f) => renderFieldRow(f, false))}
                </div>
              </AccordionPanel>
            </AccordionItem>
          ))}
        </Accordion>
      </div>
    </div>
  );
};
