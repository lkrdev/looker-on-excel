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
  Select,
  Switch,
} from "@fluentui/react-components";
import {
  SearchRegular,
  DismissRegular,
  Info16Regular,
  ChevronRight16Regular,
  ChevronDown16Regular,
  Folder16Regular,
  ReorderRegular,
} from "@fluentui/react-icons";
import { FieldDefinition } from "../services/lookerClient";
import { WriteOptions, RowGroupDisplayLevel } from "../services/excelWriter";
import {
  resolveExploreHierarchies,
  getFieldHierarchyContext,
  expandFieldInSelection,
  applyHierarchyLadderToSelection,
} from "../services/hierarchyResolver";

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
  sorts?: string[];
  onToggleSort?: (fieldName: string) => void;
  pivots?: string[];
  onTogglePivot?: (fieldName: string) => void;
  onSetPivots?: (newPivots: string[]) => void;
  onReorderFields?: (newFields: string[]) => void;
  options?: WriteOptions;
  onOptionsChange?: (updates: Partial<WriteOptions>) => void;
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

/**
 * Checks if a field matches the search query across name, label, short label,
 * view label, field group, and synonyms.
 */
export function matchesFieldSearch(f: FieldDefinition, searchTerm: string): boolean {
  if (!searchTerm.trim()) return true;
  const term = searchTerm.toLowerCase();
  return (
    f.name.toLowerCase().includes(term) ||
    (f.label ? f.label.toLowerCase().includes(term) : false) ||
    (f.label_short ? f.label_short.toLowerCase().includes(term) : false) ||
    (f.view_label ? f.view_label.toLowerCase().includes(term) : false) ||
    (f.field_group_label ? f.field_group_label.toLowerCase().includes(term) : false) ||
    (f.field_group_variant ? f.field_group_variant.toLowerCase().includes(term) : false) ||
    (f.synonyms ? f.synonyms.some((s) => s.toLowerCase().includes(term)) : false)
  );
}

/**
 * Returns the specific synonym that matched if the field matched ONLY via synonym.
 */
export function getMatchedSynonym(f: FieldDefinition, searchTerm: string): string | null {
  if (!searchTerm.trim()) return null;
  const term = searchTerm.toLowerCase();
  const directMatch =
    f.name.toLowerCase().includes(term) ||
    (f.label ? f.label.toLowerCase().includes(term) : false) ||
    (f.label_short ? f.label_short.toLowerCase().includes(term) : false) ||
    (f.view_label ? f.view_label.toLowerCase().includes(term) : false) ||
    (f.field_group_label ? f.field_group_label.toLowerCase().includes(term) : false) ||
    (f.field_group_variant ? f.field_group_variant.toLowerCase().includes(term) : false);

  if (directMatch) return null;
  return f.synonyms?.find((s) => s.toLowerCase().includes(term)) || null;
}

export const FieldPicker: React.FC<FieldPickerProps> = ({
  dimensions,
  measures,
  selectedFields,
  onToggleField,
  onClearSelected,
  sorts = [],
  onToggleSort,
  pivots = [],
  onTogglePivot,
  onSetPivots,
  onReorderFields,
  options,
  onOptionsChange,
}) => {
  const styles = useStyles();
  const [searchTerm, setSearchTerm] = useState("");
  const [openSubgroups, setOpenSubgroups] = useState<Record<string, boolean>>({});
  const [draggedIndex, setDraggedIndex] = useState<number | null>(null);
  const [dragOverIndex, setDragOverIndex] = useState<number | null>(null);

  const toggleSubgroup = (key: string) => {
    setOpenSubgroups((prev) => ({ ...prev, [key]: !prev[key] }));
  };

  const allFields = useMemo(
    () => [...dimensions, ...measures],
    [dimensions, measures]
  );

  const hierarchyLadders = useMemo(
    () => resolveExploreHierarchies(dimensions),
    [dimensions]
  );

  const nonPivotSelectedDims = useMemo(() => {
    const dimSet = new Set(dimensions.map((d) => d.name));
    return selectedFields.filter((f) => dimSet.has(f) && !pivots.includes(f));
  }, [selectedFields, dimensions, pivots]);

  const orderedRowGroupDims = useMemo(() => {
    if (options?.rowGroupFields && options.rowGroupFields.length >= 2) {
      const matched = options.rowGroupFields.filter((f) =>
        nonPivotSelectedDims.includes(f)
      );
      const extras = nonPivotSelectedDims.filter((f) => !matched.includes(f));
      return [...matched, ...extras];
    }
    return nonPivotSelectedDims;
  }, [options?.rowGroupFields, nonPivotSelectedDims]);

  // Filter fields based on search (including LookML synonyms)
  const filteredFields = useMemo(() => {
    if (!searchTerm.trim()) return allFields;
    return allFields.filter((f) => matchesFieldSearch(f, searchTerm));
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
    const matchedSynonym = getMatchedSynonym(f, searchTerm);

    return (
      <div key={f.name} className={styles.fieldRow}>
        <div className={styles.fieldLabel}>
          <Checkbox
            checked={isSelected}
            onChange={() => onToggleField(f.name)}
            label={
              <div style={{ display: "flex", flexDirection: "column", gap: "1px" }}>
                <Body2 style={{ fontSize: "12px" }}>{displayName}</Body2>
                {matchedSynonym && (
                  <Caption1 style={{ fontSize: "10px", color: tokens.colorBrandForeground1 }}>
                    synonym: &quot;{matchedSynonym}&quot;
                  </Caption1>
                )}
              </div>
            }
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
                {f.synonyms && f.synonyms.length > 0 && (
                  <div
                    style={{
                      marginTop: "4px",
                      paddingTop: "4px",
                      borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
                    }}
                  >
                    <Caption1
                      style={{
                        fontWeight: 600,
                        color: tokens.colorNeutralForeground3,
                        display: "block",
                        marginBottom: "2px",
                      }}
                    >
                      Synonyms:
                    </Caption1>
                    <div style={{ display: "flex", flexWrap: "wrap", gap: "4px" }}>
                      {f.synonyms.map((s, idx) => (
                        <Badge key={idx} size="small" appearance="tint" color="informative">
                          {s}
                        </Badge>
                      ))}
                    </div>
                  </div>
                )}
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
            {selectedFields.map((fieldKey, idx) => {
              const def = allFields.find((f) => f.name === fieldKey);
              const isDimension = def?.category === "dimension";
              const isPivoted = pivots.includes(fieldKey);
              const sortEntry = sorts.find((s) => s === fieldKey || s.startsWith(`${fieldKey} `));
              const isAsc = sortEntry === fieldKey || sortEntry === `${fieldKey} asc`;
              const isDesc = sortEntry === `${fieldKey} desc`;
              const isDragging = draggedIndex === idx;
              const isDropTarget = dragOverIndex === idx;

              return (
                <div
                  key={fieldKey}
                  className={styles.selectedChip}
                  draggable={!!onReorderFields}
                  onDragStart={(e) => {
                    setDraggedIndex(idx);
                    e.dataTransfer.effectAllowed = "move";
                    e.dataTransfer.setData("text/plain", String(idx));
                  }}
                  onDragOver={(e) => {
                    e.preventDefault();
                    e.dataTransfer.dropEffect = "move";
                  }}
                  onDragEnter={() => setDragOverIndex(idx)}
                  onDragEnd={() => {
                    setDraggedIndex(null);
                    setDragOverIndex(null);
                  }}
                  onDrop={(e) => {
                    e.preventDefault();
                    if (draggedIndex !== null && draggedIndex !== idx && onReorderFields) {
                      const next = [...selectedFields];
                      const [moved] = next.splice(draggedIndex, 1);
                      next.splice(idx, 0, moved);
                      onReorderFields(next);
                    }
                    setDraggedIndex(null);
                    setDragOverIndex(null);
                  }}
                  style={{
                    cursor: onReorderFields ? "grab" : "default",
                    opacity: isDragging ? 0.4 : 1,
                    border: isDropTarget ? `1px dashed ${tokens.colorBrandStroke1}` : undefined,
                    backgroundColor: isDropTarget ? tokens.colorBrandBackground2 : undefined,
                    transition: "border 0.15s ease, background-color 0.15s ease",
                  }}
                  title={onReorderFields ? "Drag to reorder column order" : undefined}
                >
                  {onReorderFields && (
                    <ReorderRegular
                      style={{
                        fontSize: "12px",
                        color: tokens.colorNeutralForeground4,
                        cursor: "grab",
                        marginRight: "2px",
                      }}
                    />
                  )}
                  <span style={{ fontWeight: 500 }}>{def?.label || fieldKey}</span>

                  {/* Pivot Toggle for Dimensions */}
                  {isDimension && onTogglePivot && (
                    <span
                      title={isPivoted ? "Remove pivot (Row)" : "Pivot across columns (Column)"}
                      style={{
                        cursor: "pointer",
                        fontSize: "10px",
                        fontWeight: 600,
                        padding: "1px 5px",
                        borderRadius: "8px",
                        backgroundColor: isPivoted ? tokens.colorBrandBackground : tokens.colorNeutralBackground3,
                        color: isPivoted ? tokens.colorNeutralForegroundOnBrand : tokens.colorNeutralForeground2,
                      }}
                      onClick={() => onTogglePivot(fieldKey)}
                    >
                      {isPivoted ? "PIVOT" : "+Pivot"}
                    </span>
                  )}

                  {/* Sort Direction Toggle */}
                  {onToggleSort && (
                    <span
                      title={
                        isAsc
                          ? "Sorted Ascending. Click to sort Descending."
                          : isDesc
                          ? "Sorted Descending. Click to remove sort."
                          : "Click to sort Ascending."
                      }
                      style={{
                        cursor: "pointer",
                        fontSize: "11px",
                        fontWeight: 700,
                        padding: "1px 4px",
                        borderRadius: "4px",
                        color: isAsc || isDesc ? tokens.colorBrandForeground1 : tokens.colorNeutralForeground4,
                        backgroundColor: isAsc || isDesc ? tokens.colorBrandBackground2 : "transparent",
                      }}
                      onClick={() => onToggleSort(fieldKey)}
                    >
                      {isAsc ? "↑" : isDesc ? "↓" : "↕"}
                    </span>
                  )}

                  <DismissRegular
                    style={{ cursor: "pointer", fontSize: "12px", marginLeft: "2px" }}
                    onClick={() => onToggleField(fieldKey)}
                  />
                </div>
              );
            })}
          </div>

          {pivots.length > 0 && (
            <div
              style={{
                marginTop: "6px",
                padding: "6px 8px",
                backgroundColor: tokens.colorBrandBackground2,
                borderRadius: "6px",
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                gap: "8px",
              }}
            >
              <Caption1 style={{ fontSize: "11px" }}>
                <strong>Crosstab Active:</strong> {pivots.length} column pivot
                {pivots.length > 1 ? "s" : ""} (
                {pivots
                  .map((p) => allFields.find((f) => f.name === p)?.label_short || p)
                  .join(" → ")}
                )
              </Caption1>
              {onSetPivots && (
                <Button
                  size="small"
                  appearance="outline"
                  title="Swap row dimensions and column pivot dimensions"
                  onClick={() => {
                    const selectedDims = selectedFields.filter(
                      (fk) => allFields.find((f) => f.name === fk)?.category === "dimension"
                    );
                    const rowDims = selectedDims.filter((d) => !pivots.includes(d));
                    if (rowDims.length > 0) {
                      onSetPivots(rowDims);
                    }
                  }}
                >
                  ⇄ Swap Rows &amp; Columns
                </Button>
              )}
            </div>
          )}
        </div>
      )}

      {/* Grouped Tree Accordion (Views) */}
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

      {/* Collapsed-by-default Hierarchies & Row Grouping Section (Below Views) */}
      {(hierarchyLadders.length > 0 || nonPivotSelectedDims.length >= 2 || options?.enableRowGroups) &&
        !searchTerm.trim() && (
          <Accordion collapsible defaultOpenItems={[]}>
            <AccordionItem value="hierarchies">
              <AccordionHeader size="small">
                Hierarchies & Row Grouping
                {hierarchyLadders.length > 0 ? ` (${hierarchyLadders.length})` : ""}
                {options?.enableRowGroups ? " • Active" : ""}
              </AccordionHeader>
              <AccordionPanel>
                <div style={{ display: "flex", flexDirection: "column", gap: "8px" }}>
                  {/* LookML Hierarchies Presets */}
                  {hierarchyLadders.length > 0 && onReorderFields && (
                    <div style={{ display: "flex", flexDirection: "column", gap: "4px" }}>
                      {hierarchyLadders.map((ladder) => {
                        const activeLevels = ladder.levels.filter((lv) =>
                          selectedFields.includes(lv)
                        );
                        const isHierarchyActive = activeLevels.length >= 2;
                        const levelLabels = ladder.levels.map((lv) => {
                          const d = dimensions.find((dim) => dim.name === lv);
                          return d?.label_short || d?.label || lv;
                        });

                        return (
                          <div
                            key={ladder.id}
                            style={{
                              display: "flex",
                              flexDirection: "column",
                              gap: "2px",
                              padding: "4px 6px",
                              borderRadius: "4px",
                              backgroundColor: tokens.colorNeutralBackground2,
                            }}
                          >
                            <div
                              style={{
                                display: "flex",
                                alignItems: "center",
                                justifyContent: "space-between",
                                gap: "6px",
                              }}
                            >
                              <Checkbox
                                checked={isHierarchyActive}
                                onChange={(_, data) => {
                                  if (data.checked) {
                                    const { fields, rowGroupFields } =
                                      applyHierarchyLadderToSelection(
                                        selectedFields,
                                        ladder,
                                        allFields,
                                        Math.min(3, ladder.levels.length)
                                      );
                                    onReorderFields(fields);
                                    if (onOptionsChange) {
                                      onOptionsChange({
                                        enableRowGroups: true,
                                        includeSubtotals: true,
                                        rowGroupFields,
                                        showGroupCounts: true,
                                        showGroupSubtotals: true,
                                        defaultDisplayLevel: "level_1",
                                      });
                                    }
                                  } else {
                                    const ladderSet = new Set(ladder.levels);
                                    const next = selectedFields.filter((f) => !ladderSet.has(f));
                                    onReorderFields(next);
                                    if (onOptionsChange) {
                                      onOptionsChange({
                                        enableRowGroups: false,
                                        rowGroupFields: [],
                                      });
                                    }
                                  }
                                }}
                                label={
                                  <Body2 style={{ fontSize: "12px" }}>{ladder.label}</Body2>
                                }
                              />
                              <Badge size="small" appearance="outline" color="subtle">
                                {ladder.source === "tag"
                                  ? "Tag"
                                  : ladder.source === "dimension_group"
                                  ? "Time"
                                  : "Drill"}
                              </Badge>
                            </div>
                            <Caption1
                              style={{
                                fontSize: "10px",
                                color: tokens.colorNeutralForeground3,
                                paddingLeft: "26px",
                              }}
                            >
                              {levelLabels.join(" ▸ ")}
                            </Caption1>
                            {isHierarchyActive && ladder.levels.length > 2 && (
                              <div
                                style={{
                                  display: "flex",
                                  gap: "4px",
                                  paddingLeft: "26px",
                                  flexWrap: "wrap",
                                  marginTop: "2px",
                                }}
                              >
                                {[2, 3, ladder.levels.length]
                                  .filter(
                                    (v, idx, arr) =>
                                      v <= ladder.levels.length && arr.indexOf(v) === idx
                                  )
                                  .map((depth) => (
                                    <Button
                                      key={depth}
                                      size="small"
                                      appearance={
                                        activeLevels.length === depth ? "primary" : "subtle"
                                      }
                                      style={{
                                        fontSize: "10px",
                                        padding: "1px 6px",
                                        minWidth: 0,
                                      }}
                                      onClick={() => {
                                        const { fields, rowGroupFields } =
                                          applyHierarchyLadderToSelection(
                                            selectedFields,
                                            ladder,
                                            allFields,
                                            depth
                                          );
                                        onReorderFields(fields);
                                        if (onOptionsChange) {
                                          onOptionsChange({
                                            enableRowGroups: true,
                                            includeSubtotals: true,
                                            rowGroupFields,
                                            showGroupCounts: true,
                                            showGroupSubtotals: true,
                                          });
                                        }
                                      }}
                                    >
                                      {depth === ladder.levels.length
                                        ? `All (${depth})`
                                        : `${depth} Levels`}
                                    </Button>
                                  ))}
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  )}

                  {/* Row Groups / Indented Hierarchy Tree Configuration */}
                  {(nonPivotSelectedDims.length >= 2 || options?.enableRowGroups) &&
                    onOptionsChange && (
                      <div
                        style={{
                          padding: "8px 10px",
                          borderRadius: "4px",
                          border: `1px solid ${tokens.colorNeutralStroke2}`,
                          backgroundColor: tokens.colorNeutralBackground2,
                          display: "flex",
                          flexDirection: "column",
                          gap: "8px",
                        }}
                      >
                        <div
                          style={{
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "space-between",
                          }}
                        >
                          <Body2 style={{ fontWeight: 600, fontSize: "12px" }}>
                            Enable Row Groups
                          </Body2>
                          <Switch
                            checked={Boolean(options?.enableRowGroups)}
                            onChange={(_, data) =>
                              onOptionsChange({
                                enableRowGroups: data.checked,
                                includeSubtotals: data.checked ? true : options?.includeSubtotals,
                                rowGroupFields: orderedRowGroupDims,
                                showGroupCounts: options?.showGroupCounts ?? true,
                                showGroupSubtotals: options?.showGroupSubtotals ?? true,
                                defaultDisplayLevel:
                                  options?.defaultDisplayLevel || "all_expanded",
                              })
                            }
                          />
                        </div>

                        {options?.enableRowGroups && (
                          <div
                            style={{
                              display: "flex",
                              flexDirection: "column",
                              gap: "8px",
                              paddingTop: "6px",
                              borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
                            }}
                          >
                            {/* Ordered Grouped Dimension Fields */}
                            <div style={{ display: "flex", flexDirection: "column", gap: "4px" }}>
                              <Caption1 style={{ fontWeight: 600, fontSize: "11px" }}>
                                Grouped Hierarchy Levels ({orderedRowGroupDims.length})
                              </Caption1>
                              {orderedRowGroupDims.map((fieldKey, gIdx) => {
                                const def = allFields.find((f) => f.name === fieldKey);
                                return (
                                  <div
                                    key={fieldKey}
                                    style={{
                                      display: "flex",
                                      alignItems: "center",
                                      justifyContent: "space-between",
                                      padding: "3px 8px",
                                      borderRadius: "4px",
                                      backgroundColor: tokens.colorNeutralBackground1,
                                      border: `1px solid ${tokens.colorNeutralStroke2}`,
                                      fontSize: "12px",
                                    }}
                                  >
                                    <span style={{ fontWeight: 500 }}>
                                      {gIdx + 1}. {def?.label || fieldKey}
                                    </span>
                                    <div
                                      style={{ display: "flex", alignItems: "center", gap: "4px" }}
                                    >
                                      {gIdx > 0 && (
                                        <span
                                          title="Move up hierarchy"
                                          style={{
                                            cursor: "pointer",
                                            padding: "0 4px",
                                            fontWeight: 700,
                                          }}
                                          onClick={() => {
                                            const nextGroup = [...orderedRowGroupDims];
                                            const [moved] = nextGroup.splice(gIdx, 1);
                                            nextGroup.splice(gIdx - 1, 0, moved);
                                            onOptionsChange({ rowGroupFields: nextGroup });
                                            if (onReorderFields) {
                                              const others = selectedFields.filter(
                                                (f) => !nextGroup.includes(f)
                                              );
                                              onReorderFields([...nextGroup, ...others]);
                                            }
                                          }}
                                        >
                                          ↑
                                        </span>
                                      )}
                                      {gIdx < orderedRowGroupDims.length - 1 && (
                                        <span
                                          title="Move down hierarchy"
                                          style={{
                                            cursor: "pointer",
                                            padding: "0 4px",
                                            fontWeight: 700,
                                          }}
                                          onClick={() => {
                                            const nextGroup = [...orderedRowGroupDims];
                                            const [moved] = nextGroup.splice(gIdx, 1);
                                            nextGroup.splice(gIdx + 1, 0, moved);
                                            onOptionsChange({ rowGroupFields: nextGroup });
                                            if (onReorderFields) {
                                              const others = selectedFields.filter(
                                                (f) => !nextGroup.includes(f)
                                              );
                                              onReorderFields([...nextGroup, ...others]);
                                            }
                                          }}
                                        >
                                          ↓
                                        </span>
                                      )}
                                      <DismissRegular
                                        style={{ cursor: "pointer", fontSize: "12px" }}
                                        onClick={() => onToggleField(fieldKey)}
                                      />
                                    </div>
                                  </div>
                                );
                              })}

                              {/* Quick button to add next LookML hierarchy level if available */}
                              {orderedRowGroupDims.length > 0 &&
                                onReorderFields &&
                                (() => {
                                  const deepestKey =
                                    orderedRowGroupDims[orderedRowGroupDims.length - 1];
                                  const hCtx = getFieldHierarchyContext(
                                    deepestKey,
                                    hierarchyLadders
                                  );
                                  if (
                                    !hCtx ||
                                    !hCtx.childField ||
                                    selectedFields.includes(hCtx.childField)
                                  ) {
                                    return null;
                                  }
                                  const childDef = allFields.find(
                                    (f) => f.name === hCtx.childField
                                  );
                                  return (
                                    <Button
                                      size="small"
                                      appearance="outline"
                                      onClick={() => {
                                        const { fields } = expandFieldInSelection(
                                          selectedFields,
                                          deepestKey,
                                          hierarchyLadders,
                                          "insert_child"
                                        );
                                        onReorderFields(fields);
                                        onOptionsChange({
                                          rowGroupFields: [
                                            ...orderedRowGroupDims,
                                            hCtx.childField!,
                                          ],
                                        });
                                      }}
                                    >
                                      + Add Level:{" "}
                                      {childDef?.label_short ||
                                        childDef?.label ||
                                        hCtx.childField}
                                    </Button>
                                  );
                                })()}
                            </div>

                            {/* Default Display Level */}
                            <div style={{ display: "flex", flexDirection: "column", gap: "2px" }}>
                              <Caption1 style={{ fontWeight: 600, fontSize: "11px" }}>
                                Default Display Level
                              </Caption1>
                              <Select
                                size="small"
                                value={options.defaultDisplayLevel || "all_expanded"}
                                onChange={(_, data) =>
                                  onOptionsChange({
                                    defaultDisplayLevel: data.value as RowGroupDisplayLevel,
                                  })
                                }
                              >
                                <option value="all_collapsed">All Collapsed</option>
                                <option value="level_1">Level 1</option>
                                <option value="level_2">Level 2</option>
                                <option value="level_3">Level 3</option>
                                <option value="all_expanded">All Expanded</option>
                              </Select>
                            </div>

                            {/* Group Column Header */}
                            <div style={{ display: "flex", flexDirection: "column", gap: "2px" }}>
                              <Caption1 style={{ fontWeight: 600, fontSize: "11px" }}>
                                Group Column Header
                              </Caption1>
                              <Input
                                size="small"
                                placeholder="Leave blank for default ('Group')"
                                value={options.groupColumnHeader || ""}
                                onChange={(_, data) =>
                                  onOptionsChange({ groupColumnHeader: data.value })
                                }
                              />
                            </div>

                            {/* Show Group Counts & Show Subtotals */}
                            <div style={{ display: "flex", gap: "12px", flexWrap: "wrap" }}>
                              <Checkbox
                                size="medium"
                                checked={options.showGroupCounts !== false}
                                onChange={(_, data) =>
                                  onOptionsChange({ showGroupCounts: Boolean(data.checked) })
                                }
                                label={
                                  <Caption1 style={{ fontWeight: 600 }}>Show Group Counts</Caption1>
                                }
                              />
                              <Checkbox
                                size="medium"
                                checked={options.showGroupSubtotals !== false}
                                onChange={(_, data) =>
                                  onOptionsChange({ showGroupSubtotals: Boolean(data.checked) })
                                }
                                label={
                                  <Caption1 style={{ fontWeight: 600 }}>Show Subtotals</Caption1>
                                }
                              />
                            </div>
                          </div>
                        )}
                      </div>
                    )}
                </div>
              </AccordionPanel>
            </AccordionItem>
          </Accordion>
        )}
    </div>
  );
};
