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
  AddRegular,
} from "@fluentui/react-icons";
import { FieldDefinition } from "../services/lookerClient";
import { WriteOptions, RowGroupDisplayLevel, PivotGroupMode } from "../services/excelWriter";
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
  filterSlot?: React.ReactNode;
}

export type ColumnBucketItem =
  | { type: "measures"; id: "__measures__"; measureKeys: string[] }
  | { type: "pivot"; id: string; fieldKey: string };

export interface DesignerBucketsPartition {
  rowDimensionKeys: string[];
  columnPivotKeys: string[];
  measureKeys: string[];
  columnItems: ColumnBucketItem[];
}

/**
 * Partitions selected LookML fields into SAC-style Designer buckets:
 * - `rowDimensionKeys`: non-pivoted selected Dimensions
 * - `columnPivotKeys`: pivoted Dimensions in pivot order
 * - `measureKeys`: selected Measures in selection order
 * - `columnItems`: ordered items inside the `Columns` bucket (draggable `Measures` token + pivot dimensions)
 */
export function partitionFieldsIntoBuckets(
  selectedFields: string[],
  pivots: string[],
  dimensions: FieldDefinition[],
  measures: FieldDefinition[],
  pivotGroupMode: PivotGroupMode = "measure_first"
): DesignerBucketsPartition {
  const dimSet = new Set(dimensions.map((d) => d.name));
  const measureSet = new Set(measures.map((m) => m.name));

  const rowDimensionKeys = selectedFields.filter(
    (fk) => dimSet.has(fk) && !pivots.includes(fk)
  );
  const columnPivotKeys = pivots.filter((pk) => selectedFields.includes(pk));
  for (const fk of selectedFields) {
    if (pivots.includes(fk) && !columnPivotKeys.includes(fk)) {
      columnPivotKeys.push(fk);
    }
  }
  const measureKeys = selectedFields.filter(
    (fk) => measureSet.has(fk) || (!dimSet.has(fk) && !pivots.includes(fk))
  );

  const pivotItems: ColumnBucketItem[] = columnPivotKeys.map((pk) => ({
    type: "pivot",
    id: pk,
    fieldKey: pk,
  }));
  const measuresItem: ColumnBucketItem | null =
    measureKeys.length > 0
      ? { type: "measures", id: "__measures__", measureKeys }
      : null;

  let columnItems: ColumnBucketItem[] = [];
  if (measuresItem) {
    if (pivotGroupMode === "cohort_first" && pivotItems.length > 0) {
      columnItems = [...pivotItems, measuresItem];
    } else {
      columnItems = [measuresItem, ...pivotItems];
    }
  } else {
    columnItems = [...pivotItems];
  }

  return {
    rowDimensionKeys,
    columnPivotKeys,
    measureKeys,
    columnItems,
  };
}

/**
 * Given a reordered list of `ColumnBucketItem`s, computes the updated `pivots` array
 * and `pivotGroupMode` (`"measure_first"` if `Measures` precedes the first pivot dimension,
 * or `"cohort_first"` if a pivot dimension precedes `Measures`).
 */
export function resolveColumnBucketReorder(items: ColumnBucketItem[]): {
  pivots: string[];
  pivotGroupMode: PivotGroupMode;
} {
  const nextPivots: string[] = [];
  let measuresIdx = -1;
  let firstPivotIdx = -1;

  items.forEach((item, idx) => {
    if (item.type === "measures") {
      measuresIdx = idx;
    } else {
      if (firstPivotIdx === -1) {
        firstPivotIdx = idx;
      }
      nextPivots.push(item.fieldKey);
    }
  });

  const pivotGroupMode: PivotGroupMode =
    measuresIdx !== -1 && firstPivotIdx !== -1 && firstPivotIdx < measuresIdx
      ? "cohort_first"
      : "measure_first";

  return { pivots: nextPivots, pivotGroupMode };
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
  filterSlot,
}) => {
  const styles = useStyles();
  const [searchTerm, setSearchTerm] = useState("");
  const [catalogCategoryFilter, setCatalogCategoryFilter] = useState<
    "all" | "dimension" | "measure"
  >("all");
  const [addingToPivots, setAddingToPivots] = useState(false);
  const [openSubgroups, setOpenSubgroups] = useState<Record<string, boolean>>({});
  const [draggedRowIdx, setDraggedRowIdx] = useState<number | null>(null);
  const [dragOverRowIdx, setDragOverRowIdx] = useState<number | null>(null);
  const [draggedColIdx, setDraggedColIdx] = useState<number | null>(null);
  const [dragOverColIdx, setDragOverColIdx] = useState<number | null>(null);

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

  const buckets = useMemo(
    () =>
      partitionFieldsIntoBuckets(
        selectedFields,
        pivots,
        dimensions,
        measures,
        options?.pivotGroupMode || "measure_first"
      ),
    [selectedFields, pivots, dimensions, measures, options?.pivotGroupMode]
  );

  const nonPivotSelectedDims = buckets.rowDimensionKeys;

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

  // Filter fields based on search + category filter (including LookML synonyms)
  const filteredFields = useMemo(() => {
    const base =
      catalogCategoryFilter === "dimension"
        ? dimensions
        : catalogCategoryFilter === "measure"
        ? measures
        : allFields;
    if (!searchTerm.trim()) return base;
    return base.filter((f) => matchesFieldSearch(f, searchTerm));
  }, [allFields, dimensions, measures, catalogCategoryFilter, searchTerm]);

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

  const handleCatalogFieldToggle = (f: FieldDefinition) => {
    const isAlreadySelected = selectedFields.includes(f.name);
    onToggleField(f.name);
    if (
      !isAlreadySelected &&
      addingToPivots &&
      f.category === "dimension" &&
      !pivots.includes(f.name)
    ) {
      if (onSetPivots) {
        onSetPivots([...pivots, f.name]);
      } else if (onTogglePivot) {
        onTogglePivot(f.name);
      }
    }
  };

  const handleReorderRowDimensions = (fromIdx: number, toIdx: number) => {
    if (!onReorderFields || fromIdx === toIdx) return;
    const nextRowDims = [...buckets.rowDimensionKeys];
    const [moved] = nextRowDims.splice(fromIdx, 1);
    nextRowDims.splice(toIdx, 0, moved);
    const nextAllFields = [
      ...nextRowDims,
      ...buckets.columnPivotKeys,
      ...buckets.measureKeys,
    ];
    onReorderFields(nextAllFields);
    if (options?.enableRowGroups && onOptionsChange) {
      onOptionsChange({ rowGroupFields: nextRowDims });
    }
  };

  const handleReorderColumnBucket = (fromIdx: number, toIdx: number) => {
    if (fromIdx === toIdx) return;
    const nextItems = [...buckets.columnItems];
    const [moved] = nextItems.splice(fromIdx, 1);
    nextItems.splice(toIdx, 0, moved);
    const { pivots: nextPivots, pivotGroupMode } = resolveColumnBucketReorder(nextItems);
    if (onSetPivots) {
      onSetPivots(nextPivots);
    }
    if (onOptionsChange) {
      onOptionsChange({ pivotGroupMode });
    }
  };

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
            onChange={() => handleCatalogFieldToggle(f)}
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
      {/* SAC-Style Designer Buckets: Rows & Columns */}
      <div
        style={{
          display: "flex",
          flexDirection: "column",
          gap: "8px",
        }}
      >
        {/* ROWS BUCKET */}
        <div
          style={{
            border: `1px solid ${tokens.colorNeutralStroke2}`,
            borderRadius: "6px",
            backgroundColor: tokens.colorNeutralBackground2,
            padding: "8px 10px",
            display: "flex",
            flexDirection: "column",
            gap: "6px",
          }}
        >
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
            }}
          >
            <Caption1 style={{ fontWeight: 700, fontSize: "12px" }}>
              Rows ({buckets.rowDimensionKeys.length})
            </Caption1>
            <Button
              size="small"
              appearance="subtle"
              icon={<AddRegular />}
              style={{ fontSize: "11px", padding: "2px 6px", minWidth: 0 }}
              onClick={() => {
                setCatalogCategoryFilter("dimension");
                setAddingToPivots(false);
              }}
            >
              Add Row Dim
            </Button>
          </div>

          {buckets.rowDimensionKeys.length === 0 ? (
            <Caption1
              style={{
                color: tokens.colorNeutralForeground4,
                fontStyle: "italic",
                fontSize: "11px",
              }}
            >
              Select dimensions below to group rows.
            </Caption1>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: "4px" }}>
              {buckets.rowDimensionKeys.map((fieldKey, idx) => {
                const def = allFields.find((f) => f.name === fieldKey);
                const sortEntry = sorts.find(
                  (s) => s === fieldKey || s.startsWith(`${fieldKey} `)
                );
                const isAsc = sortEntry === fieldKey || sortEntry === `${fieldKey} asc`;
                const isDesc = sortEntry === `${fieldKey} desc`;
                const isDragging = draggedRowIdx === idx;
                const isDropTarget = dragOverRowIdx === idx;

                return (
                  <div
                    key={fieldKey}
                    draggable={!!onReorderFields}
                    onDragStart={(e) => {
                      setDraggedRowIdx(idx);
                      e.dataTransfer.effectAllowed = "move";
                      e.dataTransfer.setData("text/plain", String(idx));
                    }}
                    onDragOver={(e) => {
                      e.preventDefault();
                      e.dataTransfer.dropEffect = "move";
                    }}
                    onDragEnter={() => setDragOverRowIdx(idx)}
                    onDragEnd={() => {
                      setDraggedRowIdx(null);
                      setDragOverRowIdx(null);
                    }}
                    onDrop={(e) => {
                      e.preventDefault();
                      if (draggedRowIdx !== null && draggedRowIdx !== idx) {
                        handleReorderRowDimensions(draggedRowIdx, idx);
                      }
                      setDraggedRowIdx(null);
                      setDragOverRowIdx(null);
                    }}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "space-between",
                      gap: "6px",
                      padding: "4px 8px",
                      backgroundColor: isDropTarget
                        ? tokens.colorBrandBackground2
                        : tokens.colorNeutralBackground1,
                      border: isDropTarget
                        ? `1px dashed ${tokens.colorBrandStroke1}`
                        : `1px solid ${tokens.colorNeutralStroke1}`,
                      borderRadius: "4px",
                      fontSize: "12px",
                      opacity: isDragging ? 0.4 : 1,
                      cursor: onReorderFields ? "grab" : "default",
                    }}
                  >
                    <div
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: "6px",
                        minWidth: 0,
                        flex: 1,
                      }}
                    >
                      {onReorderFields && (
                        <ReorderRegular
                          style={{
                            fontSize: "12px",
                            color: tokens.colorNeutralForeground4,
                            flexShrink: 0,
                          }}
                        />
                      )}
                      <span
                        style={{
                          fontWeight: 500,
                          overflow: "hidden",
                          textOverflow: "ellipsis",
                          whiteSpace: "nowrap",
                        }}
                        title={def?.label || fieldKey}
                      >
                        {def?.label || fieldKey}
                      </span>
                    </div>

                    <div
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: "4px",
                        flexShrink: 0,
                      }}
                    >
                      {onTogglePivot && (
                        <span
                          title="Move dimension to Columns (Pivot)"
                          style={{
                            cursor: "pointer",
                            fontSize: "10px",
                            fontWeight: 600,
                            padding: "1px 5px",
                            borderRadius: "6px",
                            backgroundColor: tokens.colorNeutralBackground3,
                            color: tokens.colorNeutralForeground2,
                          }}
                          onClick={() => onTogglePivot(fieldKey)}
                        >
                          → Col
                        </span>
                      )}
                      {onToggleSort && (
                        <span
                          title={
                            isAsc
                              ? "Sorted Ascending. Click for Descending."
                              : isDesc
                              ? "Sorted Descending. Click to clear sort."
                              : "Click to sort Ascending."
                          }
                          style={{
                            cursor: "pointer",
                            fontSize: "11px",
                            fontWeight: 700,
                            padding: "1px 4px",
                            borderRadius: "4px",
                            color:
                              isAsc || isDesc
                                ? tokens.colorBrandForeground1
                                : tokens.colorNeutralForeground4,
                            backgroundColor:
                              isAsc || isDesc
                                ? tokens.colorBrandBackground2
                                : "transparent",
                          }}
                          onClick={() => onToggleSort(fieldKey)}
                        >
                          {isAsc ? "↑" : isDesc ? "↓" : "↕"}
                        </span>
                      )}
                      <DismissRegular
                        style={{ cursor: "pointer", fontSize: "12px" }}
                        title="Remove from Rows"
                        onClick={() => onToggleField(fieldKey)}
                      />
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>

        {/* COLUMNS BUCKET (Measures + Pivoted Dimensions) */}
        <div
          style={{
            border: `1px solid ${tokens.colorNeutralStroke2}`,
            borderRadius: "6px",
            backgroundColor: tokens.colorNeutralBackground2,
            padding: "8px 10px",
            display: "flex",
            flexDirection: "column",
            gap: "6px",
          }}
        >
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              gap: "4px",
              flexWrap: "wrap",
            }}
          >
            <Caption1 style={{ fontWeight: 700, fontSize: "12px" }}>
              Columns ({buckets.columnItems.length})
            </Caption1>
            <div style={{ display: "flex", alignItems: "center", gap: "4px" }}>
              <Button
                size="small"
                appearance={addingToPivots ? "primary" : "subtle"}
                icon={<AddRegular />}
                style={{ fontSize: "11px", padding: "2px 6px", minWidth: 0 }}
                onClick={() => {
                  setCatalogCategoryFilter("dimension");
                  setAddingToPivots((prev) => !prev);
                }}
                title="When active, checking a dimension in the catalog adds it directly as a Column Pivot"
              >
                Pivot Dim
              </Button>
              <Button
                size="small"
                appearance="subtle"
                icon={<AddRegular />}
                style={{ fontSize: "11px", padding: "2px 6px", minWidth: 0 }}
                onClick={() => {
                  setCatalogCategoryFilter("measure");
                  setAddingToPivots(false);
                }}
              >
                Measure
              </Button>
            </div>
          </div>

          {buckets.columnItems.length === 0 ? (
            <Caption1
              style={{
                color: tokens.colorNeutralForeground4,
                fontStyle: "italic",
                fontSize: "11px",
              }}
            >
              Select measures or pivot dimensions to populate columns.
            </Caption1>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: "4px" }}>
              {buckets.columnItems.map((item, colIdx) => {
                const isDragging = draggedColIdx === colIdx;
                const isDropTarget = dragOverColIdx === colIdx;

                if (item.type === "measures") {
                  return (
                    <div
                      key={item.id}
                      draggable={buckets.columnItems.length > 1}
                      onDragStart={(e) => {
                        setDraggedColIdx(colIdx);
                        e.dataTransfer.effectAllowed = "move";
                        e.dataTransfer.setData("text/plain", String(colIdx));
                      }}
                      onDragOver={(e) => {
                        e.preventDefault();
                        e.dataTransfer.dropEffect = "move";
                      }}
                      onDragEnter={() => setDragOverColIdx(colIdx)}
                      onDragEnd={() => {
                        setDraggedColIdx(null);
                        setDragOverColIdx(null);
                      }}
                      onDrop={(e) => {
                        e.preventDefault();
                        if (draggedColIdx !== null && draggedColIdx !== colIdx) {
                          handleReorderColumnBucket(draggedColIdx, colIdx);
                        }
                        setDraggedColIdx(null);
                        setDragOverColIdx(null);
                      }}
                      style={{
                        display: "flex",
                        flexDirection: "column",
                        gap: "4px",
                        padding: "5px 8px",
                        backgroundColor: isDropTarget
                          ? tokens.colorBrandBackground2
                          : tokens.colorNeutralBackground1,
                        border: isDropTarget
                          ? `1px dashed ${tokens.colorBrandStroke1}`
                          : `1px solid ${tokens.colorNeutralStroke1}`,
                        borderRadius: "4px",
                        opacity: isDragging ? 0.4 : 1,
                        cursor: buckets.columnItems.length > 1 ? "grab" : "default",
                      }}
                      title={
                        buckets.columnItems.length > 1
                          ? "Drag above or below Pivot Dimensions to switch Measure-first vs Cohort-first column grouping"
                          : undefined
                      }
                    >
                      <div
                        style={{
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "space-between",
                          gap: "6px",
                        }}
                      >
                        <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                          {buckets.columnItems.length > 1 && (
                            <ReorderRegular
                              style={{
                                fontSize: "12px",
                                color: tokens.colorNeutralForeground4,
                              }}
                            />
                          )}
                          <Badge size="small" appearance="tint" color="warning">
                            M
                          </Badge>
                          <span style={{ fontWeight: 600, fontSize: "12px" }}>
                            Measures ({item.measureKeys.length})
                          </span>
                        </div>
                        <div style={{ display: "flex", alignItems: "center", gap: "4px" }}>
                          {colIdx > 0 && (
                            <span
                              title="Move Measures before Pivot Dimension (Group by Measure)"
                              style={{
                                cursor: "pointer",
                                fontSize: "11px",
                                fontWeight: 700,
                                padding: "0 4px",
                              }}
                              onClick={() => handleReorderColumnBucket(colIdx, colIdx - 1)}
                            >
                              ↑
                            </span>
                          )}
                          {colIdx < buckets.columnItems.length - 1 && (
                            <span
                              title="Move Measures after Pivot Dimension (Group by Cohort)"
                              style={{
                                cursor: "pointer",
                                fontSize: "11px",
                                fontWeight: 700,
                                padding: "0 4px",
                              }}
                              onClick={() => handleReorderColumnBucket(colIdx, colIdx + 1)}
                            >
                              ↓
                            </span>
                          )}
                          <DismissRegular
                            style={{ cursor: "pointer", fontSize: "12px" }}
                            title="Remove all measures"
                            onClick={() => {
                              item.measureKeys.forEach((mk) => onToggleField(mk));
                            }}
                          />
                        </div>
                      </div>

                      {/* Individual Measure Sub-Chips */}
                      <div
                        style={{
                          display: "flex",
                          flexWrap: "wrap",
                          gap: "4px",
                          paddingLeft: buckets.columnItems.length > 1 ? "18px" : "2px",
                        }}
                      >
                        {item.measureKeys.map((mKey) => {
                          const def = allFields.find((f) => f.name === mKey);
                          const sortEntry = sorts.find(
                            (s) => s === mKey || s.startsWith(`${mKey} `)
                          );
                          const isAsc = sortEntry === mKey || sortEntry === `${mKey} asc`;
                          const isDesc = sortEntry === `${mKey} desc`;
                          return (
                            <div
                              key={mKey}
                              className={styles.selectedChip}
                              style={{ padding: "1px 6px", fontSize: "11px" }}
                            >
                              <span>{def?.label_short || def?.label || mKey}</span>
                              {onToggleSort && (
                                <span
                                  title="Toggle sort direction"
                                  style={{
                                    cursor: "pointer",
                                    fontWeight: 700,
                                    color:
                                      isAsc || isDesc
                                        ? tokens.colorBrandForeground1
                                        : tokens.colorNeutralForeground4,
                                  }}
                                  onClick={() => onToggleSort(mKey)}
                                >
                                  {isAsc ? "↑" : isDesc ? "↓" : "↕"}
                                </span>
                              )}
                              <DismissRegular
                                style={{ cursor: "pointer", fontSize: "10px" }}
                                onClick={() => onToggleField(mKey)}
                              />
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  );
                }

                // Pivot Dimension item
                const fieldKey = item.fieldKey;
                const def = allFields.find((f) => f.name === fieldKey);
                const sortEntry = sorts.find(
                  (s) => s === fieldKey || s.startsWith(`${fieldKey} `)
                );
                const isAsc = sortEntry === fieldKey || sortEntry === `${fieldKey} asc`;
                const isDesc = sortEntry === `${fieldKey} desc`;

                return (
                  <div
                    key={item.id}
                    draggable={buckets.columnItems.length > 1}
                    onDragStart={(e) => {
                      setDraggedColIdx(colIdx);
                      e.dataTransfer.effectAllowed = "move";
                      e.dataTransfer.setData("text/plain", String(colIdx));
                    }}
                    onDragOver={(e) => {
                      e.preventDefault();
                      e.dataTransfer.dropEffect = "move";
                    }}
                    onDragEnter={() => setDragOverColIdx(colIdx)}
                    onDragEnd={() => {
                      setDraggedColIdx(null);
                      setDragOverColIdx(null);
                    }}
                    onDrop={(e) => {
                      e.preventDefault();
                      if (draggedColIdx !== null && draggedColIdx !== colIdx) {
                        handleReorderColumnBucket(draggedColIdx, colIdx);
                      }
                      setDraggedColIdx(null);
                      setDragOverColIdx(null);
                    }}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "space-between",
                      gap: "6px",
                      padding: "4px 8px",
                      backgroundColor: isDropTarget
                        ? tokens.colorBrandBackground2
                        : tokens.colorNeutralBackground1,
                      border: isDropTarget
                        ? `1px dashed ${tokens.colorBrandStroke1}`
                        : `1px solid ${tokens.colorNeutralStroke1}`,
                      borderRadius: "4px",
                      fontSize: "12px",
                      opacity: isDragging ? 0.4 : 1,
                      cursor: buckets.columnItems.length > 1 ? "grab" : "default",
                    }}
                  >
                    <div
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: "6px",
                        minWidth: 0,
                        flex: 1,
                      }}
                    >
                      {buckets.columnItems.length > 1 && (
                        <ReorderRegular
                          style={{
                            fontSize: "12px",
                            color: tokens.colorNeutralForeground4,
                            flexShrink: 0,
                          }}
                        />
                      )}
                      <Badge size="small" appearance="filled" color="brand">
                        Pivot
                      </Badge>
                      <span
                        style={{
                          fontWeight: 500,
                          overflow: "hidden",
                          textOverflow: "ellipsis",
                          whiteSpace: "nowrap",
                        }}
                        title={def?.label || fieldKey}
                      >
                        {def?.label || fieldKey}
                      </span>
                    </div>

                    <div
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: "4px",
                        flexShrink: 0,
                      }}
                    >
                      {colIdx > 0 && (
                        <span
                          title="Move up in Columns"
                          style={{
                            cursor: "pointer",
                            fontSize: "11px",
                            fontWeight: 700,
                            padding: "0 3px",
                          }}
                          onClick={() => handleReorderColumnBucket(colIdx, colIdx - 1)}
                        >
                          ↑
                        </span>
                      )}
                      {colIdx < buckets.columnItems.length - 1 && (
                        <span
                          title="Move down in Columns"
                          style={{
                            cursor: "pointer",
                            fontSize: "11px",
                            fontWeight: 700,
                            padding: "0 3px",
                          }}
                          onClick={() => handleReorderColumnBucket(colIdx, colIdx + 1)}
                        >
                          ↓
                        </span>
                      )}
                      {onTogglePivot && (
                        <span
                          title="Move back to Rows"
                          style={{
                            cursor: "pointer",
                            fontSize: "10px",
                            fontWeight: 600,
                            padding: "1px 5px",
                            borderRadius: "6px",
                            backgroundColor: tokens.colorNeutralBackground3,
                            color: tokens.colorNeutralForeground2,
                          }}
                          onClick={() => onTogglePivot(fieldKey)}
                        >
                          ← Row
                        </span>
                      )}
                      {onToggleSort && (
                        <span
                          title="Toggle sort direction"
                          style={{
                            cursor: "pointer",
                            fontSize: "11px",
                            fontWeight: 700,
                            padding: "1px 4px",
                            borderRadius: "4px",
                            color:
                              isAsc || isDesc
                                ? tokens.colorBrandForeground1
                                : tokens.colorNeutralForeground4,
                            backgroundColor:
                              isAsc || isDesc
                                ? tokens.colorBrandBackground2
                                : "transparent",
                          }}
                          onClick={() => onToggleSort(fieldKey)}
                        >
                          {isAsc ? "↑" : isDesc ? "↓" : "↕"}
                        </span>
                      )}
                      <DismissRegular
                        style={{ cursor: "pointer", fontSize: "12px" }}
                        title="Remove dimension"
                        onClick={() => onToggleField(fieldKey)}
                      />
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          {pivots.length > 0 && buckets.rowDimensionKeys.length > 0 && onSetPivots && (
            <div style={{ display: "flex", justifyContent: "flex-end", marginTop: "2px" }}>
              <Button
                size="small"
                appearance="outline"
                style={{ fontSize: "11px", padding: "2px 8px" }}
                title="Swap row dimensions and column pivot dimensions"
                onClick={() => {
                  onSetPivots([...buckets.rowDimensionKeys]);
                }}
              >
                ⇄ Swap Rows &amp; Columns
              </Button>
            </div>
          )}
        </div>

        {/* FILTER BUCKET SLOT */}
        {filterSlot && (
          <div
            style={{
              border: `1px solid ${tokens.colorNeutralStroke2}`,
              borderRadius: "6px",
              backgroundColor: tokens.colorNeutralBackground2,
              padding: "8px 10px",
            }}
          >
            {filterSlot}
          </div>
        )}
      </div>

      {/* Field Catalog Search & Filter Pills */}
      <div style={{ display: "flex", flexDirection: "column", gap: "6px", marginTop: "2px" }}>
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
          }}
        >
          <Caption1 style={{ fontWeight: 700, fontSize: "12px" }}>
            Field Catalog {addingToPivots ? "(Adding to Column Pivots)" : ""}
          </Caption1>
          {selectedFields.length > 0 && (
            <Button size="small" appearance="subtle" onClick={onClearSelected}>
              Clear All ({selectedFields.length})
            </Button>
          )}
        </div>

        <Input
          size="small"
          contentBefore={<SearchRegular />}
          placeholder="Search dimensions & measures..."
          value={searchTerm}
          onChange={(_, data) => setSearchTerm(data.value)}
        />

        <div style={{ display: "flex", gap: "4px", flexWrap: "wrap" }}>
          <Button
            size="small"
            appearance={catalogCategoryFilter === "all" ? "primary" : "outline"}
            style={{ fontSize: "11px", padding: "2px 8px", minWidth: 0 }}
            onClick={() => {
              setCatalogCategoryFilter("all");
              setAddingToPivots(false);
            }}
          >
            All ({allFields.length})
          </Button>
          <Button
            size="small"
            appearance={catalogCategoryFilter === "dimension" ? "primary" : "outline"}
            style={{ fontSize: "11px", padding: "2px 8px", minWidth: 0 }}
            onClick={() => setCatalogCategoryFilter("dimension")}
          >
            Dimensions ({dimensions.length})
          </Button>
          <Button
            size="small"
            appearance={catalogCategoryFilter === "measure" ? "primary" : "outline"}
            style={{ fontSize: "11px", padding: "2px 8px", minWidth: 0 }}
            onClick={() => {
              setCatalogCategoryFilter("measure");
              setAddingToPivots(false);
            }}
          >
            Measures ({measures.length})
          </Button>
        </div>
      </div>

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
