import { PivotDescriptor } from "./queryTaskRunner";

export type PivotGroupMode = "measure_first" | "cohort_first";

export interface ColumnDefinition {
  fieldKey: string;
  label: string;
  type?: string;
  excelFormat?: string;
  category?: "dimension" | "measure" | "parameter";
  isRowTotal?: boolean;
  measureKey?: string;
  measureLabel?: string;
  pivotKey?: string;
  pivotValues?: string[];
  isColumnGroupSummary?: boolean;
  columnGroupNodeId?: string;
  columnGroupHeaderRowOffset?: number;
}

export interface BoundingBox {
  top: number;
  bottom: number;
  left: number;
  right: number;
}

export type RowGroupDisplayLevel =
  | "all_expanded"
  | "all_collapsed"
  | "level_1"
  | "level_2"
  | "level_3";

export interface HierarchyTreeNodeMeta {
  rowOffset: number;
  level: number;
  label: string;
  childCount: number;
  descendantStartOffset: number;
  descendantEndOffset: number;
  directChildOffsets: number[];
  isExpanded: boolean;
}

export interface ColumnTreeNodeMeta {
  id: string;
  label: string;
  level: number;
  headerRowOffset: number;
  summaryColOffset: number;
  summaryColCount: number;
  childColStartOffset: number;
  childColEndOffset: number;
  isExpanded: boolean;
}

export interface WriteOptions {
  useExcelTable?: boolean;
  tableStyle?: string;
  preserveUserFormatting?: boolean;
  autoExpandFormulas?: boolean;
  freezeHeader?: boolean;
  nullDisplay?: string;
  destination?: "active" | "new";
  sheetName?: string;
  targetSheetId?: string;
  pivots?: string[];
  pivotsMeta?: PivotDescriptor[] | null;
  pivotGroupMode?: PivotGroupMode;
  useMultiRowCrosstabHeader?: boolean;
  prevRowCount?: number;
  prevColCount?: number;
  includeTotals?: boolean;
  includeRowTotals?: boolean;
  includeSubtotals?: boolean;
  totalsData?: Record<string, any> | null;
  subtotalsData?: Record<string, Record<string, any>[]> | null;
  anchorCell?: string;
  tableId?: string;
  siblingTableRanges?: BoundingBox[];
  enableRowGroups?: boolean;
  rowGroupFields?: string[];
  defaultDisplayLevel?: RowGroupDisplayLevel;
  groupColumnHeader?: string;
  showGroupCounts?: boolean;
  showGroupSubtotals?: boolean;
  treeNodes?: HierarchyTreeNodeMeta[];
  rowLevels?: number[];
  showSourceBanner?: boolean;
  sourceBannerTitle?: string;
  pivotDimensionLabels?: string[];
  enableColumnGroups?: boolean;
  defaultColumnDisplayLevel?: RowGroupDisplayLevel;
  columnTreeNodes?: ColumnTreeNodeMeta[];
  measureFormatOverrides?: Record<string, string>;
}

export function resolveEffectiveColumnNumberFormat(
  col: ColumnDefinition,
  options?: WriteOptions
): string | undefined {
  const baseKey = col.measureKey || col.fieldKey;
  const override = options?.measureFormatOverrides?.[baseKey];
  if (override && override !== "default" && override.trim() !== "") {
    return override.trim();
  }
  return col.excelFormat;
}

export function getEffectiveHeaderRowCount(options?: WriteOptions): {
  bannerRowCount: number;
  tableHeaderRowCount: number;
  totalHeaderRowCount: number;
} {
  const bannerRowCount = options?.showSourceBanner ? 1 : 0;
  const useMultiRowHeader = Boolean(
    (options?.useMultiRowCrosstabHeader ?? true) &&
      options?.pivots &&
      options.pivots.length > 0
  );
  const tableHeaderRowCount = useMultiRowHeader ? (options?.pivots?.length || 1) + 1 : 1;
  return {
    bannerRowCount,
    tableHeaderRowCount,
    totalHeaderRowCount: bannerRowCount + tableHeaderRowCount,
  };
}

/**
 * Converts an Excel A1-style cell reference (e.g. "A1", "B4", "$H$12", "Sheet1!C5")
 * into 0-based { startRow, startCol } coordinates.
 */
export function parseCellAddress(address?: string | null): { startRow: number; startCol: number } {
  if (!address || typeof address !== "string") {
    return { startRow: 0, startCol: 0 };
  }
  const bangIdx = address.lastIndexOf("!");
  const clean = (bangIdx >= 0 ? address.slice(bangIdx + 1) : address)
    .split(":")[0]
    .replace(/\$/g, "")
    .trim()
    .toUpperCase();

  const match = /^([A-Z]+)(\d+)$/.exec(clean);
  if (!match) {
    return { startRow: 0, startCol: 0 };
  }

  const colLetters = match[1];
  const rowNum = parseInt(match[2], 10);
  if (!Number.isFinite(rowNum) || rowNum < 1) {
    return { startRow: 0, startCol: 0 };
  }

  let colIdx = 0;
  for (let i = 0; i < colLetters.length; i++) {
    colIdx = colIdx * 26 + (colLetters.charCodeAt(i) - 64);
  }

  return {
    startRow: rowNum - 1,
    startCol: Math.max(0, colIdx - 1),
  };
}

/**
 * Converts 0-based (rowIndex, colIndex) into an Excel A1 cell reference (e.g., (3, 1) -> "B4").
 */
export function formatCellAddress(rowIndex: number, colIndex: number): string {
  let c = Math.max(0, Math.floor(colIndex)) + 1;
  let letters = "";
  while (c > 0) {
    const rem = (c - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    c = Math.floor((c - 1) / 26);
  }
  const r = Math.max(0, Math.floor(rowIndex)) + 1;
  return `${letters}${r}`;
}

export function calculateBatchSize(colCount: number, targetCellBudget: number = 35000): number {
  return Math.max(500, Math.floor(targetCellBudget / Math.max(1, colCount)));
}

export interface CrosstabHeaderMerge {
  rowOffset: number;
  colOffset: number;
  rowCount: number;
  colCount: number;
}

export interface CrosstabHeaderMatrix {
  headerRows: string[][];
  merges: CrosstabHeaderMerge[];
}

const CELL_METADATA_KEYS = new Set([
  "value",
  "rendered",
  "html",
  "has_custom_html",
  "links",
  "filterable_value",
  "drill_fields",
]);

function isNestedPivotObject(val: any): boolean {
  if (val && typeof val === "object" && !Array.isArray(val)) {
    const entries = Object.entries(val).filter(([k]) => k !== "$$$_row_total_$$$");
    if (entries.length === 0) return false;
    // Never treat a Looker json_detail leaf cell object as a pivot dimension map
    if (entries.some(([k]) => CELL_METADATA_KEYS.has(k) || k === "__pivot_key__")) {
      return false;
    }
    const firstInner = entries[0][1];
    if (!firstInner || typeof firstInner !== "object" || Array.isArray(firstInner)) {
      return false;
    }
    const innerKeys = Object.keys(firstInner);
    if (innerKeys.some((ik) => CELL_METADATA_KEYS.has(ik))) {
      return false;
    }
    return true;
  }
  return false;
}

/**
 * Resolves clean, human-readable pivot dimension values from a Looker `json_detail` pivot key
 * (e.g. `"1 - Processing|FIELD|0|FIELD|Processing"` -> `["1 - Processing", "Processing"]`).
 * Uses `pivotsMeta` (`labels` / `data`) when available, and falls back to stripping Looker's
 * internal `|FIELD|` delimiters and injected `case:` / `order_by_field` sort indices.
 */
export function resolveCleanPivotValues(
  rawPivotKey: string,
  pivots?: string[],
  pivotsMeta?: PivotDescriptor[] | null
): string[] {
  if (pivotsMeta && pivotsMeta.length > 0) {
    const pm = pivotsMeta.find((p) => p.key === rawPivotKey);
    if (pm && (pm.labels || pm.data) && pivots && pivots.length > 0) {
      return pivots.map((pf) => {
        const v = pm.labels?.[pf] ?? pm.data?.[pf];
        return v === null || v === undefined || v === "null" ? "None" : String(v);
      });
    }
  }

  if (rawPivotKey.includes("|FIELD|")) {
    const parts = rawPivotKey
      .split("|FIELD|")
      .map((s) => (s === "null" ? "None" : s.trim()));
    const expectedCount = pivots && pivots.length > 0 ? pivots.length : parts.length;
    if (parts.length > expectedCount) {
      const filtered = parts.filter(
        (token, idx) => !(idx > 0 && idx < parts.length - 1 && /^\d+$/.test(token))
      );
      if (filtered.length >= expectedCount) {
        return filtered.slice(0, expectedCount);
      }
    }
    return parts;
  }

  if (pivots && pivots.length > 1 && rawPivotKey.includes("|")) {
    return rawPivotKey.split("|").map((s) => (s === "null" ? "None" : s.trim()));
  }

  return [rawPivotKey === "null" ? "None" : rawPivotKey];
}

/**
 * Extracts all flat pivot cohort keys, their per-dimension pivot values, and scalar cell values
 * from a measure cell object (handling both 1-dim / json_detail flat maps and multi-level nested JSON maps).
 */
function extractMeasurePivotLeaves(
  measureObj: Record<string, any>,
  pivotField: string,
  pivots?: string[],
  pivotsMeta?: PivotDescriptor[] | null
): Array<{ pivotKey: string; pivotValues: string[]; value: any; path: string[] }> {
  const inner = measureObj[pivotField];
  if (!inner || typeof inner !== "object") {
    return [];
  }

  const leaves: Array<{ pivotKey: string; pivotValues: string[]; value: any; path: string[] }> = [];

  const walk = (node: Record<string, any>, currentVals: string[], currentPath: string[]) => {
    for (const [k, v] of Object.entries(node)) {
      if (
        v &&
        typeof v === "object" &&
        !Array.isArray(v) &&
        Object.keys(v).length > 0 &&
        pivots &&
        pivots.includes(k)
      ) {
        walk(v, currentVals, [...currentPath, k]);
      } else if (v && typeof v === "object" && !Array.isArray(v) && Object.keys(v).length > 0) {
        walk(v, [...currentVals, k === "null" ? "None" : k], [...currentPath, k]);
      } else {
        const rawVals = [...currentVals, k === "null" ? "None" : k];
        const pivotKey = currentVals.length === 0 ? k : rawVals.join("|");
        const resolvedValues =
          rawVals.length === 1
            ? resolveCleanPivotValues(pivotKey, pivots, pivotsMeta)
            : rawVals;

        leaves.push({
          pivotKey,
          pivotValues: resolvedValues,
          value: v,
          path: [...currentPath, k],
        });
      }
    }
  };

  walk(inner, [], [pivotField]);
  return leaves;
}

function readNestedPath(obj: any, path: string[]): any {
  let curr = obj;
  for (const key of path) {
    if (curr === null || curr === undefined || typeof curr !== "object") {
      return undefined;
    }
    curr = curr[key];
  }
  return curr;
}

const ADDITIVE_MEASURE_TYPES = new Set([
  "sum",
  "sum_distinct",
  "count",
  "running_total",
]);

export function isAdditiveMeasure(col: ColumnDefinition): boolean {
  const t = (col.type || "").toLowerCase();
  return ADDITIVE_MEASURE_TYPES.has(t);
}

/**
 * Detects nested pivot structures in Looker JSON / json_detail query results and expands them
 * into tabular Crosstab columns (including multi-dimension pivots, cohort vs. measure grouping,
 * collapsible parent cohort summary columns, and Looker-computed Row Totals) with appropriate
 * labels and number formatting.
 */
export function expandPivotedColumns(
  columns: ColumnDefinition[],
  dataRows: Record<string, any>[],
  pivots?: string[],
  includeRowTotals?: boolean,
  pivotsMeta?: PivotDescriptor[] | null,
  pivotGroupMode: PivotGroupMode = "measure_first",
  enableColumnGroups: boolean = false,
  defaultColumnDisplayLevel: RowGroupDisplayLevel = "all_expanded"
): {
  expandedColumns: ColumnDefinition[];
  getValue: (row: Record<string, any>, col: ColumnDefinition) => any;
  columnTreeNodes: ColumnTreeNodeMeta[];
} {
  if (dataRows.length === 0) {
    return {
      expandedColumns: columns,
      getValue: (row, col) => row[col.fieldKey],
      columnTreeNodes: [],
    };
  }

  const pivotedColSample = new Map<string, Record<string, any>>();
  const discoveredPivotDims = new Set<string>(pivots || []);

  for (const col of columns) {
    for (const row of dataRows) {
      const val = row[col.fieldKey];
      if (isNestedPivotObject(val)) {
        pivotedColSample.set(col.fieldKey, val);
        Object.keys(val)
          .filter((pf) => pf !== "$$$_row_total_$$$")
          .forEach((pf) => discoveredPivotDims.add(pf));
        break;
      }
    }
  }

  if (pivotedColSample.size === 0) {
    return {
      expandedColumns: columns,
      getValue: (row, col) => {
        const val = row[col.fieldKey];
        if (typeof val === "object" && val !== null) {
          return JSON.stringify(val);
        }
        return val;
      },
      columnTreeNodes: [],
    };
  }

  const pivotAccessorMap = new Map<
    string,
    {
      measureKey: string;
      path: string[];
      isRowTotal?: boolean;
      summaryChildPaths?: string[][];
      isAdditive?: boolean;
    }
  >();

  interface CohortInfo {
    pivotKey: string;
    pivotValues: string[];
    path: string[];
  }

  const nonPivotedColumns: ColumnDefinition[] = [];
  const pivotedMeasureCols: ColumnDefinition[] = [];
  const cohortsByMeasure = new Map<string, Map<string, CohortInfo>>();
  const globalCohortOrder: CohortInfo[] = [];
  const seenGlobalCohortKeys = new Set<string>();
  let anyRowTotal = Boolean(includeRowTotals);

  columns.forEach((col) => {
    if (discoveredPivotDims.has(col.fieldKey)) {
      return;
    }

    const val = pivotedColSample.get(col.fieldKey);
    if (!val) {
      nonPivotedColumns.push(col);
      return;
    }

    pivotedMeasureCols.push(col);
    const pivotFields = Object.keys(val).filter((pf) => pf !== "$$$_row_total_$$$");
    const pivotField = pivotFields[0];
    const measureCohorts = new Map<string, CohortInfo>();

    // Seed cohort ordering from `pivotsMeta` if present
    if (pivotsMeta && pivotsMeta.length > 0) {
      for (const pm of pivotsMeta) {
        if (pm.is_total || pm.key === "$$$_row_total_$$$") continue;
        const resolvedValues = resolveCleanPivotValues(pm.key, pivots, pivotsMeta);
        const info: CohortInfo = {
          pivotKey: pm.key,
          pivotValues: resolvedValues,
          path: [pivotField, pm.key],
        };
        measureCohorts.set(pm.key, info);
        if (!seenGlobalCohortKeys.has(pm.key)) {
          seenGlobalCohortKeys.add(pm.key);
          globalCohortOrder.push(info);
        }
      }
    }

    dataRows.forEach((row) => {
      const measureObj = row[col.fieldKey];
      if (measureObj && typeof measureObj === "object") {
        if ("$$$_row_total_$$$" in measureObj) {
          anyRowTotal = true;
        }
        const leaves = extractMeasurePivotLeaves(measureObj, pivotField, pivots, pivotsMeta);
        for (const leaf of leaves) {
          if (leaf.pivotKey === "null" && leaf.value === 0) {
            continue;
          }
          if (!measureCohorts.has(leaf.pivotKey)) {
            const info: CohortInfo = {
              pivotKey: leaf.pivotKey,
              pivotValues: leaf.pivotValues,
              path: leaf.path,
            };
            measureCohorts.set(leaf.pivotKey, info);
            if (!seenGlobalCohortKeys.has(leaf.pivotKey)) {
              seenGlobalCohortKeys.add(leaf.pivotKey);
              globalCohortOrder.push(info);
            }
          }
        }
      }
    });

    cohortsByMeasure.set(col.fieldKey, measureCohorts);
  });

  const expandedColumns: ColumnDefinition[] = [...nonPivotedColumns];
  const columnTreeNodes: ColumnTreeNodeMeta[] = [];

  const pivotDepth = Math.max(
    pivots?.length || 0,
    globalCohortOrder.reduce((maxD, c) => Math.max(maxD, c.pivotValues.length), 0)
  );

  const createPivotCol = (col: ColumnDefinition, cohort: CohortInfo): ColumnDefinition => {
    const key = `__pivot__:${col.fieldKey}:::${cohort.pivotKey}`;
    const displayCohort = cohort.pivotValues.join(" · ");
    pivotAccessorMap.set(key, {
      measureKey: col.fieldKey,
      path: cohort.path,
    });
    return {
      fieldKey: key,
      label: `${col.label} (${displayCohort})`,
      type: col.type,
      excelFormat: col.excelFormat,
      category: col.category || "measure",
      measureKey: col.fieldKey,
      measureLabel: col.label,
      pivotKey: cohort.pivotKey,
      pivotValues: cohort.pivotValues,
    };
  };

  const createGroupSummaryCol = (
    col: ColumnDefinition,
    nodeId: string,
    headerRowOffset: number,
    groupLabel: string,
    isExpanded: boolean,
    parentValues: string[],
    childCohorts: CohortInfo[]
  ): ColumnDefinition => {
    const key = `__pivot_col_summary__:${col.fieldKey}:::${nodeId}`;
    const prefix = isExpanded ? "▾" : "▸";
    const summaryPivotValues: string[] = [];
    for (let r = 0; r < pivotDepth; r++) {
      if (r < headerRowOffset) {
        summaryPivotValues.push(parentValues[r] ?? "");
      } else if (r === headerRowOffset) {
        summaryPivotValues.push(`${prefix} ${groupLabel}`);
      } else {
        summaryPivotValues.push("");
      }
    }

    const mCohorts = cohortsByMeasure.get(col.fieldKey);
    const summaryChildPaths = childCohorts.map(
      (c) => (mCohorts?.get(c.pivotKey) || c).path
    );

    pivotAccessorMap.set(key, {
      measureKey: col.fieldKey,
      path: [],
      summaryChildPaths,
      isAdditive: isAdditiveMeasure(col),
    });

    const displayPrefix = [...parentValues.slice(0, headerRowOffset), `${prefix} ${groupLabel}`]
      .filter(Boolean)
      .join(" · ");

    return {
      fieldKey: key,
      label: `${col.label} (${displayPrefix})`,
      type: col.type,
      excelFormat: col.excelFormat,
      category: col.category || "measure",
      measureKey: col.fieldKey,
      measureLabel: col.label,
      pivotKey: nodeId,
      pivotValues: summaryPivotValues,
      isColumnGroupSummary: true,
      columnGroupNodeId: nodeId,
      columnGroupHeaderRowOffset: headerRowOffset,
    };
  };

  const createRowTotalCol = (col: ColumnDefinition): ColumnDefinition => {
    const rtKey = `__pivot_row_total__:${col.fieldKey}`;
    pivotAccessorMap.set(rtKey, {
      measureKey: col.fieldKey,
      path: ["$$$_row_total_$$$"],
      isRowTotal: true,
    });
    return {
      fieldKey: rtKey,
      label: `${col.label} (Row Total)`,
      type: col.type,
      excelFormat: col.excelFormat,
      category: col.category || "measure",
      isRowTotal: true,
      measureKey: col.fieldKey,
      measureLabel: col.label,
    };
  };

  const getMaxColVisibleLevel = (dl: RowGroupDisplayLevel): number => {
    switch (dl) {
      case "all_collapsed":
        return 0;
      case "level_1":
        return 1;
      case "level_2":
        return 2;
      case "level_3":
        return 3;
      case "all_expanded":
      default:
        return Number.POSITIVE_INFINITY;
    }
  };
  const maxVisibleColLevel = getMaxColVisibleLevel(defaultColumnDisplayLevel);

  const emitCohortHierarchy = (
    cohorts: CohortInfo[],
    headerRowOffset: number,
    treeDepthLevel: number,
    scopeId: string,
    measuresToEmit: ColumnDefinition[]
  ) => {
    if (headerRowOffset >= pivotDepth - 1) {
      for (const cohort of cohorts) {
        for (const mCol of measuresToEmit) {
          const mCohorts = cohortsByMeasure.get(mCol.fieldKey);
          const match = mCohorts?.get(cohort.pivotKey) || cohort;
          expandedColumns.push(createPivotCol(mCol, match));
        }
      }
      return;
    }

    // Group contiguous cohorts sharing the same value at `headerRowOffset`
    const groups: { label: string; items: CohortInfo[] }[] = [];
    for (const c of cohorts) {
      const val = c.pivotValues[headerRowOffset] ?? "";
      const last = groups[groups.length - 1];
      if (last && last.label === val) {
        last.items.push(c);
      } else {
        groups.push({ label: val, items: [c] });
      }
    }

    for (let gIdx = 0; gIdx < groups.length; gIdx++) {
      const grp = groups[gIdx];
      const distinctNextLevel = new Set(
        grp.items.map((item) => item.pivotValues[headerRowOffset + 1] ?? "")
      ).size;

      // Always emit a collapsible summary node at the immediate parent of leaf level (`pivotDepth - 2`),
      // or at higher levels when there are >= 2 distinct child groups underneath
      const shouldEmitSummaryNode =
        headerRowOffset === pivotDepth - 2 || distinctNextLevel >= 2;

      if (shouldEmitSummaryNode) {
        const nodeId = `colgrp:${scopeId}:${headerRowOffset}:${gIdx}:${grp.label}`;
        const isExpanded = treeDepthLevel < maxVisibleColLevel;
        const summaryColOffset = expandedColumns.length;
        const parentVals = grp.items[0]?.pivotValues.slice(0, headerRowOffset) || [];

        for (const mCol of measuresToEmit) {
          expandedColumns.push(
            createGroupSummaryCol(
              mCol,
              nodeId,
              headerRowOffset,
              grp.label,
              isExpanded,
              parentVals,
              grp.items
            )
          );
        }

        const summaryColCount = measuresToEmit.length;
        const childColStartOffset = expandedColumns.length;

        emitCohortHierarchy(
          grp.items,
          headerRowOffset + 1,
          treeDepthLevel + 1,
          `${scopeId}_${gIdx}`,
          measuresToEmit
        );

        const childColEndOffset = expandedColumns.length - 1;
        if (childColEndOffset >= childColStartOffset) {
          columnTreeNodes.push({
            id: nodeId,
            label: grp.label,
            level: treeDepthLevel,
            headerRowOffset,
            summaryColOffset,
            summaryColCount,
            childColStartOffset,
            childColEndOffset,
            isExpanded,
          });
        }
      } else {
        emitCohortHierarchy(
          grp.items,
          headerRowOffset + 1,
          treeDepthLevel,
          `${scopeId}_${gIdx}`,
          measuresToEmit
        );
      }
    }
  };

  if (pivotGroupMode === "cohort_first") {
    if (enableColumnGroups && pivotDepth >= 2) {
      emitCohortHierarchy(globalCohortOrder, 0, 0, "cohort", pivotedMeasureCols);
    } else {
      for (const cohort of globalCohortOrder) {
        for (const mCol of pivotedMeasureCols) {
          const mCohorts = cohortsByMeasure.get(mCol.fieldKey);
          const match = mCohorts?.get(cohort.pivotKey) || cohort;
          expandedColumns.push(createPivotCol(mCol, match));
        }
      }
    }
    if (anyRowTotal) {
      for (const mCol of pivotedMeasureCols) {
        expandedColumns.push(createRowTotalCol(mCol));
      }
    }
  } else {
    // Default "measure_first" grouping
    for (const mCol of pivotedMeasureCols) {
      const mCohorts = cohortsByMeasure.get(mCol.fieldKey);
      const cohortList = mCohorts ? Array.from(mCohorts.values()) : [];
      if (enableColumnGroups && pivotDepth >= 2 && cohortList.length > 0) {
        emitCohortHierarchy(cohortList, 0, 0, mCol.fieldKey, [mCol]);
      } else {
        for (const cohort of cohortList) {
          expandedColumns.push(createPivotCol(mCol, cohort));
        }
      }
      if (anyRowTotal) {
        expandedColumns.push(createRowTotalCol(mCol));
      }
    }
  }

  const unwrapNumericCell = (raw: any): number | null => {
    if (raw === null || raw === undefined) return null;
    if (typeof raw === "number") return Number.isFinite(raw) ? raw : null;
    if (typeof raw === "object" && "value" in raw) {
      const inner = raw.value;
      if (typeof inner === "number" && Number.isFinite(inner)) return inner;
      if (typeof inner === "string" && inner.trim() !== "") {
        const parsed = Number(inner);
        return Number.isFinite(parsed) ? parsed : null;
      }
      return null;
    }
    if (typeof raw === "string" && raw.trim() !== "") {
      const parsed = Number(raw);
      return Number.isFinite(parsed) ? parsed : null;
    }
    return null;
  };

  return {
    expandedColumns,
    columnTreeNodes,
    getValue: (row, col) => {
      const acc = pivotAccessorMap.get(col.fieldKey);
      if (acc) {
        if (acc.isRowTotal) {
          return row[acc.measureKey]?.["$$$_row_total_$$$"];
        }
        if (acc.summaryChildPaths && acc.summaryChildPaths.length > 0) {
          if (!acc.isAdditive) return null;
          const measureObj = row[acc.measureKey];
          if (!measureObj || typeof measureObj !== "object") return null;
          const nums: number[] = [];
          for (const childPath of acc.summaryChildPaths) {
            const rawLeaf = readNestedPath(measureObj, childPath);
            const num = unwrapNumericCell(rawLeaf);
            if (num !== null) {
              nums.push(num);
            }
          }
          if (nums.length === 0) return null;
          return nums.reduce((accSum, v) => accSum + v, 0);
        }
        return readNestedPath(row[acc.measureKey], acc.path);
      }
      const val = row[col.fieldKey];
      if (typeof val === "object" && val !== null) {
        return JSON.stringify(val);
      }
      return val;
    },
  };
}

/**
 * Builds a multi-row Crosstab header matrix (1 row per pivoted dimension + 1 bottom row for Measure/Dimension labels)
 * along with horizontal merge descriptors for contiguous identical pivot cohort headers.
 */
export function buildCrosstabHeaderMatrix(
  expandedColumns: ColumnDefinition[],
  pivots?: string[],
  trailingHeaders: string[] = [],
  pivotDimensionLabels?: string[]
): CrosstabHeaderMatrix {
  const pivotDepth = Math.max(
    pivots?.length || 0,
    expandedColumns.reduce((maxD, c) => Math.max(maxD, c.pivotValues?.length || 0), 0)
  );

  if (pivotDepth === 0) {
    return {
      headerRows: [[...expandedColumns.map((c) => c.label || c.fieldKey), ...trailingHeaders]],
      merges: [],
    };
  }

  const firstPivotColIdx = expandedColumns.findIndex(
    (c) => Boolean(c.pivotValues && c.pivotValues.length > 0) || Boolean(c.isRowTotal)
  );
  const nonPivotColCount = firstPivotColIdx >= 0 ? firstPivotColIdx : 0;

  const totalCols = expandedColumns.length + trailingHeaders.length;
  const headerRows: string[][] = [];

  for (let r = 0; r < pivotDepth; r++) {
    const rowCells: string[] = [];
    for (let c = 0; c < expandedColumns.length; c++) {
      const col = expandedColumns[c];
      if (col.isRowTotal) {
        rowCells.push(r === 0 ? "Row Total" : "");
      } else if (col.pivotValues && col.pivotValues.length > 0) {
        rowCells.push(col.pivotValues[r] ?? "");
      } else if (
        c === nonPivotColCount - 1 &&
        pivotDimensionLabels &&
        pivotDimensionLabels[r]
      ) {
        rowCells.push(pivotDimensionLabels[r]);
      } else {
        rowCells.push("");
      }
    }
    for (let t = 0; t < trailingHeaders.length; t++) {
      rowCells.push("");
    }
    headerRows.push(rowCells);
  }

  // Bottom row: Measure labels for pivoted/row-total columns, and Dimension labels for row dimensions
  const bottomRow: string[] = expandedColumns.map((col) => {
    if (col.isRowTotal) {
      return col.measureLabel ? `${col.measureLabel} (Row Total)` : col.label;
    }
    if (col.measureLabel && col.pivotValues && col.pivotValues.length > 0) {
      return col.measureLabel;
    }
    return col.label || col.fieldKey;
  });
  for (const th of trailingHeaders) {
    bottomRow.push(th);
  }
  headerRows.push(bottomRow);

  // Compute horizontal merges on each pivot cohort row (0 .. pivotDepth - 1) across pivoted columns
  const merges: CrosstabHeaderMerge[] = [];
  for (let r = 0; r < pivotDepth; r++) {
    let runStart = Math.max(0, nonPivotColCount);
    while (runStart < totalCols) {
      const val = headerRows[r][runStart];
      if (!val) {
        runStart++;
        continue;
      }
      const startColDef = expandedColumns[runStart];
      const startIsSummaryAtOrBelow =
        Boolean(startColDef?.isColumnGroupSummary) &&
        (startColDef?.columnGroupHeaderRowOffset ?? 0) <= r;
      let runEnd = runStart + 1;
      while (
        runEnd < totalCols &&
        headerRows[r][runEnd] === val &&
        // Never merge across a summary-to-child column boundary on the summary's own row or below
        Boolean(
          expandedColumns[runEnd]?.isColumnGroupSummary &&
            (expandedColumns[runEnd]?.columnGroupHeaderRowOffset ?? 0) <= r
        ) === startIsSummaryAtOrBelow &&
        (!startIsSummaryAtOrBelow ||
          expandedColumns[runEnd]?.columnGroupNodeId === startColDef?.columnGroupNodeId) &&
        // Only merge if all parent pivot rows above (0 .. r-1) also match
        Array.from({ length: r }, (_, pr) => headerRows[pr][runEnd] === headerRows[pr][runStart]).every(Boolean)
      ) {
        runEnd++;
      }
      const runLen = runEnd - runStart;
      if (runLen >= 2) {
        merges.push({
          rowOffset: r,
          colOffset: runStart,
          rowCount: 1,
          colCount: runLen,
        });
      }
      runStart = runEnd;
    }
  }

  return { headerRows, merges };
}

export interface WorksheetRowMatrixResult {
  rows: Record<string, any>[];
  subtotalRowIndices: number[];
  grandTotalRowIndex: number | null;
  outlineGroups: Array<{ startRowOffset: number; endRowOffset: number }>;
  treeNodes?: HierarchyTreeNodeMeta[];
  rowLevels?: number[];
}

export function getMaxVisibleLevel(displayLevel?: RowGroupDisplayLevel): number {
  switch (displayLevel) {
    case "all_collapsed":
      return 0;
    case "level_1":
      return 1;
    case "level_2":
      return 2;
    case "level_3":
      return 3;
    case "all_expanded":
    default:
      return 999;
  }
}

export function formatTreeGroupCellLabel(
  rawLabel: string,
  level: number,
  isParent: boolean,
  isExpanded: boolean,
  childCount: number,
  showGroupCounts: boolean = true
): string {
  const indent = "    ".repeat(Math.max(0, level));
  if (isParent) {
    const caret = isExpanded ? "▾ " : "▸ ";
    const countSuffix = showGroupCounts ? ` (${childCount})` : "";
    return `${indent}${caret}${rawLabel}${countSuffix}`;
  }
  return `${indent}  ${rawLabel}`;
}

export function resolveRowGroupFieldKeys(
  columns: ColumnDefinition[],
  options?: WriteOptions
): string[] {
  const pivotSet = new Set<string>(options?.pivots || []);
  const nonPivotDimKeys = columns
    .filter((c) => !pivotSet.has(c.fieldKey) && c.category !== "measure" && !c.isRowTotal)
    .map((c) => c.fieldKey);

  if (options?.rowGroupFields && options.rowGroupFields.length >= 2) {
    const matched = options.rowGroupFields.filter((f) => nonPivotDimKeys.includes(f));
    if (matched.length >= 2) {
      return matched;
    }
  }
  return nonPivotDimKeys;
}

/**
 * When `enableRowGroups` is enabled across >= 2 non-pivoted dimensions, collapses all
 * grouped dimensions into a single leading `Group` column (`__row_group__`), matching
 * Looker Explore's "Row Groups" and SAP Analysis for Office Compact Hierarchy display.
 */
export function resolveEffectiveColumnsForWrite(
  columns: ColumnDefinition[],
  options?: WriteOptions
): ColumnDefinition[] {
  if (!options?.enableRowGroups) {
    return columns;
  }
  const groupKeys = resolveRowGroupFieldKeys(columns, options);
  if (groupKeys.length < 2) {
    return columns;
  }
  const groupKeySet = new Set<string>(groupKeys);
  const groupHeader = options.groupColumnHeader?.trim() || "Group";
  const result: ColumnDefinition[] = [
    {
      fieldKey: "__row_group__",
      label: groupHeader,
      type: "string",
      category: "dimension",
    },
  ];
  for (const col of columns) {
    if (!groupKeySet.has(col.fieldKey)) {
      result.push(col);
    }
  }
  return result;
}

/**
 * Computes per-row hidden flags (`true` = hidden, `false` = visible) for a Row Groups
 * hierarchy tree based on the current `isExpanded` state of each parent in `treeNodes`.
 * A row `r` is visible iff every ancestor parent node covering `r` has `isExpanded === true`.
 */
export function computeTreeRowHiddenFlags(
  totalRows: number,
  treeNodes: HierarchyTreeNodeMeta[],
  grandTotalRowIndex: number | null = null
): boolean[] {
  const hidden = new Array<boolean>(totalRows).fill(false);
  for (const node of treeNodes) {
    if (!node.isExpanded) {
      for (
        let r = node.descendantStartOffset;
        r <= Math.min(totalRows - 1, node.descendantEndOffset);
        r++
      ) {
        if (r !== grandTotalRowIndex) {
          hidden[r] = true;
        }
      }
    }
  }
  return hidden;
}

function normalizeDimVal(v: any): string {
  if (v === null || v === undefined || v === "") return "__NULL__";
  return String(v);
}

function buildRowGroupsTreeMatrix(
  columns: ColumnDefinition[],
  dataRows: Record<string, any>[],
  subtotalsData?: Record<string, Record<string, any>[]> | null,
  totalsData?: Record<string, any> | null,
  options?: WriteOptions
): WorksheetRowMatrixResult {
  const groupFieldKeys = resolveRowGroupFieldKeys(columns, options);
  const K = groupFieldKeys.length;
  const maxVisibleLevel = getMaxVisibleLevel(options?.defaultDisplayLevel);
  const showCounts = options?.showGroupCounts !== false;
  const showSubtotals = options?.showGroupSubtotals !== false;

  const allSubtotalRows: Record<string, any>[] = subtotalsData
    ? Object.values(subtotalsData).flat()
    : [];
  const usesSentinel = allSubtotalRows.some((sr) =>
    groupFieldKeys.some((gk) => sr[gk] === "$$$_subtotal_$$$")
  );

  const outRows: Record<string, any>[] = [];
  const subtotalRowIndices: number[] = [];
  const outlineGroups: Array<{ startRowOffset: number; endRowOffset: number }> = [];
  const treeNodes: HierarchyTreeNodeMeta[] = [];
  const rowLevels: number[] = [];

  const findMatchingSubtotal = (
    level: number,
    pathVals: any[]
  ): Record<string, any> | undefined => {
    if (!showSubtotals || allSubtotalRows.length === 0) return undefined;
    return allSubtotalRows.find((sr) => {
      for (let j = 0; j <= level; j++) {
        const gk = groupFieldKeys[j];
        if (sr[gk] === "$$$_subtotal_$$$") return false;
        if (normalizeDimVal(sr[gk]) !== normalizeDimVal(pathVals[j])) return false;
      }
      for (let j = level + 1; j < K; j++) {
        const gk = groupFieldKeys[j];
        if (usesSentinel) {
          if (sr[gk] !== "$$$_subtotal_$$$") return false;
        } else {
          if (sr[gk] !== "$$$_subtotal_$$$" && sr[gk] !== null && sr[gk] !== undefined) {
            return false;
          }
        }
      }
      return true;
    });
  };

  const emitLevel = (level: number, subsetRows: Record<string, any>[], pathVals: any[]): number[] => {
    const dimKey = groupFieldKeys[level];
    const groups: Array<{ rawVal: any; rows: Record<string, any>[] }> = [];
    const groupMap = new Map<string, { rawVal: any; rows: Record<string, any>[] }>();

    for (const r of subsetRows) {
      const val = r[dimKey];
      const normKey = normalizeDimVal(val);
      let entry = groupMap.get(normKey);
      if (!entry) {
        entry = { rawVal: val, rows: [] };
        groupMap.set(normKey, entry);
        groups.push(entry);
      }
      entry.rows.push(r);
    }

    const levelNodeOffsets: number[] = [];

    for (const grp of groups) {
      const currentPath = [...pathVals, grp.rawVal];
      const displayVal =
        grp.rawVal === null || grp.rawVal === undefined || grp.rawVal === ""
          ? "∅"
          : String(grp.rawVal);

      // Compute direct child count (distinct values at level + 1, or leaf row count if level + 1 === K - 1)
      let directChildCount = grp.rows.length;
      if (level + 1 < K - 1) {
        const nextDimKey = groupFieldKeys[level + 1];
        const distinctNext = new Set(grp.rows.map((r) => normalizeDimVal(r[nextDimKey])));
        directChildCount = distinctNext.size;
      }

      const isExpanded = level < maxVisibleLevel;
      const matchingSub = findMatchingSubtotal(level, currentPath);
      const parentRow: Record<string, any> = matchingSub ? { ...matchingSub } : {};
      parentRow.__row_group__ = formatTreeGroupCellLabel(
        displayVal,
        level,
        true,
        isExpanded,
        directChildCount,
        showCounts
      );

      const parentRowOffset = outRows.length;
      outRows.push(parentRow);
      subtotalRowIndices.push(parentRowOffset);
      rowLevels.push(level);
      levelNodeOffsets.push(parentRowOffset);

      const nodeMeta: HierarchyTreeNodeMeta = {
        rowOffset: parentRowOffset,
        level,
        label: displayVal,
        childCount: directChildCount,
        descendantStartOffset: parentRowOffset + 1,
        descendantEndOffset: parentRowOffset + 1,
        directChildOffsets: [],
        isExpanded,
      };
      treeNodes.push(nodeMeta);

      if (level + 1 < K - 1) {
        nodeMeta.directChildOffsets = emitLevel(level + 1, grp.rows, currentPath);
      } else {
        const leafKey = groupFieldKeys[K - 1];
        const leafOffsets: number[] = [];
        for (const leaf of grp.rows) {
          const leafRaw = leaf[leafKey];
          const leafDisplay =
            leafRaw === null || leafRaw === undefined || leafRaw === ""
              ? "∅"
              : String(leafRaw);
          const formattedLeaf: Record<string, any> = {
            ...leaf,
            __row_group__: formatTreeGroupCellLabel(
              leafDisplay,
              K - 1,
              false,
              false,
              0,
              showCounts
            ),
          };
          const leafOffset = outRows.length;
          outRows.push(formattedLeaf);
          rowLevels.push(K - 1);
          leafOffsets.push(leafOffset);
        }
        nodeMeta.directChildOffsets = leafOffsets;
      }

      nodeMeta.descendantEndOffset = outRows.length - 1;
      if (level === 0 && nodeMeta.descendantEndOffset >= nodeMeta.descendantStartOffset) {
        outlineGroups.push({
          startRowOffset: nodeMeta.descendantStartOffset,
          endRowOffset: nodeMeta.descendantEndOffset,
        });
      }
    }

    return levelNodeOffsets;
  };

  if (dataRows.length > 0 && K >= 2) {
    emitLevel(0, dataRows, []);
  }

  let grandTotalRowIndex: number | null = null;
  if (totalsData && Object.keys(totalsData).length > 0) {
    const grandTotalRow: Record<string, any> = {
      ...totalsData,
      __row_group__: "Grand Total",
    };
    outRows.push(grandTotalRow);
    rowLevels.push(0);
    grandTotalRowIndex = outRows.length - 1;
  }

  return {
    rows: outRows,
    subtotalRowIndices,
    grandTotalRowIndex,
    outlineGroups,
    treeNodes,
    rowLevels,
  };
}

/**
 * Interleaves Looker-computed `subtotalsData` rows at dimension group breaks (or builds a
 * single-column indented Row Groups hierarchy tree when `options.enableRowGroups` is enabled)
 * and appends the Looker-computed `totalsData` Grand Total row at the bottom.
 */
export function buildWorksheetRowMatrix(
  columns: ColumnDefinition[],
  dataRows: Record<string, any>[],
  subtotalsData?: Record<string, Record<string, any>[]> | null,
  totalsData?: Record<string, any> | null,
  pivots?: string[],
  options?: WriteOptions
): WorksheetRowMatrixResult {
  const effectiveOpts: WriteOptions | undefined = options
    ? { ...options, pivots: options.pivots || pivots }
    : pivots
    ? { pivots }
    : undefined;

  if (effectiveOpts?.enableRowGroups) {
    const groupFieldKeys = resolveRowGroupFieldKeys(columns, effectiveOpts);
    if (groupFieldKeys.length >= 2) {
      return buildRowGroupsTreeMatrix(
        columns,
        dataRows,
        subtotalsData,
        totalsData,
        effectiveOpts
      );
    }
  }

  const pivotSet = new Set<string>(pivots || []);
  const nonPivotDims = columns.filter(
    (c) => !pivotSet.has(c.fieldKey) && c.category !== "measure" && !c.isRowTotal
  );

  const outRows: Record<string, any>[] = [];
  const subtotalRowIndices: number[] = [];
  const outlineGroups: Array<{ startRowOffset: number; endRowOffset: number }> = [];

  const allSubtotalRows: Record<string, any>[] = subtotalsData
    ? Object.values(subtotalsData).flat()
    : [];

  if (allSubtotalRows.length > 0 && nonPivotDims.length >= 2 && dataRows.length > 0) {
    const primaryDimKey = nonPivotDims[0].fieldKey;
    const childDimKeys = nonPivotDims.slice(1).map((d) => d.fieldKey);

    let groupStartOffset = 0;
    for (let i = 0; i < dataRows.length; i++) {
      const row = dataRows[i];
      outRows.push(row);

      const currentVal = row[primaryDimKey];
      const nextVal = i + 1 < dataRows.length ? dataRows[i + 1][primaryDimKey] : Symbol("END");

      if (currentVal !== nextVal) {
        const groupEndOffset = outRows.length - 1;
        const matchingSubtotal = allSubtotalRows.find(
          (sr) =>
            sr[primaryDimKey] === currentVal &&
            childDimKeys.every(
              (ck) => sr[ck] === "$$$_subtotal_$$$" || sr[ck] === null || sr[ck] === undefined
            )
        );

        if (matchingSubtotal) {
          const formattedSub: Record<string, any> = { ...matchingSubtotal };
          const labelVal =
            currentVal === null || currentVal === undefined || currentVal === ""
              ? "None"
              : String(currentVal);
          formattedSub[primaryDimKey] = `Subtotal (${labelVal})`;
          childDimKeys.forEach((ck) => {
            formattedSub[ck] = "";
          });

          outRows.push(formattedSub);
          subtotalRowIndices.push(outRows.length - 1);
          outlineGroups.push({
            startRowOffset: groupStartOffset,
            endRowOffset: groupEndOffset,
          });
        }
        groupStartOffset = outRows.length;
      }
    }
  } else {
    outRows.push(...dataRows);
  }

  let grandTotalRowIndex: number | null = null;
  if (totalsData && Object.keys(totalsData).length > 0) {
    const grandTotalRow: Record<string, any> = { ...totalsData };
    const firstColKey = nonPivotDims[0]?.fieldKey || columns[0]?.fieldKey;
    if (firstColKey) {
      grandTotalRow[firstColKey] = "Grand Total";
    }
    nonPivotDims.slice(1).forEach((d) => {
      grandTotalRow[d.fieldKey] = "";
    });
    outRows.push(grandTotalRow);
    grandTotalRowIndex = outRows.length - 1;
  }

  return {
    rows: outRows,
    subtotalRowIndices,
    grandTotalRowIndex,
    outlineGroups,
  };
}

/**
 * Calculates residual ranges (anchored at startRow, startCol) that must be cleared when
 * a refreshed query has fewer rows or fewer columns than its previous run.
 */
export function computeResidualRanges(
  prevRows: number,
  prevCols: number,
  newRows: number,
  newCols: number,
  startRow: number = 0,
  startCol: number = 0
): {
  extraColRange: { startRow: number; startCol: number; rowCount: number; colCount: number } | null;
  extraRowRange: { startRow: number; startCol: number; rowCount: number; colCount: number } | null;
} {
  const extraColRange =
    prevCols > newCols
      ? {
          startRow,
          startCol: startCol + newCols,
          rowCount: Math.max(prevRows, newRows),
          colCount: prevCols - newCols,
        }
      : null;

  const extraRowRange =
    prevRows > newRows
      ? {
          startRow: startRow + newRows,
          startCol,
          rowCount: prevRows - newRows,
          colCount: Math.max(newCols, prevCols),
        }
      : null;

  return { extraColRange, extraRowRange };
}

interface SavedFormulaColumn {
  offset: number;
  header: string;
  formula: string;
  numberFormat?: string;
}

export async function writeLookerDataToWorksheet(
  columns: ColumnDefinition[],
  dataRows: Record<string, any>[],
  options: WriteOptions,
  onProgress: (percent: number, currentRow: number) => void
): Promise<string> {
  const { startRow, startCol } = parseCellAddress(options.anchorCell);
  const matrix = buildWorksheetRowMatrix(
    columns,
    dataRows,
    options.subtotalsData,
    options.totalsData,
    options.pivots,
    options
  );
  if (matrix.treeNodes) {
    options.treeNodes = matrix.treeNodes;
  }
  if (matrix.rowLevels) {
    options.rowLevels = matrix.rowLevels;
  }
  const effectiveRows = matrix.rows;
  const baseColumns = resolveEffectiveColumnsForWrite(columns, options);
  const { expandedColumns, getValue, columnTreeNodes } = expandPivotedColumns(
    baseColumns,
    effectiveRows,
    options.pivots,
    options.includeRowTotals,
    options.pivotsMeta,
    options.pivotGroupMode || "measure_first",
    Boolean(options.enableColumnGroups),
    options.defaultColumnDisplayLevel || "all_expanded"
  );
  if (columnTreeNodes.length > 0) {
    options.columnTreeNodes = columnTreeNodes;
  } else {
    options.columnTreeNodes = undefined;
  }
  const colCount = Math.max(1, expandedColumns.length);
  const totalRows = effectiveRows.length;
  const batchSize = calculateBatchSize(colCount);
  const nullVal = options.nullDisplay ?? "";
  const hasSubtotals = matrix.subtotalRowIndices.length > 0;
  const useMultiRowHeader = Boolean(
    (options.useMultiRowCrosstabHeader ?? true) &&
      options.pivots &&
      options.pivots.length > 0
  );
  const {
    bannerRowCount,
    tableHeaderRowCount,
    totalHeaderRowCount: headerRowCount,
  } = getEffectiveHeaderRowCount(options);
  const tableHeaderStartRow = startRow + bannerRowCount;
  let writtenSheetId = "";

  await Excel.run(async (context) => {
    const createNewLookerSheet = async (): Promise<Excel.Worksheet> => {
      const sheets = context.workbook.worksheets;
      sheets.load(["items/name"]);
      await context.sync();

      const existingNames = sheets.items.map((s) => s.name);
      const rawName = (options.sheetName || "Looker Data").replace(/[\\/?*:[\]]/g, "").trim();
      const baseName = rawName.slice(0, 25) || "Looker Data";
      let targetName = baseName;
      let counter = 1;
      while (existingNames.some((n) => n.toLowerCase() === targetName.toLowerCase())) {
        targetName = `${baseName.slice(0, 20)} (${counter})`;
        counter++;
      }

      const newSheet = context.workbook.worksheets.add(targetName);
      newSheet.activate();
      newSheet.load(["id"]);
      await context.sync();
      return newSheet;
    };

    let sheet: Excel.Worksheet;
    let isCreatingNewSheet = options.destination === "new" && !options.targetSheetId;

    if (options.targetSheetId) {
      const candidate = context.workbook.worksheets.getItemOrNullObject(options.targetSheetId);
      candidate.load(["id", "isNullObject"]);
      await context.sync();
      if (!candidate.isNullObject) {
        sheet = candidate;
      } else {
        sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load(["id"]);
        await context.sync();
      }
    } else if (isCreatingNewSheet) {
      sheet = await createNewLookerSheet();
    } else {
      sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load(["id"]);
      await context.sync();
    }

    writtenSheetId = sheet.id;
    const app = context.workbook.application;

    // 1. Suspend costly calculation mode during bulk write
    let prevCalcMode:
      | Excel.CalculationMode
      | "Automatic"
      | "AutomaticExceptTables"
      | "Manual"
      | undefined = undefined;
    try {
      app.load("calculationMode");
      await context.sync();
      prevCalcMode = app.calculationMode;
      app.calculationMode = Excel.CalculationMode.manual;
      await context.sync();
    } catch (e) {
      console.warn("Calculation mode adjustment skipped or not supported:", e);
    }

    try {
      // 2. Inspect existing Looker table (matching tableId or startRow/startCol) & PivotTables on this sheet
      let prevTotalRows = options.prevRowCount ? options.prevRowCount + 1 : 0;
      let prevTotalCols = options.prevColCount || 0;
      let existingTable: Excel.Table | null = null;
      let savedTableName: string | null = options.tableId || null;
      let pivotRanges: BoundingBox[] = [...(options.siblingTableRanges || [])];

      if (!isCreatingNewSheet && (prevTotalRows > 1 || prevTotalCols > 1)) {
        try {
          const prevDataRng = sheet.getRangeByIndexes(
            startRow + 1,
            startCol,
            Math.max(1, prevTotalRows),
            Math.max(1, prevTotalCols)
          );
          prevDataRng.rowHidden = false;
          prevDataRng.columnHidden = false;
          prevDataRng.ungroup(Excel.GroupOption.byRows);
          await context.sync();
        } catch {
          // Ignore if no row groups existed
        }
        try {
          const prevColRng = sheet.getRangeByIndexes(
            startRow,
            startCol,
            Math.max(1, prevTotalRows + headerRowCount),
            Math.max(1, prevTotalCols)
          );
          prevColRng.columnHidden = false;
          prevColRng.ungroup(Excel.GroupOption.byColumns);
          await context.sync();
        } catch {
          // Ignore if no column groups existed
        }
      }

      if (!isCreatingNewSheet) {
        try {
          const existingTables = sheet.tables;
          const sheetPivots = sheet.pivotTables;
          existingTables.load(["items/name"]);
          sheetPivots.load(["items/name"]);
          await context.sync();

          if (sheetPivots.items.length > 0) {
            const pLayouts = sheetPivots.items.map((pt) => {
              const rng = pt.layout.getRange();
              rng.load(["rowIndex", "rowCount", "columnIndex", "columnCount"]);
              return rng;
            });
            await context.sync();
            pivotRanges.push(
              ...pLayouts.map((r) => ({
                top: r.rowIndex,
                bottom: r.rowIndex + r.rowCount - 1,
                left: r.columnIndex,
                right: r.columnIndex + r.columnCount - 1,
              }))
            );
          }

          if (existingTables.items.length > 0) {
            const tableRanges = existingTables.items.map((tbl) => {
              const rng = tbl.getRange();
              rng.load(["rowIndex", "rowCount", "columnIndex", "columnCount"]);
              return { tbl, rng };
            });
            await context.sync();

            for (const { tbl, rng } of tableRanges) {
              const matchesById = options.tableId && tbl.name === options.tableId;
              const matchesByAnchor =
                (rng.rowIndex === startRow || rng.rowIndex === startRow + 1) &&
                rng.columnIndex === startCol;

              if (matchesById || matchesByAnchor) {
                existingTable = tbl;
                savedTableName = tbl.name;
                prevTotalRows = Math.max(
                  prevTotalRows,
                  rng.rowCount + Math.max(0, rng.rowIndex - startRow)
                );
                prevTotalCols = Math.max(prevTotalCols, rng.columnCount);
              } else {
                // Protect other Excel Tables on the same worksheet from being overwritten or cleared
                pivotRanges.push({
                  top: rng.rowIndex,
                  bottom: rng.rowIndex + rng.rowCount - 1,
                  left: rng.columnIndex,
                  right: rng.columnIndex + rng.columnCount - 1,
                });
              }
            }

            if (existingTable && (!options.useExcelTable || hasSubtotals || useMultiRowHeader)) {
              existingTable.convertToRange();
              existingTable = null;
              await context.sync();
            }
          }
        } catch (err) {
          console.warn("Could not inspect existing tables/pivots:", err);
        }
      }

      const overlapsProtectedRange = (
        r0: number,
        c0: number,
        rowCount: number,
        colCount: number
      ): boolean => {
        const r1 = r0 + rowCount - 1;
        const c1 = c0 + colCount - 1;
        return pivotRanges.some(
          (p) => !(r1 < p.top || r0 > p.bottom || c1 < p.left || c0 > p.right)
        );
      };

      // 2.2 Guard against writing onto a PivotTable-only worksheet when anchor is A1
      if (
        !isCreatingNewSheet &&
        !existingTable &&
        startRow === 0 &&
        startCol === 0 &&
        overlapsProtectedRange(startRow, startCol, Math.max(1, totalRows + headerRowCount), colCount)
      ) {
        const allSheets = context.workbook.worksheets;
        allSheets.load(["items/id", "items/name"]);
        await context.sync();

        const propLookups = allSheets.items
          .filter((s) => s.id !== sheet.id)
          .map((s) => {
            const props = s.customProperties;
            const tbls = s.tables;
            props.load(["items"]);
            tbls.load(["items/name"]);
            return { s, props, tbls };
          });
        await context.sync();

        const sourceMatch = propLookups.find(({ props }) =>
          props.items.some((p) => p.key === "looker_query_config" || p.key === "looker_tables_v2")
        );

        if (sourceMatch) {
          sheet = sourceMatch.s;
          writtenSheetId = sheet.id;
          pivotRanges = [...(options.siblingTableRanges || [])];
          if (sourceMatch.tbls.items.length > 0) {
            existingTable = sourceMatch.tbls.items[0];
            savedTableName = existingTable.name;
            const tableRange = existingTable.getRange();
            tableRange.load(["rowIndex", "rowCount", "columnIndex", "columnCount"]);
            await context.sync();
            prevTotalRows = Math.max(prevTotalRows, tableRange.rowCount);
            prevTotalCols = Math.max(prevTotalCols, tableRange.columnCount);
          }
        } else {
          sheet = await createNewLookerSheet();
          writtenSheetId = sheet.id;
          isCreatingNewSheet = true;
          pivotRanges = [];
        }
      }

      // 2.5 Inspect adjacent columns for user formulas & shift whole columns if Looker schema changed
      const savedFormulaCols: SavedFormulaColumn[] = [];
      const maxFormulaColsToScan = 15;

      if (!isCreatingNewSheet) {
        try {
          if (options.prevColCount && options.prevColCount > 0 && options.prevColCount !== colCount) {
            const deltaCols = colCount - options.prevColCount;
            const boundaryCol = startCol + options.prevColCount;
            const checkFormulaCell = sheet.getRangeByIndexes(
              startRow + headerRowCount,
              boundaryCol,
              1,
              1
            );
            checkFormulaCell.load(["formulas"]);
            await context.sync();
            const hasFormulaAtPrevBoundary =
              typeof checkFormulaCell.formulas?.[0]?.[0] === "string" &&
              checkFormulaCell.formulas[0][0].startsWith("=");
            const hasProtectedToRight = pivotRanges.some((p) => p.left >= boundaryCol);

            if (
              deltaCols > 0 &&
              (hasFormulaAtPrevBoundary ||
                overlapsProtectedRange(
                  startRow,
                  boundaryCol,
                  Math.max(prevTotalRows, totalRows + headerRowCount),
                  deltaCols
                ))
            ) {
              const insertColRng = sheet
                .getRangeByIndexes(startRow, boundaryCol, 1, deltaCols)
                .getEntireColumn();
              insertColRng.insert(Excel.InsertShiftDirection.right);
              await context.sync();
              prevTotalCols += deltaCols;
              pivotRanges = pivotRanges.map((p) =>
                p.left >= boundaryCol
                  ? { ...p, left: p.left + deltaCols, right: p.right + deltaCols }
                  : p
              );
            } else if (deltaCols < 0 && hasFormulaAtPrevBoundary && !hasProtectedToRight) {
              const colsToRemove = Math.abs(deltaCols);
              if (!overlapsProtectedRange(startRow, startCol + colCount, prevTotalRows, colsToRemove)) {
                const deleteColRng = sheet
                  .getRangeByIndexes(startRow, startCol + colCount, 1, colsToRemove)
                  .getEntireColumn();
                deleteColRng.delete(Excel.DeleteShiftDirection.left);
                await context.sync();
                prevTotalCols = Math.max(colCount, prevTotalCols - colsToRemove);
              }
            }
          }

          if (options.autoExpandFormulas) {
            const scanRange = sheet.getRangeByIndexes(
              startRow + headerRowCount - 1,
              startCol + colCount,
              2,
              maxFormulaColsToScan
            );
            scanRange.load(["formulas", "values", "numberFormat"]);
            await context.sync();

            for (let offset = 0; offset < maxFormulaColsToScan; offset++) {
              if (
                overlapsProtectedRange(
                  startRow + headerRowCount - 1,
                  startCol + colCount + offset,
                  2,
                  1
                )
              ) {
                break;
              }
              const formulaCell = scanRange.formulas?.[1]?.[offset];
              const headerVal = scanRange.values?.[0]?.[offset];
              const numFmt = scanRange.numberFormat?.[1]?.[offset];

              if (typeof formulaCell === "string" && formulaCell.startsWith("=")) {
                savedFormulaCols.push({
                  offset,
                  header:
                    headerVal !== null && headerVal !== undefined && String(headerVal).trim() !== ""
                      ? String(headerVal)
                      : `Calculation ${offset + 1}`,
                  formula: formulaCell,
                  numberFormat:
                    typeof numFmt === "string" && numFmt !== "General" ? numFmt : undefined,
                });
              } else {
                break;
              }
            }
          }
        } catch (e) {
          console.warn("Formula inspection / column shift skipped:", e);
        }
      }

      const protectedCols = colCount + savedFormulaCols.length;
      const listObjectBodyRows =
        matrix.grandTotalRowIndex !== null ? totalRows - 1 : totalRows;

      // 2.7 Shrink-resize existing native Excel Table before clearing residual cells
      if (
        existingTable &&
        options.useExcelTable &&
        !hasSubtotals &&
        !useMultiRowHeader &&
        listObjectBodyRows > 0 &&
        (listObjectBodyRows + 1 < prevTotalRows || protectedCols < prevTotalCols)
      ) {
        try {
          const shrinkRows = Math.min(prevTotalRows, listObjectBodyRows + 1);
          const shrinkCols = Math.min(prevTotalCols, protectedCols);
          const shrinkRange = sheet.getRangeByIndexes(
            tableHeaderStartRow,
            startCol,
            shrinkRows,
            shrinkCols
          );
          existingTable.resize(shrinkRange);
          await context.sync();
        } catch (e) {
          console.warn("Pre-clear table shrink resize skipped:", e);
        }
      }

      // 2.8 Clear residual rows and columns from previous larger runs (anchored at startRow, startCol)
      if (!isCreatingNewSheet && (prevTotalRows > 0 || prevTotalCols > 0)) {
        // Unmerge any previously merged multi-row header cells before clearing/writing
        try {
          const prevHeaderArea = sheet.getRangeByIndexes(
            startRow,
            startCol,
            Math.max(headerRowCount, 5),
            Math.max(protectedCols, prevTotalCols)
          );
          prevHeaderArea.unmerge();
          await context.sync();
        } catch {
          // Ignore if no merged cells existed
        }

        const { extraColRange, extraRowRange } = computeResidualRanges(
          prevTotalRows,
          prevTotalCols,
          totalRows + headerRowCount,
          protectedCols,
          startRow,
          startCol
        );

        if (
          extraColRange &&
          !overlapsProtectedRange(
            extraColRange.startRow,
            extraColRange.startCol,
            extraColRange.rowCount,
            extraColRange.colCount
          )
        ) {
          try {
            const colRng = sheet.getRangeByIndexes(
              extraColRange.startRow,
              extraColRange.startCol,
              extraColRange.rowCount,
              extraColRange.colCount
            );
            colRng.clear(Excel.ClearApplyTo.all);
            await context.sync();
          } catch (e) {
            console.warn("Could not clear residual columns:", e);
          }
        }

        if (
          extraRowRange &&
          !overlapsProtectedRange(
            extraRowRange.startRow,
            extraRowRange.startCol,
            extraRowRange.rowCount,
            extraRowRange.colCount
          )
        ) {
          try {
            const rowRng = sheet.getRangeByIndexes(
              extraRowRange.startRow,
              extraRowRange.startCol,
              extraRowRange.rowCount,
              extraRowRange.colCount
            );
            rowRng.clear(Excel.ClearApplyTo.all);
            await context.sync();
          } catch (e) {
            console.warn("Could not clear residual rows:", e);
          }
        }
      }

      // 3. Write Source Information Banner (if enabled) and Table / Crosstab Header
      if (bannerRowCount > 0) {
        const bannerText =
          options.sourceBannerTitle || options.sheetName || "Looker Data";
        const bannerCells = Array.from({ length: protectedCols }, (_, idx) =>
          idx === 0 ? bannerText : ""
        );
        const bannerRange = sheet.getRangeByIndexes(startRow, startCol, 1, protectedCols);
        bannerRange.values = [bannerCells];
        bannerRange.format.fill.color = "#E8EAED";
        bannerRange.format.font.color = "#3C4043";
        bannerRange.format.font.size = 13;
        bannerRange.format.font.bold = false;
        bannerRange.format.font.italic = false;
        bannerRange.format.rowHeight = 24;
        bannerRange.format.verticalAlignment = Excel.VerticalAlignment.center;
        const bannerBottomBorder = bannerRange.format.borders.getItem(
          Excel.BorderIndex.edgeBottom
        );
        bannerBottomBorder.style = Excel.BorderLineStyle.continuous;
        bannerBottomBorder.color = "#BDC1C6";
      }

      const seenHeaders = new Map<string, number>();
      const deduplicateHeader = (raw: string): string => {
        const base = (raw || "Column").trim();
        const count = seenHeaders.get(base) || 0;
        seenHeaders.set(base, count + 1);
        return count === 0 ? base : `${base} (${count + 1})`;
      };

      const uniqueHeaderLabels = expandedColumns.map((c) =>
        deduplicateHeader(c.label || c.fieldKey || "Column")
      );
      const uniqueFormulaHeaders = savedFormulaCols.map((fCol) =>
        deduplicateHeader(fCol.header)
      );

      if (useMultiRowHeader) {
        const crosstabHeader = buildCrosstabHeaderMatrix(
          expandedColumns,
          options.pivots,
          uniqueFormulaHeaders,
          options.pivotDimensionLabels
        );
        const hRows = crosstabHeader.headerRows.length;
        const headerRange = sheet.getRangeByIndexes(
          tableHeaderStartRow,
          startCol,
          hRows,
          protectedCols
        );
        headerRange.values = crosstabHeader.headerRows;

        const lookerTopHeader = sheet.getRangeByIndexes(
          tableHeaderStartRow,
          startCol,
          hRows - 1,
          colCount
        );
        lookerTopHeader.format.font.bold = true;
        lookerTopHeader.format.fill.color = "#1D5288";
        lookerTopHeader.format.font.color = "#FFFFFF";
        lookerTopHeader.format.horizontalAlignment = Excel.HorizontalAlignment.left;

        // Right-align top-left corner cells holding the pivoted dimension names (e.g. Version, Time)
        const firstPivotColIdx = expandedColumns.findIndex(
          (c) => Boolean(c.pivotValues && c.pivotValues.length > 0) || Boolean(c.isRowTotal)
        );
        if (firstPivotColIdx > 0 && hRows > 1) {
          const cornerHeader = sheet.getRangeByIndexes(
            tableHeaderStartRow,
            startCol,
            hRows - 1,
            firstPivotColIdx
          );
          cornerHeader.format.horizontalAlignment = Excel.HorizontalAlignment.right;
        }

        const lookerBottomHeader = sheet.getRangeByIndexes(
          tableHeaderStartRow + hRows - 1,
          startCol,
          1,
          colCount
        );
        lookerBottomHeader.format.font.bold = true;
        lookerBottomHeader.format.fill.color = "#2B6CB0";
        lookerBottomHeader.format.font.color = "#FFFFFF";

        for (const m of crosstabHeader.merges) {
          const mergeRng = sheet.getRangeByIndexes(
            tableHeaderStartRow + m.rowOffset,
            startCol + m.colOffset,
            m.rowCount,
            m.colCount
          );
          mergeRng.merge();
        }
      } else {
        const allHeaderLabels = [...uniqueHeaderLabels, ...uniqueFormulaHeaders];
        const headerRange = sheet.getRangeByIndexes(
          tableHeaderStartRow,
          startCol,
          1,
          protectedCols
        );
        headerRange.values = [allHeaderLabels];
        if (!options.preserveUserFormatting) {
          const lookerHeaderRange = sheet.getRangeByIndexes(
            tableHeaderStartRow,
            startCol,
            1,
            colCount
          );
          lookerHeaderRange.format.font.bold = true;
          lookerHeaderRange.format.fill.color = "#1D5288"; // Looker Brand Blue
          lookerHeaderRange.format.font.color = "#FFFFFF";
        }
      }

      if (options.freezeHeader !== false && startRow === 0) {
        sheet.freezePanes.freezeRows(headerRowCount);
      }
      await context.sync();

      // 4. Batch write data in dynamic chunks starting at (startRow + headerRowCount, startCol)
      let currentRowOffset = headerRowCount;
      for (let i = 0; i < totalRows; i += batchSize) {
        const slice = effectiveRows.slice(i, i + batchSize);
        const rowBlock: (string | number | boolean)[][] = slice.map((item) =>
          expandedColumns.map((col) => {
            const val = getValue(item, col);
            if (val === null || val === undefined) {
              return nullVal;
            }
            if (typeof val === "object") {
              if ("value" in val && (val.value === null || typeof val.value !== "object")) {
                return val.value ?? nullVal;
              }
              return nullVal;
            }
            return val;
          })
        );

        const chunkRange = sheet.getRangeByIndexes(
          startRow + currentRowOffset,
          startCol,
          rowBlock.length,
          colCount
        );
        chunkRange.values = rowBlock;

        currentRowOffset += rowBlock.length;
        await context.sync();

        const rowsWritten = currentRowOffset - headerRowCount;
        onProgress(
          Math.min(100, Math.round((rowsWritten / Math.max(1, totalRows)) * 100)),
          rowsWritten
        );
        await new Promise((r) => setTimeout(r, 0));
      }

      // 5. Resize or Create Native Excel Table (ListObject) anchored at (tableHeaderStartRow, startCol)
      try {
        if (options.useExcelTable && !hasSubtotals && !useMultiRowHeader && listObjectBodyRows > 0) {
          const fullRange = sheet.getRangeByIndexes(
            tableHeaderStartRow,
            startCol,
            listObjectBodyRows + 1,
            protectedCols
          );
          if (existingTable) {
            existingTable.resize(fullRange);
            if (options.tableStyle) {
              existingTable.style = options.tableStyle;
            }
          } else {
            const table = sheet.tables.add(fullRange, true /* hasHeaders */);
            table.name = savedTableName || `Looker_${Date.now().toString().slice(-6)}`;
            table.style = options.tableStyle || "TableStyleLight1";
          }
          await context.sync();
        }
      } catch (err) {
        console.warn("Could not configure/resize native Excel table (ListObject):", err);
      }

      // 5.5 Apply Semantic Styling for Row Totals, Column Summary Groups, Subtotals, Row Outline Groups, and Grand Total
      try {
        if (totalRows > 0) {
          for (let c = 0; c < colCount; c++) {
            if (expandedColumns[c]?.isRowTotal) {
              const rtColRng = sheet.getRangeByIndexes(
                startRow + headerRowCount,
                startCol + c,
                totalRows,
                1
              );
              rtColRng.format.font.bold = true;
              rtColRng.format.fill.color = "#F8F9FA";
            } else if (expandedColumns[c]?.isColumnGroupSummary) {
              const sumColRng = sheet.getRangeByIndexes(
                startRow + headerRowCount,
                startCol + c,
                totalRows,
                1
              );
              sumColRng.format.font.bold = true;
              sumColRng.format.fill.color = "#F4F6F9";
            }
          }
        }

        if (columnTreeNodes.length > 0) {
          for (const colNode of columnTreeNodes) {
            const childWidth = colNode.childColEndOffset - colNode.childColStartOffset + 1;
            if (childWidth > 0) {
              const colGrpRng = sheet.getRangeByIndexes(
                tableHeaderStartRow,
                startCol + colNode.childColStartOffset,
                Math.max(1, tableHeaderRowCount + totalRows),
                childWidth
              );
              colGrpRng.group(Excel.GroupOption.byColumns);
            }
          }

          const colHiddenFlags = computeColumnTreeHiddenFlags(colCount, columnTreeNodes);
          let cRunStart = 0;
          while (cRunStart < colCount) {
            const state = colHiddenFlags[cRunStart];
            let cRunEnd = cRunStart;
            while (cRunEnd + 1 < colCount && colHiddenFlags[cRunEnd + 1] === state) {
              cRunEnd++;
            }
            if (state) {
              const hideColRng = sheet.getRangeByIndexes(
                tableHeaderStartRow,
                startCol + cRunStart,
                Math.max(1, tableHeaderRowCount + totalRows),
                cRunEnd - cRunStart + 1
              );
              hideColRng.columnHidden = true;
            }
            cRunStart = cRunEnd + 1;
          }
        }

        if (hasSubtotals) {
          const nodeByOffset = new Map<number, HierarchyTreeNodeMeta>();
          if (matrix.treeNodes) {
            matrix.treeNodes.forEach((n) => nodeByOffset.set(n.rowOffset, n));
          }

          for (const subRowIdx of matrix.subtotalRowIndices) {
            const sheetRow = startRow + headerRowCount + subRowIdx;
            const subRng = sheet.getRangeByIndexes(sheetRow, startCol, 1, colCount);
            const treeNode = nodeByOffset.get(subRowIdx);
            subRng.format.font.bold = true;
            if (treeNode) {
              subRng.format.fill.color = treeNode.level === 0 ? "#EDF2FA" : "#F6F8FA";
              if (treeNode.level === 0) {
                subRng.format.borders.getItem("EdgeTop").style = Excel.BorderLineStyle.continuous;
              }
            } else {
              subRng.format.fill.color = "#F3F2F1";
              subRng.format.borders.getItem("EdgeTop").style = Excel.BorderLineStyle.continuous;
              subRng.format.borders.getItem("EdgeBottom").style = Excel.BorderLineStyle.continuous;
            }
          }

          for (const grp of matrix.outlineGroups) {
            const rowCountGrp = grp.endRowOffset - grp.startRowOffset + 1;
            if (rowCountGrp > 0) {
              const grpRng = sheet.getRangeByIndexes(
                startRow + headerRowCount + grp.startRowOffset,
                startCol,
                rowCountGrp,
                colCount
              );
              grpRng.group(Excel.GroupOption.byRows);
            }
          }

          if (matrix.treeNodes && matrix.treeNodes.length > 0 && totalRows > 0) {
            const hiddenFlags = computeTreeRowHiddenFlags(
              totalRows,
              matrix.treeNodes,
              matrix.grandTotalRowIndex
            );
            let runStart = 0;
            while (runStart < totalRows) {
              const state = hiddenFlags[runStart];
              let runEnd = runStart;
              while (runEnd + 1 < totalRows && hiddenFlags[runEnd + 1] === state) {
                runEnd++;
              }
              if (state) {
                const hideRng = sheet.getRangeByIndexes(
                  startRow + headerRowCount + runStart,
                  startCol,
                  runEnd - runStart + 1,
                  colCount
                );
                hideRng.rowHidden = true;
              }
              runStart = runEnd + 1;
            }
          }
        }

        if (matrix.grandTotalRowIndex !== null) {
          const gtSheetRow = startRow + headerRowCount + matrix.grandTotalRowIndex;
          const gtRng = sheet.getRangeByIndexes(gtSheetRow, startCol, 1, colCount);
          gtRng.format.font.bold = true;
          gtRng.format.fill.color = "#E8F0FE";
          gtRng.format.borders.getItem("EdgeTop").style = Excel.BorderLineStyle.continuous;
          gtRng.format.borders.getItem("EdgeBottom").style = Excel.BorderLineStyle.double;
        }

        await context.sync();
      } catch (e) {
        console.warn("Could not apply totals/subtotals styling:", e);
      }

      // 6. Apply Column-Level Number Formats (isolated sync)
      try {
        if (totalRows > 0) {
          for (let c = 0; c < colCount; c++) {
            const colDef = expandedColumns[c];
            const fmt = colDef ? resolveEffectiveColumnNumberFormat(colDef, options) : undefined;
            if (fmt) {
              const colRange = sheet.getRangeByIndexes(
                startRow + headerRowCount,
                startCol + c,
                totalRows,
                1
              );
              colRange.numberFormat = [[fmt]];
            }
          }
          await context.sync();
        }
      } catch (e) {
        console.warn("Could not apply number formats:", e);
      }

      // 7. Restore & Auto-expand adjacent formula columns down to listObjectBodyRows
      const formulaRowCount = listObjectBodyRows > 0 ? listObjectBodyRows : totalRows;
      if (options.autoExpandFormulas && savedFormulaCols.length > 0 && formulaRowCount >= 1) {
        for (let idx = 0; idx < savedFormulaCols.length; idx++) {
          const fCol = savedFormulaCols[idx];
          try {
            const targetColIndex = startCol + colCount + fCol.offset;
            const headerCell = sheet.getRangeByIndexes(
              startRow + headerRowCount - 1,
              targetColIndex,
              1,
              1
            );
            headerCell.values = [[uniqueFormulaHeaders[idx]]];

            const formulaOrigin = sheet.getRangeByIndexes(
              startRow + headerRowCount,
              targetColIndex,
              1,
              1
            );
            formulaOrigin.formulas = [[fCol.formula]];
            await context.sync();

            if (formulaRowCount > 1) {
              const formulaTarget = sheet.getRangeByIndexes(
                startRow + headerRowCount,
                targetColIndex,
                formulaRowCount,
                1
              );
              try {
                formulaTarget.copyFrom(formulaOrigin, Excel.RangeCopyType.formulas);
                await context.sync();
              } catch {
                formulaOrigin.autoFill(formulaTarget, Excel.AutoFillType.fillDefault);
                await context.sync();
              }
            }

            if (fCol.numberFormat) {
              const fmtRange = sheet.getRangeByIndexes(
                startRow + headerRowCount,
                targetColIndex,
                formulaRowCount,
                1
              );
              fmtRange.numberFormat = [[fCol.numberFormat]];
              await context.sync();
            }
          } catch (e) {
            console.warn(`Could not auto-expand formula column at offset ${fCol.offset}:`, e);
          }
        }
      }

      // 8. Sampled Autofit (Top 150 rows) across Looker + formula columns
      try {
        const sampleRows = Math.min(totalRows + headerRowCount, 150);
        const sampleRange = sheet.getRangeByIndexes(startRow, startCol, sampleRows, protectedCols);
        sampleRange.format.autofitColumns();
        await context.sync();
      } catch (e) {
        console.warn("Autofit skipped:", e);
      }

      // 9. Automatically refresh all PivotTables in the workbook that depend on this data
      try {
        context.workbook.pivotTables.refreshAll();
        await context.sync();
      } catch (e) {
        console.warn("PivotTable auto-refresh skipped:", e);
      }
    } finally {
      // 10. Always restore calculation mode
      try {
        if (prevCalcMode) {
          app.calculationMode = prevCalcMode;
        }
        await context.sync();
      } catch (e) {
        console.warn("Could not restore calculationMode:", e);
      }
    }
  });

  return writtenSheetId;
}

/**
 * Applies `rowHidden` runs and updates `▸` / `▾` group cell labels in-place on the active
 * or target worksheet when a parent node or display level is toggled.
 */
async function syncTreeVisibilityToWorksheet(
  targetSheetId: string | undefined,
  anchorCell: string | undefined,
  headerRowCount: number,
  totalRows: number,
  colCount: number,
  treeNodes: HierarchyTreeNodeMeta[],
  showGroupCounts: boolean,
  hasGrandTotal: boolean,
  focusAfterRowAndCol?: { row: number; col: number }
): Promise<void> {
  const { startRow, startCol } = parseCellAddress(anchorCell);
  const grandTotalRowIndex = hasGrandTotal ? totalRows - 1 : null;
  const hiddenFlags = computeTreeRowHiddenFlags(totalRows, treeNodes, grandTotalRowIndex);

  await Excel.run(async (context) => {
    const sheet = targetSheetId
      ? context.workbook.worksheets.getItem(targetSheetId)
      : context.workbook.worksheets.getActiveWorksheet();

    // 1. Update parent `▸` / `▾` cell labels in the Group column
    for (const node of treeNodes) {
      const cell = sheet.getRangeByIndexes(
        startRow + headerRowCount + node.rowOffset,
        startCol,
        1,
        1
      );
      cell.values = [
        [
          formatTreeGroupCellLabel(
            node.label,
            node.level,
            true,
            node.isExpanded,
            node.childCount,
            showGroupCounts
          ),
        ],
      ];
    }

    // 2. Apply contiguous rowHidden runs
    let runStart = 0;
    while (runStart < totalRows) {
      const state = hiddenFlags[runStart];
      let runEnd = runStart;
      while (runEnd + 1 < totalRows && hiddenFlags[runEnd + 1] === state) {
        runEnd++;
      }
      const rng = sheet.getRangeByIndexes(
        startRow + headerRowCount + runStart,
        startCol,
        runEnd - runStart + 1,
        Math.max(1, colCount)
      );
      rng.rowHidden = state;
      runStart = runEnd + 1;
    }

    if (focusAfterRowAndCol) {
      const nextFocus = sheet.getRangeByIndexes(
        focusAfterRowAndCol.row,
        focusAfterRowAndCol.col,
        1,
        1
      );
      nextFocus.select();
    }

    await context.sync();
  });
}

/**
 * Toggles a single parent hierarchy node in the `Group` column (`startCol`) at `clickedSheetRow` (0-based)
 * or by explicit `rowOffset`, expanding/collapsing its descendants in-place without re-querying Looker.
 * Returns updated `HierarchyTreeNodeMeta[]` if a parent node was toggled, or `null` if the row is not a parent node.
 */
export async function toggleWorksheetTreeNode(
  targetSheetId: string | undefined,
  options: WriteOptions,
  totalRows: number,
  clickedSheetRow0Based: number,
  clickedSheetCol0Based: number
): Promise<HierarchyTreeNodeMeta[] | null> {
  if (!options.enableRowGroups || !options.treeNodes || options.treeNodes.length === 0) {
    return null;
  }
  const { startRow, startCol } = parseCellAddress(options.anchorCell);
  const { totalHeaderRowCount: headerRowCount } = getEffectiveHeaderRowCount(options);
  if (clickedSheetCol0Based !== startCol || clickedSheetRow0Based < startRow + headerRowCount) {
    return null;
  }
  const clickedOffset = clickedSheetRow0Based - (startRow + headerRowCount);
  const targetIdx = options.treeNodes.findIndex((n) => n.rowOffset === clickedOffset);
  if (targetIdx === -1) {
    return null;
  }

  const updatedNodes = options.treeNodes.map((n, i) =>
    i === targetIdx ? { ...n, isExpanded: !n.isExpanded } : { ...n }
  );

  const colCount = Math.max(2, options.prevColCount || 2);
  await syncTreeVisibilityToWorksheet(
    targetSheetId,
    options.anchorCell,
    headerRowCount,
    totalRows,
    colCount,
    updatedNodes,
    options.showGroupCounts !== false,
    Boolean(options.includeTotals),
    { row: clickedSheetRow0Based, col: startCol + 1 }
  );

  return updatedNodes;
}

/**
 * Expands or collapses all hierarchy nodes in a Row Groups worksheet table to the specified
 * `displayLevel` (`all_collapsed`, `level_1`, `level_2`, `level_3`, `all_expanded`) in-place.
 */
export async function applyWorksheetTreeDisplayLevel(
  targetSheetId: string | undefined,
  options: WriteOptions,
  totalRows: number,
  displayLevel: RowGroupDisplayLevel
): Promise<HierarchyTreeNodeMeta[] | null> {
  if (!options.enableRowGroups || !options.treeNodes || options.treeNodes.length === 0) {
    return null;
  }
  const { totalHeaderRowCount: headerRowCount } = getEffectiveHeaderRowCount(options);
  const maxVisibleLevel = getMaxVisibleLevel(displayLevel);
  const updatedNodes = options.treeNodes.map((n) => ({
    ...n,
    isExpanded: n.level < maxVisibleLevel,
  }));

  const colCount = Math.max(2, options.prevColCount || 2);
  await syncTreeVisibilityToWorksheet(
    targetSheetId,
    options.anchorCell,
    headerRowCount,
    totalRows,
    colCount,
    updatedNodes,
    options.showGroupCounts !== false,
    Boolean(options.includeTotals)
  );

  return updatedNodes;
}

/**
 * Updates the in-sheet Source Information Banner row to display `Loading... <Title>` during
 * query refresh, or restores `<Title>` upon completion (matching SAC's in-sheet loading banner).
 */
export async function setWorksheetBannerLoadingState(
  targetSheetId: string | undefined,
  options: WriteOptions,
  isLoading: boolean
): Promise<void> {
  if (!options.showSourceBanner) {
    return;
  }
  const { startRow, startCol } = parseCellAddress(options.anchorCell);
  const baseTitle = options.sourceBannerTitle || options.sheetName || "Looker Data";
  const displayText = isLoading ? `Loading... ${baseTitle}` : baseTitle;

  try {
    await Excel.run(async (context) => {
      const sheet = targetSheetId
        ? context.workbook.worksheets.getItemOrNullObject(targetSheetId)
        : context.workbook.worksheets.getActiveWorksheet();
      sheet.load(["isNullObject"]);
      await context.sync();
      if (sheet.isNullObject) return;

      const cell = sheet.getRangeByIndexes(startRow, startCol, 1, 1);
      cell.values = [[displayText]];
      cell.format.font.italic = isLoading;
      cell.format.font.color = isLoading ? "#5F6368" : "#3C4043";
      await context.sync();
    });
  } catch {
    // Non-fatal if Excel runtime is unavailable
  }
}

/**
 * Pure helper computing a boolean array of length `totalCols` indicating whether each
 * column offset (`0 .. totalCols - 1`) should have `columnHidden = true` given the
 * current expansion state of `columnTreeNodes`.
 */
export function computeColumnTreeHiddenFlags(
  totalCols: number,
  columnTreeNodes?: ColumnTreeNodeMeta[]
): boolean[] {
  const hidden = new Array(Math.max(0, totalCols)).fill(false);
  if (!columnTreeNodes || columnTreeNodes.length === 0 || totalCols <= 0) {
    return hidden;
  }

  const ordered = [...columnTreeNodes].sort((a, b) => a.level - b.level);
  for (const node of ordered) {
    if (!node.isExpanded) {
      const start = Math.max(0, node.childColStartOffset);
      const end = Math.min(totalCols - 1, node.childColEndOffset);
      for (let c = start; c <= end; c++) {
        hidden[c] = true;
      }
    }
  }

  return hidden;
}

/**
 * Applies `columnHidden` runs and updates `▸` / `▾` parent cohort header cells in-place on the
 * worksheet when a column hierarchy node or column display level is toggled.
 */
async function syncColumnTreeVisibilityToWorksheet(
  targetSheetId: string | undefined,
  options: WriteOptions,
  totalRows: number,
  colCount: number,
  columnTreeNodes: ColumnTreeNodeMeta[],
  focusAfterRowAndCol?: { row: number; col: number }
): Promise<void> {
  const { startRow, startCol } = parseCellAddress(options.anchorCell);
  const {
    bannerRowCount,
    tableHeaderRowCount,
    totalHeaderRowCount: headerRowCount,
  } = getEffectiveHeaderRowCount(options);
  const tableHeaderStartRow = startRow + bannerRowCount;
  const useMultiRowHeader = Boolean(
    options.useMultiRowCrosstabHeader && options.pivots && options.pivots.length > 0
  );
  const hiddenFlags = computeColumnTreeHiddenFlags(colCount, columnTreeNodes);

  await Excel.run(async (context) => {
    const sheet = targetSheetId
      ? context.workbook.worksheets.getItem(targetSheetId)
      : context.workbook.worksheets.getActiveWorksheet();

    // 1. Update `▸` / `▾` labels in the crosstab or single-row header cells
    for (const node of columnTreeNodes) {
      const prefix = node.isExpanded ? "▾" : "▸";
      if (useMultiRowHeader) {
        const cell = sheet.getRangeByIndexes(
          tableHeaderStartRow + node.headerRowOffset,
          startCol + node.summaryColOffset,
          1,
          1
        );
        cell.values = [[`${prefix} ${node.label}`]];
      } else {
        const cell = sheet.getRangeByIndexes(
          tableHeaderStartRow,
          startCol + node.summaryColOffset,
          1,
          1
        );
        cell.load(["values"]);
        await context.sync();
        const currentHeader = String(cell.values?.[0]?.[0] ?? "");
        const updatedHeader = currentHeader.replace(/[▾▸]\s*/g, `${prefix} `);
        cell.values = [[updatedHeader !== currentHeader ? updatedHeader : `${prefix} ${node.label}`]];
      }
    }

    // 2. Apply contiguous columnHidden runs across table columns
    let runStart = 0;
    while (runStart < colCount) {
      const state = hiddenFlags[runStart];
      let runEnd = runStart;
      while (runEnd + 1 < colCount && hiddenFlags[runEnd + 1] === state) {
        runEnd++;
      }
      const rng = sheet.getRangeByIndexes(
        tableHeaderStartRow,
        startCol + runStart,
        Math.max(1, tableHeaderRowCount + totalRows),
        runEnd - runStart + 1
      );
      rng.columnHidden = state;
      runStart = runEnd + 1;
    }

    if (focusAfterRowAndCol) {
      const nextFocus = sheet.getRangeByIndexes(
        focusAfterRowAndCol.row,
        focusAfterRowAndCol.col,
        1,
        1
      );
      nextFocus.select();
    }

    await context.sync();
  });
}

/**
 * Toggles a single parent column hierarchy cohort when the user clicks its `▾` / `▸` header cell
 * in the worksheet, expanding/collapsing its child columns in-place without re-querying Looker.
 */
export async function toggleWorksheetColumnTreeNode(
  targetSheetId: string | undefined,
  options: WriteOptions,
  totalRows: number,
  clickedSheetRow0Based: number,
  clickedSheetCol0Based: number
): Promise<ColumnTreeNodeMeta[] | null> {
  if (
    !options.enableColumnGroups ||
    !options.columnTreeNodes ||
    options.columnTreeNodes.length === 0
  ) {
    return null;
  }

  const { startRow, startCol } = parseCellAddress(options.anchorCell);
  const {
    bannerRowCount,
    totalHeaderRowCount: headerRowCount,
  } = getEffectiveHeaderRowCount(options);
  const tableHeaderStartRow = startRow + bannerRowCount;

  if (
    clickedSheetRow0Based < tableHeaderStartRow ||
    clickedSheetRow0Based >= startRow + headerRowCount ||
    clickedSheetCol0Based < startCol
  ) {
    return null;
  }

  const clickedHeaderRowOffset = clickedSheetRow0Based - tableHeaderStartRow;
  const clickedColOffset = clickedSheetCol0Based - startCol;

  let targetIdx = options.columnTreeNodes.findIndex(
    (n) =>
      n.headerRowOffset === clickedHeaderRowOffset &&
      clickedColOffset >= n.summaryColOffset &&
      clickedColOffset < n.summaryColOffset + Math.max(1, n.summaryColCount)
  );

  // Fallback: if single-row header or user clicked the summary column header cell on any header row
  if (targetIdx === -1) {
    targetIdx = options.columnTreeNodes.findIndex(
      (n) =>
        clickedColOffset >= n.summaryColOffset &&
        clickedColOffset < n.summaryColOffset + Math.max(1, n.summaryColCount)
    );
  }

  if (targetIdx === -1) {
    return null;
  }

  const updatedNodes = options.columnTreeNodes.map((n, i) =>
    i === targetIdx ? { ...n, isExpanded: !n.isExpanded } : { ...n }
  );

  const colCount = Math.max(2, options.prevColCount || 2);
  await syncColumnTreeVisibilityToWorksheet(
    targetSheetId,
    options,
    totalRows,
    colCount,
    updatedNodes,
    {
      row: Math.min(startRow + headerRowCount + Math.max(0, totalRows - 1), clickedSheetRow0Based + 1),
      col: clickedSheetCol0Based,
    }
  );

  return updatedNodes;
}

/**
 * Expands or collapses all column hierarchy nodes in a pivoted crosstab worksheet table to the
 * specified `displayLevel` (`all_collapsed`, `level_1`, `level_2`, `level_3`, `all_expanded`) in-place.
 */
export async function applyWorksheetColumnDisplayLevel(
  targetSheetId: string | undefined,
  options: WriteOptions,
  totalRows: number,
  displayLevel: RowGroupDisplayLevel
): Promise<ColumnTreeNodeMeta[] | null> {
  if (
    !options.enableColumnGroups ||
    !options.columnTreeNodes ||
    options.columnTreeNodes.length === 0
  ) {
    return null;
  }

  const maxVisibleLevel = getMaxVisibleLevel(displayLevel);
  const updatedNodes = options.columnTreeNodes.map((n) => ({
    ...n,
    isExpanded: n.level < maxVisibleLevel,
  }));

  const colCount = Math.max(2, options.prevColCount || 2);
  await syncColumnTreeVisibilityToWorksheet(
    targetSheetId,
    options,
    totalRows,
    colCount,
    updatedNodes
  );

  return updatedNodes;
}
