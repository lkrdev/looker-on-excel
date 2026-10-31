export interface ColumnDefinition {
  fieldKey: string;
  label: string;
  type?: string;
  excelFormat?: string;
  category?: "dimension" | "measure" | "parameter";
  isRowTotal?: boolean;
}

export interface BoundingBox {
  top: number;
  bottom: number;
  left: number;
  right: number;
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

function isNestedPivotObject(val: any): boolean {
  if (val && typeof val === "object" && !Array.isArray(val)) {
    const entries = Object.entries(val).filter(([k]) => k !== "$$$_row_total_$$$");
    if (entries.length === 0) return false;
    const firstInner = entries[0][1];
    return Boolean(firstInner && typeof firstInner === "object");
  }
  return false;
}

/**
 * Detects nested pivot structures in Looker JSON / json_detail query results and expands them
 * into tabular Crosstab columns (including Looker-computed Row Totals when present)
 * with appropriate labels and number formatting.
 */
export function expandPivotedColumns(
  columns: ColumnDefinition[],
  dataRows: Record<string, any>[],
  pivots?: string[],
  includeRowTotals?: boolean
): {
  expandedColumns: ColumnDefinition[];
  getValue: (row: Record<string, any>, col: ColumnDefinition) => any;
} {
  if (dataRows.length === 0) {
    return {
      expandedColumns: columns,
      getValue: (row, col) => row[col.fieldKey],
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
    };
  }

  const expandedColumns: ColumnDefinition[] = [];
  const pivotAccessorMap = new Map<
    string,
    { measureKey: string; pivotField: string; pivotVal: string; isRowTotal?: boolean }
  >();

  columns.forEach((col) => {
    if (discoveredPivotDims.has(col.fieldKey)) {
      return;
    }

    const val = pivotedColSample.get(col.fieldKey);
    if (val) {
      const pivotFields = Object.keys(val).filter((pf) => pf !== "$$$_row_total_$$$");
      const pivotField = pivotFields[0];

      const seenPivotVals = new Set<string>();
      let hasRowTotalInRows = Boolean(includeRowTotals);

      dataRows.forEach((row) => {
        const measureObj = row[col.fieldKey];
        if (measureObj && typeof measureObj === "object") {
          if ("$$$_row_total_$$$" in measureObj) {
            hasRowTotalInRows = true;
          }
          const inner = measureObj[pivotField];
          if (inner && typeof inner === "object") {
            Object.keys(inner).forEach((pv) => {
              if (pv !== "null" || inner[pv] !== 0) {
                seenPivotVals.add(pv);
              }
            });
          }
        }
      });

      seenPivotVals.forEach((pVal) => {
        const key = `__pivot__:${col.fieldKey}:::${pivotField}:::${pVal}`;
        const displayPivotVal = pVal === "null" ? "None" : pVal;
        expandedColumns.push({
          fieldKey: key,
          label: `${col.label} (${displayPivotVal})`,
          type: col.type,
          excelFormat: col.excelFormat,
          category: col.category || "measure",
        });
        pivotAccessorMap.set(key, { measureKey: col.fieldKey, pivotField, pivotVal: pVal });
      });

      if (hasRowTotalInRows) {
        const rtKey = `__pivot_row_total__:${col.fieldKey}`;
        expandedColumns.push({
          fieldKey: rtKey,
          label: `${col.label} (Row Total)`,
          type: col.type,
          excelFormat: col.excelFormat,
          category: col.category || "measure",
          isRowTotal: true,
        });
        pivotAccessorMap.set(rtKey, {
          measureKey: col.fieldKey,
          pivotField,
          pivotVal: "$$$_row_total_$$$",
          isRowTotal: true,
        });
      }
    } else {
      expandedColumns.push(col);
    }
  });

  return {
    expandedColumns,
    getValue: (row, col) => {
      const acc = pivotAccessorMap.get(col.fieldKey);
      if (acc) {
        if (acc.isRowTotal) {
          return row[acc.measureKey]?.["$$$_row_total_$$$"];
        }
        return row[acc.measureKey]?.[acc.pivotField]?.[acc.pivotVal];
      }
      const val = row[col.fieldKey];
      if (typeof val === "object" && val !== null) {
        return JSON.stringify(val);
      }
      return val;
    },
  };
}

export interface WorksheetRowMatrixResult {
  rows: Record<string, any>[];
  subtotalRowIndices: number[];
  grandTotalRowIndex: number | null;
  outlineGroups: Array<{ startRowOffset: number; endRowOffset: number }>;
}

/**
 * Interleaves Looker-computed `subtotalsData` rows at dimension group breaks and appends
 * the Looker-computed `totalsData` Grand Total row at the bottom.
 */
export function buildWorksheetRowMatrix(
  columns: ColumnDefinition[],
  dataRows: Record<string, any>[],
  subtotalsData?: Record<string, Record<string, any>[]> | null,
  totalsData?: Record<string, any> | null,
  pivots?: string[]
): WorksheetRowMatrixResult {
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
    options.pivots
  );
  const effectiveRows = matrix.rows;
  const { expandedColumns, getValue } = expandPivotedColumns(
    columns,
    effectiveRows,
    options.pivots,
    options.includeRowTotals
  );
  const colCount = Math.max(1, expandedColumns.length);
  const totalRows = effectiveRows.length;
  const batchSize = calculateBatchSize(colCount);
  const nullVal = options.nullDisplay ?? "";
  const hasSubtotals = matrix.subtotalRowIndices.length > 0;
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

      if (!isCreatingNewSheet && prevTotalRows > 1) {
        try {
          const prevDataRng = sheet.getRangeByIndexes(
            startRow + 1,
            startCol,
            prevTotalRows,
            Math.max(1, prevTotalCols)
          );
          prevDataRng.ungroup(Excel.GroupOption.byRows);
          await context.sync();
        } catch {
          // Ignore if no row groups existed
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
              const matchesByAnchor = rng.rowIndex === startRow && rng.columnIndex === startCol;

              if (matchesById || matchesByAnchor) {
                existingTable = tbl;
                savedTableName = tbl.name;
                prevTotalRows = Math.max(prevTotalRows, rng.rowCount);
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

            if (existingTable && (!options.useExcelTable || hasSubtotals)) {
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
        overlapsProtectedRange(startRow, startCol, Math.max(1, totalRows + 1), colCount)
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
            const checkFormulaCell = sheet.getRangeByIndexes(startRow + 1, boundaryCol, 1, 1);
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
                  Math.max(prevTotalRows, totalRows + 1),
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
              startRow,
              startCol + colCount,
              2,
              maxFormulaColsToScan
            );
            scanRange.load(["formulas", "values", "numberFormat"]);
            await context.sync();

            for (let offset = 0; offset < maxFormulaColsToScan; offset++) {
              if (overlapsProtectedRange(startRow, startCol + colCount + offset, 2, 1)) {
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
        listObjectBodyRows > 0 &&
        (listObjectBodyRows + 1 < prevTotalRows || protectedCols < prevTotalCols)
      ) {
        try {
          const shrinkRows = Math.min(prevTotalRows, listObjectBodyRows + 1);
          const shrinkCols = Math.min(prevTotalCols, protectedCols);
          const shrinkRange = sheet.getRangeByIndexes(startRow, startCol, shrinkRows, shrinkCols);
          existingTable.resize(shrinkRange);
          await context.sync();
        } catch (e) {
          console.warn("Pre-clear table shrink resize skipped:", e);
        }
      }

      // 2.8 Clear residual rows and columns from previous larger runs (anchored at startRow, startCol)
      if (!isCreatingNewSheet && (prevTotalRows > 0 || prevTotalCols > 0)) {
        const { extraColRange, extraRowRange } = computeResidualRanges(
          prevTotalRows,
          prevTotalCols,
          totalRows + 1,
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

      // 3. Write Header at (startRow, startCol)
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

      const allHeaderLabels = [...uniqueHeaderLabels, ...uniqueFormulaHeaders];
      const headerRange = sheet.getRangeByIndexes(startRow, startCol, 1, protectedCols);
      headerRange.values = [allHeaderLabels];
      if (!options.preserveUserFormatting) {
        const lookerHeaderRange = sheet.getRangeByIndexes(startRow, startCol, 1, colCount);
        lookerHeaderRange.format.font.bold = true;
        lookerHeaderRange.format.fill.color = "#1D5288"; // Looker Brand Blue
        lookerHeaderRange.format.font.color = "#FFFFFF";
      }

      if (options.freezeHeader !== false && startRow === 0) {
        sheet.freezePanes.freezeRows(1);
      }
      await context.sync();

      // 4. Batch write data in dynamic chunks starting at (startRow + 1, startCol)
      let currentRowOffset = 1;
      for (let i = 0; i < totalRows; i += batchSize) {
        const slice = effectiveRows.slice(i, i + batchSize);
        const rowBlock: (string | number | boolean)[][] = slice.map((item) =>
          expandedColumns.map((col) => {
            const val = getValue(item, col);
            if (val === null || val === undefined) {
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

        onProgress(
          Math.min(100, Math.round((currentRowOffset / totalRows) * 100)),
          currentRowOffset
        );
        await new Promise((r) => setTimeout(r, 0));
      }

      // 5. Resize or Create Native Excel Table (ListObject) anchored at (startRow, startCol)
      try {
        if (options.useExcelTable && !hasSubtotals && listObjectBodyRows > 0) {
          const fullRange = sheet.getRangeByIndexes(
            startRow,
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

      // 5.5 Apply Semantic Styling for Row Totals, Subtotals, Row Outline Groups, and Grand Total
      try {
        if (totalRows > 0) {
          for (let c = 0; c < colCount; c++) {
            if (expandedColumns[c]?.isRowTotal) {
              const rtColRng = sheet.getRangeByIndexes(startRow + 1, startCol + c, totalRows, 1);
              rtColRng.format.font.bold = true;
              rtColRng.format.fill.color = "#F8F9FA";
            }
          }
        }

        if (hasSubtotals) {
          for (const subRowIdx of matrix.subtotalRowIndices) {
            const sheetRow = startRow + subRowIdx + 1;
            const subRng = sheet.getRangeByIndexes(sheetRow, startCol, 1, colCount);
            subRng.format.font.bold = true;
            subRng.format.fill.color = "#F3F2F1";
            subRng.format.borders.getItem("EdgeTop").style = Excel.BorderLineStyle.continuous;
            subRng.format.borders.getItem("EdgeBottom").style = Excel.BorderLineStyle.continuous;
          }

          for (const grp of matrix.outlineGroups) {
            const rowCountGrp = grp.endRowOffset - grp.startRowOffset + 1;
            if (rowCountGrp > 0) {
              const grpRng = sheet.getRangeByIndexes(
                startRow + grp.startRowOffset + 1,
                startCol,
                rowCountGrp,
                colCount
              );
              grpRng.group(Excel.GroupOption.byRows);
            }
          }
        }

        if (matrix.grandTotalRowIndex !== null) {
          const gtSheetRow = startRow + matrix.grandTotalRowIndex + 1;
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
            const fmt = expandedColumns[c]?.excelFormat;
            if (fmt) {
              const colRange = sheet.getRangeByIndexes(startRow + 1, startCol + c, totalRows, 1);
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
            const headerCell = sheet.getRangeByIndexes(startRow, targetColIndex, 1, 1);
            headerCell.values = [[uniqueFormulaHeaders[idx]]];

            const formulaOrigin = sheet.getRangeByIndexes(startRow + 1, targetColIndex, 1, 1);
            formulaOrigin.formulas = [[fCol.formula]];
            await context.sync();

            if (formulaRowCount > 1) {
              const formulaTarget = sheet.getRangeByIndexes(
                startRow + 1,
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
                startRow + 1,
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
        const sampleRows = Math.min(totalRows + 1, 150);
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
      if (prevCalcMode) {
        try {
          app.calculationMode = prevCalcMode;
        } catch (e) {
          console.warn("Could not restore calculationMode:", e);
        }
      }
      await context.sync();
    }
  });

  return writtenSheetId;
}
