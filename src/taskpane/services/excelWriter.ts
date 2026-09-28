export interface ColumnDefinition {
  fieldKey: string;
  label: string;
  type?: string;
  excelFormat?: string;
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
}

export function calculateBatchSize(colCount: number, targetCellBudget: number = 35000): number {
  return Math.max(500, Math.floor(targetCellBudget / Math.max(1, colCount)));
}

function isNestedPivotObject(val: any): boolean {
  if (val && typeof val === "object" && !Array.isArray(val)) {
    const firstInner = Object.values(val)[0];
    return Boolean(firstInner && typeof firstInner === "object");
  }
  return false;
}

/**
 * Detects nested pivot structures in Looker JSON query results and expands them
 * into tabular Crosstab columns with appropriate labels and number formatting.
 */
export function expandPivotedColumns(
  columns: ColumnDefinition[],
  dataRows: Record<string, any>[],
  pivots?: string[]
): { expandedColumns: ColumnDefinition[]; getValue: (row: Record<string, any>, col: ColumnDefinition) => any } {
  if (dataRows.length === 0) {
    return {
      expandedColumns: columns,
      getValue: (row, col) => row[col.fieldKey],
    };
  }

  // Scan rows to find which columns are pivoted measures and collect discovered pivot dimension keys
  const pivotedColSample = new Map<string, Record<string, any>>();
  const discoveredPivotDims = new Set<string>(pivots || []);

  for (const col of columns) {
    for (const row of dataRows) {
      const val = row[col.fieldKey];
      if (isNestedPivotObject(val)) {
        pivotedColSample.set(col.fieldKey, val);
        Object.keys(val).forEach((pf) => discoveredPivotDims.add(pf));
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
  const pivotAccessorMap = new Map<string, { measureKey: string; pivotField: string; pivotVal: string }>();

  columns.forEach((col) => {
    // Skip columns that are pivoted dimensions (since their values become column headers)
    if (discoveredPivotDims.has(col.fieldKey)) {
      return;
    }

    const val = pivotedColSample.get(col.fieldKey);
    if (val) {
      const pivotFields = Object.keys(val);
      const pivotField = pivotFields[0];

      // Discover all unique pivot values across all dataRows in order
      const seenPivotVals = new Set<string>();
      dataRows.forEach((row) => {
        const inner = row[col.fieldKey]?.[pivotField];
        if (inner && typeof inner === "object") {
          Object.keys(inner).forEach((pv) => {
            if (pv !== "null" || inner[pv] !== 0) {
              seenPivotVals.add(pv);
            }
          });
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
        });
        pivotAccessorMap.set(key, { measureKey: col.fieldKey, pivotField, pivotVal: pVal });
      });
    } else {
      // Standard dimension/measure column
      expandedColumns.push(col);
    }
  });

  return {
    expandedColumns,
    getValue: (row, col) => {
      const acc = pivotAccessorMap.get(col.fieldKey);
      if (acc) {
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

/**
 * Calculates residual ranges that must be cleared when a new query has fewer rows
 * or fewer columns than the previously populated range.
 */
export function computeResidualRanges(
  prevRows: number,
  prevCols: number,
  newRows: number,
  newCols: number
): {
  extraColRange: { startRow: number; startCol: number; rowCount: number; colCount: number } | null;
  extraRowRange: { startRow: number; startCol: number; rowCount: number; colCount: number } | null;
} {
  const extraColRange =
    prevCols > newCols
      ? {
          startRow: 0,
          startCol: newCols,
          rowCount: Math.max(prevRows, newRows),
          colCount: prevCols - newCols,
        }
      : null;

  const extraRowRange =
    prevRows > newRows
      ? {
          startRow: newRows,
          startCol: 0,
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
  const { expandedColumns, getValue } = expandPivotedColumns(columns, dataRows, options.pivots);
  const colCount = Math.max(1, expandedColumns.length);
  const totalRows = dataRows.length;
  const batchSize = calculateBatchSize(colCount);
  const nullVal = options.nullDisplay ?? "";
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

    // Resolve target worksheet based on targetSheetId or destination option
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
      // 2. Inspect existing Looker table & PivotTables on this sheet.
      // IMPORTANT: Never use sheet.getUsedRange() for prevTotalRows/Cols because usedRange
      // includes any PivotTables or side-by-side analysis the user placed to the right or below
      // the Looker table, which would cause clear() to throw:
      // "The operation failed because it conflicts with a pivot table range."
      let prevTotalRows = options.prevRowCount ? options.prevRowCount + 1 : 0;
      let prevTotalCols = options.prevColCount || 0;
      let existingTable: Excel.Table | null = null;
      let savedTableName: string | null = null;
      let pivotRanges: Array<{ top: number; bottom: number; left: number; right: number }> = [];

      if (!isCreatingNewSheet) {
        try {
          const existingTables = sheet.tables;
          const sheetPivots = sheet.pivotTables;
          existingTables.load(["items/name"]);
          sheetPivots.load(["items/name"]);
          await context.sync();

          // Record bounding boxes of any PivotTables on the same worksheet so we never clear into them
          if (sheetPivots.items.length > 0) {
            const pLayouts = sheetPivots.items.map((pt) => {
              const rng = pt.layout.getRange();
              rng.load(["rowIndex", "rowCount", "columnIndex", "columnCount"]);
              return rng;
            });
            await context.sync();
            pivotRanges = pLayouts.map((r) => ({
              top: r.rowIndex,
              bottom: r.rowIndex + r.rowCount - 1,
              left: r.columnIndex,
              right: r.columnIndex + r.columnCount - 1,
            }));
          }

          if (existingTables.items.length > 0) {
            existingTable = existingTables.items[0];
            savedTableName = existingTable.name;
            const tableRange = existingTable.getRange();
            tableRange.load(["rowIndex", "rowCount", "columnIndex", "columnCount"]);
            await context.sync();

            prevTotalRows = Math.max(prevTotalRows, tableRange.rowIndex + tableRange.rowCount);
            prevTotalCols = Math.max(prevTotalCols, tableRange.columnIndex + tableRange.columnCount);

            // Only convert to range if the user explicitly unchecked "Create Native Excel Table"
            if (!options.useExcelTable) {
              existingTable.convertToRange();
              existingTable = null;
              await context.sync();
            }
          }
        } catch (err) {
          console.warn("Could not inspect existing tables/pivots:", err);
        }
      }

      // Helper to check if a target range [r0..r1, c0..c1] overlaps any PivotTable on this worksheet
      const overlapsPivotTable = (startRow: number, startCol: number, rowCount: number, colCount: number): boolean => {
        const r1 = startRow + rowCount - 1;
        const c1 = startCol + colCount - 1;
        return pivotRanges.some(
          (p) => !(r1 < p.top || startRow > p.bottom || c1 < p.left || startCol > p.right)
        );
      };

      // 2.2 Guard against writing onto a PivotTable worksheet:
      // If the resolved sheet has no Looker table and a PivotTable occupies the top-left write range
      // (e.g., user created a PivotTable on a New Worksheet and triggered Import/Refresh while that
      // PivotTable sheet was active), resolve the Looker source sheet or create a new sheet.
      if (!isCreatingNewSheet && !existingTable && overlapsPivotTable(0, 0, Math.max(1, totalRows + 1), colCount)) {
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
          props.items.some((p) => p.key === "looker_query_config")
        );

        if (sourceMatch) {
          sheet = sourceMatch.s;
          writtenSheetId = sheet.id;
          pivotRanges = [];
          if (sourceMatch.tbls.items.length > 0) {
            existingTable = sourceMatch.tbls.items[0];
            savedTableName = existingTable.name;
            const tableRange = existingTable.getRange();
            tableRange.load(["rowIndex", "rowCount", "columnIndex", "columnCount"]);
            await context.sync();
            prevTotalRows = Math.max(prevTotalRows, tableRange.rowIndex + tableRange.rowCount);
            prevTotalCols = Math.max(prevTotalCols, tableRange.columnIndex + tableRange.columnCount);
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
          // If Looker column count grew or shrank since the previous run, check if we need to shift
          // whole columns so adjacent user formulas or side-by-side PivotTables move cleanly.
          if (options.prevColCount && options.prevColCount > 0 && options.prevColCount !== colCount) {
            const deltaCols = colCount - options.prevColCount;
            const checkFormulaCell = sheet.getRangeByIndexes(1, options.prevColCount, 1, 1);
            checkFormulaCell.load(["formulas"]);
            await context.sync();
            const hasFormulaAtPrevBoundary =
              typeof checkFormulaCell.formulas?.[0]?.[0] === "string" &&
              checkFormulaCell.formulas[0][0].startsWith("=");
            const hasPivotToRight = pivotRanges.some((p) => p.left >= options.prevColCount!);

            if (deltaCols > 0 && (hasFormulaAtPrevBoundary || overlapsPivotTable(0, options.prevColCount, Math.max(prevTotalRows, totalRows + 1), deltaCols))) {
              // Insert entire worksheet columns so Excel shifts formulas AND any PivotTables to the right as a whole unit
              const insertColRng = sheet
                .getRangeByIndexes(0, options.prevColCount, 1, deltaCols)
                .getEntireColumn();
              insertColRng.insert(Excel.InsertShiftDirection.right);
              await context.sync();
              prevTotalCols += deltaCols;
              pivotRanges = pivotRanges.map((p) =>
                p.left >= options.prevColCount!
                  ? { ...p, left: p.left + deltaCols, right: p.right + deltaCols }
                  : p
              );
            } else if (deltaCols < 0 && hasFormulaAtPrevBoundary && !hasPivotToRight) {
              const colsToRemove = Math.abs(deltaCols);
              if (!overlapsPivotTable(0, colCount, prevTotalRows, colsToRemove)) {
                const deleteColRng = sheet
                  .getRangeByIndexes(0, colCount, 1, colsToRemove)
                  .getEntireColumn();
                deleteColRng.delete(Excel.DeleteShiftDirection.left);
                await context.sync();
                prevTotalCols = Math.max(colCount, prevTotalCols - colsToRemove);
              }
            }
          }

          if (options.autoExpandFormulas) {
            const scanRange = sheet.getRangeByIndexes(0, colCount, 2, maxFormulaColsToScan);
            scanRange.load(["formulas", "values", "numberFormat"]);
            await context.sync();

            for (let offset = 0; offset < maxFormulaColsToScan; offset++) {
              // Do not scan into a PivotTable
              if (overlapsPivotTable(0, colCount + offset, 2, 1)) {
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

      // 2.7 If an existing native Excel Table is shrinking in rows or columns,
      // shrink-resize the Table BEFORE clearing residual cells so the residual cells are outside the Table.
      if (
        existingTable &&
        options.useExcelTable &&
        totalRows > 0 &&
        (totalRows + 1 < prevTotalRows || protectedCols < prevTotalCols)
      ) {
        try {
          const shrinkRows = Math.min(prevTotalRows, totalRows + 1);
          const shrinkCols = Math.min(prevTotalCols, protectedCols);
          const shrinkRange = sheet.getRangeByIndexes(0, 0, shrinkRows, shrinkCols);
          existingTable.resize(shrinkRange);
          await context.sync();
        } catch (e) {
          console.warn("Pre-clear table shrink resize skipped:", e);
        }
      }

      // 2.8 Clear residual rows and columns from previous larger runs (skipping any PivotTable range)
      if (!isCreatingNewSheet && (prevTotalRows > 0 || prevTotalCols > 0)) {
        const { extraColRange, extraRowRange } = computeResidualRanges(
          prevTotalRows,
          prevTotalCols,
          totalRows + 1,
          protectedCols
        );

        if (
          extraColRange &&
          !overlapsPivotTable(
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
          !overlapsPivotTable(
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

      // 3. Write Header (deduplicate labels across Looker + preserved formula columns so Excel Table headers are always unique)
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
      const headerRange = sheet.getRangeByIndexes(0, 0, 1, protectedCols);
      headerRange.values = [allHeaderLabels];
      if (!options.preserveUserFormatting) {
        const lookerHeaderRange = sheet.getRangeByIndexes(0, 0, 1, colCount);
        lookerHeaderRange.format.font.bold = true;
        lookerHeaderRange.format.fill.color = "#1D5288"; // Looker Brand Blue
        lookerHeaderRange.format.font.color = "#FFFFFF";
      }

      if (options.freezeHeader !== false) {
        sheet.freezePanes.freezeRows(1);
      }
      await context.sync();

      // 4. Batch write data in dynamic chunks
      let currentRow = 1;
      for (let i = 0; i < totalRows; i += batchSize) {
        const slice = dataRows.slice(i, i + batchSize);
        const rowBlock: (string | number | boolean)[][] = slice.map((item) =>
          expandedColumns.map((col) => {
            const val = getValue(item, col);
            if (val === null || val === undefined) {
              return nullVal;
            }
            return val;
          })
        );

        const chunkRange = sheet.getRangeByIndexes(currentRow, 0, rowBlock.length, colCount);
        chunkRange.values = rowBlock;

        currentRow += rowBlock.length;
        await context.sync();

        onProgress(Math.min(100, Math.round((currentRow / totalRows) * 100)), currentRow);
        // Yield to browser event loop for smooth UI rendering
        await new Promise((r) => setTimeout(r, 0));
      }

      // 5. Resize or Create Native Excel Table (ListObject) BEFORE formula expansion
      // so Table Structured References (e.g. =[@[Sale Price]]*0.2) remain valid inside the Table!
      try {
        if (options.useExcelTable && totalRows > 0) {
          const fullRange = sheet.getRangeByIndexes(0, 0, totalRows + 1, protectedCols);
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

      // 6. Apply Column-Level Number Formats (isolated sync)
      try {
        if (totalRows > 0) {
          for (let c = 0; c < colCount; c++) {
            const fmt = expandedColumns[c]?.excelFormat;
            if (fmt) {
              const colRange = sheet.getRangeByIndexes(1, c, totalRows, 1);
              colRange.numberFormat = [[fmt]];
            }
          }
          await context.sync();
        }
      } catch (e) {
        console.warn("Could not apply number formats:", e);
      }

      // 7. Restore & Auto-expand adjacent formula columns down to totalRows (isolated sync)
      if (options.autoExpandFormulas && savedFormulaCols.length > 0 && totalRows >= 1) {
        for (let idx = 0; idx < savedFormulaCols.length; idx++) {
          const fCol = savedFormulaCols[idx];
          try {
            const targetColIndex = colCount + fCol.offset;
            const headerCell = sheet.getRangeByIndexes(0, targetColIndex, 1, 1);
            headerCell.values = [[uniqueFormulaHeaders[idx]]];

            const formulaOrigin = sheet.getRangeByIndexes(1, targetColIndex, 1, 1);
            formulaOrigin.formulas = [[fCol.formula]];
            await context.sync();

            if (totalRows > 1) {
              const formulaTarget = sheet.getRangeByIndexes(1, targetColIndex, totalRows, 1);
              try {
                formulaTarget.copyFrom(formulaOrigin, Excel.RangeCopyType.formulas);
                await context.sync();
              } catch {
                formulaOrigin.autoFill(formulaTarget, Excel.AutoFillType.fillDefault);
                await context.sync();
              }
            }

            if (fCol.numberFormat) {
              const fmtRange = sheet.getRangeByIndexes(1, targetColIndex, totalRows, 1);
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
        const sampleRange = sheet.getRangeByIndexes(0, 0, sampleRows, protectedCols);
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
