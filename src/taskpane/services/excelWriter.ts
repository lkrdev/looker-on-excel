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
): Promise<void> {
  const { expandedColumns, getValue } = expandPivotedColumns(columns, dataRows, options.pivots);
  const colCount = Math.max(1, expandedColumns.length);
  const totalRows = dataRows.length;
  const batchSize = calculateBatchSize(colCount);
  const nullVal = options.nullDisplay ?? "";

  await Excel.run(async (context) => {
    let sheet: Excel.Worksheet;
    const isCreatingNewSheet = options.destination === "new" && !options.targetSheetId;

    // Resolve target worksheet based on targetSheetId or destination option
    if (options.targetSheetId) {
      sheet = context.workbook.worksheets.getItem(options.targetSheetId);
    } else if (isCreatingNewSheet) {
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

      sheet = context.workbook.worksheets.add(targetName);
      sheet.activate();
      await context.sync();
    } else {
      sheet = context.workbook.worksheets.getActiveWorksheet();
    }
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
      // 2. Inspect existing tables & used range FIRST, and convert tables to ranges
      // CRITICAL: Converting existing Excel Tables to ranges BEFORE reading adjacent formulas
      // forces Excel to automatically translate Table Structured References (e.g. =[@[Price]]*1.1)
      // into standard A1 cell references (e.g. =C2*1.1) so they can be cleanly copied/expanded.
      let prevTotalRows = options.prevRowCount ? options.prevRowCount + 1 : 0;
      let prevTotalCols = options.prevColCount || 0;
      let savedTableName: string | null = null;

      if (!isCreatingNewSheet) {
        try {
          const existingTables = sheet.tables;
          const usedRange = sheet.getUsedRangeOrNullObject(true);
          existingTables.load(["items/name"]);
          usedRange.load(["rowIndex", "rowCount", "columnIndex", "columnCount"]);
          await context.sync();

          if (!usedRange.isNullObject) {
            prevTotalRows = Math.max(prevTotalRows, usedRange.rowIndex + usedRange.rowCount);
            prevTotalCols = Math.max(prevTotalCols, usedRange.columnIndex + usedRange.columnCount);
          }

          if (existingTables.items.length > 0) {
            savedTableName = existingTables.items[0].name;
            const tableRanges = existingTables.items.map((t) => {
              const tableRange = t.getRange();
              tableRange.load(["rowIndex", "rowCount", "columnIndex", "columnCount"]);
              return { table: t, tableRange };
            });
            await context.sync();

            for (const { table, tableRange } of tableRanges) {
              try {
                prevTotalRows = Math.max(prevTotalRows, tableRange.rowIndex + tableRange.rowCount);
                prevTotalCols = Math.max(prevTotalCols, tableRange.columnIndex + tableRange.columnCount);
                table.convertToRange();
              } catch (e) {
                console.warn("Could not inspect/convert existing table:", e);
              }
            }
            await context.sync();
          }
        } catch (err) {
          console.warn("Could not inspect existing tables/usedRange:", err);
        }
      }

      // 2.5 Inspect adjacent columns for user formulas (supports multiple contiguous formula columns)
      const savedFormulaCols: SavedFormulaColumn[] = [];
      const formulaScanStartCol = options.prevColCount && options.prevColCount > 0 ? options.prevColCount : colCount;
      const maxFormulaColsToScan = 15;

      if (!isCreatingNewSheet && options.autoExpandFormulas) {
        try {
          // Shift adjacent formula columns if Looker column count changed since previous run
          if (options.prevColCount && options.prevColCount > 0 && options.prevColCount !== colCount && prevTotalCols > options.prevColCount) {
            const checkRange = sheet.getRangeByIndexes(1, options.prevColCount, 1, 1);
            checkRange.load(["formulas"]);
            await context.sync();
            const firstAdjFormula = checkRange.formulas?.[0]?.[0];
            if (typeof firstAdjFormula === "string" && firstAdjFormula.startsWith("=")) {
              const rowsToShift = Math.max(prevTotalRows, totalRows + 1, 2);
              if (colCount > options.prevColCount) {
                const insertRng = sheet.getRangeByIndexes(0, options.prevColCount, rowsToShift, colCount - options.prevColCount);
                insertRng.insert(Excel.InsertShiftDirection.right);
                await context.sync();
                prevTotalCols += colCount - options.prevColCount;
              } else if (colCount < options.prevColCount) {
                const deleteRng = sheet.getRangeByIndexes(0, colCount, rowsToShift, options.prevColCount - colCount);
                deleteRng.delete(Excel.DeleteShiftDirection.left);
                await context.sync();
                prevTotalCols = Math.max(colCount, prevTotalCols - (options.prevColCount - colCount));
              }
            }
          }

          const scanRange = sheet.getRangeByIndexes(0, colCount, 2, maxFormulaColsToScan);
          scanRange.load(["formulas", "values", "numberFormat"]);
          await context.sync();

          for (let offset = 0; offset < maxFormulaColsToScan; offset++) {
            const formulaCell = scanRange.formulas?.[1]?.[offset];
            const headerVal = scanRange.values?.[0]?.[offset];
            const numFmt = scanRange.numberFormat?.[1]?.[offset];

            if (typeof formulaCell === "string" && formulaCell.startsWith("=")) {
              savedFormulaCols.push({
                offset,
                header: headerVal !== null && headerVal !== undefined && String(headerVal).trim() !== ""
                  ? String(headerVal)
                  : `Calculation ${offset + 1}`,
                formula: formulaCell,
                numberFormat: typeof numFmt === "string" && numFmt !== "General" ? numFmt : undefined,
              });
            } else {
              // Stop at first non-formula column
              break;
            }
          }
        } catch (e) {
          console.warn("Formula inspection skipped:", e);
        }
      }

      // 2.8 Clear residual rows and columns from previous larger runs
      const protectedCols = colCount + savedFormulaCols.length;
      if (!isCreatingNewSheet && (prevTotalRows > 0 || prevTotalCols > 0)) {
        const { extraColRange, extraRowRange } = computeResidualRanges(
          prevTotalRows,
          prevTotalCols,
          totalRows + 1,
          protectedCols
        );

        if (extraColRange) {
          try {
            const colRng = sheet.getRangeByIndexes(
              extraColRange.startRow,
              extraColRange.startCol,
              extraColRange.rowCount,
              extraColRange.colCount
            );
            colRng.clear(Excel.ClearApplyTo.all);
          } catch (e) {
            console.warn("Could not clear residual columns:", e);
          }
        }

        if (extraRowRange) {
          try {
            const rowRng = sheet.getRangeByIndexes(
              extraRowRange.startRow,
              extraRowRange.startCol,
              extraRowRange.rowCount,
              extraRowRange.colCount
            );
            rowRng.clear(Excel.ClearApplyTo.all);
          } catch (e) {
            console.warn("Could not clear residual rows:", e);
          }
        }
        await context.sync();
      }

      // 3. Write Header (deduplicate labels if needed so Excel Table creation never fails on duplicate column names)
      const seenHeaders = new Map<string, number>();
      const uniqueHeaderLabels = expandedColumns.map((c) => {
        const base = (c.label || c.fieldKey || "Column").trim();
        const count = seenHeaders.get(base) || 0;
        seenHeaders.set(base, count + 1);
        return count === 0 ? base : `${base} (${count + 1})`;
      });

      const headerRange = sheet.getRangeByIndexes(0, 0, 1, colCount);
      headerRange.values = [uniqueHeaderLabels];
      if (!options.preserveUserFormatting) {
        headerRange.format.font.bold = true;
        headerRange.format.fill.color = "#1D5288"; // Looker Brand Blue
        headerRange.format.font.color = "#FFFFFF";
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

      // 5. Apply Column-Level Number Formats (isolated sync)
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

      // 6. Restore & Auto-expand adjacent formula columns down to totalRows (isolated sync)
      if (options.autoExpandFormulas && savedFormulaCols.length > 0 && totalRows >= 1) {
        for (const fCol of savedFormulaCols) {
          try {
            const targetColIndex = colCount + fCol.offset;
            const headerCell = sheet.getRangeByIndexes(0, targetColIndex, 1, 1);
            headerCell.values = [[fCol.header]];

            const formulaOrigin = sheet.getRangeByIndexes(1, targetColIndex, 1, 1);
            formulaOrigin.formulas = [[fCol.formula]];
            await context.sync();

            if (totalRows > 1) {
              const formulaTarget = sheet.getRangeByIndexes(1, targetColIndex, totalRows, 1);
              try {
                // copyFrom with RangeCopyType.formulas reliably shifts relative references (C2 -> C3..CN)
                // across the entire target range even if cells previously contained values
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

      // 7. Sampled Autofit (Top 150 rows) across Looker + formula columns
      try {
        const sampleRows = Math.min(totalRows + 1, 150);
        const sampleRange = sheet.getRangeByIndexes(0, 0, sampleRows, protectedCols);
        sampleRange.format.autofitColumns();
        await context.sync();
      } catch (e) {
        console.warn("Autofit skipped:", e);
      }

      // 8. Native Excel Table (ListObject) creation across Looker + adjacent formula columns
      try {
        if (options.useExcelTable && totalRows > 0) {
          const fullRange = sheet.getRangeByIndexes(0, 0, totalRows + 1, protectedCols);
          const table = sheet.tables.add(fullRange, true /* hasHeaders */);
          table.name = savedTableName || `Looker_${Date.now().toString().slice(-6)}`;
          table.style = options.tableStyle || "TableStyleLight1";
          await context.sync();
        }
      } catch (err) {
        console.warn("Could not configure native Excel table (ListObject):", err);
      }
    } finally {
      // 9. Always restore calculation mode
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
}
