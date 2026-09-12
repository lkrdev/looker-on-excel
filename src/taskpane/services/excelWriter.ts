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
}

/**
 * Detects nested pivot structures in Looker JSON query results and expands them
 * into tabular Crosstab columns with appropriate labels and number formatting.
 */
export function expandPivotedColumns(
  columns: ColumnDefinition[],
  dataRows: Record<string, any>[]
): { expandedColumns: ColumnDefinition[]; getValue: (row: Record<string, any>, col: ColumnDefinition) => any } {
  if (dataRows.length === 0) {
    return {
      expandedColumns: columns,
      getValue: (row, col) => row[col.fieldKey],
    };
  }

  // Check if any column contains nested pivot object { [pivotField]: { [pivotVal]: number } }
  const sample = dataRows[0];
  const hasNestedPivots = columns.some((col) => {
    const val = sample[col.fieldKey];
    if (val && typeof val === "object" && !Array.isArray(val)) {
      const firstInner = Object.values(val)[0];
      return firstInner && typeof firstInner === "object";
    }
    return false;
  });

  if (!hasNestedPivots) {
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
    const val = sample[col.fieldKey];
    if (val && typeof val === "object" && !Array.isArray(val)) {
      // It's a pivoted measure! e.g. val = { "order_items.status": { "Complete": 123 } }
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

export async function writeLookerDataToWorksheet(
  columns: ColumnDefinition[],
  dataRows: Record<string, any>[],
  options: WriteOptions,
  onProgress: (percent: number, currentRow: number) => void
): Promise<void> {
  const TARGET_CELL_BUDGET = 35000;
  const { expandedColumns, getValue } = expandPivotedColumns(columns, dataRows);
  const colCount = expandedColumns.length;
  const totalRows = dataRows.length;
  const batchSize = Math.max(500, Math.floor(TARGET_CELL_BUDGET / Math.max(1, colCount)));
  const nullVal = options.nullDisplay ?? "";

  await Excel.run(async (context) => {
    let sheet: Excel.Worksheet;

    // Resolve target worksheet based on destination option
    if (options.destination === "new") {
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

    // 1. Suspend costly calculation & grid repainting
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
    } catch (e) {
      console.warn("Calculation mode adjustment skipped or not supported:", e);
    }
    app.suspendScreenUpdatingUntilNextSync();
    app.suspendApiCalculationUntilNextSync();

    // 2. Inspect adjacent column for user formulas (SAC Formula Auto-Expansion)
    let hasAdjacentFormula = false;
    let savedAdjacentFormula: string | null = null;
    let savedAdjacentHeader: string | null = null;

    if (options.autoExpandFormulas) {
      try {
        const adjacentCell = sheet.getRangeByIndexes(1, colCount, 1, 1);
        const adjacentHeaderCell = sheet.getRangeByIndexes(0, colCount, 1, 1);
        adjacentCell.load(["formulas"]);
        adjacentHeaderCell.load(["values"]);
        await context.sync();
        const formulaVal = adjacentCell.formulas?.[0]?.[0];
        if (typeof formulaVal === "string" && formulaVal.startsWith("=")) {
          hasAdjacentFormula = true;
          savedAdjacentFormula = formulaVal;
          savedAdjacentHeader = (adjacentHeaderCell.values?.[0]?.[0] as string) || "Calculation";
        }
      } catch {
        hasAdjacentFormula = false;
      }
    }

    // 2.5 Clean existing tables & residual cell ranges (prevents ghost rows/columns on active sheet)
    let prevTableRows = 0;
    let prevTableCols = 0;

    try {
      const existingTables = sheet.tables;
      existingTables.load(["items/name"]);
      await context.sync();

      if (existingTables.items.length > 0) {
        for (const t of existingTables.items) {
          try {
            const tableRange = t.getRange();
            tableRange.load(["rowCount", "columnCount"]);
            await context.sync();
            prevTableRows = Math.max(prevTableRows, tableRange.rowCount);
            prevTableCols = Math.max(prevTableCols, tableRange.columnCount);
            t.convertToRange();
          } catch (e) {
            console.warn("Could not inspect/convert existing table:", e);
          }
        }
        await context.sync();
      }
    } catch (err) {
      console.warn("Could not inspect existing tables:", err);
    }

    // If writing to existing active worksheet, clear any residual rows or columns from previous larger runs.
    // If the user added an adjacent calculation column, exclude it from residual column wiping.
    if (options.destination !== "new" && (prevTableRows > 0 || prevTableCols > 0)) {
      const residualColsToProtect = hasAdjacentFormula ? 1 : 0;
      const { extraColRange, extraRowRange } = computeResidualRanges(
        prevTableRows,
        prevTableCols,
        totalRows + 1,
        colCount + residualColsToProtect
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

    // 3. Write Header
    const headerRange = sheet.getRangeByIndexes(0, 0, 1, colCount);
    headerRange.values = [expandedColumns.map((c) => c.label)];
    if (!options.preserveUserFormatting) {
      headerRange.format.font.bold = true;
      headerRange.format.fill.color = "#1D5288"; // Looker Brand Blue
      headerRange.format.font.color = "#FFFFFF";
    }

    if (options.freezeHeader !== false) {
      sheet.freezePanes.freezeRows(1);
    }

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

    // 5. Apply Column-Level Number Formats (Dual-Channel Contract)
    for (let c = 0; c < colCount; c++) {
      const fmt = expandedColumns[c].excelFormat;
      if (fmt && totalRows > 0) {
        const colRange = sheet.getRangeByIndexes(1, c, totalRows, 1);
        colRange.numberFormat = [[fmt]];
      }
    }

    // 6. Sampled Autofit (Top 150 rows) to prevent freezing on 2.1M cells
    const sampleRows = Math.min(totalRows + 1, 150);
    const sampleRange = sheet.getRangeByIndexes(0, 0, sampleRows, colCount);
    sampleRange.format.autofitColumns();

    // 7. Auto-fill adjacent formula down if detected
    if (options.autoExpandFormulas && hasAdjacentFormula && savedAdjacentFormula && totalRows >= 1) {
      try {
        if (savedAdjacentHeader) {
          const headerCell = sheet.getRangeByIndexes(0, colCount, 1, 1);
          headerCell.values = [[savedAdjacentHeader]];
        }
        const formulaOrigin = sheet.getRangeByIndexes(1, colCount, 1, 1);
        formulaOrigin.formulas = [[savedAdjacentFormula]];
        if (totalRows > 1) {
          const formulaTarget = sheet.getRangeByIndexes(1, colCount, totalRows, 1);
          formulaOrigin.autoFill(formulaTarget, Excel.AutoFillType.fillCopy);
        }
      } catch (e) {
        console.warn("Could not auto-fill adjacent formula:", e);
      }
    }

    // 8. Native Excel Table (ListObject) creation
    try {
      if (options.useExcelTable && totalRows > 0) {
        const fullRange = sheet.getRangeByIndexes(0, 0, totalRows + 1, colCount);
        const table = sheet.tables.add(fullRange, true /* hasHeaders */);
        table.name = `Looker_${Date.now().toString().slice(-6)}`;
        table.style = options.tableStyle || "TableStyleLight1";
        await context.sync();
      }
    } catch (err) {
      console.warn("Could not configure native Excel table (ListObject):", err);
    }

    // 9. Restore calculation mode
    if (prevCalcMode) {
      try {
        app.calculationMode = prevCalcMode;
      } catch (e) {
        console.warn("Could not restore calculationMode:", e);
      }
    }
    await context.sync();
  });
}
