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

export async function writeLookerDataToWorksheet(
  columns: ColumnDefinition[],
  dataRows: Record<string, any>[],
  options: WriteOptions,
  onProgress: (percent: number, currentRow: number) => void
): Promise<void> {
  const TARGET_CELL_BUDGET = 35000;
  const colCount = columns.length;
  const totalRows = dataRows.length;
  const batchSize = Math.max(500, Math.floor(TARGET_CELL_BUDGET / colCount));
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
    if (options.autoExpandFormulas) {
      try {
        const adjacentCell = sheet.getRangeByIndexes(1, colCount, 1, 1);
        adjacentCell.load(["formulas"]);
        await context.sync();
        const formulaVal = adjacentCell.formulas?.[0]?.[0];
        hasAdjacentFormula = typeof formulaVal === "string" && formulaVal.startsWith("=");
      } catch {
        hasAdjacentFormula = false;
      }
    }

    // 3. Write Header
    const headerRange = sheet.getRangeByIndexes(0, 0, 1, colCount);
    headerRange.values = [columns.map((c) => c.label)];
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
        columns.map((col) => {
          const val = item[col.fieldKey];
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
      const fmt = columns[c].excelFormat;
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
    if (options.autoExpandFormulas && hasAdjacentFormula && totalRows > 1) {
      try {
        const formulaOrigin = sheet.getRangeByIndexes(1, colCount, 1, 1);
        const formulaTarget = sheet.getRangeByIndexes(1, colCount, totalRows, 1);
        formulaOrigin.autoFill(formulaTarget, Excel.AutoFillType.fillCopy);
      } catch (e) {
        console.warn("Could not auto-fill adjacent formula:", e);
      }
    }

    // 8. Native Excel Table (ListObject) handling
    try {
      const existingTables = sheet.tables;
      existingTables.load(["items/name"]);
      await context.sync();

      if (existingTables.items.length > 0) {
        for (const t of existingTables.items) {
          try {
            t.convertToRange();
          } catch (e) {
            console.warn("Could not convert existing table to range:", e);
          }
        }
        await context.sync();
      }

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
