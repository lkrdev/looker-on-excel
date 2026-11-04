import assert from "node:assert/strict";
import {
  buildCrosstabHeaderMatrix,
  expandPivotedColumns,
  getEffectiveHeaderRowCount,
  ColumnDefinition,
} from "../src/taskpane/services/excelWriter";
import { normalizeQueryResult } from "../src/taskpane/services/queryTaskRunner";
import { normalizeTableConfigBounds, StoredSheetConfig } from "../src/taskpane/services/sheetMetadata";

console.log("Running Feature 7 (Source Information Banner Row & Crosstab Corner Pivot Labels) unit tests...");

// 1. getEffectiveHeaderRowCount calculates bannerRowCount, tableHeaderRowCount, and totalHeaderRowCount
{
  const flatNoBanner = getEffectiveHeaderRowCount({
    showSourceBanner: false,
    useMultiRowCrosstabHeader: false,
  });
  assert.deepEqual(flatNoBanner, {
    bannerRowCount: 0,
    tableHeaderRowCount: 1,
    totalHeaderRowCount: 1,
  });

  const flatWithBanner = getEffectiveHeaderRowCount({
    showSourceBanner: true,
    useMultiRowCrosstabHeader: false,
  });
  assert.deepEqual(flatWithBanner, {
    bannerRowCount: 1,
    tableHeaderRowCount: 1,
    totalHeaderRowCount: 2,
  });

  const crosstabWithBanner = getEffectiveHeaderRowCount({
    showSourceBanner: true,
    pivots: ["orders.version", "orders.quarter"],
  });
  assert.deepEqual(crosstabWithBanner, {
    bannerRowCount: 1,
    tableHeaderRowCount: 3, // 2 pivot rows + 1 measure row (useMultiRowCrosstabHeader defaults to true)
    totalHeaderRowCount: 4,
  });

  const crosstabExplicitOff = getEffectiveHeaderRowCount({
    showSourceBanner: true,
    useMultiRowCrosstabHeader: false,
    pivots: ["orders.version", "orders.quarter"],
  });
  assert.deepEqual(crosstabExplicitOff, {
    bannerRowCount: 1,
    tableHeaderRowCount: 1,
    totalHeaderRowCount: 2,
  });
}

// 2. buildCrosstabHeaderMatrix populates top-left corner pivot dimension labels in nonPivotColCount - 1
{
  const columns: ColumnDefinition[] = [
    { fieldKey: "account.entity", label: "Entity", category: "dimension" },
    { fieldKey: "account.name", label: "Operating Account", category: "dimension" },
    { fieldKey: "orders.version", label: "Version", category: "dimension" },
    { fieldKey: "orders.time", label: "Time", category: "dimension" },
    { fieldKey: "orders.amount", label: "Amount", category: "measure" },
  ];
  const pivots = ["orders.version", "orders.time"];
  const mockJsonDetail = {
    pivots: [
      {
        key: "Strategic Plan|Q1 (2021)",
        data: { "orders.version": "Strategic Plan", "orders.time": "Q1 (2021)" },
      },
      {
        key: "Strategic Plan|Jan (2021)",
        data: { "orders.version": "Strategic Plan", "orders.time": "Jan (2021)" },
      },
    ],
    data: [
      {
        "account.entity": { value: "Finance" },
        "account.name": { value: "Gross Profit" },
        "orders.amount": {
          "Strategic Plan|Q1 (2021)": { value: 635466650.23 },
          "Strategic Plan|Jan (2021)": { value: 183267622.53 },
        },
      },
    ],
  };

  const bundle = normalizeQueryResult(mockJsonDetail, pivots, "json_detail");

  const { expandedColumns } = expandPivotedColumns(
    columns,
    bundle.rows,
    pivots,
    false,
    bundle.pivotsMeta,
    "cohort_first"
  );

  const matrix = buildCrosstabHeaderMatrix(
    expandedColumns,
    pivots,
    undefined,
    ["Version", "Time"]
  );

  // 2 pivot rows + 1 bottom row = 3 header rows
  assert.equal(matrix.headerRows.length, 3);
  // Top-left corner cells at col 0 are blank, col 1 (nonPivotColCount - 1) hold pivot labels
  assert.equal(matrix.headerRows[0][0], "");
  assert.equal(matrix.headerRows[0][1], "Version");
  assert.equal(matrix.headerRows[0][2], "Strategic Plan");
  assert.equal(matrix.headerRows[0][3], "Strategic Plan");
  assert.equal(matrix.headerRows[1][0], "");
  assert.equal(matrix.headerRows[1][1], "Time");
  assert.equal(matrix.headerRows[1][2], "Q1 (2021)");
  assert.equal(matrix.headerRows[1][3], "Jan (2021)");
  // Measure/leaf row holds non-pivot dimension labels
  assert.equal(matrix.headerRows[2][0], "Entity");
  assert.equal(matrix.headerRows[2][1], "Operating Account");
  assert.equal(matrix.headerRows[2][2], "Amount");
  // Horizontal merges never include non-pivot corner columns (< 2)
  for (const span of matrix.merges) {
    assert.ok(
      span.colOffset >= 2,
      `Expected merge span colOffset (${span.colOffset}) >= nonPivotColCount (2)`
    );
  }
}

// 3. normalizeTableConfigBounds accounts for showSourceBanner + crosstab header rows
{
  const baseCfg: StoredSheetConfig = {
    queryPayload: {
      model: "finance",
      view: "operating_income",
      fields: ["account.name", "orders.version", "orders.amount"],
      pivots: ["orders.version"],
    },
    columns: [
      { fieldKey: "account.name", label: "Operating Account", category: "dimension" },
      { fieldKey: "orders.version", label: "Version", category: "dimension" },
      { fieldKey: "orders.amount", label: "Amount", category: "measure" },
    ],
    options: {
      anchorCell: "A1",
      showSourceBanner: true,
      useMultiRowCrosstabHeader: true,
      pivots: ["orders.version"],
      prevColCount: 4,
    },
    lastRefreshed: "2026-11-03T20:00:00Z",
    rowCount: 10,
  };

  const normalized = normalizeTableConfigBounds(baseCfg);
  // 1 banner row + 2 crosstab header rows (1 pivot + 1 measure) + 10 data rows = 13 rows total (indices 0..12)
  assert.equal(normalized.startRow, 0);
  assert.equal(normalized.endRow, 12);
  assert.equal(normalized.startCol, 0);
  assert.equal(normalized.endCol, 3);
}

console.log("All Feature 7 unit tests passed!");
