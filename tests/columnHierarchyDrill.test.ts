import assert from "node:assert/strict";
import {
  expandPivotedColumns,
  buildCrosstabHeaderMatrix,
  computeColumnTreeHiddenFlags,
  isAdditiveMeasure,
  ColumnDefinition,
} from "../src/taskpane/services/excelWriter";
import { normalizeQueryResult } from "../src/taskpane/services/queryTaskRunner";

console.log("Running Feature 8 (Column Hierarchy Expand / Collapse) unit tests...");

const mockColumns: ColumnDefinition[] = [
  { fieldKey: "account.name", label: "Operating Account", category: "dimension" },
  { fieldKey: "orders.version", label: "Version", category: "dimension" },
  { fieldKey: "orders.quarter", label: "Quarter", category: "dimension" },
  { fieldKey: "orders.month", label: "Time", category: "dimension" },
  {
    fieldKey: "orders.amount",
    label: "Amount",
    type: "sum",
    category: "measure",
    excelFormat: "$#,##0.00",
  },
  {
    fieldKey: "orders.margin_pct",
    label: "Margin %",
    type: "average",
    category: "measure",
    excelFormat: "0.0%",
  },
];

// 1. 3-level column hierarchy (Version -> Quarter -> Month) matching SAC Screenshot 1
{
  const pivots = ["orders.version", "orders.quarter", "orders.month"];
  const mockJsonDetail = {
    pivots: [
      {
        key: "Strategic Plan|Q1 (2021)|Jan (2021)",
        data: {
          "orders.version": "Strategic Plan",
          "orders.quarter": "Q1 (2021)",
          "orders.month": "Jan (2021)",
        },
      },
      {
        key: "Strategic Plan|Q1 (2021)|Feb (2021)",
        data: {
          "orders.version": "Strategic Plan",
          "orders.quarter": "Q1 (2021)",
          "orders.month": "Feb (2021)",
        },
      },
      {
        key: "Strategic Plan|Q1 (2021)|Mar (2021)",
        data: {
          "orders.version": "Strategic Plan",
          "orders.quarter": "Q1 (2021)",
          "orders.month": "Mar (2021)",
        },
      },
    ],
    data: [
      {
        "account.name": { value: "Operating Income" },
        "orders.amount": {
          "Strategic Plan|Q1 (2021)|Jan (2021)": { value: 125808538.17 },
          "Strategic Plan|Q1 (2021)|Feb (2021)": { value: 169113903.5 },
          "Strategic Plan|Q1 (2021)|Mar (2021)": { value: 155755793.25 },
        },
        "orders.margin_pct": {
          "Strategic Plan|Q1 (2021)|Jan (2021)": { value: 0.3 },
          "Strategic Plan|Q1 (2021)|Feb (2021)": { value: 0.4 },
          "Strategic Plan|Q1 (2021)|Mar (2021)": { value: 0.5 },
        },
      },
    ],
  };

  const bundle = normalizeQueryResult(mockJsonDetail, pivots, "json_detail");
  const singleMeasureCols = mockColumns.filter((c) => c.fieldKey !== "orders.margin_pct");

  const { expandedColumns, getValue, columnTreeNodes } = expandPivotedColumns(
    singleMeasureCols,
    bundle.rows,
    pivots,
    false,
    bundle.pivotsMeta,
    "cohort_first",
    true, // enableColumnGroups
    "all_expanded"
  );

  // Col 0: Operating Account
  // Col 1: Summary column for Q1 (2021)
  // Col 2..4: Jan (2021), Feb (2021), Mar (2021)
  assert.equal(expandedColumns.length, 5);
  assert.equal(expandedColumns[1].isColumnGroupSummary, true);
  assert.deepEqual(expandedColumns[1].pivotValues, ["Strategic Plan", "▾ Q1 (2021)", ""]);

  // Verify aggregated sum for Q1 (2021) summary column
  const q1Sum = getValue(bundle.rows[0], expandedColumns[1]);
  assert.ok(Math.abs(q1Sum - 450678234.92) < 1e-4, `Expected ~450678234.92, got ${q1Sum}`);

  // Verify columnTreeNodes metadata
  assert.equal(columnTreeNodes.length, 1);
  assert.equal(columnTreeNodes[0].label, "Q1 (2021)");
  assert.equal(columnTreeNodes[0].headerRowOffset, 1);
  assert.equal(columnTreeNodes[0].summaryColOffset, 1);
  assert.equal(columnTreeNodes[0].summaryColCount, 1);
  assert.equal(columnTreeNodes[0].childColStartOffset, 2);
  assert.equal(columnTreeNodes[0].childColEndOffset, 4);
  assert.equal(columnTreeNodes[0].isExpanded, true);

  // Build Crosstab Header Matrix and verify top-left labels + unmerged summary cell
  const matrix = buildCrosstabHeaderMatrix(
    expandedColumns,
    pivots,
    undefined,
    ["Version", "Quarter", "Time"]
  );
  assert.equal(matrix.headerRows[0][0], "Version");
  assert.equal(matrix.headerRows[1][0], "Quarter");
  assert.equal(matrix.headerRows[2][0], "Time");
  assert.equal(matrix.headerRows[1][1], "▾ Q1 (2021)");
  assert.equal(matrix.headerRows[2][1], "");
  assert.equal(matrix.headerRows[2][2], "Jan (2021)");
  assert.equal(matrix.headerRows[2][3], "Feb (2021)");
  assert.equal(matrix.headerRows[2][4], "Mar (2021)");

  // Verify computeColumnTreeHiddenFlags when expanded vs collapsed
  const expandedFlags = computeColumnTreeHiddenFlags(expandedColumns.length, columnTreeNodes);
  assert.deepEqual(expandedFlags, [false, false, false, false, false]);

  const collapsedNodes = columnTreeNodes.map((n) => ({ ...n, isExpanded: false }));
  const collapsedFlags = computeColumnTreeHiddenFlags(expandedColumns.length, collapsedNodes);
  // Summary column (index 1) stays visible; child columns (indices 2..4) are hidden!
  assert.deepEqual(collapsedFlags, [false, false, true, true, true]);
}

// 2. Additive measure summation vs. non-additive measure `null` safety & `all_collapsed` defaultColumnDisplayLevel
{
  assert.equal(isAdditiveMeasure({ fieldKey: "m1", label: "M1", type: "sum" }), true);
  assert.equal(isAdditiveMeasure({ fieldKey: "m2", label: "M2", type: "count" }), true);
  assert.equal(isAdditiveMeasure({ fieldKey: "m3", label: "M3", type: "sum_distinct" }), true);
  assert.equal(isAdditiveMeasure({ fieldKey: "m4", label: "M4", type: "average" }), false);
  assert.equal(isAdditiveMeasure({ fieldKey: "m5", label: "M5", type: "count_distinct" }), false);
  assert.equal(isAdditiveMeasure({ fieldKey: "m6", label: "M6", type: "number" }), false);

  const pivots = ["orders.quarter", "orders.month"];
  const mockJsonDetail = {
    pivots: [
      {
        key: "Q1|Jan",
        data: { "orders.quarter": "Q1", "orders.month": "Jan" },
      },
      {
        key: "Q1|Feb",
        data: { "orders.quarter": "Q1", "orders.month": "Feb" },
      },
    ],
    data: [
      {
        "account.name": { value: "Gross Margin" },
        "orders.amount": {
          "Q1|Jan": { value: 100 },
          "Q1|Feb": { value: 200 },
        },
        "orders.margin_pct": {
          "Q1|Jan": { value: 0.2 },
          "Q1|Feb": { value: 0.4 },
        },
      },
    ],
  };

  const bundle = normalizeQueryResult(mockJsonDetail, pivots, "json_detail");
  const { expandedColumns, getValue, columnTreeNodes } = expandPivotedColumns(
    mockColumns.filter((c) => c.fieldKey !== "orders.version"),
    bundle.rows,
    pivots,
    false,
    bundle.pivotsMeta,
    "cohort_first",
    true,
    "all_collapsed"
  );

  // Col 0: Operating Account
  // Col 1: Amount (▸ Q1)
  // Col 2: Margin % (▸ Q1)
  // Col 3..6: leaf columns
  assert.equal(expandedColumns.length, 7);
  assert.equal(columnTreeNodes.length, 1);
  assert.equal(columnTreeNodes[0].isExpanded, false);
  assert.equal(columnTreeNodes[0].summaryColOffset, 1);
  assert.equal(columnTreeNodes[0].summaryColCount, 2);
  assert.equal(columnTreeNodes[0].childColStartOffset, 3);
  assert.equal(columnTreeNodes[0].childColEndOffset, 6);
  assert.deepEqual(expandedColumns[1].pivotValues, ["▸ Q1", ""]);

  // Additive `sum` measure sums (100 + 200 = 300); non-additive `average` measure returns `null`
  assert.equal(getValue(bundle.rows[0], expandedColumns[1]), 300);
  assert.equal(getValue(bundle.rows[0], expandedColumns[2]), null);

  const hiddenFlags = computeColumnTreeHiddenFlags(expandedColumns.length, columnTreeNodes);
  assert.deepEqual(hiddenFlags, [false, false, false, true, true, true, true]);
}

console.log("All Feature 8 unit tests passed!");
