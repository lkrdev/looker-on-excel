import assert from "node:assert";
import {
  expandPivotedColumns,
  buildCrosstabHeaderMatrix,
  ColumnDefinition,
} from "../src/taskpane/services/excelWriter";
import { normalizeQueryResult } from "../src/taskpane/services/queryTaskRunner";

const mockColumns: ColumnDefinition[] = [
  { fieldKey: "users.country", label: "Country", category: "dimension" },
  { fieldKey: "orders.created_year", label: "Created Year", category: "dimension" },
  { fieldKey: "orders.created_quarter", label: "Created Quarter", category: "dimension" },
  {
    fieldKey: "orders.total_revenue",
    label: "Total Revenue",
    category: "measure",
    excelFormat: "$#,##0",
  },
  {
    fieldKey: "orders.gross_margin_pct",
    label: "Gross Margin %",
    category: "measure",
    excelFormat: "0.0%",
  },
];

const mockMultiPivotJsonDetail = {
  pivots: [
    {
      key: "2024|Q1",
      data: { "orders.created_year": "2024", "orders.created_quarter": "Q1" },
    },
    {
      key: "2024|Q2",
      data: { "orders.created_year": "2024", "orders.created_quarter": "Q2" },
    },
    {
      key: "2025|Q1",
      data: { "orders.created_year": "2025", "orders.created_quarter": "Q1" },
    },
  ],
  data: [
    {
      "users.country": { value: "USA" },
      "orders.total_revenue": {
        "2024|Q1": { value: 1000 },
        "2024|Q2": { value: 1200 },
        "2025|Q1": { value: 1500 },
        "$$$_row_total_$$$": { value: 3700 },
      },
      "orders.gross_margin_pct": {
        "2024|Q1": { value: 0.42 },
        "2024|Q2": { value: 0.45 },
        "2025|Q1": { value: 0.48 },
        "$$$_row_total_$$$": { value: 0.454 },
      },
    },
  ],
};

const pivots = ["orders.created_year", "orders.created_quarter"];
const bundle = normalizeQueryResult(mockMultiPivotJsonDetail, pivots, "json_detail");

// Test 1: "measure_first" grouping mode (default)
const measureFirst = expandPivotedColumns(
  mockColumns,
  bundle.rows,
  pivots,
  true,
  bundle.pivotsMeta,
  "measure_first"
);

assert.deepStrictEqual(
  measureFirst.expandedColumns.map((c) => c.label),
  [
    "Country",
    "Total Revenue (2024 · Q1)",
    "Total Revenue (2024 · Q2)",
    "Total Revenue (2025 · Q1)",
    "Total Revenue (Row Total)",
    "Gross Margin % (2024 · Q1)",
    "Gross Margin % (2024 · Q2)",
    "Gross Margin % (2025 · Q1)",
    "Gross Margin % (Row Total)",
  ]
);

assert.strictEqual(
  measureFirst.getValue(bundle.rows[0], measureFirst.expandedColumns[1]),
  1000
);
assert.strictEqual(
  measureFirst.getValue(bundle.rows[0], measureFirst.expandedColumns[4]),
  3700
);
assert.strictEqual(
  measureFirst.getValue(bundle.rows[0], measureFirst.expandedColumns[8]),
  0.454
);

// Test 2: "cohort_first" grouping mode (Looker Explore Crosstab layout)
const cohortFirst = expandPivotedColumns(
  mockColumns,
  bundle.rows,
  pivots,
  true,
  bundle.pivotsMeta,
  "cohort_first"
);

assert.deepStrictEqual(
  cohortFirst.expandedColumns.map((c) => c.label),
  [
    "Country",
    "Total Revenue (2024 · Q1)",
    "Gross Margin % (2024 · Q1)",
    "Total Revenue (2024 · Q2)",
    "Gross Margin % (2024 · Q2)",
    "Total Revenue (2025 · Q1)",
    "Gross Margin % (2025 · Q1)",
    "Total Revenue (Row Total)",
    "Gross Margin % (Row Total)",
  ]
);

// Test 3: buildCrosstabHeaderMatrix generates 3 header rows (Year, Quarter, Measure) and merges contiguous cohorts
const matrix = buildCrosstabHeaderMatrix(cohortFirst.expandedColumns, pivots);
assert.strictEqual(matrix.headerRows.length, 3, "2 pivot dims + 1 measure row = 3 header rows");

// Row 0: Year cohort header row
assert.deepStrictEqual(matrix.headerRows[0], [
  "",
  "2024",
  "2024",
  "2024",
  "2024",
  "2025",
  "2025",
  "Row Total",
  "Row Total",
]);

// Row 1: Quarter cohort header row
assert.deepStrictEqual(matrix.headerRows[1], [
  "",
  "Q1",
  "Q1",
  "Q2",
  "Q2",
  "Q1",
  "Q1",
  "",
  "",
]);

// Row 2: Bottom Measure / Dimension row
assert.deepStrictEqual(matrix.headerRows[2], [
  "Country",
  "Total Revenue",
  "Gross Margin %",
  "Total Revenue",
  "Gross Margin %",
  "Total Revenue",
  "Gross Margin %",
  "Total Revenue (Row Total)",
  "Gross Margin % (Row Total)",
]);

// Verify horizontal merges:
// On Row 0: "2024" spans cols 1..4 (colCount 4), "2025" spans cols 5..6 (colCount 2), "Row Total" spans cols 7..8 (colCount 2)
// On Row 1: "Q1" under 2024 spans cols 1..2 (colCount 2), "Q2" under 2024 spans cols 3..4 (colCount 2), "Q1" under 2025 spans cols 5..6 (colCount 2)
assert.deepStrictEqual(matrix.merges, [
  { rowOffset: 0, colOffset: 1, rowCount: 1, colCount: 4 },
  { rowOffset: 0, colOffset: 5, rowCount: 1, colCount: 2 },
  { rowOffset: 0, colOffset: 7, rowCount: 1, colCount: 2 },
  { rowOffset: 1, colOffset: 1, rowCount: 1, colCount: 2 },
  { rowOffset: 1, colOffset: 3, rowCount: 1, colCount: 2 },
  { rowOffset: 1, colOffset: 5, rowCount: 1, colCount: 2 },
]);

console.log("All crosstabMultiPivot tests passed!");
