import { expandPivotedColumns, computeResidualRanges, ColumnDefinition } from "../src/taskpane/services/excelWriter";

function assertEqual(actual: any, expected: any, testName: string) {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`[FAIL] ${testName}: Expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
  }
  console.log(`[PASS] ${testName}`);
}

export function runPivotTests() {
  console.log("=== Running Pivot & Crosstab Expansion Tests ===");

  // 1. Flat Data (No pivots)
  const flatColumns: ColumnDefinition[] = [
    { fieldKey: "products.category", label: "Category", type: "string" },
    { fieldKey: "order_items.total_sale_price", label: "Total Sale Price", type: "number", excelFormat: "$#,##0.00" },
  ];

  const flatData = [
    { "products.category": "Jeans", "order_items.total_sale_price": 500.5 },
    { "products.category": "Swim", "order_items.total_sale_price": 120.0 },
  ];

  const flatResult = expandPivotedColumns(flatColumns, flatData);
  assertEqual(flatResult.expandedColumns.length, 2, "Flat data keeps column count");
  assertEqual(flatResult.getValue(flatData[0], flatResult.expandedColumns[0]), "Jeans", "Flat data dim value");
  assertEqual(flatResult.getValue(flatData[0], flatResult.expandedColumns[1]), 500.5, "Flat data measure value");

  // 2. Pivoted Data from Looker
  // In Looker API, pivoted measures are nested: { [pivotField]: { [val1]: num1, [val2]: num2 } }
  const pivotedColumns: ColumnDefinition[] = [
    { fieldKey: "products.category", label: "Category", type: "string" },
    { fieldKey: "order_items.total_sale_price", label: "Total Sale Price", type: "number", excelFormat: "$#,##0.00" },
  ];

  const pivotedData = [
    {
      "products.category": "Jeans",
      "order_items.total_sale_price": {
        "order_items.status": {
          "Complete": 315725.22,
          "Cancelled": 185741.06,
          "Shipped": 370546.01,
        },
      },
    },
    {
      "products.category": "Sweaters",
      "order_items.total_sale_price": {
        "order_items.status": {
          "Complete": 213184.51,
          "Cancelled": 121839.50,
          "Shipped": 235471.79,
        },
      },
    },
  ];

  const pivotResult = expandPivotedColumns(pivotedColumns, pivotedData);

  // Expected columns: 1 dim + 3 pivoted status columns = 4 columns
  assertEqual(pivotResult.expandedColumns.length, 4, "Pivoted data expands to 4 columns");
  assertEqual(pivotResult.expandedColumns[0].label, "Category", "First column is Category");
  assertEqual(pivotResult.expandedColumns[1].label, "Total Sale Price (Complete)", "Second column is Complete");
  assertEqual(pivotResult.expandedColumns[2].label, "Total Sale Price (Cancelled)", "Third column is Cancelled");
  assertEqual(pivotResult.expandedColumns[3].label, "Total Sale Price (Shipped)", "Fourth column is Shipped");

  // Number format and type should be preserved on expanded columns
  assertEqual(pivotResult.expandedColumns[1].excelFormat, "$#,##0.00", "Expanded column preserves excelFormat");
  assertEqual(pivotResult.expandedColumns[1].type, "number", "Expanded column preserves type");

  // Values extraction
  assertEqual(pivotResult.getValue(pivotedData[0], pivotResult.expandedColumns[0]), "Jeans", "Row 0 Category value");
  assertEqual(pivotResult.getValue(pivotedData[0], pivotResult.expandedColumns[1]), 315725.22, "Row 0 Complete sale price");
  assertEqual(pivotResult.getValue(pivotedData[0], pivotResult.expandedColumns[2]), 185741.06, "Row 0 Cancelled sale price");
  assertEqual(pivotResult.getValue(pivotedData[1], pivotResult.expandedColumns[3]), 235471.79, "Row 1 Shipped sale price");

  // 3. Residual Range Calculation (Overwrite & Cleaning safety)
  const shrinkRows = computeResidualRanges(500, 5, 200, 5);
  assertEqual(shrinkRows.extraColRange, null, "No extra columns when col count unchanged");
  assertEqual(
    shrinkRows.extraRowRange,
    { startRow: 200, startCol: 0, rowCount: 300, colCount: 5 },
    "Extra rows calculated properly when row count shrinks"
  );

  const shrinkCols = computeResidualRanges(200, 10, 200, 4);
  assertEqual(shrinkCols.extraRowRange, null, "No extra rows when row count unchanged");
  assertEqual(
    shrinkCols.extraColRange,
    { startRow: 0, startCol: 4, rowCount: 200, colCount: 6 },
    "Extra columns calculated properly when col count shrinks"
  );

  const growBoth = computeResidualRanges(100, 4, 300, 8);
  assertEqual(growBoth.extraRowRange, null, "No extra rows when dataset expands");
  assertEqual(growBoth.extraColRange, null, "No extra cols when dataset expands");

  console.log("=== All Pivot & Residual Cleaning Tests Passed! ===");
}

runPivotTests();
