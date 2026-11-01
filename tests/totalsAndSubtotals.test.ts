import {
  resolveResultFormatStrategy,
  normalizeQueryResult,
  QueryPayload,
} from "../src/taskpane/services/queryTaskRunner";
import {
  buildWorksheetRowMatrix,
  expandPivotedColumns,
  ColumnDefinition,
} from "../src/taskpane/services/excelWriter";

function assertEqual(actual: any, expected: any, testName: string) {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(
      `[FAIL] ${testName}: Expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`
    );
  }
  console.log(`[PASS] ${testName}`);
}

export function runTotalsAndSubtotalsTests() {
  console.log("=== Running Looker Totals, Subtotals & Conditional json_detail Tests ===");

  // 1. Conditional ResultFormat Strategy
  const flatSmall: QueryPayload = {
    model: "thelook",
    view: "orders",
    fields: ["orders.region", "orders.unique_users"],
    limit: "500",
  };
  assertEqual(
    resolveResultFormatStrategy(flatSmall),
    "json",
    "Standard query without totals/subtotals uses lightweight json"
  );

  const totalSmall: QueryPayload = {
    ...flatSmall,
    total: true,
    limit: "500",
  };
  assertEqual(
    resolveResultFormatStrategy(totalSmall),
    "json_detail",
    "Small query (<5000 rows) with total=true uses single json_detail request"
  );

  const totalHighVolume: QueryPayload = {
    ...flatSmall,
    total: true,
    limit: "10000",
  };
  assertEqual(
    resolveResultFormatStrategy(totalHighVolume),
    "split_json_and_total",
    "High-volume query (5000+ rows) with total=true uses split_json_and_total to avoid payload bloat"
  );

  const totalUnlimited: QueryPayload = {
    ...flatSmall,
    total: true,
    limit: "-1",
  };
  assertEqual(
    resolveResultFormatStrategy(totalUnlimited),
    "split_json_and_total",
    "Unlimited (-1) query with total=true uses split_json_and_total"
  );

  const subtotalsQuery: QueryPayload = {
    ...flatSmall,
    subtotals: ["orders.region"],
    limit: "10000",
  };
  assertEqual(
    resolveResultFormatStrategy(subtotalsQuery),
    "json_detail",
    "Query with subtotals uses json_detail"
  );

  // 2. Normalize json_detail with non-additive measures, pivots, row_total, totals_data, and subtotals_data
  const rawJsonDetail = {
    pivots: [
      { key: "2024", is_total: false, data: { "orders.created_year": "2024" } },
      { key: "2025", is_total: false, data: { "orders.created_year": "2025" } },
      { key: "$$$_row_total_$$$", is_total: true, label: "Row Total" },
    ],
    data: [
      {
        "orders.region": { value: "West" },
        "orders.category": { value: "Jeans" },
        "orders.unique_users": {
          "2024": { value: 100 },
          "2025": { value: 120 },
          "$$$_row_total_$$$": { value: 180 }, // Non-additive row total (180 != 100+120)
        },
      },
      {
        "orders.region": { value: "West" },
        "orders.category": { value: "Swim" },
        "orders.unique_users": {
          "2024": { value: 50 },
          "2025": { value: 60 },
          "$$$_row_total_$$$": { value: 90 },
        },
      },
      {
        "orders.region": { value: "East" },
        "orders.category": { value: "Jeans" },
        "orders.unique_users": {
          "2024": { value: 70 },
          "2025": { value: 80 },
          "$$$_row_total_$$$": { value: 130 },
        },
      },
    ],
    subtotals_data: {
      "1": [
        {
          "orders.region": { value: "West" },
          "orders.category": { value: "$$$_subtotal_$$$" },
          "orders.unique_users": {
            "2024": { value: 135 }, // Non-additive subtotal (135 != 100+50)
            "2025": { value: 160 },
            "$$$_row_total_$$$": { value: 240 },
          },
        },
        {
          "orders.region": { value: "East" },
          "orders.category": { value: "$$$_subtotal_$$$" },
          "orders.unique_users": {
            "2024": { value: 70 },
            "2025": { value: 80 },
            "$$$_row_total_$$$": { value: 130 },
          },
        },
      ],
    },
    totals_data: {
      "orders.unique_users": {
        "2024": { value: 195 }, // Non-additive grand total (195 != 135+70)
        "2025": { value: 225 },
        "$$$_row_total_$$$": { value: 350 },
      },
    },
  };

  const bundle = normalizeQueryResult(rawJsonDetail, ["orders.created_year"], "json_detail");
  assertEqual(bundle.rows.length, 3, "Unwrapped 3 data rows from json_detail");
  assertEqual(
    bundle.rows[0]["orders.unique_users"]["$$$_row_total_$$$"],
    180,
    "Unwrapped non-additive row total (180)"
  );

  // 3. Interleave Subtotals and Grand Total in Worksheet Row Matrix
  const cols: ColumnDefinition[] = [
    { fieldKey: "orders.region", label: "Region", category: "dimension" },
    { fieldKey: "orders.category", label: "Category", category: "dimension" },
    { fieldKey: "orders.created_year", label: "Created Year", category: "dimension" },
    { fieldKey: "orders.unique_users", label: "Unique Users", category: "measure" },
  ];

  const matrix = buildWorksheetRowMatrix(
    cols,
    bundle.rows,
    bundle.subtotalsData,
    bundle.totalsData,
    ["orders.created_year"]
  );

  // 3 data rows + 2 subtotal rows + 1 grand total row = 6 rows
  assertEqual(matrix.rows.length, 6, "Matrix interleaves 2 subtotals and 1 grand total (6 rows)");
  assertEqual(matrix.subtotalRowIndices, [2, 4], "Subtotal rows placed at indices 2 and 4");
  assertEqual(matrix.grandTotalRowIndex, 5, "Grand total placed at index 5");
  assertEqual(
    matrix.outlineGroups,
    [
      { startRowOffset: 0, endRowOffset: 1 },
      { startRowOffset: 3, endRowOffset: 3 },
    ],
    "Outline groups cover West (0..1) and East (3..3)"
  );

  // 4. Expand Pivoted Columns with Row Totals
  const expanded = expandPivotedColumns(cols, matrix.rows, ["orders.created_year"], true);
  // 2 non-pivoted dims + 2 pivot years + 1 Row Total = 5 columns
  assertEqual(expanded.expandedColumns.length, 5, "Expands 2 dims + 2 pivot cols + 1 Row Total");
  assertEqual(
    expanded.expandedColumns[4].label,
    "Unique Users (Row Total)",
    "5th column is Row Total"
  );
  assertEqual(
    expanded.getValue(matrix.rows[2], expanded.expandedColumns[0]),
    "Subtotal (West)",
    "Subtotal label for West"
  );
  assertEqual(
    expanded.getValue(matrix.rows[2], expanded.expandedColumns[2]),
    135,
    "Subtotal West 2024 uses Looker non-additive SQL value (135)"
  );
  assertEqual(
    expanded.getValue(matrix.rows[5], expanded.expandedColumns[4]),
    350,
    "Grand Total Row Total uses Looker non-additive SQL value (350)"
  );

  // 5. Non-pivoted query with Column Grand Totals (`has_custom_html`, `html`, `links` in totals_data)
  const rawNonPivotedTotals = {
    data: [
      {
        "basic_products.department": { value: "Men" },
        "basic_users.unique_users": {
          value: 40000,
          links: [{ label: "Show All 40,000", url: "/explore/basic_ecomm" }],
        },
        "basic_users.sum_distinct_age": {
          value: 1520000,
          links: [{ label: "Show All", url: "/explore/basic_ecomm" }],
        },
      },
      {
        "basic_products.department": { value: "Women" },
        "basic_users.unique_users": {
          value: 42000,
          links: [{ label: "Show All 42,000", url: "/explore/basic_ecomm" }],
        },
        "basic_users.sum_distinct_age": {
          value: 1610000,
          links: [{ label: "Show All", url: "/explore/basic_ecomm" }],
        },
      },
    ],
    totals_data: {
      "basic_products.department": {},
      "basic_users.unique_users": {
        value: 79748,
        has_custom_html: false,
        html: "79,748",
        links: [{ label: "Show All 79,748", url: "/explore/basic_ecomm" }],
      },
      "basic_users.sum_distinct_age": {
        value: 3010500,
        has_custom_html: false,
        html: "3,010,500",
        links: [{ label: "Show All", url: "/explore/basic_ecomm" }],
      },
    },
  };

  const nonPivotedBundle = normalizeQueryResult(rawNonPivotedTotals, null, "json_detail");
  const nonPivotCols: ColumnDefinition[] = [
    { fieldKey: "basic_products.department", label: "Basic Products Department", category: "dimension" },
    { fieldKey: "basic_users.unique_users", label: "Basic Users Unique Users", category: "measure" },
    { fieldKey: "basic_users.sum_distinct_age", label: "Basic Users Sum Distinct Age", category: "measure" },
  ];

  const nonPivotMatrix = buildWorksheetRowMatrix(
    nonPivotCols,
    nonPivotedBundle.rows,
    nonPivotedBundle.subtotalsData,
    nonPivotedBundle.totalsData,
    undefined
  );
  const nonPivotExpanded = expandPivotedColumns(nonPivotCols, nonPivotMatrix.rows, undefined, false);

  assertEqual(
    nonPivotExpanded.expandedColumns.map((c) => c.label),
    [
      "Basic Products Department",
      "Basic Users Unique Users",
      "Basic Users Sum Distinct Age",
    ],
    "Non-pivoted totals_data with has_custom_html/html/links never expands into fake (value)/(has_custom_html)/(html) columns"
  );
  assertEqual(
    nonPivotExpanded.getValue(nonPivotMatrix.rows[0], nonPivotExpanded.expandedColumns[1]),
    40000,
    "Row 0 Men Unique Users scalar value preserved (40000)"
  );
  assertEqual(
    nonPivotExpanded.getValue(nonPivotMatrix.rows[2], nonPivotExpanded.expandedColumns[0]),
    "Grand Total",
    "Row 2 Grand Total label"
  );
  assertEqual(
    nonPivotExpanded.getValue(nonPivotMatrix.rows[2], nonPivotExpanded.expandedColumns[1]),
    79748,
    "Row 2 Grand Total Unique Users scalar value (79748)"
  );
  assertEqual(
    nonPivotExpanded.getValue(nonPivotMatrix.rows[2], nonPivotExpanded.expandedColumns[2]),
    3010500,
    "Row 2 Grand Total Sum Distinct Age scalar value (3010500)"
  );

  console.log("=== All Totals, Subtotals & Conditional json_detail Tests Passed! ===");
}

runTotalsAndSubtotalsTests();
