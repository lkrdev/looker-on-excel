import {
  convertWriteQueryToPayload,
  extractTileQueryAndListeners,
  applyDashboardFiltersToTilePayload,
  orderFieldsDimensionsFirst,
  DashboardElementDef,
  FieldDefinition,
} from "../src/taskpane/services/lookerClient";
import { resolveCleanPivotValues, expandPivotedColumns } from "../src/taskpane/services/excelWriter";
import { normalizeQueryResult } from "../src/taskpane/services/queryTaskRunner";

function assertEqual(actual: any, expected: any, testName: string) {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(
      `[FAIL] ${testName}: Expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`
    );
  }
  console.log(`[PASS] ${testName}`);
}

export function runLooksAndDashboardsTests() {
  console.log("=== Running Saved Looks & Dashboards Filter Mapping Tests ===");

  // 1. Convert LookerWriteQuery from a Saved Look into QueryPayload
  const lookPayload = convertWriteQueryToPayload({
    model: "thelook",
    view: "order_items",
    fields: ["orders.region", "order_items.total_revenue"],
    pivots: ["orders.created_year"],
    filters: { "orders.status": "Complete" },
    sorts: ["order_items.total_revenue desc"],
    limit: "1000",
    total: true,
    row_total: "right",
    subtotals: ["orders.region"],
  });

  assertEqual(lookPayload.model, "thelook", "Look query preserves model");
  assertEqual(lookPayload.total, true, "Look query preserves total=true");
  assertEqual(lookPayload.row_total, "right", "Look query preserves row_total='right'");
  assertEqual(lookPayload.subtotals, ["orders.region"], "Look query preserves subtotals");

  // 2. Extract Query & Filter Listeners from Dashboard Tile (`result_maker.query` + `filterables.listen`)
  const tileElement: DashboardElementDef = {
    id: "tile_101",
    title: "Revenue by Category",
    type: "vis",
    result_maker: {
      query: {
        model: "thelook",
        view: "order_items",
        fields: ["products.category", "order_items.total_revenue"],
        filters: { "orders.status": "Complete" },
        limit: "500",
      },
      filterables: [
        {
          listen: [
            { dashboard_filter_name: "Date Range", field: "orders.created_date" },
            { dashboard_filter_name: "Region", field: "users.state" },
          ],
        },
      ],
    },
  };

  const { query, listeners } = extractTileQueryAndListeners(tileElement);
  assertEqual(Boolean(query), true, "Extracts query from result_maker.query");
  assertEqual(listeners.length, 2, "Extracts 2 dashboard filter listeners");

  const baseTilePayload = convertWriteQueryToPayload(query!);
  const filteredPayload = applyDashboardFiltersToTilePayload(baseTilePayload, listeners, {
    "Date Range": "last 90 days",
    Region: "California, New York",
  });

  assertEqual(
    filteredPayload.filters,
    {
      "orders.status": "Complete",
      "orders.created_date": "last 90 days",
      "users.state": "California, New York",
    },
    "Dashboard filter values map onto tile field filters while preserving base tile filters"
  );

  // 3. Verify `orderFieldsDimensionsFirst` places non-pivoted dimensions -> pivoted dimensions -> measures
  const catalogFields: FieldDefinition[] = [
    {
      name: "orders.region",
      label: "Region",
      label_short: "Region",
      view_label: "Orders",
      type: "string",
      category: "dimension",
    },
    {
      name: "users.country",
      label: "Country",
      label_short: "Country",
      view_label: "Users",
      type: "string",
      category: "dimension",
    },
    {
      name: "orders.created_year",
      label: "Created Year",
      label_short: "Year",
      view_label: "Orders",
      type: "date_year",
      category: "dimension",
    },
    {
      name: "order_items.total_revenue",
      label: "Total Revenue",
      label_short: "Total Revenue",
      view_label: "Order Items",
      type: "sum",
      category: "measure",
    },
    {
      name: "orders.count",
      label: "Order Count",
      label_short: "Count",
      view_label: "Orders",
      type: "count",
      category: "measure",
    },
  ];

  // Simulate Looker API `query.fields` where user clicked a measure first, then dimensions & pivots
  const rawInterleavedFields = [
    "order_items.total_revenue",
    "orders.region",
    "orders.created_year",
    "orders.count",
    "users.country",
  ];
  const ordered = orderFieldsDimensionsFirst(rawInterleavedFields, catalogFields, [
    "orders.created_year",
  ]);
  assertEqual(
    ordered,
    [
      "orders.region",
      "users.country",
      "orders.created_year",
      "order_items.total_revenue",
      "orders.count",
    ],
    "orderFieldsDimensionsFirst partitions non-pivoted dimensions first, then pivoted dimensions, then measures while preserving relative order within each group"
  );

  // 4. Verify multi-pivot LookML `case:` dimension keys (`1 - Processing|FIELD|0|FIELD|Processing`)
  // are resolved to clean human-readable labels via `pivotsMeta` (and fallback `|FIELD|` stripping)
  const pivots = ["order_items.status_in_order", "order_items.status"];
  const rawJsonDetail = {
    pivots: [
      {
        key: "1 - Processing|FIELD|0|FIELD|Processing",
        data: { "order_items.status_in_order": "1 - Processing", "order_items.status": "Processing" },
        labels: { "order_items.status_in_order": "1 - Processing", "order_items.status": "Processing" },
        is_total: false,
      },
      {
        key: "1x - Cancelled|FIELD|1|FIELD|Cancelled",
        data: { "order_items.status_in_order": "1x - Cancelled", "order_items.status": "Cancelled" },
        labels: { "order_items.status_in_order": "1x - Cancelled", "order_items.status": "Cancelled" },
        is_total: false,
      },
    ],
    data: [
      {
        "order_items.validation_status": { value: "Valid" },
        "order_items.count": {
          "1 - Processing|FIELD|0|FIELD|Processing": { value: 9132 },
          "1x - Cancelled|FIELD|1|FIELD|Cancelled": { value: null },
        },
      },
    ],
  };

  const bundle = normalizeQueryResult(rawJsonDetail, pivots, "json_detail");
  const { expandedColumns } = expandPivotedColumns(
    [
      { fieldKey: "order_items.validation_status", label: "Order Items Validation Status", category: "dimension" },
      { fieldKey: "order_items.status_in_order", label: "Order Items Status In Order", category: "dimension" },
      { fieldKey: "order_items.status", label: "Order Items Status", category: "dimension" },
      { fieldKey: "order_items.count", label: "Order Items # of Order Items", category: "measure" },
    ],
    bundle.rows,
    pivots,
    false,
    bundle.pivotsMeta
  );

  assertEqual(
    expandedColumns.map((c) => c.label),
    [
      "Order Items Validation Status",
      "Order Items # of Order Items (1 - Processing · Processing)",
      "Order Items # of Order Items (1x - Cancelled · Cancelled)",
    ],
    "Multi-pivot LookML case dimension keys resolve to clean human-readable labels without |FIELD|0|FIELD|"
  );

  // Verify fallback without pivotsMeta also strips |FIELD| and injected numeric sort token
  assertEqual(
    resolveCleanPivotValues("1 - Processing|FIELD|0|FIELD|Processing", pivots, null),
    ["1 - Processing", "Processing"],
    "Fallback resolveCleanPivotValues strips |FIELD| and numeric case sort index"
  );

  console.log("=== All Saved Looks & Dashboards Tests Passed! ===");
}

runLooksAndDashboardsTests();
