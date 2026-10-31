import {
  convertWriteQueryToPayload,
  extractTileQueryAndListeners,
  applyDashboardFiltersToTilePayload,
  DashboardElementDef,
} from "../src/taskpane/services/lookerClient";

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

  console.log("=== All Saved Looks & Dashboards Tests Passed! ===");
}

runLooksAndDashboardsTests();
