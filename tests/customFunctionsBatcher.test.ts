import {
  LookerValueBatcher,
  parseFormulaFilters,
  lookerQuery,
  clearCustomFunctionCache,
} from "../src/functions/functions";
import { QueryPayload } from "../src/taskpane/services/queryTaskRunner";

function assertEqual(actual: any, expected: any, testName: string) {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(
      `[FAIL] ${testName}: Expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`
    );
  }
  console.log(`[PASS] ${testName}`);
}

export async function runCustomFunctionsTests() {
  console.log("=== Running Custom Functions (=LOOKER.VALUE & =LOOKER.QUERY) Tests ===");
  clearCustomFunctionCache();

  // 1. Flexible Filter Parsing
  assertEqual(
    parseFormulaFilters(["orders.region", "West", "orders.year", "2025"]),
    { "orders.region": "West", "orders.year": "2025" },
    "Parses alternating field/value arguments"
  );
  assertEqual(
    parseFormulaFilters(["orders.region=East;orders.status=Complete"]),
    { "orders.region": "East", "orders.status": "Complete" },
    "Parses semicolon-separated key=value string"
  );
  assertEqual(
    parseFormulaFilters(['{"orders.region":"South","orders.year":"2024"}']),
    { "orders.region": "South", "orders.year": "2024" },
    "Parses JSON filter string"
  );
  assertEqual(
    parseFormulaFilters([
      [
        ["orders.region", "North"],
        ["orders.category", "Outerwear"],
      ],
    ]),
    { "orders.region": "North", "orders.category": "Outerwear" },
    "Parses 2D Excel cell range array"
  );

  // 2. Micro-batch Coalescing for =LOOKER.VALUE across 4 concurrent cells
  let executedPayloads: QueryPayload[] = [];
  const mockRunner = async (_baseUrl: string, _token: string, payload: QueryPayload) => {
    executedPayloads.push(payload);
    return [
      { "orders.region": "West", "orders.unique_customers": 420 },
      { "orders.region": "East", "orders.unique_customers": 310 },
      { "orders.region": "South", "orders.unique_customers": 195 },
      { "orders.region": "North", "orders.unique_customers": 280 },
    ];
  };

  const batcher = new LookerValueBatcher({
    windowMs: 25,
    runner: mockRunner,
    baseUrlProvider: async () => ({
      baseUrl: "https://looker.example.com",
      accessToken: "mock_token",
    }),
  });

  const [vWest, vEast, vSouth, vNorth] = await Promise.all([
    batcher.requestValue("thelook", "orders", "orders.unique_customers", {
      "orders.region": "West",
      "orders.created_year": "2025",
    }),
    batcher.requestValue("thelook", "orders", "orders.unique_customers", {
      "orders.region": "East",
      "orders.created_year": "2025",
    }),
    batcher.requestValue("thelook", "orders", "orders.unique_customers", {
      "orders.region": "South",
      "orders.created_year": "2025",
    }),
    batcher.requestValue("thelook", "orders", "orders.unique_customers", {
      "orders.region": "North",
      "orders.created_year": "2025",
    }),
  ]);

  assertEqual(executedPayloads.length, 1, "4 concurrent cell formulas coalesced into 1 Looker query");
  assertEqual(
    executedPayloads[0].filters,
    { "orders.created_year": "2025", "orders.region": "West,East,South,North" },
    "Coalesced query combines primary filter values and preserves static filters"
  );
  assertEqual([vWest, vEast, vSouth, vNorth], [420, 310, 195, 280], "Each cell resolves its slice");

  // 3. =LOOKER.QUERY Spill Matrix
  const spillMatrix = await lookerQuery(
    "thelook",
    "orders",
    "orders.region, orders.unique_customers",
    "orders.created_year=2025",
    "orders.unique_customers desc",
    10,
    mockRunner,
    async () => ({ baseUrl: "https://looker.example.com", accessToken: "mock_token" })
  );

  assertEqual(spillMatrix.length, 5, "Spill matrix has 1 header row + 4 data rows");
  assertEqual(
    spillMatrix[0],
    ["orders.region", "orders.unique_customers"],
    "Spill matrix header row"
  );
  assertEqual(spillMatrix[1], ["West", 420], "Spill matrix first data row");

  console.log("=== All Custom Functions Tests Passed! ===");
}

runCustomFunctionsTests();
