import { ParameterDefinition, ExploreDetail } from "../src/taskpane/services/lookerClient";

function assertEqual(actual: any, expected: any, testName: string) {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`[FAIL] ${testName}: Expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
  }
  console.log(`[PASS] ${testName}:`, actual);
}

/**
 * Pure helper mirroring the query filter compilation in App.tsx
 */
export function compileFiltersPayload(
  promptValues: Record<string, string>,
  alwaysFilters?: Array<{ field: string; values: string[] }>,
  userFilters: Array<{ field: string; expr: string | null }> = []
): Record<string, string> | null {
  const compiled: Record<string, string> = {};

  // 1. Only include non-empty prompt / parameter values
  Object.entries(promptValues).forEach(([k, v]) => {
    if (v !== undefined && v !== null && v.trim() !== "") {
      compiled[k] = v.trim();
    }
  });

  // 2. Ensure always_filter default values are applied if not explicitly in promptValues
  if (alwaysFilters) {
    alwaysFilters.forEach((af) => {
      if (!compiled[af.field] && af.values && af.values.length > 0) {
        compiled[af.field] = af.values.join(", ");
      }
    });
  }

  // 3. User filters (combine multiple filters on the same field with ', ')
  userFilters.forEach((f) => {
    if (f.field && f.expr !== null && f.expr !== "") {
      compiled[f.field] = compiled[f.field]
        ? `${compiled[f.field]}, ${f.expr}`
        : f.expr;
    }
  });

  return Object.keys(compiled).length > 0 ? compiled : null;
}

export function runPromptsTests() {
  console.log("=== Running Prompts and Parameter Compilation Tests ===");

  // Case 1: Empty parameter input does NOT inject filter
  const emptyParamResult = compileFiltersPayload(
    { "order_items.kpi_selector": "" },
    undefined,
    []
  );
  assertEqual(emptyParamResult, null, "Empty parameter does not emit filter");

  // Case 2: Whitespace-only parameter input does NOT inject filter
  const whitespaceParamResult = compileFiltersPayload(
    { "order_items.kpi_selector": "   " },
    undefined,
    []
  );
  assertEqual(whitespaceParamResult, null, "Whitespace parameter does not emit filter");

  // Case 3: Filled parameter input emits clean filter
  const filledParamResult = compileFiltersPayload(
    { "order_items.kpi_selector": "gross_margin" },
    undefined,
    []
  );
  assertEqual(filledParamResult, { "order_items.kpi_selector": "gross_margin" }, "Explicit parameter value emits filter");

  // Case 4: always_filter defaults applied when promptValues empty
  const alwaysFilterResult = compileFiltersPayload(
    {},
    [{ field: "users.country", values: ["USA", "Canada"] }],
    []
  );
  assertEqual(alwaysFilterResult, { "users.country": "USA, Canada" }, "always_filter default applied");

  // Case 5: always_filter overridden by user
  const alwaysFilterOverridden = compileFiltersPayload(
    { "users.country": "UK" },
    [{ field: "users.country", values: ["USA", "Canada"] }],
    []
  );
  assertEqual(alwaysFilterOverridden, { "users.country": "UK" }, "always_filter overridden by user prompt");

  // Case 6: Combined parameters, always_filter, and user filters
  const combined = compileFiltersPayload(
    { "order_items.kpi_selector": "revenue" },
    [{ field: "users.country", values: ["USA"] }],
    [
      { field: "orders.status", expr: "-cancelled" },
      { field: "orders.status", expr: "-returned" },
    ]
  );
  assertEqual(
    combined,
    {
      "order_items.kpi_selector": "revenue",
      "users.country": "USA",
      "orders.status": "-cancelled, -returned",
    },
    "Combined parameters and multiple filters on the same field compile properly"
  );

  console.log("=== All Prompts and Parameter Tests Passed! ===");
}

runPromptsTests();
