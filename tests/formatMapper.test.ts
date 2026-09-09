import { resolveExcelNumberFormat, LOOKER_EXCEL_FORMAT_MAP } from "../src/taskpane/services/formatMapper";

function assertEqual(actual: any, expected: any, testName: string) {
  if (actual !== expected) {
    throw new Error(`[FAIL] ${testName}: Expected "${expected}", got "${actual}"`);
  }
  console.log(`[PASS] ${testName}: ${actual}`);
}

export function runFormatMapperTests() {
  console.log("=== Running formatMapper Tests ===");

  // 1. Explicit value_format
  assertEqual(
    resolveExcelNumberFormat({ value_format: "$#,##0.00" }),
    "$#,##0.00",
    "Explicit LookML value_format passes through"
  );
  assertEqual(
    resolveExcelNumberFormat({ value_format: "00000" }),
    "00000",
    "LookML ID format 00000 passes through"
  );

  // 2. Named format translations
  assertEqual(
    resolveExcelNumberFormat({ value_format_name: "usd" }),
    LOOKER_EXCEL_FORMAT_MAP.usd,
    "Named format usd"
  );
  assertEqual(
    resolveExcelNumberFormat({ value_format_name: "percent_2" }),
    "0.00%",
    "Named format percent_2"
  );
  assertEqual(
    resolveExcelNumberFormat({ value_format_name: "decimal_2" }),
    "#,##0.00",
    "Named format decimal_2"
  );

  // 3. Fallbacks by field type
  assertEqual(
    resolveExcelNumberFormat({ type: "date_date" }),
    "yyyy-mm-dd",
    "Type date_date fallback"
  );
  assertEqual(
    resolveExcelNumberFormat({ type: "date_time" }),
    "yyyy-mm-dd hh:mm:ss",
    "Type date_time fallback"
  );
  assertEqual(
    resolveExcelNumberFormat({ type: "count" }),
    "#,##0",
    "Type count fallback"
  );
  assertEqual(
    resolveExcelNumberFormat({ type: "sum" }),
    "#,##0.00",
    "Type sum fallback"
  );
  assertEqual(
    resolveExcelNumberFormat({ type: "string" }),
    undefined,
    "Type string returns undefined"
  );

  console.log("=== All formatMapper Tests Passed! ===\n");
}

if (require.main === module) {
  runFormatMapperTests();
}
