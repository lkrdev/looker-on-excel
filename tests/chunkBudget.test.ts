import { calculateBatchSize } from "../src/taskpane/services/excelWriter";

function assertEqual(actual: any, expected: any, testName: string) {
  if (actual !== expected) {
    throw new Error(`[FAIL] ${testName}: Expected ${expected}, got ${actual}`);
  }
  console.log(`[PASS] ${testName}: batchSize=${actual}`);
}

export function runChunkBudgetTests() {
  console.log("=== Running Chunk Budget Calculator Tests ===");

  // 1. 5 columns -> 35000 / 5 = 7000 rows
  assertEqual(calculateBatchSize(5), 7000, "5 columns scales to 7,000 rows per batch");

  // 2. 30 columns -> floor(35000 / 30) = 1166 rows
  assertEqual(calculateBatchSize(30), 1166, "30 columns scales to 1,166 rows per batch (~700KB payload)");

  // 3. 70 columns -> floor(35000 / 70) = 500 rows (minimum floor)
  assertEqual(calculateBatchSize(70), 500, "70 columns hits 500 rows minimum floor");

  // 4. 100 columns -> floor(35000 / 100) = 350 -> clamped to 500
  assertEqual(calculateBatchSize(100), 500, "100 columns clamps to 500 rows minimum floor");

  console.log("=== All Chunk Budget Tests Passed! ===\n");
}

if (require.main === module) {
  runChunkBudgetTests();
}
