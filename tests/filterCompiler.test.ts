import {
  isDateField,
  isNumberField,
  isStringField,
  compileFilterToLookerExpression,
  FilterCondition,
} from "../src/taskpane/components/FilterBuilder";
import { FieldDefinition } from "../src/taskpane/services/lookerClient";

function assertEqual(actual: any, expected: any, testName: string) {
  if (actual !== expected) {
    throw new Error(`[FAIL] ${testName}: Expected "${expected}", got "${actual}"`);
  }
  console.log(`[PASS] ${testName}: ${actual}`);
}

export function runFilterCompilerTests() {
  console.log("=== Running Filter Compiler Tests ===");

  const stringField: FieldDefinition = { name: "users.status", label: "Status", type: "string" };
  const dateField: FieldDefinition = {
    name: "orders.created_date",
    label: "Created Date",
    type: "date_date",
    dimension_group: "orders.created",
  };
  const numberField: FieldDefinition = {
    name: "orders.total_amount",
    label: "Total Amount",
    type: "number",
    category: "measure",
  };

  // 1. Type detection
  assertEqual(isStringField(stringField), true, "stringField isString");
  assertEqual(isDateField(stringField), false, "stringField not date");
  assertEqual(isNumberField(stringField), false, "stringField not number");

  assertEqual(isDateField(dateField), true, "dateField isDate");
  assertEqual(isStringField(dateField), false, "dateField not string");
  assertEqual(isNumberField(dateField), false, "dateField not number");

  assertEqual(isNumberField(numberField), true, "numberField isNumber");
  assertEqual(isStringField(numberField), false, "numberField not string");
  assertEqual(isDateField(numberField), false, "numberField not date");

  // 2. String Filter Expressions
  assertEqual(
    compileFilterToLookerExpression({ id: "1", field: "users.status", operator: "is", value: "ACTIVE" }, stringField),
    "ACTIVE",
    "String is ACTIVE"
  );
  assertEqual(
    compileFilterToLookerExpression({ id: "2", field: "users.status", operator: "is_not", value: "INACTIVE" }, stringField),
    "-INACTIVE",
    "String is_not INACTIVE"
  );
  assertEqual(
    compileFilterToLookerExpression({ id: "3", field: "users.status", operator: "contains", value: "gold" }, stringField),
    "%gold%",
    "String contains gold"
  );
  assertEqual(
    compileFilterToLookerExpression({ id: "4", field: "users.status", operator: "is_null", value: "" }, stringField),
    "EMPTY",
    "String is_null EMPTY"
  );

  // 3. Number Filter Expressions
  assertEqual(
    compileFilterToLookerExpression({ id: "5", field: "orders.total_amount", operator: "is", value: "50" }, numberField),
    "50",
    "Number is 50"
  );
  assertEqual(
    compileFilterToLookerExpression({ id: "6", field: "orders.total_amount", operator: "greater_than", value: "100" }, numberField),
    ">100",
    "Number > 100"
  );
  assertEqual(
    compileFilterToLookerExpression({ id: "7", field: "orders.total_amount", operator: "is_between", value: "", startValue: "10", endValue: "50" }, numberField),
    "[10, 50]",
    "Number between 10 and 50"
  );
  assertEqual(
    compileFilterToLookerExpression({ id: "8", field: "orders.total_amount", operator: "is_null", value: "" }, numberField),
    "null",
    "Number is_null"
  );

  // 4. Date Filter Expressions
  assertEqual(
    compileFilterToLookerExpression({ id: "9", field: "orders.created_date", operator: "is_in_the_last", value: "7", unit: "days" }, dateField),
    "7 days",
    "Date is in the last 7 days"
  );
  assertEqual(
    compileFilterToLookerExpression({ id: "10", field: "orders.created_date", operator: "is_in_range", value: "", startValue: "2026-09-01", endValue: "2026-09-10" }, dateField),
    "2026-09-01 to 2026-09-10",
    "Date in range"
  );
  assertEqual(
    compileFilterToLookerExpression({ id: "11", field: "orders.created_date", operator: "is_on_the_day", value: "2026-09-10" }, dateField),
    "2026-09-10",
    "Date on the day"
  );
  assertEqual(
    compileFilterToLookerExpression({ id: "12", field: "orders.created_date", operator: "is_before", value: "2026-09-10" }, dateField),
    "before 2026-09-10",
    "Date before"
  );
  assertEqual(
    compileFilterToLookerExpression({ id: "13", field: "orders.created_date", operator: "is_this", value: "", unit: "month" }, dateField),
    "this month",
    "Date is this month"
  );
  assertEqual(
    compileFilterToLookerExpression({ id: "14", field: "orders.created_date", operator: "is_previous", value: "", unit: "year" }, dateField),
    "previous year",
    "Date is previous year"
  );
  assertEqual(
    compileFilterToLookerExpression({ id: "15", field: "orders.created_date", operator: "is_next", value: "2", unit: "weeks" }, dateField),
    "next 2 weeks",
    "Date is next 2 weeks"
  );
  assertEqual(
    compileFilterToLookerExpression({ id: "16", field: "orders.created_date", operator: "is_in_the_month", value: "2026-09" }, dateField),
    "2026/09",
    "Date is in the month"
  );
  assertEqual(
    compileFilterToLookerExpression({ id: "17", field: "orders.created_date", operator: "is_any_time", value: "" }, dateField),
    null,
    "Date is any time returns null (no filter)"
  );

  console.log("=== All Filter Compiler Tests Passed! ===\n");
}

if (require.main === module) {
  runFilterCompilerTests();
}
