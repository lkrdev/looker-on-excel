import assert from "assert";
import { FieldDefinition } from "../src/taskpane/services/lookerClient";
import { matchesFieldSearch, getMatchedSynonym } from "../src/taskpane/components/FieldPicker";

console.log("=== Running Synonyms & Field Search Tests ===");

const sampleFields: FieldDefinition[] = [
  {
    name: "orders.sale_price",
    label: "Sale Price",
    label_short: "Sale Price",
    view_label: "Orders",
    category: "measure",
    type: "sum",
    synonyms: ["revenue", "turnover", "gross sales", "top line"],
  },
  {
    name: "users.status",
    label: "User Status",
    view_label: "Users",
    category: "dimension",
    type: "string",
    synonyms: ["account state", "standing"],
  },
  {
    name: "orders.created_at_date",
    label: "Created Date",
    view_label: "Orders",
    category: "dimension",
    field_group_label: "Created Date",
    field_group_variant: "Date",
    synonyms: ["order date", "purchase date"],
  },
  {
    name: "inventory.sku",
    label: "Stock Keeping Unit",
    view_label: "Inventory",
    category: "dimension",
    type: "string",
  },
];

// Test 1: Direct name and label matches
assert.strictEqual(matchesFieldSearch(sampleFields[0], "sale"), true, "Matches by label substring");
assert.strictEqual(matchesFieldSearch(sampleFields[0], "orders.sale"), true, "Matches by name substring");
assert.strictEqual(getMatchedSynonym(sampleFields[0], "sale"), null, "Direct match does not flag as synonym-only");

// Test 2: Matches by Synonym
assert.strictEqual(matchesFieldSearch(sampleFields[0], "revenue"), true, "Matches by synonym revenue");
assert.strictEqual(getMatchedSynonym(sampleFields[0], "revenue"), "revenue", "Detects matched synonym revenue");
assert.strictEqual(matchesFieldSearch(sampleFields[0], "TURNOVER"), true, "Matches by case-insensitive synonym");
assert.strictEqual(getMatchedSynonym(sampleFields[0], "TURNOVER"), "turnover", "Detects matched synonym turnover");
assert.strictEqual(matchesFieldSearch(sampleFields[0], "gross"), true, "Matches by partial synonym gross");
assert.strictEqual(getMatchedSynonym(sampleFields[0], "gross"), "gross sales", "Detects partial matched synonym gross sales");

// Test 3: Synonyms on other fields
assert.strictEqual(matchesFieldSearch(sampleFields[1], "standing"), true, "Matches users.status by standing");
assert.strictEqual(getMatchedSynonym(sampleFields[1], "standing"), "standing", "Returns standing synonym");
assert.strictEqual(matchesFieldSearch(sampleFields[2], "purchase"), true, "Matches created_at by purchase date synonym");
assert.strictEqual(getMatchedSynonym(sampleFields[2], "purchase"), "purchase date", "Returns purchase date synonym");

// Test 4: Field without synonyms
assert.strictEqual(matchesFieldSearch(sampleFields[3], "revenue"), false, "Field without synonyms does not match unrelated term");
assert.strictEqual(getMatchedSynonym(sampleFields[3], "revenue"), null, "Field without synonyms returns null matched synonym");

// Test 5: Empty search string
assert.strictEqual(matchesFieldSearch(sampleFields[0], ""), true, "Empty search matches all");
assert.strictEqual(getMatchedSynonym(sampleFields[0], ""), null, "Empty search returns null matched synonym");

console.log("[PASS] Matches by label substring");
console.log("[PASS] Matches by name substring");
console.log("[PASS] Matches by case-insensitive synonym");
console.log("[PASS] Detects matched synonym when not a direct match");
console.log("[PASS] Field without synonyms behaves safely");
console.log("[PASS] Empty search query handled correctly");
console.log("=== All Synonyms & Search Tests Passed! ===");
