import assert from "node:assert";
import {
  resolveExploreHierarchies,
  getFieldHierarchyContext,
  expandFieldInSelection,
  collapseFieldInSelection,
  buildRowDrillQueryPayload,
} from "../src/taskpane/services/hierarchyResolver";
import { FieldDefinition } from "../src/taskpane/services/lookerClient";
import { QueryPayload } from "../src/taskpane/services/queryTaskRunner";

const mockDimensions: FieldDefinition[] = [
  // 1. Time dimension_group (Created)
  {
    name: "orders.created_date",
    label: "Orders Created Date",
    label_short: "Created Date",
    type: "date_date",
    category: "dimension",
    dimension_group: "orders.created",
    field_group_label: "Created Date",
  },
  {
    name: "orders.created_year",
    label: "Orders Created Year",
    label_short: "Created Year",
    type: "date_year",
    category: "dimension",
    dimension_group: "orders.created",
    field_group_label: "Created Date",
  },
  {
    name: "orders.created_quarter",
    label: "Orders Created Quarter",
    label_short: "Created Quarter",
    type: "date_quarter",
    category: "dimension",
    dimension_group: "orders.created",
    field_group_label: "Created Date",
  },
  {
    name: "orders.created_month",
    label: "Orders Created Month",
    label_short: "Created Month",
    type: "date_month",
    category: "dimension",
    dimension_group: "orders.created",
    field_group_label: "Created Date",
  },
  // 2. Custom LookML tag hierarchy (`hierarchy:geo:1..3`)
  {
    name: "users.country",
    label: "Users Country",
    label_short: "Country",
    type: "string",
    category: "dimension",
    tags: ["hierarchy:geo:1"],
  },
  {
    name: "users.state",
    label: "Users State",
    label_short: "State",
    type: "string",
    category: "dimension",
    tags: ["hierarchy:geo:2"],
  },
  {
    name: "users.city",
    label: "Users City",
    label_short: "City",
    type: "string",
    category: "dimension",
    tags: ["hierarchy:geo:3"],
  },
  // 3. LookML `drill_next:<field>` tag chain
  {
    name: "products.department",
    label: "Department",
    type: "string",
    category: "dimension",
    tags: ["drill_next:products.category"],
  },
  {
    name: "products.category",
    label: "Category",
    type: "string",
    category: "dimension",
    tags: ["drill_next:products.brand"],
  },
  {
    name: "products.brand",
    label: "Brand",
    type: "string",
    category: "dimension",
  },
  // 4. LookML `drill_fields` fallback
  {
    name: "inventory_items.distribution_center",
    label: "Distribution Center",
    type: "string",
    category: "dimension",
    drill_fields: ["inventory_items.aisle", "inventory_items.bin"],
  },
  {
    name: "inventory_items.aisle",
    label: "Aisle",
    type: "string",
    category: "dimension",
  },
  {
    name: "inventory_items.bin",
    label: "Bin",
    type: "string",
    category: "dimension",
  },
];

// Test 1: resolveExploreHierarchies discovers tag, dimension_group, and drill_fields ladders in proper order
const ladders = resolveExploreHierarchies(mockDimensions);
const geoLadder = ladders.find((l) => l.id === "tag:geo");
assert.ok(geoLadder, "Should discover custom LookML tag hierarchy `hierarchy:geo:*`");
assert.deepStrictEqual(geoLadder!.levels, ["users.country", "users.state", "users.city"]);

const deptLadder = ladders.find((l) => l.id === "tag_chain:products.department");
assert.ok(deptLadder, "Should discover `drill_next:` tag chain");
assert.deepStrictEqual(deptLadder!.levels, [
  "products.department",
  "products.category",
  "products.brand",
]);

const timeLadder = ladders.find((l) => l.id === "time:orders.created");
assert.ok(timeLadder, "Should discover `dimension_group` time hierarchy");
assert.deepStrictEqual(
  timeLadder!.levels,
  [
    "orders.created_year",
    "orders.created_quarter",
    "orders.created_month",
    "orders.created_date",
  ],
  "Timeframes should be ordered coarsest to finest: year -> quarter -> month -> date"
);

const drillFieldsLadder = ladders.find(
  (l) => l.id === "drill:inventory_items.distribution_center"
);
assert.ok(drillFieldsLadder, "Should discover `drill_fields` hierarchy fallback");
assert.deepStrictEqual(drillFieldsLadder!.levels, [
  "inventory_items.distribution_center",
  "inventory_items.aisle",
  "inventory_items.bin",
]);

// Test 2: getFieldHierarchyContext & expandFieldInSelection / collapseFieldInSelection
const yearCtx = getFieldHierarchyContext("orders.created_year", ladders);
assert.strictEqual(yearCtx?.parentField, null);
assert.strictEqual(yearCtx?.childField, "orders.created_quarter");

const expanded1 = expandFieldInSelection(
  ["orders.created_year", "orders.total_revenue"],
  "orders.created_year",
  ladders,
  "insert_child"
);
assert.deepStrictEqual(expanded1.fields, [
  "orders.created_year",
  "orders.created_quarter",
  "orders.total_revenue",
]);

const expanded2 = expandFieldInSelection(
  expanded1.fields,
  "orders.created_quarter",
  ladders,
  "insert_child"
);
assert.deepStrictEqual(expanded2.fields, [
  "orders.created_year",
  "orders.created_quarter",
  "orders.created_month",
  "orders.total_revenue",
]);

// Collapse from parent removes deepest expanded descendant first
const collapsed1 = collapseFieldInSelection(
  expanded2.fields,
  "orders.created_year",
  ladders
);
assert.deepStrictEqual(collapsed1.fields, [
  "orders.created_year",
  "orders.created_quarter",
  "orders.total_revenue",
]);
assert.strictEqual(collapsed1.removedField, "orders.created_month");

// Collapse when only a child is selected rolls it up to its parent
const rolledUp = collapseFieldInSelection(
  ["users.state", "orders.total_revenue"],
  "users.state",
  ladders
);
assert.deepStrictEqual(rolledUp.fields, ["users.country", "orders.total_revenue"]);

// Test 3: buildRowDrillQueryPayload filters parent value and injects child dimension
const basePayload: QueryPayload = {
  model: "thelook",
  view: "orders",
  fields: ["users.country", "orders.total_revenue"],
  filters: { "orders.status": "complete" },
  sorts: ["orders.total_revenue desc"],
  limit: "500",
};

const drilledPayload = buildRowDrillQueryPayload(
  basePayload,
  "users.country",
  "USA",
  "users.state",
  true
);
assert.deepStrictEqual(drilledPayload.fields, [
  "users.country",
  "users.state",
  "orders.total_revenue",
]);
assert.deepStrictEqual(drilledPayload.filters, {
  "orders.status": "complete",
  "users.country": "USA",
});

console.log("All hierarchyResolver tests passed!");
