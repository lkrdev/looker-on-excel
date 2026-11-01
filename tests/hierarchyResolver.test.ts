import assert from "node:assert";
import {
  resolveExploreHierarchies,
  getFieldHierarchyContext,
  expandFieldInSelection,
  collapseFieldInSelection,
  buildRowDrillQueryPayload,
  buildEffectiveSortsForSubtotals,
  applyHierarchyLadderToSelection,
} from "../src/taskpane/services/hierarchyResolver";
import {
  buildWorksheetRowMatrix,
  resolveEffectiveColumnsForWrite,
  computeTreeRowHiddenFlags,
  ColumnDefinition,
} from "../src/taskpane/services/excelWriter";
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

// Test 4: applyHierarchyLadderToSelection & buildEffectiveSortsForSubtotals
const allFieldsWithMeasure: FieldDefinition[] = [
  ...mockDimensions,
  {
    name: "orders.total_revenue",
    label: "Total Revenue",
    type: "sum",
    category: "measure",
  },
];
const appliedGeo = applyHierarchyLadderToSelection(
  ["orders.total_revenue"],
  geoLadder!,
  allFieldsWithMeasure,
  3
);
assert.deepStrictEqual(
  appliedGeo.fields,
  ["users.country", "users.state", "users.city", "orders.total_revenue"],
  "applyHierarchyLadderToSelection places hierarchy dimensions in order before measures"
);
assert.deepStrictEqual(appliedGeo.rowGroupFields, [
  "users.country",
  "users.state",
  "users.city",
]);

const effSorts = buildEffectiveSortsForSubtotals(
  ["orders.total_revenue desc", "users.state desc"],
  ["users.country", "users.state"]
);
assert.deepStrictEqual(
  effSorts,
  ["users.country asc", "users.state desc", "orders.total_revenue desc"],
  "buildEffectiveSortsForSubtotals places grouped dimensions first while preserving explicit sort directions and measure sorts"
);

// Test 5: Row Groups / Indented Hierarchy Tree Matrix (matching Looker Explore & SAP AfO Compact Display)
const treeColumns: ColumnDefinition[] = [
  { fieldKey: "products.brand", label: "Products Brand", category: "dimension" },
  { fieldKey: "products.category", label: "Products Category", category: "dimension" },
  { fieldKey: "products.department", label: "Products Department", category: "dimension" },
  { fieldKey: "inventory_items.total_cost", label: "Total Cost", category: "measure" },
  { fieldKey: "inventory_items.count", label: "Number On Hand", category: "measure" },
];

const treeDataRows = [
  {
    "products.brand": "!it Jeans",
    "products.category": "Jeans",
    "products.department": "Women",
    "inventory_items.total_cost": 3051.05,
    "inventory_items.count": 56,
  },
  {
    "products.brand": "!it Jeans",
    "products.category": "Maternity",
    "products.department": "Women",
    "inventory_items.total_cost": 1101.19,
    "inventory_items.count": 11,
  },
  {
    "products.brand": "!it Jeans",
    "products.category": "Shorts",
    "products.department": "Women",
    "inventory_items.total_cost": 1007.02,
    "inventory_items.count": 46,
  },
];

const treeSubtotalsData = {
  "1": [
    {
      "products.brand": "!it Jeans",
      "products.category": "$$$_subtotal_$$$",
      "products.department": "$$$_subtotal_$$$",
      "inventory_items.total_cost": 5159.26,
      "inventory_items.count": 113,
    },
  ],
  "2": [
    {
      "products.brand": "!it Jeans",
      "products.category": "Jeans",
      "products.department": "$$$_subtotal_$$$",
      "inventory_items.total_cost": 3051.05,
      "inventory_items.count": 56,
    },
    {
      "products.brand": "!it Jeans",
      "products.category": "Maternity",
      "products.department": "$$$_subtotal_$$$",
      "inventory_items.total_cost": 1101.19,
      "inventory_items.count": 11,
    },
    {
      "products.brand": "!it Jeans",
      "products.category": "Shorts",
      "products.department": "$$$_subtotal_$$$",
      "inventory_items.total_cost": 1007.02,
      "inventory_items.count": 46,
    },
  ],
};

const effCols = resolveEffectiveColumnsForWrite(treeColumns, {
  enableRowGroups: true,
  groupColumnHeader: "Product Hierarchy",
});
assert.strictEqual(effCols.length, 3, "3 grouped dimensions collapse into 1 Group column + 2 measures");
assert.strictEqual(effCols[0].fieldKey, "__row_group__");
assert.strictEqual(effCols[0].label, "Product Hierarchy");

const treeMatrix = buildWorksheetRowMatrix(
  treeColumns,
  treeDataRows,
  treeSubtotalsData,
  { "inventory_items.total_cost": 5159.26, "inventory_items.count": 113 },
  undefined,
  {
    enableRowGroups: true,
    defaultDisplayLevel: "level_1",
    showGroupCounts: true,
    showGroupSubtotals: true,
  }
);

// 1 Level-0 brand node + 3 Level-1 category nodes + 3 Level-2 department leaves + 1 Grand Total = 8 rows
assert.strictEqual(treeMatrix.rows.length, 8, "Should emit 7 tree rows + 1 Grand Total");
assert.strictEqual(
  treeMatrix.rows[0].__row_group__,
  "▾ !it Jeans (3)",
  "Level 0 root node is expanded at defaultDisplayLevel=level_1 and shows direct child count (3)"
);
assert.strictEqual(
  treeMatrix.rows[0]["inventory_items.total_cost"],
  5159.26,
  "Level 0 root node carries Looker SQL subtotal"
);
assert.strictEqual(
  treeMatrix.rows[1].__row_group__,
  "    ▸ Jeans (1)",
  "Level 1 category node is collapsed at defaultDisplayLevel=level_1 with 4-space indent"
);
assert.strictEqual(
  treeMatrix.rows[2].__row_group__,
  "          Women",
  "Level 2 leaf row has 8-space + 2-space indent"
);

// Verify computeTreeRowHiddenFlags at defaultDisplayLevel = "level_1"
const hiddenAtLevel1 = computeTreeRowHiddenFlags(
  treeMatrix.rows.length,
  treeMatrix.treeNodes!,
  treeMatrix.grandTotalRowIndex
);
// Row 0 (!it Jeans): visible (false)
// Row 1 (Jeans): visible (false)
// Row 2 (Women leaf): hidden (true) because Jeans is collapsed at level_1!
// Row 3 (Maternity): visible (false)
// Row 4 (Women leaf): hidden (true)
// Row 5 (Shorts): visible (false)
// Row 6 (Women leaf): hidden (true)
// Row 7 (Grand Total): visible (false)
assert.deepStrictEqual(hiddenAtLevel1, [false, false, true, false, true, false, true, false]);

// Simulate clicking "    ▸ Jeans (1)" (treeNodes[1]) to expand it in-sheet
treeMatrix.treeNodes![1].isExpanded = true;
const hiddenAfterExpandJeans = computeTreeRowHiddenFlags(
  treeMatrix.rows.length,
  treeMatrix.treeNodes!,
  treeMatrix.grandTotalRowIndex
);
assert.deepStrictEqual(
  hiddenAfterExpandJeans,
  [false, false, false, false, true, false, true, false],
  "Expanding Jeans reveals its leaf row (index 2) while Maternity and Shorts leaves remain hidden"
);

console.log("All hierarchyResolver & Row Groups Tree tests passed!");
