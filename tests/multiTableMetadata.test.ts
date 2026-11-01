import {
  parseCellAddress,
  formatCellAddress,
  computeResidualRanges,
} from "../src/taskpane/services/excelWriter";
import {
  normalizeTableConfigBounds,
  upsertTableInCollection,
  hitTestTableAtCell,
  getSiblingTableBoundingBoxes,
  StoredSheetConfig,
} from "../src/taskpane/services/sheetMetadata";

function assertEqual(actual: any, expected: any, testName: string) {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(
      `[FAIL] ${testName}: Expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`
    );
  }
  console.log(`[PASS] ${testName}`);
}

export function runMultiTableMetadataTests() {
  console.log("=== Running Multi-Table Worksheet & Custom Start Cell Tests ===");

  // 1. Cell Address Parsing & Formatting
  assertEqual(parseCellAddress("A1"), { startRow: 0, startCol: 0 }, "Parse A1 -> (0, 0)");
  assertEqual(parseCellAddress("B4"), { startRow: 3, startCol: 1 }, "Parse B4 -> (3, 1)");
  assertEqual(parseCellAddress("$H$12"), { startRow: 11, startCol: 7 }, "Parse $H$12 -> (11, 7)");
  assertEqual(
    parseCellAddress("Sheet1!AA10:AC20"),
    { startRow: 9, startCol: 26 },
    "Parse Sheet1!AA10:AC20 -> (9, 26)"
  );
  assertEqual(formatCellAddress(0, 0), "A1", "Format (0, 0) -> A1");
  assertEqual(formatCellAddress(3, 1), "B4", "Format (3, 1) -> B4");
  assertEqual(formatCellAddress(9, 26), "AA10", "Format (9, 26) -> AA10");

  // 2. Offset Residual Ranges for Custom Start Cell
  const resAtB4 = computeResidualRanges(100, 6, 40, 4, 3, 1);
  assertEqual(
    resAtB4.extraColRange,
    { startRow: 3, startCol: 5, rowCount: 100, colCount: 2 },
    "Extra columns offset by startRow=3, startCol=1"
  );
  assertEqual(
    resAtB4.extraRowRange,
    { startRow: 43, startCol: 1, rowCount: 60, colCount: 6 },
    "Extra rows offset by startRow=3, startCol=1"
  );

  // 3. Multi-Table Bounding Boxes, Upsert, and Hit-Testing
  const table1: StoredSheetConfig = {
    tableId: "Looker_B4",
    anchorCell: "B4",
    queryPayload: { model: "thelook", view: "orders", fields: ["orders.region", "orders.count"] },
    columns: [
      { fieldKey: "orders.region", label: "Region" },
      { fieldKey: "orders.count", label: "Count" },
    ],
    options: { anchorCell: "B4", prevColCount: 2 },
    lastRefreshed: "2026-10-31T00:00:00Z",
    rowCount: 10, // 1 header + 10 data rows = 11 rows (rows 3..13, cols 1..2)
  };

  const table2: StoredSheetConfig = {
    tableId: "Looker_H4",
    anchorCell: "H4",
    queryPayload: { model: "thelook", view: "products", fields: ["products.brand", "products.sku_count", "products.retail"] },
    columns: [
      { fieldKey: "products.brand", label: "Brand" },
      { fieldKey: "products.sku_count", label: "SKU Count" },
      { fieldKey: "products.retail", label: "Retail" },
    ],
    options: { anchorCell: "H4", prevColCount: 3 },
    lastRefreshed: "2026-10-31T00:01:00Z",
    rowCount: 20, // 1 header + 20 data rows = 21 rows (rows 3..23, cols 7..9)
  };

  const bounds1 = normalizeTableConfigBounds(table1);
  assertEqual(
    { startRow: bounds1.startRow, endRow: bounds1.endRow, startCol: bounds1.startCol, endCol: bounds1.endCol },
    { startRow: 3, endRow: 13, startCol: 1, endCol: 2 },
    "Table 1 at B4 spans rows 3..13 and cols 1..2"
  );

  let collection = upsertTableInCollection([], table1);
  collection = upsertTableInCollection(collection, table2);
  assertEqual(collection.length, 2, "Collection stores 2 distinct tables on same worksheet");

  // Update table 1 in place
  collection = upsertTableInCollection(collection, { ...table1, rowCount: 15 });
  assertEqual(collection.length, 2, "Upserting table1 updates in-place without duplicating");
  assertEqual(collection[0].rowCount, 15, "Table 1 rowCount updated to 15");

  // Hit-testing selection inside Table 1 (e.g. B6 -> row 5, col 1)
  const hit1 = hitTestTableAtCell(collection, 5, 1);
  assertEqual(hit1?.tableId, "Looker_B4", "Clicking B6 selects Table 1 (Looker_B4)");

  // Hit-testing selection inside Table 2 (e.g. I10 -> row 9, col 8)
  const hit2 = hitTestTableAtCell(collection, 9, 8);
  assertEqual(hit2?.tableId, "Looker_H4", "Clicking I10 selects Table 2 (Looker_H4)");

  // Sibling bounding boxes for Table 1 returns Table 2's bounding box
  const siblingsFor1 = getSiblingTableBoundingBoxes(collection, "Looker_B4", "B4");
  assertEqual(
    siblingsFor1,
    [{ top: 3, bottom: 23, left: 7, right: 9 }],
    "Sibling bounding box for Table 1 protects Table 2 at H4 (rows 3..23, cols 7..9)"
  );

  console.log("=== All Multi-Table Worksheet & Custom Start Cell Tests Passed! ===");
}

runMultiTableMetadataTests();
