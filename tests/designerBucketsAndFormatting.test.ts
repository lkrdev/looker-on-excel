import assert from "node:assert/strict";
import {
  partitionFieldsIntoBuckets,
  resolveColumnBucketReorder,
} from "../src/taskpane/components/FieldPicker";
import {
  resolveEffectiveColumnNumberFormat,
  ColumnDefinition,
} from "../src/taskpane/services/excelWriter";
import { FieldDefinition } from "../src/taskpane/services/lookerClient";

const dimensions: FieldDefinition[] = [
  { name: "products.category", label: "Category", category: "dimension" },
  { name: "products.brand", label: "Brand", category: "dimension" },
  { name: "orders.created_year", label: "Created Year", category: "dimension" },
  { name: "orders.created_quarter", label: "Created Quarter", category: "dimension" },
];

const measures: FieldDefinition[] = [
  { name: "orders.total_revenue", label: "Total Revenue", category: "measure" },
  { name: "orders.count", label: "Order Count", category: "measure" },
];

// 1. partitionFieldsIntoBuckets separates Row Dimensions, Column Pivots, and Measures
{
  const selected = [
    "products.category",
    "products.brand",
    "orders.created_year",
    "orders.created_quarter",
    "orders.total_revenue",
    "orders.count",
  ];
  const pivots = ["orders.created_year", "orders.created_quarter"];

  const measureFirstBuckets = partitionFieldsIntoBuckets(
    selected,
    pivots,
    dimensions,
    measures,
    "measure_first"
  );
  assert.deepEqual(measureFirstBuckets.rowDimensionKeys, [
    "products.category",
    "products.brand",
  ]);
  assert.deepEqual(measureFirstBuckets.columnPivotKeys, [
    "orders.created_year",
    "orders.created_quarter",
  ]);
  assert.deepEqual(measureFirstBuckets.measureKeys, [
    "orders.total_revenue",
    "orders.count",
  ]);
  assert.equal(measureFirstBuckets.columnItems[0].type, "measures");
  assert.equal(measureFirstBuckets.columnItems[1].id, "orders.created_year");
  assert.equal(measureFirstBuckets.columnItems[2].id, "orders.created_quarter");

  const cohortFirstBuckets = partitionFieldsIntoBuckets(
    selected,
    pivots,
    dimensions,
    measures,
    "cohort_first"
  );
  assert.equal(cohortFirstBuckets.columnItems[0].id, "orders.created_year");
  assert.equal(cohortFirstBuckets.columnItems[1].id, "orders.created_quarter");
  assert.equal(cohortFirstBuckets.columnItems[2].type, "measures");
}

// 2. resolveColumnBucketReorder maps Measures position relative to Pivot dimensions
{
  const reorderedCohortFirst = resolveColumnBucketReorder([
    { type: "pivot", id: "orders.created_year", fieldKey: "orders.created_year" },
    { type: "measures", id: "__measures__", measureKeys: ["orders.total_revenue"] },
    { type: "pivot", id: "orders.created_quarter", fieldKey: "orders.created_quarter" },
  ]);
  assert.equal(reorderedCohortFirst.pivotGroupMode, "cohort_first");
  assert.deepEqual(reorderedCohortFirst.pivots, [
    "orders.created_year",
    "orders.created_quarter",
  ]);

  const reorderedMeasureFirst = resolveColumnBucketReorder([
    { type: "measures", id: "__measures__", measureKeys: ["orders.total_revenue"] },
    { type: "pivot", id: "orders.created_quarter", fieldKey: "orders.created_quarter" },
    { type: "pivot", id: "orders.created_year", fieldKey: "orders.created_year" },
  ]);
  assert.equal(reorderedMeasureFirst.pivotGroupMode, "measure_first");
  assert.deepEqual(reorderedMeasureFirst.pivots, [
    "orders.created_quarter",
    "orders.created_year",
  ]);
}

// 3. resolveEffectiveColumnNumberFormat applies per-measure overrides across unpivoted and pivoted columns
{
  const unpivotedMeasureCol: ColumnDefinition = {
    fieldKey: "orders.total_revenue",
    label: "Total Revenue",
    category: "measure",
    excelFormat: "$#,##0.00",
  };
  const pivotedMeasureCol: ColumnDefinition = {
    fieldKey: "orders.total_revenue::__pivot__::2024",
    measureKey: "orders.total_revenue",
    label: "Total Revenue (2024)",
    category: "measure",
    excelFormat: "$#,##0.00",
  };

  assert.equal(
    resolveEffectiveColumnNumberFormat(unpivotedMeasureCol, {}),
    "$#,##0.00"
  );
  assert.equal(
    resolveEffectiveColumnNumberFormat(unpivotedMeasureCol, {
      measureFormatOverrides: { "orders.total_revenue": "$#,##0" },
    }),
    "$#,##0"
  );
  assert.equal(
    resolveEffectiveColumnNumberFormat(pivotedMeasureCol, {
      measureFormatOverrides: { "orders.total_revenue": "0.0%" },
    }),
    "0.0%"
  );
}

console.log("All Feature 9 Designer Buckets & Number Format tests passed!");
