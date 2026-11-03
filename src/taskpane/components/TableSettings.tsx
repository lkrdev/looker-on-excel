import React from "react";
import {
  makeStyles,
  tokens,
  Select,
  Checkbox,
  Label,
  RadioGroup,
  Radio,
  Caption1,
  Input,
  Button,
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
} from "@fluentui/react-components";
import { TableSettingsRegular, TargetRegular } from "@fluentui/react-icons";
import { WriteOptions, PivotGroupMode } from "../services/excelWriter";
import { getActiveCellAddress } from "../services/sheetMetadata";

const useStyles = makeStyles({
  container: {
    padding: "0 16px 12px 16px",
  },
  section: {
    display: "flex",
    flexDirection: "column",
    gap: "10px",
  },
  row: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
  },
});

interface TableSettingsProps {
  rowLimit: string;
  onRowLimitChange: (limit: string) => void;
  destination: "active" | "new";
  onDestinationChange: (dest: "active" | "new") => void;
  options: WriteOptions;
  onOptionsChange: (updates: Partial<WriteOptions>) => void;
  pivots?: string[];
  selectedMeasureLabels?: string[];
}

export const TableSettings: React.FC<TableSettingsProps> = ({
  rowLimit,
  onRowLimitChange,
  destination,
  onDestinationChange,
  options,
  onOptionsChange,
  pivots = [],
  selectedMeasureLabels = [],
}) => {
  const styles = useStyles();
  const PRESET_LIMITS = ["500", "2000", "10000", "25000", "50000", "100000", "-1"];
  const isCustom = !PRESET_LIMITS.includes(rowLimit);
  const [showCustomInput, setShowCustomInput] = React.useState(isCustom);
  const m1 = selectedMeasureLabels[0] || "Measure 1";
  const m2 = selectedMeasureLabels[1] || "Measure 2";
  const hasMultipleMeasures = selectedMeasureLabels.length >= 2;

  return (
    <div className={styles.container}>
      <Accordion collapsible defaultOpenItems={[]}>
        <AccordionItem value="settings">
          <AccordionHeader size="small" icon={<TableSettingsRegular />}>
            Table & Streaming Settings
          </AccordionHeader>
          <AccordionPanel>
            <div className={styles.section}>
              {/* Row Limit */}
              <div className={styles.row}>
                <Label size="small" weight="semibold">
                  Row Limit
                </Label>
                <Select
                  size="small"
                  value={isCustom || showCustomInput ? "custom" : rowLimit}
                  onChange={(_, data) => {
                    if (data.value === "custom") {
                      setShowCustomInput(true);
                      if (!isCustom) {
                        onRowLimitChange("1000");
                      }
                    } else {
                      setShowCustomInput(false);
                      onRowLimitChange(data.value);
                    }
                  }}
                >
                  <option value="500">500 rows</option>
                  <option value="2000">2,000 rows</option>
                  <option value="10000">10,000 rows</option>
                  <option value="25000">25,000 rows</option>
                  <option value="50000">50,000 rows</option>
                  <option value="100000">100,000 rows (High-Volume)</option>
                  <option value="-1">All / Max Results (-1)</option>
                  <option value="custom">Custom...</option>
                </Select>
                {(isCustom || showCustomInput) && (
                  <Input
                    size="small"
                    type="number"
                    min={1}
                    max={100000}
                    placeholder="Enter custom row limit (e.g. 1500)"
                    value={rowLimit === "-1" ? "" : rowLimit}
                    onChange={(_, data) => onRowLimitChange(data.value)}
                    style={{ marginTop: "4px" }}
                  />
                )}
              </div>

              {/* Destination */}
              <div className={styles.row}>
                <Label size="small" weight="semibold">
                  Destination & Anchor Cell
                </Label>
                <RadioGroup
                  value={destination}
                  onChange={(_, data) => onDestinationChange(data.value as "active" | "new")}
                >
                  <Radio value="active" label="Insert into Active Worksheet" />
                  <Radio value="new" label="Create New Worksheet" />
                </RadioGroup>
                {destination === "active" && (
                  <div
                    style={{
                      marginLeft: "28px",
                      marginTop: "4px",
                      display: "flex",
                      flexDirection: "column",
                      gap: "4px",
                    }}
                  >
                    <Label size="small">Start Cell (Top-Left Anchor)</Label>
                    <div style={{ display: "flex", gap: "6px", alignItems: "center" }}>
                      <Input
                        size="small"
                        placeholder="A1"
                        value={options.anchorCell || "A1"}
                        onChange={(_, data) =>
                          onOptionsChange({
                            anchorCell: data.value.toUpperCase().trim() || "A1",
                          })
                        }
                        style={{ width: "84px" }}
                      />
                      <Button
                        size="small"
                        appearance="secondary"
                        icon={<TargetRegular />}
                        onClick={async () => {
                          const addr = await getActiveCellAddress();
                          onOptionsChange({ anchorCell: addr, tableId: `Looker_${addr}` });
                        }}
                      >
                        Use Selected Cell
                      </Button>
                    </div>
                    <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                      Place multiple Looker tables side-by-side (e.g. A1, H1, B25).
                    </Caption1>
                  </div>
                )}
                {destination === "new" && (
                  <div style={{ marginLeft: "28px", marginTop: "4px", display: "flex", flexDirection: "column", gap: "2px" }}>
                    <Label size="small">Worksheet Name</Label>
                    <Input
                      size="small"
                      placeholder="e.g. Q3 Orders Analysis"
                      value={options.sheetName || ""}
                      onChange={(_, data) => onOptionsChange({ sheetName: data.value })}
                    />
                    <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                      Leave blank to default to Explore name.
                    </Caption1>
                  </div>
                )}
              </div>

              {/* Source Information Banner Row */}
              <div className={styles.row}>
                <Checkbox
                  checked={options.showSourceBanner ?? true}
                  onChange={(_, data) => onOptionsChange({ showSourceBanner: !!data.checked })}
                  label={
                    <div>
                      <span style={{ fontSize: "12px", fontWeight: 600 }}>
                        Show Source Information Banner Row
                      </span>
                      <Caption1 style={{ color: tokens.colorNeutralForeground3, display: "block" }}>
                        Renders a header row above the table displaying the Explore / Look source name.
                      </Caption1>
                    </div>
                  }
                />
                {(options.showSourceBanner ?? true) && (
                  <div style={{ marginLeft: "28px", marginTop: "2px", display: "flex", flexDirection: "column", gap: "2px" }}>
                    <Label size="small">Banner Title Override (Optional)</Label>
                    <Input
                      size="small"
                      placeholder="Leave blank to use Explore / Look title"
                      value={options.sourceBannerTitle || ""}
                      onChange={(_, data) => onOptionsChange({ sourceBannerTitle: data.value })}
                    />
                  </div>
                )}
              </div>

              {/* Table Style */}
              <div className={styles.row}>
                <Checkbox
                  checked={options.useExcelTable ?? true}
                  onChange={(_, data) => onOptionsChange({ useExcelTable: !!data.checked })}
                  label={
                    <div>
                      <span style={{ fontSize: "12px", fontWeight: 600 }}>Create Native Excel Table (ListObject)</span>
                      <Caption1 style={{ color: tokens.colorNeutralForeground3, display: "block" }}>
                        Enables built-in Excel auto-filters and seamless PivotTable integration.
                      </Caption1>
                    </div>
                  }
                />
              </div>

              {options.useExcelTable && (
                <div className={styles.row} style={{ marginLeft: "24px" }}>
                  <Label size="small">Excel Table Style</Label>
                  <Select
                    size="small"
                    value={options.tableStyle || "TableStyleMedium2"}
                    onChange={(_, data) => onOptionsChange({ tableStyle: data.value })}
                  >
                    <option value="TableStyleLight1">Light 1 (Minimal)</option>
                    <option value="TableStyleMedium2">Medium 2 (Office Blue)</option>
                    <option value="TableStyleMedium9">Medium 9 (Navy / Slate)</option>
                    <option value="TableStyleDark1">Dark 1</option>
                  </Select>
                </div>
              )}

              {/* Looker Semantic Totals & Subtotals */}
              <div className={styles.row}>
                <Label size="small" weight="semibold">
                  Semantic Totals & Subtotals (Looker SQL)
                </Label>
                <Checkbox
                  checked={options.includeTotals ?? false}
                  onChange={(_, data) => onOptionsChange({ includeTotals: !!data.checked })}
                  label={
                    <div>
                      <span style={{ fontSize: "12px", fontWeight: 600 }}>Include Column Grand Totals</span>
                      <Caption1 style={{ color: tokens.colorNeutralForeground3, display: "block" }}>
                        Evaluates non-additive totals (count_distinct, ratios) in Looker SQL.
                      </Caption1>
                    </div>
                  }
                />
                <Checkbox
                  checked={options.includeRowTotals ?? false}
                  onChange={(_, data) => onOptionsChange({ includeRowTotals: !!data.checked })}
                  label={
                    <div>
                      <span style={{ fontSize: "12px", fontWeight: 600 }}>Include Row Totals (when Pivoted)</span>
                      <Caption1 style={{ color: tokens.colorNeutralForeground3, display: "block" }}>
                        Adds a Looker-evaluated Row Total column across pivoted dimensions.
                      </Caption1>
                    </div>
                  }
                />
                <Checkbox
                  checked={options.includeSubtotals ?? false}
                  onChange={(_, data) => onOptionsChange({ includeSubtotals: !!data.checked })}
                  label={
                    <div>
                      <span style={{ fontSize: "12px", fontWeight: 600 }}>Include Dimension Subtotals</span>
                      <Caption1 style={{ color: tokens.colorNeutralForeground3, display: "block" }}>
                        Computes Looker SQL subtotals with collapsible Excel row groups.
                      </Caption1>
                    </div>
                  }
                />
              </div>

              {/* Crosstab & Multi-Pivot Layout */}
              {(pivots.length > 0 || options.useMultiRowCrosstabHeader) && (
                <div className={styles.row}>
                  <Label size="small" weight="semibold">
                    Crosstab &amp; Pivot Layout
                  </Label>
                  {hasMultipleMeasures && (
                    <Select
                      size="small"
                      value={options.pivotGroupMode || "measure_first"}
                      onChange={(_, data) =>
                        onOptionsChange({ pivotGroupMode: data.value as PivotGroupMode })
                      }
                    >
                      <option value="measure_first">
                        Group by Measure ({m1} [Pivots], {m2} [Pivots])
                      </option>
                      <option value="cohort_first">
                        Group by Pivot Value (Pivot [{m1}, {m2}])
                      </option>
                    </Select>
                  )}
                  <Checkbox
                    checked={options.useMultiRowCrosstabHeader ?? false}
                    onChange={(_, data) =>
                      onOptionsChange({ useMultiRowCrosstabHeader: !!data.checked })
                    }
                    label={
                      <div>
                        <span style={{ fontSize: "12px", fontWeight: 600 }}>
                          Multi-Row Merged Crosstab Header
                        </span>
                        <Caption1 style={{ color: tokens.colorNeutralForeground3, display: "block" }}>
                          Renders stacked pivot cohort rows with merged header cells above measure names.
                        </Caption1>
                      </div>
                    }
                  />
                </div>
              )}

              {/* Enterprise SAC Benchmarks */}
              <div className={styles.row}>
                <Checkbox
                  checked={options.preserveUserFormatting ?? true}
                  onChange={(_, data) => onOptionsChange({ preserveUserFormatting: !!data.checked })}
                  label={
                    <div>
                      <span style={{ fontSize: "12px", fontWeight: 600 }}>Preserve cell formatting on refresh</span>
                      <Caption1 style={{ color: tokens.colorNeutralForeground3, display: "block" }}>
                        Maintains manual fonts, borders, fills, and row heights when refreshed.
                      </Caption1>
                    </div>
                  }
                />
              </div>

              <div className={styles.row}>
                <Checkbox
                  checked={options.autoExpandFormulas ?? true}
                  onChange={(_, data) => onOptionsChange({ autoExpandFormulas: !!data.checked })}
                  label={
                    <div>
                      <span style={{ fontSize: "12px", fontWeight: 600 }}>Auto-expand adjacent custom formulas</span>
                      <Caption1 style={{ color: tokens.colorNeutralForeground3, display: "block" }}>
                        Auto-fills custom formulas in adjacent columns down when row count grows.
                      </Caption1>
                    </div>
                  }
                />
              </div>

              <div className={styles.row}>
                <Checkbox
                  checked={options.freezeHeader ?? true}
                  onChange={(_, data) => onOptionsChange({ freezeHeader: !!data.checked })}
                  label="Freeze Header Row"
                />
              </div>
            </div>
          </AccordionPanel>
        </AccordionItem>
      </Accordion>
    </div>
  );
};
