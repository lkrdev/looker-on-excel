import React from "react";
import {
  makeStyles,
  tokens,
  Card,
  CardHeader,
  Body1,
  Caption1,
  Button,
  ProgressBar,
  Divider,
  Badge,
  Select,
  Switch,
} from "@fluentui/react-components";
import {
  ArrowSyncRegular,
  EditRegular,
  AddRegular,
  TableRegular,
  ClockRegular,
  TargetRegular,
} from "@fluentui/react-icons";
import { LookerIcon } from "./LookerLogo";
import { lookerColors } from "../theme/lookerTheme";
import { StoredSheetConfig } from "../services/sheetMetadata";
import { FieldDefinition } from "../services/lookerClient";
import { QueryPayload } from "../services/queryTaskRunner";
import { RowGroupDisplayLevel, WriteOptions } from "../services/excelWriter";
import {
  resolveExploreHierarchies,
  getFieldHierarchyContext,
  expandFieldInSelection,
  collapseFieldInSelection,
} from "../services/hierarchyResolver";

const useStyles = makeStyles({
  container: {
    padding: "16px",
    display: "flex",
    flexDirection: "column",
    gap: "16px",
  },
  switcherBar: {
    display: "flex",
    flexWrap: "wrap",
    gap: "6px",
    padding: "8px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: "6px",
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  card: {
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: "8px",
    overflow: "hidden",
    boxShadow: "0 2px 6px rgba(0,0,0,0.04)",
  },
  metaGrid: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    padding: "0 12px 12px 12px",
  },
  metaItem: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    fontSize: "12px",
    color: tokens.colorNeutralForeground2,
  },
  actions: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
});

interface RefreshViewProps {
  config: StoredSheetConfig;
  allSheetTables?: StoredSheetConfig[];
  exploreDimensions?: FieldDefinition[];
  onSelectTable?: (tableConfig: StoredSheetConfig) => void;
  onAddTableToSheet?: () => void;
  onApplyDrillPayload?: (updatedPayload: QueryPayload, updatedOptions?: Partial<WriteOptions>) => void;
  onChangeDisplayLevel?: (displayLevel: RowGroupDisplayLevel) => void;
  onRefresh: () => void;
  onCancel?: () => void;
  onEdit: () => void;
  onNewQuery: () => void;
  isRefreshing: boolean;
  statusText: string;
  progressPercent: number;
}

export const RefreshView: React.FC<RefreshViewProps> = ({
  config,
  allSheetTables = [],
  exploreDimensions = [],
  onSelectTable,
  onAddTableToSheet,
  onApplyDrillPayload,
  onChangeDisplayLevel,
  onRefresh,
  onCancel,
  onEdit,
  onNewQuery,
  isRefreshing,
  statusText,
  progressPercent,
}) => {
  const styles = useStyles();

  const formattedDate = config.lastRefreshed
    ? new Date(config.lastRefreshed).toLocaleString()
    : "Unknown";

  const anchorLabel = config.anchorCell || config.options?.anchorCell || "A1";

  const hierarchyLadders = React.useMemo(
    () => resolveExploreHierarchies(exploreDimensions),
    [exploreDimensions]
  );

  const dimByName = React.useMemo(() => {
    const map = new Map<string, FieldDefinition>();
    exploreDimensions.forEach((d) => map.set(d.name, d));
    return map;
  }, [exploreDimensions]);

  const activeTableFields = config.queryPayload.fields || [];
  const activePivots = config.queryPayload.pivots || [];
  const nonPivotDims = activeTableFields.filter((fName) => {
    const col = config.columns.find((c) => c.fieldKey === fName);
    const isDim = (col?.category || (dimByName.has(fName) ? "dimension" : "measure")) === "dimension";
    return isDim && !activePivots.includes(fName);
  });
  const isRowGroupsEnabled = Boolean(config.options?.enableRowGroups && nonPivotDims.length >= 2);
  const deepestDim = nonPivotDims.length > 0 ? nonPivotDims[nonPivotDims.length - 1] : null;
  const deepestCtx = deepestDim ? getFieldHierarchyContext(deepestDim, hierarchyLadders) : null;

  return (
    <div className={styles.container}>
      {allSheetTables.length > 1 && (
        <div>
          <Caption1 style={{ fontWeight: 600, display: "block", marginBottom: "4px" }}>
            Tables on Worksheet ({allSheetTables.length})
          </Caption1>
          <div className={styles.switcherBar}>
            {allSheetTables.map((tbl, idx) => {
              const tAnchor = tbl.anchorCell || tbl.options?.anchorCell || "A1";
              const isSelected =
                (tbl.tableId && tbl.tableId === config.tableId) || tAnchor === anchorLabel;
              return (
                <Button
                  key={tbl.tableId || `${tAnchor}_${idx}`}
                  size="small"
                  appearance={isSelected ? "primary" : "outline"}
                  onClick={() => onSelectTable && onSelectTable(tbl)}
                >
                  {tbl.exploreLabel || tbl.queryPayload.view} ({tAnchor})
                </Button>
              );
            })}
          </div>
        </div>
      )}

      <Card className={styles.card}>
        <div
          style={{
            height: "3px",
            width: "100%",
            background: lookerColors.accentGradient,
          }}
        />
        <CardHeader
          image={<LookerIcon size={28} variant="color" />}
          header={
            <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
              <Body1 style={{ fontWeight: 600 }}>
                {config.exploreLabel || config.queryPayload.view}
              </Body1>
              <Badge appearance="tint" color="brand" size="small">
                {anchorLabel}
              </Badge>
            </div>
          }
          description={
            <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
              Model: {config.modelLabel || config.queryPayload.model}
            </Caption1>
          }
        />
        <Divider />
        <div className={styles.metaGrid}>
          <div className={styles.metaItem}>
            <TableRegular style={{ color: lookerColors.navy }} />
            <span>
              <strong>{config.rowCount.toLocaleString()}</strong> rows ·{" "}
              <strong>{config.columns.length}</strong> columns
            </span>
          </div>
          <div className={styles.metaItem}>
            <TargetRegular style={{ color: tokens.colorNeutralForeground3 }} />
            <span>Start Cell Anchor: <strong>{anchorLabel}</strong></span>
          </div>
          <div className={styles.metaItem}>
            <ClockRegular style={{ color: tokens.colorNeutralForeground3 }} />
            <span>Last refreshed: {formattedDate}</span>
          </div>
        </div>
      </Card>

      {(nonPivotDims.length >= 2 || deepestCtx) && onApplyDrillPayload && (
        <Card className={styles.card}>
          <div style={{ padding: "10px 12px", display: "flex", flexDirection: "column", gap: "8px" }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
              <div style={{ display: "flex", flexDirection: "column" }}>
                <Caption1 style={{ fontWeight: 600 }}>Row Groups / Hierarchy Tree</Caption1>
                <Caption1 style={{ color: tokens.colorNeutralForeground3, fontSize: "11px" }}>
                  {nonPivotDims
                    .map((d) => dimByName.get(d)?.label_short || dimByName.get(d)?.label || d)
                    .join(" ▸ ")}
                </Caption1>
              </div>
              {nonPivotDims.length >= 2 && (
                <Switch
                  checked={isRowGroupsEnabled}
                  disabled={isRefreshing}
                  onChange={(_, data) => {
                    onApplyDrillPayload(config.queryPayload, {
                      enableRowGroups: data.checked,
                      includeSubtotals: data.checked ? true : config.options?.includeSubtotals,
                      rowGroupFields: nonPivotDims,
                      showGroupCounts: config.options?.showGroupCounts ?? true,
                      showGroupSubtotals: config.options?.showGroupSubtotals ?? true,
                      defaultDisplayLevel: config.options?.defaultDisplayLevel || "all_expanded",
                    });
                  }}
                />
              )}
            </div>

            {isRowGroupsEnabled && (
              <>
                <div style={{ display: "flex", flexDirection: "column", gap: "4px" }}>
                  <Caption1 style={{ fontWeight: 600, fontSize: "11px" }}>
                    Display Level (In-Sheet Expand / Collapse)
                  </Caption1>
                  <Select
                    size="small"
                    value={config.options?.defaultDisplayLevel || "all_expanded"}
                    onChange={(_, data) => {
                      const nextLevel = data.value as RowGroupDisplayLevel;
                      if (onChangeDisplayLevel) {
                        onChangeDisplayLevel(nextLevel);
                      }
                    }}
                  >
                    <option value="all_collapsed">All Collapsed</option>
                    <option value="level_1">Level 1</option>
                    <option value="level_2">Level 2</option>
                    <option value="level_3">Level 3</option>
                    <option value="all_expanded">All Expanded</option>
                  </Select>
                </div>
                <Caption1 style={{ color: tokens.colorNeutralForeground3, fontSize: "11px" }}>
                  Tip: Click any <strong>▸</strong> or <strong>▾</strong> group cell in the{" "}
                  <strong>{config.options?.groupColumnHeader || "Group"}</strong> column on the
                  worksheet to expand or collapse that branch in-place.
                </Caption1>
              </>
            )}

            {/* Add / Remove Hierarchy Level Buttons */}
            {deepestCtx && (
              <div style={{ display: "flex", gap: "6px", flexWrap: "wrap", paddingTop: "2px" }}>
                {deepestCtx.childField && !activeTableFields.includes(deepestCtx.childField) && (
                  <Button
                    size="small"
                    appearance="outline"
                    disabled={isRefreshing}
                    onClick={() => {
                      const { fields } = expandFieldInSelection(
                        activeTableFields,
                        deepestDim!,
                        hierarchyLadders,
                        "insert_child"
                      );
                      onApplyDrillPayload(
                        {
                          ...config.queryPayload,
                          fields,
                        },
                        {
                          enableRowGroups: true,
                          includeSubtotals: true,
                          rowGroupFields: [...nonPivotDims, deepestCtx.childField!],
                        }
                      );
                    }}
                  >
                    + Add Level:{" "}
                    {dimByName.get(deepestCtx.childField)?.label_short ||
                      dimByName.get(deepestCtx.childField)?.label ||
                      deepestCtx.childField}
                  </Button>
                )}
                {nonPivotDims.length >= 2 && deepestCtx.parentField && (
                  <Button
                    size="small"
                    appearance="subtle"
                    disabled={isRefreshing}
                    onClick={() => {
                      const { fields } = collapseFieldInSelection(
                        activeTableFields,
                        deepestCtx.parentField!,
                        hierarchyLadders
                      );
                      const nextGroupFields = nonPivotDims.slice(0, -1);
                      onApplyDrillPayload(
                        {
                          ...config.queryPayload,
                          fields,
                        },
                        {
                          enableRowGroups: nextGroupFields.length >= 2,
                          rowGroupFields: nextGroupFields,
                        }
                      );
                    }}
                  >
                    - Remove Level:{" "}
                    {dimByName.get(deepestDim!)?.label_short ||
                      dimByName.get(deepestDim!)?.label ||
                      deepestDim}
                  </Button>
                )}
              </div>
            )}
          </div>
        </Card>
      )}

      {isRefreshing && (
        <div style={{ display: "flex", flexDirection: "column", gap: "4px" }}>
          <div style={{ display: "flex", justifyContent: "space-between" }}>
            <Caption1 style={{ fontWeight: 600 }}>{statusText}</Caption1>
            {progressPercent > 0 && <Caption1>{progressPercent}%</Caption1>}
          </div>
          <ProgressBar
            value={progressPercent > 0 ? progressPercent / 100 : undefined}
            color="brand"
          />
        </div>
      )}

      <div className={styles.actions}>
        <Button
          appearance="primary"
          icon={<ArrowSyncRegular />}
          size="large"
          onClick={onRefresh}
          disabled={isRefreshing}
          style={{ backgroundColor: lookerColors.blueDark }}
        >
          {isRefreshing ? "Refreshing..." : `Refresh Table (${anchorLabel})`}
        </Button>

        {isRefreshing && onCancel && (
          <Button
            appearance="secondary"
            onClick={onCancel}
            style={{ color: lookerColors.red }}
          >
            Cancel Refresh
          </Button>
        )}

        <Button
          appearance="secondary"
          icon={<EditRegular />}
          onClick={onEdit}
          disabled={isRefreshing}
        >
          Edit Query & Filters
        </Button>

        {onAddTableToSheet && (
          <Button
            appearance="outline"
            icon={<TargetRegular />}
            onClick={onAddTableToSheet}
            disabled={isRefreshing}
          >
            Add Another Table to This Sheet
          </Button>
        )}

        <Button
          appearance="subtle"
          icon={<AddRegular />}
          onClick={onNewQuery}
          disabled={isRefreshing}
        >
          Build New Query (New Sheet)
        </Button>
      </div>
    </div>
  );
};
