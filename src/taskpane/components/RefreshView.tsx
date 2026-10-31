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
  Input,
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
import {
  resolveExploreHierarchies,
  getFieldHierarchyContext,
  expandFieldInSelection,
  collapseFieldInSelection,
  buildRowDrillQueryPayload,
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
  onApplyDrillPayload?: (updatedPayload: QueryPayload) => void;
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
  onRefresh,
  onCancel,
  onEdit,
  onNewQuery,
  isRefreshing,
  statusText,
  progressPercent,
}) => {
  const styles = useStyles();
  const [drillFilterValue, setDrillFilterValue] = React.useState("");

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
  const hierarchicalFields = activeTableFields
    .map((fName) => ({
      fieldName: fName,
      label: dimByName.get(fName)?.label_short || dimByName.get(fName)?.label || fName,
      ctx: getFieldHierarchyContext(fName, hierarchyLadders),
    }))
    .filter((item) => item.ctx !== null);

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

      {hierarchicalFields.length > 0 && onApplyDrillPayload && (
        <Card className={styles.card}>
          <div style={{ padding: "10px 12px", display: "flex", flexDirection: "column", gap: "8px" }}>
            <Caption1 style={{ fontWeight: 600 }}>Hierarchy Expand / Collapse & Drill</Caption1>
            <Input
              size="small"
              placeholder="Optional: Filter value for row drill (e.g. 2025)"
              value={drillFilterValue}
              onChange={(_, d) => setDrillFilterValue(d.value)}
            />
            {hierarchicalFields.map(({ fieldName, label, ctx }) => {
              if (!ctx) return null;
              const childField = ctx.childField;
              const childLabel = childField
                ? dimByName.get(childField)?.label_short ||
                  dimByName.get(childField)?.label ||
                  childField
                : null;
              const canExpand = Boolean(childField && !activeTableFields.includes(childField));
              const hasExpandedDescendant = ctx.ladder.levels
                .slice(ctx.levelIndex + 1)
                .some((lv) => activeTableFields.includes(lv));
              const canCollapse = hasExpandedDescendant || Boolean(ctx.parentField);

              return (
                <div
                  key={fieldName}
                  style={{
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "space-between",
                    gap: "6px",
                    fontSize: "12px",
                  }}
                >
                  <span style={{ fontWeight: 500, overflow: "hidden", textOverflow: "ellipsis" }}>
                    {label}
                  </span>
                  <div style={{ display: "flex", gap: "4px", flexShrink: 0 }}>
                    {canExpand && childField && (
                      <Button
                        size="small"
                        appearance="outline"
                        disabled={isRefreshing}
                        onClick={() => {
                          if (drillFilterValue.trim()) {
                            const nextPayload = buildRowDrillQueryPayload(
                              config.queryPayload,
                              fieldName,
                              drillFilterValue.trim(),
                              childField,
                              true
                            );
                            onApplyDrillPayload(nextPayload);
                          } else {
                            const { fields } = expandFieldInSelection(
                              activeTableFields,
                              fieldName,
                              hierarchyLadders,
                              "insert_child"
                            );
                            onApplyDrillPayload({
                              ...config.queryPayload,
                              fields,
                            });
                          }
                        }}
                      >
                        + {childLabel}
                      </Button>
                    )}
                    {canCollapse && (
                      <Button
                        size="small"
                        appearance="subtle"
                        disabled={isRefreshing}
                        onClick={() => {
                          const { fields } = collapseFieldInSelection(
                            activeTableFields,
                            fieldName,
                            hierarchyLadders
                          );
                          onApplyDrillPayload({
                            ...config.queryPayload,
                            fields,
                          });
                        }}
                      >
                        - Collapse
                      </Button>
                    )}
                  </div>
                </div>
              );
            })}
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
