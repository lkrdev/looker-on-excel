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
  onSelectTable?: (tableConfig: StoredSheetConfig) => void;
  onAddTableToSheet?: () => void;
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
  onSelectTable,
  onAddTableToSheet,
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
