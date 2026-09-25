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
} from "@fluentui/react-components";
import {
  ArrowSyncRegular,
  EditRegular,
  AddRegular,
  TableRegular,
  ClockRegular,
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

  return (
    <div className={styles.container}>
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
            <Body1 style={{ fontWeight: 600 }}>
              {config.exploreLabel || config.queryPayload.view}
            </Body1>
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
          {isRefreshing ? "Refreshing..." : "Refresh Data"}
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

        <Button
          appearance="subtle"
          icon={<AddRegular />}
          onClick={onNewQuery}
          disabled={isRefreshing}
        >
          Build New Query
        </Button>
      </div>
    </div>
  );
};
