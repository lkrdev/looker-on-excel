import React from "react";
import {
  makeStyles,
  tokens,
  Button,
  ProgressBar,
  Caption1,
} from "@fluentui/react-components";
import { ArrowDownloadRegular, DismissCircleRegular } from "@fluentui/react-icons";
import { lookerColors } from "../theme/lookerTheme";

const useStyles = makeStyles({
  container: {
    padding: "12px 16px",
    borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground1,
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  progressInfo: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
  },
  actionRow: {
    display: "flex",
    gap: "8px",
  },
});

interface ExecutionBarProps {
  isExecuting: boolean;
  statusText: string;
  progressPercent: number;
  canExecute: boolean;
  onExecute: () => void;
  onCancel: () => void;
}

export const ExecutionBar: React.FC<ExecutionBarProps> = ({
  isExecuting,
  statusText,
  progressPercent,
  canExecute,
  onExecute,
  onCancel,
}) => {
  const styles = useStyles();

  return (
    <div className={styles.container}>
      {isExecuting && (
        <div style={{ display: "flex", flexDirection: "column", gap: "4px" }}>
          <div className={styles.progressInfo}>
            <Caption1 style={{ fontWeight: 600 }}>{statusText}</Caption1>
            {progressPercent > 0 && <Caption1>{progressPercent}%</Caption1>}
          </div>
          <ProgressBar
            value={progressPercent > 0 ? progressPercent / 100 : undefined}
            color="brand"
          />
        </div>
      )}

      <div className={styles.actionRow}>
        {!isExecuting ? (
          <Button
            appearance="primary"
            icon={<ArrowDownloadRegular />}
            onClick={onExecute}
            disabled={!canExecute}
            style={{
              flex: 1,
              backgroundColor: canExecute ? lookerColors.blueDark : undefined,
              boxShadow: canExecute ? "0 1px 3px rgba(0,0,0,0.15)" : undefined,
            }}
            size="large"
          >
            Import Data into Excel
          </Button>
        ) : (
          <Button
            appearance="secondary"
            icon={<DismissCircleRegular />}
            onClick={onCancel}
            style={{ flex: 1, color: lookerColors.red }}
          >
            Cancel Query
          </Button>
        )}
      </div>
    </div>
  );
};
