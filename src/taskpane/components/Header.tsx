import React from "react";
import {
  makeStyles,
  tokens,
  Caption1,
  Button,
} from "@fluentui/react-components";
import {
  ArrowSyncRegular,
  SignOutRegular,
  CheckmarkCircleFilled,
  DatabaseSearchRegular,
  BookmarkRegular,
  BoardRegular,
} from "@fluentui/react-icons";
import { LookerIcon, LookerAccentBar } from "./LookerLogo";
import { lookerColors } from "../theme/lookerTheme";

export type NavigationMode = "explore" | "looks" | "dashboards";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    backgroundColor: tokens.colorNeutralBackground1,
  },
  content: {
    display: "flex",
    flexDirection: "column",
    padding: "10px 16px 8px 16px",
    gap: "8px",
  },
  topRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
  },
  brand: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  brandTitle: {
    fontFamily:
      "'Google Sans', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
    fontSize: "16px",
    fontWeight: 600,
    color: "#202124",
    letterSpacing: "-0.2px",
    display: "flex",
    alignItems: "center",
  },
  excelTag: {
    fontSize: "11px",
    fontWeight: 600,
    color: lookerColors.navy,
    backgroundColor: lookerColors.blueLight,
    padding: "2px 6px",
    borderRadius: "4px",
    marginLeft: "6px",
    letterSpacing: "0.2px",
  },
  instanceRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "4px 8px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: "6px",
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  instanceInfo: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    color: tokens.colorNeutralForeground2,
    fontSize: "12px",
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  navBar: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr 1fr",
    gap: "4px",
    padding: "3px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: "6px",
  },
  navBtn: {
    fontSize: "12px",
    fontWeight: 600,
  },
});

interface HeaderProps {
  isConnected: boolean;
  instanceUrl?: string;
  userName?: string;
  onDisconnect?: () => void;
  onRefreshAll?: () => void;
  refreshingAll?: boolean;
  activeMode?: NavigationMode;
  onModeChange?: (mode: NavigationMode) => void;
}

export const Header: React.FC<HeaderProps> = ({
  isConnected,
  instanceUrl,
  userName,
  onDisconnect,
  onRefreshAll,
  refreshingAll,
  activeMode = "explore",
  onModeChange,
}) => {
  const styles = useStyles();

  const formattedUrl = instanceUrl
    ? instanceUrl.replace(/^https?:\/\//, "").replace(/\/$/, "")
    : "";

  return (
    <div className={styles.container}>
      <div className={styles.content}>
        <div className={styles.topRow}>
          <div className={styles.brand}>
            <LookerIcon size={24} variant="color" />
            <div className={styles.brandTitle}>
              Looker <span className={styles.excelTag}>Excel Add-in</span>
            </div>
          </div>
          {isConnected && onRefreshAll && (
            <Button
              size="small"
              appearance="subtle"
              icon={<ArrowSyncRegular />}
              onClick={onRefreshAll}
              disabled={refreshingAll}
              style={{ color: lookerColors.navy }}
            >
              {refreshingAll ? "Refreshing..." : "Refresh All"}
            </Button>
          )}
        </div>

        {isConnected && (
          <div className={styles.instanceRow}>
            <div className={styles.instanceInfo}>
              <CheckmarkCircleFilled style={{ color: lookerColors.green, fontSize: "14px" }} />
              <Caption1 title={instanceUrl} style={{ fontWeight: 500 }}>
                {userName ? `${userName} (${formattedUrl})` : formattedUrl}
              </Caption1>
            </div>
            {onDisconnect && (
              <Button
                size="small"
                appearance="subtle"
                icon={<SignOutRegular />}
                onClick={onDisconnect}
                title="Disconnect Looker instance"
              />
            )}
          </div>
        )}

        {isConnected && onModeChange && (
          <div className={styles.navBar} role="tablist" aria-label="Looker Mode Navigation">
            <Button
              size="small"
              appearance={activeMode === "explore" ? "primary" : "subtle"}
              icon={<DatabaseSearchRegular />}
              onClick={() => onModeChange("explore")}
              className={styles.navBtn}
            >
              Explore
            </Button>
            <Button
              size="small"
              appearance={activeMode === "looks" ? "primary" : "subtle"}
              icon={<BookmarkRegular />}
              onClick={() => onModeChange("looks")}
              className={styles.navBtn}
            >
              Looks
            </Button>
            <Button
              size="small"
              appearance={activeMode === "dashboards" ? "primary" : "subtle"}
              icon={<BoardRegular />}
              onClick={() => onModeChange("dashboards")}
              className={styles.navBtn}
            >
              Dashboards
            </Button>
          </div>
        )}
      </div>
      <LookerAccentBar height={3} mode="gradient" />
    </div>
  );
};
