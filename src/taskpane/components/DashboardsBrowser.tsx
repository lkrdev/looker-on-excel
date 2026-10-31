import React, { useState, useEffect, useMemo } from "react";
import {
  makeStyles,
  tokens,
  Input,
  Card,
  Body1,
  Body2,
  Caption1,
  Button,
  Spinner,
  Badge,
  Label,
} from "@fluentui/react-components";
import {
  SearchRegular,
  ArrowDownloadRegular,
  ArrowLeftRegular,
  EditRegular,
  Folder16Regular,
  BoardRegular,
  FilterRegular,
} from "@fluentui/react-icons";
import {
  getDashboards,
  getDashboard,
  convertWriteQueryToPayload,
  extractTileQueryAndListeners,
  applyDashboardFiltersToTilePayload,
  DashboardSummary,
  DashboardDetail,
  DashboardElementDef,
} from "../services/lookerClient";
import { QueryPayload } from "../services/queryTaskRunner";
import { lookerColors } from "../theme/lookerTheme";

const useStyles = makeStyles({
  container: {
    padding: "12px 16px",
    display: "flex",
    flexDirection: "column",
    gap: "12px",
  },
  list: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  card: {
    padding: "10px 12px",
    display: "flex",
    flexDirection: "column",
    gap: "6px",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: "6px",
    backgroundColor: tokens.colorNeutralBackground1,
  },
  filterBox: {
    padding: "10px 12px",
    borderRadius: "6px",
    backgroundColor: tokens.colorNeutralBackground2,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  actionsRow: {
    display: "flex",
    gap: "6px",
    marginTop: "4px",
  },
});

interface DashboardsBrowserProps {
  baseUrl: string;
  token: string;
  isExecuting: boolean;
  onImportPayload: (payload: QueryPayload, sheetTitle: string, forceNewSheet?: boolean) => Promise<void>;
  onOpenInExplore: (payload: QueryPayload) => void;
  onError: (msg: string) => void;
}

export const DashboardsBrowser: React.FC<DashboardsBrowserProps> = ({
  baseUrl,
  token,
  isExecuting,
  onImportPayload,
  onOpenInExplore,
  onError,
}) => {
  const styles = useStyles();
  const [dashboards, setDashboards] = useState<DashboardSummary[]>([]);
  const [loadingList, setLoadingList] = useState(false);
  const [selectedDash, setSelectedDash] = useState<DashboardDetail | null>(null);
  const [loadingDash, setLoadingDash] = useState(false);
  const [searchTerm, setSearchTerm] = useState("");
  const [dashFilterValues, setDashFilterValues] = useState<Record<string, string>>({});
  const [importingAll, setImportingAll] = useState(false);

  useEffect(() => {
    setLoadingList(true);
    getDashboards(baseUrl, token)
      .then(setDashboards)
      .catch((e) => onError(`Failed to load Dashboards: ${e.message}`))
      .finally(() => setLoadingList(false));
  }, [baseUrl, token]);

  const filteredDashboards = useMemo(() => {
    if (!searchTerm.trim()) return dashboards;
    const term = searchTerm.toLowerCase();
    return dashboards.filter(
      (d) =>
        (d.title && d.title.toLowerCase().includes(term)) ||
        (d.description && d.description.toLowerCase().includes(term)) ||
        (d.folder?.name && d.folder.name.toLowerCase().includes(term))
    );
  }, [dashboards, searchTerm]);

  const handleSelectDashboard = async (dashId: string) => {
    setLoadingDash(true);
    try {
      const detail = await getDashboard(baseUrl, token, dashId);
      setSelectedDash(detail);
      const defaults: Record<string, string> = {};
      (detail.dashboard_filters || []).forEach((df) => {
        if (df.default_value) {
          defaults[df.name] = df.default_value;
        }
      });
      setDashFilterValues(defaults);
    } catch (e: any) {
      onError(e.message || "Failed to load Dashboard details.");
    } finally {
      setLoadingDash(false);
    }
  };

  const queryableTiles = useMemo(() => {
    if (!selectedDash?.dashboard_elements) return [];
    return selectedDash.dashboard_elements.filter((el) => {
      const { query } = extractTileQueryAndListeners(el);
      return Boolean(query && query.model && query.view && query.fields && query.fields.length > 0);
    });
  }, [selectedDash]);

  const buildTilePayload = (el: DashboardElementDef): QueryPayload | null => {
    const { query, listeners } = extractTileQueryAndListeners(el);
    if (!query) return null;
    const basePayload = convertWriteQueryToPayload(query);
    return applyDashboardFiltersToTilePayload(basePayload, listeners, dashFilterValues);
  };

  const handleImportAllTiles = async () => {
    if (queryableTiles.length === 0) return;
    setImportingAll(true);
    try {
      for (const tile of queryableTiles) {
        const payload = buildTilePayload(tile);
        if (payload) {
          await onImportPayload(
            payload,
            tile.title || `${selectedDash?.title || "Dashboard"} Tile`,
            true
          );
        }
      }
    } finally {
      setImportingAll(false);
    }
  };

  if (selectedDash) {
    return (
      <div className={styles.container}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <Button
            size="small"
            appearance="subtle"
            icon={<ArrowLeftRegular />}
            onClick={() => setSelectedDash(null)}
          >
            All Dashboards
          </Button>
          {queryableTiles.length > 1 && (
            <Button
              size="small"
              appearance="primary"
              icon={<ArrowDownloadRegular />}
              disabled={isExecuting || importingAll}
              onClick={handleImportAllTiles}
              style={{ backgroundColor: lookerColors.blueDark }}
            >
              {importingAll ? "Importing All..." : `Import All Tiles (${queryableTiles.length})`}
            </Button>
          )}
        </div>

        <Body1 style={{ fontWeight: 600 }}>{selectedDash.title}</Body1>

        {/* Dashboard-level Filters */}
        {selectedDash.dashboard_filters && selectedDash.dashboard_filters.length > 0 && (
          <div className={styles.filterBox}>
            <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
              <FilterRegular style={{ color: lookerColors.navy }} />
              <Body2 style={{ fontWeight: 600, fontSize: "12px" }}>
                Dashboard Filters ({selectedDash.dashboard_filters.length})
              </Body2>
            </div>
            {selectedDash.dashboard_filters.map((df) => (
              <div key={df.name} style={{ display: "flex", flexDirection: "column", gap: "2px" }}>
                <Label size="small">{df.title || df.name}</Label>
                <Input
                  size="small"
                  placeholder="Any value"
                  value={dashFilterValues[df.name] || ""}
                  onChange={(_, data) =>
                    setDashFilterValues((prev) => ({ ...prev, [df.name]: data.value }))
                  }
                />
              </div>
            ))}
          </div>
        )}

        {/* Dashboard Tiles */}
        <div className={styles.list}>
          {queryableTiles.map((tile, idx) => {
            const { query, listeners } = extractTileQueryAndListeners(tile);
            return (
              <Card key={tile.id || idx} className={styles.card}>
                <Body1 style={{ fontWeight: 600, fontSize: "13px" }}>
                  {tile.title || `Tile #${idx + 1}`}
                </Body1>
                <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                  Explore: {query?.model}.{query?.view} · {query?.fields?.length || 0} fields
                  {listeners.length > 0 ? ` · ${listeners.length} bound filter(s)` : ""}
                </Caption1>
                <div className={styles.actionsRow}>
                  <Button
                    size="small"
                    appearance="primary"
                    icon={<ArrowDownloadRegular />}
                    disabled={isExecuting || importingAll}
                    onClick={() => {
                      const payload = buildTilePayload(tile);
                      if (payload) {
                        onImportPayload(payload, tile.title || `Tile ${idx + 1}`);
                      }
                    }}
                    style={{ backgroundColor: lookerColors.blueDark }}
                  >
                    Import Tile
                  </Button>
                  <Button
                    size="small"
                    appearance="secondary"
                    icon={<EditRegular />}
                    disabled={isExecuting || importingAll}
                    onClick={() => {
                      const payload = buildTilePayload(tile);
                      if (payload) {
                        onOpenInExplore(payload);
                      }
                    }}
                  >
                    Open in Explore
                  </Button>
                </div>
              </Card>
            );
          })}
        </div>
      </div>
    );
  }

  return (
    <div className={styles.container}>
      <Input
        size="small"
        contentBefore={<SearchRegular />}
        placeholder="Search Dashboards by title or folder..."
        value={searchTerm}
        onChange={(_, data) => setSearchTerm(data.value)}
      />

      {loadingList || loadingDash ? (
        <div style={{ display: "flex", justifyContent: "center", padding: "24px" }}>
          <Spinner
            size="small"
            label={loadingDash ? "Loading Dashboard tiles..." : "Loading Dashboards..."}
          />
        </div>
      ) : filteredDashboards.length === 0 ? (
        <Caption1 style={{ color: tokens.colorNeutralForeground3, textAlign: "center", padding: "16px" }}>
          No Dashboards found matching your search.
        </Caption1>
      ) : (
        <div className={styles.list}>
          {filteredDashboards.map((dash) => (
            <Card
              key={dash.id}
              className={styles.card}
              style={{ cursor: "pointer" }}
              onClick={() => handleSelectDashboard(dash.id)}
            >
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: "8px" }}>
                <div style={{ display: "flex", alignItems: "center", gap: "6px", minWidth: 0 }}>
                  <BoardRegular style={{ color: lookerColors.blueDark, flexShrink: 0 }} />
                  <Body1
                    style={{
                      fontWeight: 600,
                      fontSize: "13px",
                      overflow: "hidden",
                      textOverflow: "ellipsis",
                      whiteSpace: "nowrap",
                    }}
                  >
                    {dash.title}
                  </Body1>
                </div>
                {dash.folder?.name && (
                  <Badge size="small" appearance="tint" color="informative" icon={<Folder16Regular />}>
                    {dash.folder.name}
                  </Badge>
                )}
              </div>
              {dash.description && (
                <Caption1 style={{ color: tokens.colorNeutralForeground2 }}>
                  {dash.description}
                </Caption1>
              )}
            </Card>
          ))}
        </div>
      )}
    </div>
  );
};
