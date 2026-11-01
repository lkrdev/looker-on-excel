import React, { useState, useEffect, useMemo } from "react";
import {
  makeStyles,
  tokens,
  Input,
  Card,
  Body1,
  Caption1,
  Button,
  Spinner,
  Badge,
} from "@fluentui/react-components";
import {
  SearchRegular,
  ArrowDownloadRegular,
  EditRegular,
  Folder16Regular,
  BookmarkRegular,
} from "@fluentui/react-icons";
import {
  getLooks,
  getLook,
  convertWriteQueryToPayload,
  LookSummary,
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
  titleRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "8px",
  },
  actionsRow: {
    display: "flex",
    gap: "6px",
    marginTop: "4px",
  },
});

interface LooksBrowserProps {
  baseUrl: string;
  token: string;
  isExecuting: boolean;
  onImportPayload: (payload: QueryPayload, sheetTitle: string) => void;
  onOpenInExplore: (payload: QueryPayload) => void;
  onError: (msg: string) => void;
}

export const LooksBrowser: React.FC<LooksBrowserProps> = ({
  baseUrl,
  token,
  isExecuting,
  onImportPayload,
  onOpenInExplore,
  onError,
}) => {
  const styles = useStyles();
  const [looks, setLooks] = useState<LookSummary[]>([]);
  const [loading, setLoading] = useState(false);
  const [loadingLookId, setLoadingLookId] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState("");

  useEffect(() => {
    setLoading(true);
    getLooks(baseUrl, token)
      .then(setLooks)
      .catch((e) => onError(`Failed to load Saved Looks: ${e.message}`))
      .finally(() => setLoading(false));
  }, [baseUrl, token]);

  const filteredLooks = useMemo(() => {
    if (!searchTerm.trim()) return looks;
    const term = searchTerm.toLowerCase();
    return looks.filter(
      (l) =>
        (l.title && l.title.toLowerCase().includes(term)) ||
        (l.description && l.description.toLowerCase().includes(term)) ||
        (l.folder?.name && l.folder.name.toLowerCase().includes(term)) ||
        (l.model?.label && l.model.label.toLowerCase().includes(term))
    );
  }, [looks, searchTerm]);

  const handleLoadLook = async (look: LookSummary, mode: "import" | "edit") => {
    setLoadingLookId(look.id);
    try {
      const detail = await getLook(baseUrl, token, look.id);
      if (!detail.query) {
        throw new Error("This Look does not contain an underlying query definition.");
      }
      const payload = convertWriteQueryToPayload(detail.query);
      if (mode === "import") {
        onImportPayload(payload, detail.title || `Look ${look.id}`);
      } else {
        onOpenInExplore(payload);
      }
    } catch (e: any) {
      onError(e.message || "Could not load Look query.");
    } finally {
      setLoadingLookId(null);
    }
  };

  return (
    <div className={styles.container}>
      <Input
        size="small"
        contentBefore={<SearchRegular />}
        placeholder="Search Saved Looks by title, folder, or model..."
        value={searchTerm}
        onChange={(_, data) => setSearchTerm(data.value)}
      />

      {loading ? (
        <div style={{ display: "flex", justifyContent: "center", padding: "24px" }}>
          <Spinner size="small" label="Loading Saved Looks..." />
        </div>
      ) : filteredLooks.length === 0 ? (
        <Caption1 style={{ color: tokens.colorNeutralForeground3, textAlign: "center", padding: "16px" }}>
          No Saved Looks found matching your search.
        </Caption1>
      ) : (
        <div className={styles.list}>
          {filteredLooks.map((look) => {
            const busy = isExecuting || loadingLookId === look.id;
            return (
              <Card key={look.id} className={styles.card}>
                <div className={styles.titleRow}>
                  <div style={{ display: "flex", alignItems: "center", gap: "6px", minWidth: 0 }}>
                    <BookmarkRegular style={{ color: lookerColors.blueDark, flexShrink: 0 }} />
                    <Body1
                      style={{
                        fontWeight: 600,
                        fontSize: "13px",
                        overflow: "hidden",
                        textOverflow: "ellipsis",
                        whiteSpace: "nowrap",
                      }}
                    >
                      {look.title}
                    </Body1>
                  </div>
                  {look.folder?.name && (
                    <Badge size="small" appearance="tint" color="informative" icon={<Folder16Regular />}>
                      {look.folder.name}
                    </Badge>
                  )}
                </div>

                {look.description && (
                  <Caption1 style={{ color: tokens.colorNeutralForeground2 }}>
                    {look.description}
                  </Caption1>
                )}

                <div className={styles.actionsRow}>
                  <Button
                    size="small"
                    appearance="primary"
                    icon={<ArrowDownloadRegular />}
                    disabled={busy}
                    onClick={() => handleLoadLook(look, "import")}
                    style={{ backgroundColor: lookerColors.blueDark }}
                  >
                    {loadingLookId === look.id ? "Loading..." : "Import Look"}
                  </Button>
                  <Button
                    size="small"
                    appearance="secondary"
                    icon={<EditRegular />}
                    disabled={busy}
                    onClick={() => handleLoadLook(look, "edit")}
                  >
                    Open in Explore
                  </Button>
                </div>
              </Card>
            );
          })}
        </div>
      )}
    </div>
  );
};
