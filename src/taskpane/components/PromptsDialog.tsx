import React from "react";
import {
  makeStyles,
  tokens,
  MessageBar,
  MessageBarBody,
  Label,
  Input,
  Caption1,
} from "@fluentui/react-components";
import { WarningRegular } from "@fluentui/react-icons";
import { ParameterDefinition } from "../services/lookerClient";

const useStyles = makeStyles({
  container: {
    margin: "0 16px 12px 16px",
    display: "flex",
    flexDirection: "column",
    gap: "12px",
  },
  section: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  sectionHeader: {
    fontWeight: 600,
    color: tokens.colorNeutralForeground2,
    textTransform: "uppercase",
    letterSpacing: "0.5px",
  },
  promptRow: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
  },
});

interface PromptsDialogProps {
  parameters?: ParameterDefinition[];
  alwaysFilters?: Array<{ field: string; values: string[] }>;
  promptValues: Record<string, string>;
  onPromptChange: (field: string, value: string) => void;
}

export const PromptsDialog: React.FC<PromptsDialogProps> = ({
  parameters = [],
  alwaysFilters = [],
  promptValues,
  onPromptChange,
}) => {
  const styles = useStyles();

  if (parameters.length === 0 && alwaysFilters.length === 0) {
    return null;
  }

  return (
    <div className={styles.container}>
      {/* 1. True Required Filters (always_filter defined in LookML) */}
      {alwaysFilters.length > 0 && (
        <div className={styles.section}>
          <MessageBar intent="warning" icon={<WarningRegular />}>
            <MessageBarBody>
              <strong>Required Filter{alwaysFilters.length > 1 ? "s" : ""}</strong>
              <Caption1 style={{ display: "block" }}>
                This Explore requires the following filter{alwaysFilters.length > 1 ? "s" : ""} before running:
              </Caption1>
            </MessageBarBody>
          </MessageBar>

          {alwaysFilters.map((f) => (
            <div key={f.field} className={styles.promptRow}>
              <Label size="small" weight="semibold" required>
                {f.field}
              </Label>
              <Input
                size="small"
                value={promptValues[f.field] ?? (f.values ? f.values.join(",") : "")}
                onChange={(_, data) => onPromptChange(f.field, data.value)}
                placeholder="e.g. last 30 days, USA, etc."
              />
            </div>
          ))}
        </div>
      )}

      {/* 2. Optional LookML Parameters */}
      {parameters.length > 0 && (
        <div className={styles.section}>
          <Caption1 className={styles.sectionHeader}>
            Parameters ({parameters.length})
          </Caption1>
          {parameters.map((p) => (
            <div key={p.name} className={styles.promptRow}>
              <Label size="small" weight="regular">
                {p.label || p.name}
              </Label>
              <Input
                size="small"
                value={promptValues[p.name] ?? ""}
                onChange={(_, data) => onPromptChange(p.name, data.value)}
                placeholder={p.default_value ? `Default: ${p.default_value}` : `Enter ${p.label || p.name}`}
              />
            </div>
          ))}
        </div>
      )}
    </div>
  );
};
