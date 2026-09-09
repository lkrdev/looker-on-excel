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
    gap: "8px",
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
      <MessageBar intent="warning" icon={<WarningRegular />}>
        <MessageBarBody>
          <strong>Required Prompts</strong>
          <Caption1 style={{ display: "block" }}>
            This Explore requires the following parameters/filters before running:
          </Caption1>
        </MessageBarBody>
      </MessageBar>

      {parameters.map((p) => (
        <div key={p.name} className={styles.promptRow}>
          <Label size="small" weight="semibold" required>
            {p.label || p.name} (Parameter)
          </Label>
          <Input
            size="small"
            value={promptValues[p.name] ?? p.default_value ?? ""}
            onChange={(_, data) => onPromptChange(p.name, data.value)}
            placeholder={`Enter ${p.label || p.name}`}
          />
        </div>
      ))}

      {alwaysFilters.map((f) => (
        <div key={f.field} className={styles.promptRow}>
          <Label size="small" weight="semibold" required>
            {f.field} (Required Filter)
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
  );
};
