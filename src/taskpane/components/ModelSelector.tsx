import React from "react";
import {
  makeStyles,
  tokens,
  Label,
  Select,
  Spinner,
} from "@fluentui/react-components";
import { ModelSummary } from "../services/lookerClient";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    padding: "0 16px 12px 16px",
  },
  row: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
  },
});

interface ModelSelectorProps {
  models: ModelSummary[];
  selectedModel: string;
  selectedExplore: string;
  onModelChange: (modelName: string) => void;
  onExploreChange: (exploreName: string) => void;
  loadingModels: boolean;
  loadingExplore: boolean;
}

export const ModelSelector: React.FC<ModelSelectorProps> = ({
  models,
  selectedModel,
  selectedExplore,
  onModelChange,
  onExploreChange,
  loadingModels,
  loadingExplore,
}) => {
  const styles = useStyles();

  const currentModel = models.find((m) => m.name === selectedModel);
  const explores = currentModel?.explores || [];

  return (
    <div className={styles.container}>
      <div className={styles.row}>
        <Label size="small" weight="semibold" htmlFor="model-select">
          LookML Model
        </Label>
        {loadingModels ? (
          <Spinner size="tiny" label="Loading models..." />
        ) : (
          <Select
            id="model-select"
            value={selectedModel}
            onChange={(_, data) => onModelChange(data.value)}
          >
            <option value="">-- Select a Model --</option>
            {models.map((m) => (
              <option key={m.name} value={m.name}>
                {m.label || m.name}
              </option>
            ))}
          </Select>
        )}
      </div>

      <div className={styles.row}>
        <Label size="small" weight="semibold" htmlFor="explore-select">
          Explore
        </Label>
        {loadingExplore ? (
          <Spinner size="tiny" label="Loading explore fields..." />
        ) : (
          <Select
            id="explore-select"
            value={selectedExplore}
            onChange={(_, data) => onExploreChange(data.value)}
            disabled={!selectedModel || explores.length === 0}
          >
            <option value="">-- Select an Explore --</option>
            {explores.map((e) => (
              <option key={e.name} value={e.name}>
                {e.label || e.name}
              </option>
            ))}
          </Select>
        )}
      </div>
    </div>
  );
};
