import React from "react";
import {
  makeStyles,
  Select,
  Input,
  Button,
  Caption1,
} from "@fluentui/react-components";
import { AddRegular, DeleteRegular } from "@fluentui/react-icons";
import { FieldDefinition } from "../services/lookerClient";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    padding: "0 16px 12px 16px",
  },
  filterRow: {
    display: "grid",
    gridTemplateColumns: "1fr 90px 1fr 30px",
    gap: "6px",
    alignItems: "center",
  },
});

export interface FilterCondition {
  id: string;
  field: string;
  operator: string;
  value: string;
}

interface FilterBuilderProps {
  fields: FieldDefinition[];
  filters: FilterCondition[];
  onAddFilter: () => void;
  onUpdateFilter: (id: string, updates: Partial<FilterCondition>) => void;
  onRemoveFilter: (id: string) => void;
}

export const FilterBuilder: React.FC<FilterBuilderProps> = ({
  fields,
  filters,
  onAddFilter,
  onUpdateFilter,
  onRemoveFilter,
}) => {
  const styles = useStyles();

  return (
    <div className={styles.container}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <Caption1 style={{ fontWeight: 600 }}>Filters ({filters.length})</Caption1>
        <Button
          size="small"
          appearance="subtle"
          icon={<AddRegular />}
          onClick={onAddFilter}
          disabled={fields.length === 0}
        >
          Add Filter
        </Button>
      </div>

      {filters.map((f) => (
        <div key={f.id} className={styles.filterRow}>
          <Select
            size="small"
            value={f.field}
            onChange={(_, data) => onUpdateFilter(f.id, { field: data.value })}
          >
            <option value="">Select Field</option>
            {fields.map((fld) => (
              <option key={fld.name} value={fld.name}>
                {fld.label || fld.name}
              </option>
            ))}
          </Select>

          <Select
            size="small"
            value={f.operator}
            onChange={(_, data) => onUpdateFilter(f.id, { operator: data.value })}
          >
            <option value="is">is</option>
            <option value="is_not">is not</option>
            <option value="contains">contains</option>
            <option value="before">before</option>
            <option value="after">after</option>
            <option value="greater_than">&gt;</option>
            <option value="less_than">&lt;</option>
          </Select>

          <Input
            size="small"
            placeholder="Value..."
            value={f.value}
            onChange={(_, data) => onUpdateFilter(f.id, { value: data.value })}
          />

          <Button
            size="small"
            appearance="subtle"
            icon={<DeleteRegular />}
            onClick={() => onRemoveFilter(f.id)}
            title="Remove Filter"
          />
        </div>
      ))}
    </div>
  );
};
