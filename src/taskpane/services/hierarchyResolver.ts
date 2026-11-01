import { FieldDefinition } from "./lookerClient";
import { QueryPayload } from "./queryTaskRunner";

export interface HierarchyLadder {
  id: string;
  label: string;
  source: "dimension_group" | "tag" | "drill_fields";
  levels: string[]; // Ordered from coarsest (index 0, e.g. Year / Country) to finest (e.g. Date / City)
}

export interface FieldHierarchyContext {
  ladder: HierarchyLadder;
  levelIndex: number;
  parentField: string | null;
  childField: string | null;
}

const TIMEFRAME_RANK: Array<{ match: (f: FieldDefinition) => boolean; rank: number }> = [
  { match: (f) => f.type === "date_year" || f.name.endsWith("_year"), rank: 10 },
  { match: (f) => f.type === "date_fiscal_year" || f.name.endsWith("_fiscal_year"), rank: 15 },
  { match: (f) => f.type === "date_quarter" || f.name.endsWith("_quarter"), rank: 20 },
  {
    match: (f) => f.type === "date_fiscal_quarter" || f.name.endsWith("_fiscal_quarter"),
    rank: 25,
  },
  { match: (f) => f.type === "date_month" || f.name.endsWith("_month"), rank: 30 },
  { match: (f) => f.type === "date_week" || f.name.endsWith("_week"), rank: 40 },
  { match: (f) => f.type === "date_date" || f.name.endsWith("_date"), rank: 50 },
  { match: (f) => f.type === "date_hour" || f.name.endsWith("_hour"), rank: 60 },
  { match: (f) => f.type === "date_minute" || f.name.endsWith("_minute"), rank: 70 },
  { match: (f) => f.type === "date_time" || f.name.endsWith("_time"), rank: 80 },
  { match: (f) => f.type === "date_raw" || f.name.endsWith("_raw"), rank: 90 },
];

function getTimeframeRank(f: FieldDefinition): number | null {
  for (const rule of TIMEFRAME_RANK) {
    if (rule.match(f)) return rule.rank;
  }
  return null;
}

/**
 * Discovers all hierarchy ladders in an Explore from:
 * 1. Custom LookML `tags` (`hierarchy:<group>:<level>` or `drill_next:<field>`)
 * 2. Time `dimension_group` timeframes (`year -> quarter -> month -> week -> date -> time`)
 * 3. LookML `drill_fields` dimension chains
 */
export function resolveExploreHierarchies(dimensions: FieldDefinition[]): HierarchyLadder[] {
  const ladders: HierarchyLadder[] = [];
  const dimByName = new Map<string, FieldDefinition>();
  dimensions.forEach((d) => dimByName.set(d.name, d));

  // 1. Custom LookML Tag Hierarchies: `hierarchy:<groupName>:<levelNumber>`
  const tagGroups = new Map<string, Array<{ fieldName: string; level: number }>>();
  const drillNextPairs = new Map<string, string>();

  for (const d of dimensions) {
    if (!Array.isArray(d.tags)) continue;
    for (const rawTag of d.tags) {
      const tag = String(rawTag).trim();
      const hMatch = /^hierarchy:([^:]+):(\d+)$/i.exec(tag);
      if (hMatch) {
        const groupName = hMatch[1].toLowerCase();
        const levelNum = parseInt(hMatch[2], 10);
        if (!tagGroups.has(groupName)) {
          tagGroups.set(groupName, []);
        }
        tagGroups.get(groupName)!.push({ fieldName: d.name, level: levelNum });
      }
      const nextMatch = /^drill_next:(.+)$/i.exec(tag);
      if (nextMatch) {
        const target = nextMatch[1].trim();
        if (dimByName.has(target)) {
          drillNextPairs.set(d.name, target);
        }
      }
    }
  }

  for (const [groupName, items] of tagGroups.entries()) {
    const sorted = [...items].sort((a, b) => a.level - b.level);
    const uniqueLevels = Array.from(new Set(sorted.map((x) => x.fieldName)));
    if (uniqueLevels.length >= 2) {
      const prettyName = groupName.charAt(0).toUpperCase() + groupName.slice(1).replace(/_/g, " ");
      ladders.push({
        id: `tag:${groupName}`,
        label: `${prettyName} Hierarchy`,
        source: "tag",
        levels: uniqueLevels,
      });
    }
  }

  // Also assemble any `drill_next:<field>` chains
  if (drillNextPairs.size > 0) {
    const childTargets = new Set(drillNextPairs.values());
    const roots = Array.from(drillNextPairs.keys()).filter((k) => !childTargets.has(k));
    for (const root of roots) {
      const chain: string[] = [root];
      let curr: string | undefined = root;
      const visited = new Set<string>([root]);
      while (curr && drillNextPairs.has(curr)) {
        const nxt: string = drillNextPairs.get(curr)!;
        if (visited.has(nxt)) break;
        visited.add(nxt);
        chain.push(nxt);
        curr = nxt;
      }
      if (chain.length >= 2) {
        const rootDef = dimByName.get(root);
        ladders.push({
          id: `tag_chain:${root}`,
          label: `${rootDef?.label || root} Hierarchy`,
          source: "tag",
          levels: chain,
        });
      }
    }
  }

  // 2. Time `dimension_group` Hierarchies
  const timeGroups = new Map<string, Array<{ fieldName: string; rank: number; label: string }>>();
  for (const d of dimensions) {
    const rank = getTimeframeRank(d);
    if (rank === null) continue;
    const groupKey =
      d.dimension_group ||
      (d.field_group_label ? `${d.view_label || "view"}.${d.field_group_label}` : null);
    if (!groupKey) continue;

    if (!timeGroups.has(groupKey)) {
      timeGroups.set(groupKey, []);
    }
    timeGroups.get(groupKey)!.push({
      fieldName: d.name,
      rank,
      label: d.field_group_label || groupKey,
    });
  }

  for (const [groupKey, items] of timeGroups.entries()) {
    const sorted = [...items].sort((a, b) => a.rank - b.rank);
    const uniqueLevels = Array.from(new Set(sorted.map((x) => x.fieldName)));
    if (uniqueLevels.length >= 2) {
      ladders.push({
        id: `time:${groupKey}`,
        label: `${items[0].label} Time Hierarchy`,
        source: "dimension_group",
        levels: uniqueLevels,
      });
    }
  }

  // 3. LookML `drill_fields` dimension chains (for fields not already covered by tag or time ladders)
  const coveredFields = new Set<string>(ladders.flatMap((l) => l.levels));
  for (const d of dimensions) {
    if (coveredFields.has(d.name)) continue;
    if (!Array.isArray(d.drill_fields) || d.drill_fields.length === 0) continue;

    const validDrillDims = d.drill_fields.filter(
      (df) => df !== d.name && dimByName.has(df)
    );
    if (validDrillDims.length >= 1) {
      const levels = [d.name, ...validDrillDims];
      ladders.push({
        id: `drill:${d.name}`,
        label: `${d.label || d.name} Drill Hierarchy`,
        source: "drill_fields",
        levels,
      });
      levels.forEach((lv) => coveredFields.add(lv));
    }
  }

  return ladders;
}

/**
 * Looks up a field's position in any discovered HierarchyLadder, prioritizing tag > dimension_group > drill_fields.
 */
export function getFieldHierarchyContext(
  fieldName: string,
  ladders: HierarchyLadder[]
): FieldHierarchyContext | null {
  for (const ladder of ladders) {
    const idx = ladder.levels.indexOf(fieldName);
    if (idx !== -1) {
      return {
        ladder,
        levelIndex: idx,
        parentField: idx > 0 ? ladder.levels[idx - 1] : null,
        childField: idx + 1 < ladder.levels.length ? ladder.levels[idx + 1] : null,
      };
    }
  }
  return null;
}

/**
 * Expands a hierarchy level in `selectedFields` either by inserting the child field immediately
 * after `fieldName` ("insert_child") or replacing `fieldName` with `childField` ("replace_with_child").
 */
export function expandFieldInSelection(
  selectedFields: string[],
  fieldName: string,
  ladders: HierarchyLadder[],
  mode: "insert_child" | "replace_with_child" = "insert_child"
): { fields: string[]; addedChild: string | null } {
  const ctx = getFieldHierarchyContext(fieldName, ladders);
  if (!ctx || !ctx.childField) {
    return { fields: selectedFields, addedChild: null };
  }

  const idx = selectedFields.indexOf(fieldName);
  if (idx === -1) {
    return { fields: selectedFields, addedChild: null };
  }

  if (mode === "replace_with_child") {
    const next = [...selectedFields];
    if (next.includes(ctx.childField)) {
      next.splice(idx, 1);
    } else {
      next[idx] = ctx.childField;
    }
    return { fields: next, addedChild: ctx.childField };
  }

  // insert_child mode
  if (selectedFields.includes(ctx.childField)) {
    return { fields: selectedFields, addedChild: null };
  }
  const next = [...selectedFields];
  next.splice(idx + 1, 0, ctx.childField);
  return { fields: next, addedChild: ctx.childField };
}

/**
 * Collapses a child hierarchy level from `selectedFields` (or rolls up to `parentField` if only the child is present).
 */
export function collapseFieldInSelection(
  selectedFields: string[],
  fieldName: string,
  ladders: HierarchyLadder[]
): { fields: string[]; removedField: string | null } {
  const ctx = getFieldHierarchyContext(fieldName, ladders);
  if (!ctx) {
    return { fields: selectedFields, removedField: null };
  }

  // If the child of this field is currently in selectedFields, collapse (remove) the deepest expanded descendant
  const descendantsInSelection = ctx.ladder.levels
    .slice(ctx.levelIndex + 1)
    .filter((lv) => selectedFields.includes(lv));
  if (descendantsInSelection.length > 0) {
    const deepest = descendantsInSelection[descendantsInSelection.length - 1];
    return {
      fields: selectedFields.filter((f) => f !== deepest),
      removedField: deepest,
    };
  }

  // Otherwise if this field has a parentField, roll this field up to parentField
  if (ctx.parentField) {
    const idx = selectedFields.indexOf(fieldName);
    if (idx === -1) return { fields: selectedFields, removedField: null };
    const next = [...selectedFields];
    if (next.includes(ctx.parentField)) {
      next.splice(idx, 1);
    } else {
      next[idx] = ctx.parentField;
    }
    return { fields: next, removedField: fieldName };
  }

  return { fields: selectedFields, removedField: null };
}

/**
 * Builds a Looker QueryPayload that drills into a specific parent dimension value (e.g. "2025" or "West")
 * at the next hierarchy level (`childField`).
 */
export function buildRowDrillQueryPayload(
  basePayload: QueryPayload,
  parentField: string,
  parentValue: string,
  childField: string,
  keepParentColumn: boolean = true
): QueryPayload {
  const fields = [...(basePayload.fields || [])];
  const pIdx = fields.indexOf(parentField);
  if (!fields.includes(childField)) {
    if (pIdx >= 0) {
      if (keepParentColumn) {
        fields.splice(pIdx + 1, 0, childField);
      } else {
        fields[pIdx] = childField;
      }
    } else {
      fields.unshift(childField);
    }
  }

  const filters: Record<string, string> = {
    ...(basePayload.filters || {}),
    [parentField]: String(parentValue).trim(),
  };

  return {
    ...basePayload,
    fields,
    filters,
  };
}

/**
 * Ensures that all subtotaled/grouped dimensions (`subtotalDims`) appear at the front
 * of the Looker `sorts` array (preserving any explicit `asc`/`desc` direction already set
 * on those dimensions), followed by any remaining sorts (e.g. measure sorts).
 */
export function buildEffectiveSortsForSubtotals(
  sorts: string[],
  subtotalDims: string[]
): string[] {
  if (!subtotalDims || subtotalDims.length === 0) {
    return sorts;
  }
  const parseField = (s: string) => s.trim().split(/\s+/)[0];
  const existingMap = new Map<string, string>();
  for (const s of sorts) {
    const f = parseField(s);
    if (!existingMap.has(f)) {
      existingMap.set(f, s.includes(" ") ? s : `${f} asc`);
    }
  }

  const leading: string[] = [];
  const subSet = new Set<string>(subtotalDims);
  for (const dim of subtotalDims) {
    leading.push(existingMap.get(dim) || `${dim} asc`);
  }

  const trailing = sorts.filter((s) => !subSet.has(parseField(s)));
  return [...leading, ...trailing];
}

/**
 * Selects a LookML HierarchyLadder up to `maxDepth` (1-based number of levels, default all levels),
 * placing the hierarchy dimensions at the front of `selectedFields` (before measures) and returning
 * the ordered `rowGroupFields`.
 */
export function applyHierarchyLadderToSelection(
  selectedFields: string[],
  ladder: HierarchyLadder,
  allFields: FieldDefinition[],
  maxDepth?: number
): { fields: string[]; rowGroupFields: string[] } {
  const depth = maxDepth && maxDepth >= 2 ? Math.min(maxDepth, ladder.levels.length) : ladder.levels.length;
  const targetLevels = ladder.levels.slice(0, depth);
  const ladderAllSet = new Set<string>(ladder.levels);
  const measureSet = new Set<string>(
    allFields.filter((f) => f.category === "measure").map((f) => f.name)
  );

  // Remove any existing levels of this ladder from selectedFields before inserting targetLevels in order
  const remaining = selectedFields.filter((f) => !ladderAllSet.has(f));
  const firstMeasureIdx = remaining.findIndex((f) => measureSet.has(f));
  const nextFields = [...remaining];
  if (firstMeasureIdx >= 0) {
    nextFields.splice(firstMeasureIdx, 0, ...targetLevels);
  } else {
    nextFields.push(...targetLevels);
  }

  return {
    fields: nextFields,
    rowGroupFields: targetLevels,
  };
}
