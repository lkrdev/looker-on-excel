import { QueryPayload } from "./queryTaskRunner";
import {
  ColumnDefinition,
  WriteOptions,
  BoundingBox,
  parseCellAddress,
  formatCellAddress,
} from "./excelWriter";
import { FilterCondition } from "../components/FilterBuilder";

export interface StoredSheetConfig {
  tableId?: string;
  anchorCell?: string;
  startRow?: number;
  startCol?: number;
  endRow?: number;
  endCol?: number;
  queryPayload: QueryPayload;
  columns: ColumnDefinition[];
  options: WriteOptions;
  lastRefreshed: string;
  rowCount: number;
  modelLabel?: string;
  exploreLabel?: string;
  filters?: FilterCondition[];
  promptValues?: Record<string, string>;
}

const PROPERTY_KEY = "looker_query_config";
const MULTI_TABLE_PROPERTY_KEY = "looker_tables_v2";

/**
 * Normalizes a StoredSheetConfig with computed bounding-box coordinates [startRow..endRow, startCol..endCol].
 */
export function normalizeTableConfigBounds(config: StoredSheetConfig): StoredSheetConfig {
  const anchor = (config.anchorCell || config.options?.anchorCell || "A1").toUpperCase();
  const { startRow, startCol } = parseCellAddress(anchor);
  const colSpan = Math.max(1, config.options?.prevColCount || config.columns?.length || 1);
  const rowSpan = Math.max(1, (config.rowCount || 0) + 1); // +1 for header row
  const tableId =
    config.tableId || config.options?.tableId || `Looker_${anchor}`;

  return {
    ...config,
    tableId,
    anchorCell: anchor,
    startRow,
    startCol,
    endRow: startRow + rowSpan - 1,
    endCol: startCol + colSpan - 1,
    options: {
      ...config.options,
      anchorCell: anchor,
      tableId,
    },
  };
}

/**
 * Upserts a table configuration into a worksheet's multi-table collection by matching
 * either `tableId` or `anchorCell`.
 */
export function upsertTableInCollection(
  existingTables: StoredSheetConfig[],
  newConfig: StoredSheetConfig
): StoredSheetConfig[] {
  const normalized = normalizeTableConfigBounds(newConfig);
  const normalizedList = existingTables.map(normalizeTableConfigBounds);

  const idx = normalizedList.findIndex(
    (t) =>
      (normalized.tableId && t.tableId === normalized.tableId) ||
      (t.anchorCell || "A1").toUpperCase() === (normalized.anchorCell || "A1").toUpperCase()
  );

  if (idx >= 0) {
    const updated = [...normalizedList];
    updated[idx] = normalized;
    return updated;
  }
  return [...normalizedList, normalized];
}

/**
 * Hit-tests which StoredSheetConfig on a worksheet contains the given 0-based (rowIndex, colIndex).
 * If no table bounding box directly contains the cell, falls back to preferredTableId or the first table.
 */
export function hitTestTableAtCell(
  tables: StoredSheetConfig[],
  rowIndex: number,
  colIndex: number,
  preferredTableId?: string | null
): StoredSheetConfig | null {
  if (!tables || tables.length === 0) return null;
  const normalized = tables.map(normalizeTableConfigBounds);

  const directHit = normalized.find((t) => {
    const r0 = t.startRow ?? 0;
    const c0 = t.startCol ?? 0;
    const r1 = t.endRow ?? r0 + Math.max(1, t.rowCount);
    const c1 = t.endCol ?? c0 + Math.max(1, t.columns.length - 1);
    return rowIndex >= r0 && rowIndex <= r1 && colIndex >= c0 && colIndex <= c1;
  });

  if (directHit) return directHit;

  if (preferredTableId) {
    const byId = normalized.find((t) => t.tableId === preferredTableId);
    if (byId) return byId;
  }

  return normalized[0];
}

/**
 * Computes bounding boxes of all sibling Looker tables on the same worksheet so
 * refreshing one table never clears or shifts into another table.
 */
export function getSiblingTableBoundingBoxes(
  tables: StoredSheetConfig[],
  currentTableId?: string,
  currentAnchorCell?: string
): BoundingBox[] {
  const normAnchor = (currentAnchorCell || "A1").toUpperCase();
  return tables
    .map(normalizeTableConfigBounds)
    .filter(
      (t) =>
        !(
          (currentTableId && t.tableId === currentTableId) ||
          (t.anchorCell || "A1").toUpperCase() === normAnchor
        )
    )
    .map((t) => ({
      top: t.startRow ?? 0,
      bottom: t.endRow ?? (t.startRow ?? 0) + Math.max(1, t.rowCount),
      left: t.startCol ?? 0,
      right: t.endCol ?? (t.startCol ?? 0) + Math.max(1, t.columns.length - 1),
    }));
}

/**
 * Reads the currently selected cell address in the active worksheet (e.g., "H4").
 */
export async function getActiveCellAddress(): Promise<string> {
  try {
    return await Excel.run(async (context) => {
      const rng = context.workbook.getSelectedRange();
      rng.load(["rowIndex", "columnIndex"]);
      await context.sync();
      return formatCellAddress(rng.rowIndex, rng.columnIndex);
    });
  } catch {
    return "A1";
  }
}

function parseSheetTablesFromCustomProps(
  items: Array<{ key: string; value: string }>
): StoredSheetConfig[] {
  const v2Item = items.find((p) => p.key === MULTI_TABLE_PROPERTY_KEY);
  if (v2Item && v2Item.value) {
    try {
      const parsed = JSON.parse(v2Item.value);
      if (Array.isArray(parsed) && parsed.length > 0) {
        return parsed.map(normalizeTableConfigBounds);
      }
    } catch {}
  }

  const legacyItem = items.find((p) => p.key === PROPERTY_KEY);
  if (legacyItem && legacyItem.value) {
    try {
      const parsed = JSON.parse(legacyItem.value) as StoredSheetConfig;
      return [normalizeTableConfigBounds(parsed)];
    } catch {}
  }

  return [];
}

/**
 * Saves query metadata to the target or active worksheet's custom properties,
 * updating both the multi-table collection (`looker_tables_v2`) and legacy `looker_query_config`.
 */
export async function saveSheetMetadata(
  config: StoredSheetConfig,
  targetSheetId?: string
): Promise<StoredSheetConfig[]> {
  let updatedList: StoredSheetConfig[] = [];
  await Excel.run(async (context) => {
    const sheet = targetSheetId
      ? context.workbook.worksheets.getItem(targetSheetId)
      : context.workbook.worksheets.getActiveWorksheet();
    const customProps = sheet.customProperties;
    customProps.load(["items"]);
    await context.sync();

    const existingTables = parseSheetTablesFromCustomProps(customProps.items);
    updatedList = upsertTableInCollection(existingTables, config);
    const normalizedCurrent = normalizeTableConfigBounds(config);

    const existingV2 = customProps.items.find((p) => p.key === MULTI_TABLE_PROPERTY_KEY);
    if (existingV2) {
      existingV2.value = JSON.stringify(updatedList);
    } else {
      customProps.add(MULTI_TABLE_PROPERTY_KEY, JSON.stringify(updatedList));
    }

    const existingLegacy = customProps.items.find((p) => p.key === PROPERTY_KEY);
    if (existingLegacy) {
      existingLegacy.value = JSON.stringify(normalizedCurrent);
    } else {
      customProps.add(PROPERTY_KEY, JSON.stringify(normalizedCurrent));
    }

    await context.sync();
  });
  return updatedList;
}

export interface ActiveSheetInfo {
  activeSheetId: string | null;
  sourceSheetId: string | null;
  sourceSheetName?: string;
  config: StoredSheetConfig | null;
  allSheetTables?: StoredSheetConfig[];
  selectedCellAddress?: string;
  isPivotSheet?: boolean;
}

/**
 * Reads the active worksheet ID, selected cell, and Looker multi-table metadata in a single Excel.run pass.
 */
export async function getActiveSheetInfo(
  preferredSourceSheetId?: string | null,
  preferredTableId?: string | null
): Promise<ActiveSheetInfo> {
  try {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const selectedRange = context.workbook.getSelectedRange();
      sheet.load(["id", "name"]);
      selectedRange.load(["rowIndex", "columnIndex"]);
      const customProps = sheet.customProperties;
      const sheetPivots = sheet.pivotTables;
      const sheetTables = sheet.tables;
      customProps.load(["items"]);
      sheetPivots.load(["items/name"]);
      sheetTables.load(["items/name"]);
      await context.sync();

      const selectedCellAddress = formatCellAddress(
        selectedRange.rowIndex,
        selectedRange.columnIndex
      );
      const tablesOnSheet = parseSheetTablesFromCustomProps(customProps.items);
      const hasPivots = sheetPivots.items.length > 0;
      const hasTables = sheetTables.items.length > 0;

      if (tablesOnSheet.length > 0 && (!hasPivots || hasTables)) {
        const hitConfig = hitTestTableAtCell(
          tablesOnSheet,
          selectedRange.rowIndex,
          selectedRange.columnIndex,
          preferredTableId
        );
        return {
          activeSheetId: sheet.id,
          sourceSheetId: sheet.id,
          sourceSheetName: sheet.name,
          config: hitConfig,
          allSheetTables: tablesOnSheet,
          selectedCellAddress,
          isPivotSheet: false,
        };
      }

      if (hasPivots) {
        const sheets = context.workbook.worksheets;
        sheets.load(["items/id", "items/name"]);
        await context.sync();

        const candidates = sheets.items
          .filter((s) => s.id !== sheet.id)
          .map((s) => {
            const props = s.customProperties;
            props.load(["items"]);
            return { s, props };
          });
        await context.sync();

        const lookerSheets: Array<{
          id: string;
          name: string;
          config: StoredSheetConfig;
          allTables: StoredSheetConfig[];
        }> = [];
        for (const { s, props } of candidates) {
          const parsedTables = parseSheetTablesFromCustomProps(props.items);
          if (parsedTables.length > 0) {
            lookerSheets.push({
              id: s.id,
              name: s.name,
              config: parsedTables[0],
              allTables: parsedTables,
            });
          }
        }

        if (lookerSheets.length > 0) {
          const matched =
            (preferredSourceSheetId &&
              lookerSheets.find((ls) => ls.id === preferredSourceSheetId)) ||
            lookerSheets.sort((a, b) =>
              (b.config.lastRefreshed || "").localeCompare(a.config.lastRefreshed || "")
            )[0];

          return {
            activeSheetId: sheet.id,
            sourceSheetId: matched.id,
            sourceSheetName: matched.name,
            config: matched.config,
            allSheetTables: matched.allTables,
            selectedCellAddress,
            isPivotSheet: true,
          };
        }
      }

      return {
        activeSheetId: sheet.id,
        sourceSheetId: null,
        config: null,
        allSheetTables: [],
        selectedCellAddress,
        isPivotSheet: hasPivots,
      };
    });
  } catch (e) {
    console.warn("Could not read sheet custom properties:", e);
    return { activeSheetId: null, sourceSheetId: null, config: null, allSheetTables: [] };
  }
}

export async function getActiveSheetMetadata(): Promise<StoredSheetConfig | null> {
  const { config } = await getActiveSheetInfo();
  return config;
}

/**
 * Catalogs all Looker tables across all worksheets in the workbook using a single batched sync.
 */
export async function getAllWorkbookLookerSheets(): Promise<
  Array<{
    sheetId: string;
    sheetName: string;
    config: StoredSheetConfig;
    allTables: StoredSheetConfig[];
  }>
> {
  try {
    return await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load(["items/name", "items/id"]);
      await context.sync();

      const propCollections = sheets.items.map((s) => {
        const props = s.customProperties;
        props.load(["items"]);
        return { sheet: s, props };
      });
      await context.sync();

      const results: Array<{
        sheetId: string;
        sheetName: string;
        config: StoredSheetConfig;
        allTables: StoredSheetConfig[];
      }> = [];

      for (const { sheet, props } of propCollections) {
        const tables = parseSheetTablesFromCustomProps(props.items);
        for (const tbl of tables) {
          results.push({
            sheetId: sheet.id,
            sheetName: sheet.name,
            config: tbl,
            allTables: tables,
          });
        }
      }

      return results;
    });
  } catch {
    return [];
  }
}
