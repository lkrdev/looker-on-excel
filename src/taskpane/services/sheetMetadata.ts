import { QueryPayload } from "./queryTaskRunner";
import { ColumnDefinition, WriteOptions } from "./excelWriter";
import { FilterCondition } from "../components/FilterBuilder";

export interface StoredSheetConfig {
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

/**
 * Saves query metadata to the target or active worksheet's custom properties.
 */
export async function saveSheetMetadata(
  config: StoredSheetConfig,
  targetSheetId?: string
): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = targetSheetId
      ? context.workbook.worksheets.getItem(targetSheetId)
      : context.workbook.worksheets.getActiveWorksheet();
    const customProps = sheet.customProperties;
    customProps.load(["items"]);
    await context.sync();

    // Check if property exists
    const existing = customProps.items.find((p) => p.key === PROPERTY_KEY);
    if (existing) {
      existing.value = JSON.stringify(config);
    } else {
      customProps.add(PROPERTY_KEY, JSON.stringify(config));
    }

    await context.sync();
  });
}

export interface ActiveSheetInfo {
  activeSheetId: string | null;
  sourceSheetId: string | null;
  sourceSheetName?: string;
  config: StoredSheetConfig | null;
  isPivotSheet?: boolean;
}

/**
 * Reads the active worksheet ID and its Looker query metadata in a single Excel.run pass.
 * If the active worksheet is a dedicated PivotTable sheet (no Looker metadata of its own,
 * but contains at least one PivotTable), resolves the Looker source worksheet so Refresh
 * and Edit Query operate on the source data sheet rather than colliding with the PivotTable.
 */
export async function getActiveSheetInfo(
  preferredSourceSheetId?: string | null
): Promise<ActiveSheetInfo> {
  try {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load(["id", "name"]);
      const customProps = sheet.customProperties;
      const sheetPivots = sheet.pivotTables;
      const sheetTables = sheet.tables;
      customProps.load(["items"]);
      sheetPivots.load(["items/name"]);
      sheetTables.load(["items/name"]);
      await context.sync();

      const item = customProps.items.find((p) => p.key === PROPERTY_KEY);
      const hasPivots = sheetPivots.items.length > 0;
      const hasTables = sheetTables.items.length > 0;

      // If this worksheet has Looker metadata and is not a PivotTable-only sheet, use it directly
      if (item && item.value && (!hasPivots || hasTables)) {
        return {
          activeSheetId: sheet.id,
          sourceSheetId: sheet.id,
          sourceSheetName: sheet.name,
          config: JSON.parse(item.value) as StoredSheetConfig,
          isPivotSheet: false,
        };
      }

      // If the active sheet contains a PivotTable (e.g. created via Insert -> PivotTable -> New Worksheet),
      // resolve the backing Looker worksheet in the workbook so Refresh / Edit Query writes to the source sheet.
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

        const lookerSheets: Array<{ id: string; name: string; config: StoredSheetConfig }> = [];
        for (const { s, props } of candidates) {
          const propItem = props.items.find((p) => p.key === PROPERTY_KEY);
          if (propItem && propItem.value) {
            try {
              lookerSheets.push({
                id: s.id,
                name: s.name,
                config: JSON.parse(propItem.value) as StoredSheetConfig,
              });
            } catch {}
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
            isPivotSheet: true,
          };
        }
      }

      return {
        activeSheetId: sheet.id,
        sourceSheetId: null,
        config: null,
        isPivotSheet: hasPivots,
      };
    });
  } catch (e) {
    console.warn("Could not read sheet custom properties:", e);
    return { activeSheetId: null, sourceSheetId: null, config: null };
  }
}

/**
 * Reads query metadata from the active worksheet.
 */
export async function getActiveSheetMetadata(): Promise<StoredSheetConfig | null> {
  const { config } = await getActiveSheetInfo();
  return config;
}

/**
 * Catalogs all sheets in the workbook containing Looker queries using a single batched sync.
 */
export async function getAllWorkbookLookerSheets(): Promise<
  Array<{ sheetId: string; sheetName: string; config: StoredSheetConfig }>
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

      const results: Array<{ sheetId: string; sheetName: string; config: StoredSheetConfig }> = [];

      for (const { sheet, props } of propCollections) {
        const match = props.items.find((p) => p.key === PROPERTY_KEY);
        if (match && match.value) {
          try {
            const parsed = JSON.parse(match.value);
            results.push({
              sheetId: sheet.id,
              sheetName: sheet.name,
              config: parsed,
            });
          } catch {}
        }
      }

      return results;
    });
  } catch {
    return [];
  }
}
