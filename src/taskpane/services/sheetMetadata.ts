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

/**
 * Reads the active worksheet ID and its Looker query metadata in a single Excel.run pass.
 */
export async function getActiveSheetInfo(): Promise<{
  sheetId: string | null;
  config: StoredSheetConfig | null;
}> {
  try {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load(["id"]);
      const customProps = sheet.customProperties;
      customProps.load(["items"]);
      await context.sync();

      const item = customProps.items.find((p) => p.key === PROPERTY_KEY);
      if (!item || !item.value) {
        return { sheetId: sheet.id, config: null };
      }

      return {
        sheetId: sheet.id,
        config: JSON.parse(item.value) as StoredSheetConfig,
      };
    });
  } catch (e) {
    console.warn("Could not read sheet custom properties:", e);
    return { sheetId: null, config: null };
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
