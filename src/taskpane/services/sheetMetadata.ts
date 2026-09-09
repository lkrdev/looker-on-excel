import { QueryPayload } from "./queryTaskRunner";
import { ColumnDefinition, WriteOptions } from "./excelWriter";

export interface StoredSheetConfig {
  queryPayload: QueryPayload;
  columns: ColumnDefinition[];
  options: WriteOptions;
  lastRefreshed: string;
  rowCount: number;
  modelLabel?: string;
  exploreLabel?: string;
}

const PROPERTY_KEY = "looker_query_config";

/**
 * Saves query metadata to the active worksheet's custom properties.
 */
export async function saveSheetMetadata(config: StoredSheetConfig): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
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
 * Reads query metadata from the active worksheet.
 */
export async function getActiveSheetMetadata(): Promise<StoredSheetConfig | null> {
  try {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const customProps = sheet.customProperties;
      customProps.load(["items"]);
      await context.sync();

      const item = customProps.items.find((p) => p.key === PROPERTY_KEY);
      if (!item || !item.value) {
        return null;
      }

      return JSON.parse(item.value) as StoredSheetConfig;
    });
  } catch (e) {
    console.warn("Could not read sheet custom properties:", e);
    return null;
  }
}

/**
 * Catalogs all sheets in the workbook containing Looker queries.
 */
export async function getAllWorkbookLookerSheets(): Promise<
  Array<{ sheetId: string; sheetName: string; config: StoredSheetConfig }>
> {
  try {
    return await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load(["items/name", "items/id"]);
      await context.sync();

      const results: Array<{ sheetId: string; sheetName: string; config: StoredSheetConfig }> = [];

      for (const s of sheets.items) {
        const props = s.customProperties;
        props.load(["items"]);
        await context.sync();

        const match = props.items.find((p) => p.key === PROPERTY_KEY);
        if (match && match.value) {
          try {
            const parsed = JSON.parse(match.value);
            results.push({
              sheetId: s.id,
              sheetName: s.name,
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
