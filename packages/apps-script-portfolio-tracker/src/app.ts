// Make changes in https://github.com/righteffort/empower-poster/packages/apps-script-portfolio-tracker/src/

import type { PostPayload } from "@righteffort/empower-poster-types";

function fail(message = "Internal error"): never {
  throw new Error(message);
}

function getCacheVersion() {
  return (
    PropertiesService.getDocumentProperties().getProperty("CACHE_VERSION") ||
    "0"
  );
}

function invalidateCache() {
  PropertiesService.getDocumentProperties().setProperty(
    "CACHE_VERSION",
    String(Number(getCacheVersion()) + 1),
  );
}

export function EMPOWER_VALUE(
  tableName: string,
  columnName: string,
  rowNumber: number,
) {
  const key = Utilities.base64Encode(
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.MD5,
      JSON.stringify([getCacheVersion(), tableName, columnName]),
    ),
  );
  const cache = CacheService.getDocumentCache() ?? fail();
  const cached = cache.get(key);
  if (cached != null) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (JSON.parse(cached) as any[])[rowNumber - 1] ?? "";
  }
  const EMPOWER_SHEET_BY_TABLE: Record<string, string> = {
    Holdings: "Empower Holdings",
    Assets: "Empower Assets",
    "Asset Classes": "Empower Assets",
    "Class Categories": "Empower Assets",
    Accounts: "Empower Accounts",
    Institutions: "Empower Accounts",
  };
  const sheetName = EMPOWER_SHEET_BY_TABLE[tableName];
  if (!sheetName) {
    throw new Error(`Empower sheet for ${tableName} not known`);
  }
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`${sheetName} does not exist`);
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const columnHeader = `${tableName}:${columnName}`;
  const columnNumber = (headers?.indexOf(columnHeader) ?? -1) + 1;
  if (columnNumber <= 0) {
    throw new Error(`Column ${columnHeader} not found in ${sheetName}`);
  }
  const values = sheet
    .getRange(1, columnNumber, sheet.getLastRow(), 1)
    .getValues()
    .map((r) => r[0]);
  const serialized = JSON.stringify(values);
  if (serialized.length > 100 * 1000) {
    throw new Error(`Too much data in ${sheetName}:${columnName} to cache`);
  }
  cache.put(key, serialized);
  return values[rowNumber - 1] ?? "";
}

export function doPost(event: GoogleAppsScript.Events.DoPost) {
  try {
    const {
      version: { major, minor },
      holdings,
      classifications,
      accounts,
    } = JSON.parse(event.postData.contents) as PostPayload;
    Logger.log(`API version: ${major}.${minor}`);
    Logger.log(`${holdings.length} holdings`);
    Logger.log(`${Object.entries(classifications).length} classifications`);
    Logger.log(`${accounts.length} accounts`);
  } finally {
    invalidateCache();
  }
}
export function placeholder(): boolean {
  return true;
}
