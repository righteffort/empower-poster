// Make changes in https://github.com/righteffort/empower-poster/packages/apps-script-portfolio-tracker-sample/src/
import type {
  PostPayload,
  PostResponse,
  HoldingEntry,
  Classifications,
  Account,
} from "@righteffort/empower-poster-types";

import { makeTableHelper } from "./sheet-utils.js";

function rangeStr(range: GoogleAppsScript.Spreadsheet.Range) {
  return `${JSON.stringify({ row: range.getRow(), column: range.getColumn(), numRows: range.getNumRows(), numColumns: range.getNumColumns() })}`;
}

/**
 * Add a leading apostrophe to purely numeric account names.
 */
function formatAccountName(accountName: string): string {
  if (accountName.search(/^[0-9]+$/) == 0) {
    return `'${accountName}`;
  }
  return accountName;
}

export function playground() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Sheet1");
  if (sheet == null) {
    throw new Error("Sheet1 not found");
  }
  const helper = makeTableHelper(spreadsheet.getId(), sheet, "Table1");
  if (helper == null) {
    throw new Error("Sheet1:Table1 not found");
  }
  let range = helper.getColumnRange("a");
  if (range == null) {
    throw new Error("Sheet1:Table1:a not found");
  }
  console.log("column a", JSON.stringify(range.getValues(), null, 2));
  range = helper.getColumnRange("d");
  if (range == null) {
    throw new Error("Sheet1:Table1:d not found");
  }
  console.log("column d", JSON.stringify(range.getValues(), null, 2));
  const helper2 =
    makeTableHelper(spreadsheet.getId(), sheet, "Table2") ??
    (() => {
      throw new Error("Sheet1:Table2 not found");
    })();
  let range2 =
    helper2.getColumnRange("g") ??
    (() => {
      throw new Error("Sheet1:Table2:g not found");
    })();
  console.log(`range2 before: ${rangeStr(range2)}`);
  helper2.ensureRowCount(13);
  range2 =
    helper2.getColumnRange("g") ??
    (() => {
      throw new Error("Sheet1:Table2:g not found");
    })();
  console.log(`range2 after: ${rangeStr(range2)}`);
}

const HOLDINGS_SHEET_NAME = "Holdings";
const HOLDINGS_TABLE_NAME = "Holdings";
const ASSET_SHEET_NAME = "Asset Setup";
function updateSpreadsheet(
  holdingsArray: HoldingEntry[],
  classifications: Classifications,
  accounts: Account[],
) {
  const accountMap = new Map(accounts.map((a) => [a.id, a.name]));
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = spreadsheet.getId();
  const assetsSheet = spreadsheet.getSheetByName(ASSET_SHEET_NAME);
  const holdingsSheet = spreadsheet.getSheetByName(HOLDINGS_SHEET_NAME);

  if (!assetsSheet || !holdingsSheet) {
    throw new Error("classifications and/or holdings sheet missing");
  }
  const holdingsTableHelper =
    makeTableHelper(spreadsheetId, holdingsSheet, HOLDINGS_TABLE_NAME) ??
    (() => {
      throw new Error(
        `${HOLDINGS_SHEET_NAME}:${HOLDINGS_TABLE_NAME} not found`,
      );
    })();
  holdingsTableHelper.ensureRowCount(holdingsArray.length);
  // TODO: Do something for realz! Stomp on Account (needs a lookup), Asset, Shares; and for manual holdings, Price.
  // For everything else just write back what was there, but clear Account and Asset for excess rows (e.g. table is larger than necessary)
  // Write it back
  const getColumnRange = (columnName: string) => {
    const result = holdingsTableHelper.getColumnRange(columnName);
    if (result === undefined) {
      throw new Error(`Holdings column ${columnName} not found`);
    }
    return result;
  };
  const padding = holdingsTableHelper.getNumRows() - holdingsArray.length;
  const hAccountCol = holdingsArray.map(
    (h) => accountMap.get(h.userAccountId) ?? `Unknown: ${h.userAccountId}`,
  );
  const hAssetCol = holdingsArray.map((h) => h.ticker);
  const hSharesCol = holdingsArray.map((h) => h.quantity);
  const hPriceCol = holdingsArray.map((h) => (h.cusip ? null : h.price));
  console.log(`hPriceCol=${JSON.stringify(hPriceCol, null, 2)}`);

  getColumnRange("Account").setValues([
    ...hAccountCol.map((v) => [formatAccountName(v)]),
    ...Array(padding).fill([""]),
  ]);
  getColumnRange("Asset").setValues([
    ...hAssetCol.map((v) => [v]),
    ...Array(padding).fill([""]),
  ]);
  getColumnRange("Shares").setValues([
    ...hSharesCol.map((v) => [v]),
    ...Array(padding).fill([""]),
  ]);
  // TODO: don't change it here, change it in the Asset Tracker sheet !!!
  const PRICE_FORMULA =
    '=IF(Holdings[Asset] <> "",XLOOKUP(Holdings[Asset],Assets[Ticker Text],Assets[Price]),"")';
  const priceA1Ref = getColumnRange("Price").getA1Notation();
  Sheets.Spreadsheets.Values.update(
    {
      values: [
        ...hPriceCol.map((price) => [price == null ? PRICE_FORMULA : price]),
        ...Array(padding).fill([PRICE_FORMULA]),
      ],
    },
    spreadsheetId,
    priceA1Ref,
    { valueInputOption: "USER_ENTERED" },
  );
  // TODO: Classifications!
  console.log(Object.entries(classifications).length);
}
export function doPost(event: GoogleAppsScript.Events.DoPost) {
  try {
    const {
      version: { major, minor },
      holdings,
      classifications,
      accounts,
    } = JSON.parse(event.postData.contents) as PostPayload;
    console.log(`API version: ${major}.${minor}`);
    const supported = { major: 0, minor: 4 };
    if (major !== supported.major || minor < supported.minor) {
      throw new Error(
        `data version ${major}.${minor} not supported, expected at least ${supported.major}.${supported.minor}`,
      );
    }

    let lock: GoogleAppsScript.Lock.Lock | undefined;
    try {
      lock = LockService.getScriptLock();
      lock.waitLock(30_000);
      updateSpreadsheet(holdings, classifications, accounts);
    } finally {
      lock?.releaseLock();
    }
    const responseBody: PostResponse = {
      success: true,
      message: "Data received",
    };
    return ContentService.createTextOutput(
      JSON.stringify(responseBody),
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    const responseBody: PostResponse = {
      success: false,
      error: e instanceof Error ? e.message : String(e),
    };
    return ContentService.createTextOutput(
      JSON.stringify(responseBody),
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

export function placeholder(): boolean {
  return true;
}
