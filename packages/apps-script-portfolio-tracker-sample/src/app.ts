// Make changes in https://github.com/righteffort/empower-poster/packages/apps-script-portfolio-tracker-sample/src/
import type {
  PostPayload,
  PostResponse,
  HoldingEntry,
  Classifications,
  Account,
} from "@righteffort/empower-poster-types";

import { makeTableHelper } from "./sheet-utils.js";

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
  console.log(
    `range2 before: ${JSON.stringify({ row: range2.getRow(), column: range2.getColumn(), numRows: range2.getNumRows(), numColumns: range2.getNumColumns() })}`,
  );
  helper2.ensureRowCount(13);
  range2 =
    helper2.getColumnRange("g") ??
    (() => {
      throw new Error("Sheet1:Table2:g not found");
    })();
  console.log(
    `range2 after: ${JSON.stringify({ row: range2.getRow(), column: range2.getColumn(), numRows: range2.getNumRows(), numColumns: range2.getNumColumns() })}`,
  );
}

const HOLDINGS_SHEET_NAME = "Holdings";
const HOLDINGS_TABLE_NAME = "Holdings";
const ASSET_SHEET_NAME = "Asset Setup";
function updateSpreadsheet(
  holdingsArray: HoldingEntry[],
  classifications: Classifications,
  accounts: Account[],
) {
  console.log(`${accounts.length} accounts`); // TODO remove
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const assetsSheet = spreadsheet.getSheetByName(ASSET_SHEET_NAME);
  const holdingsSheet = spreadsheet.getSheetByName(HOLDINGS_SHEET_NAME);

  if (!assetsSheet || !holdingsSheet) {
    throw new Error("classifications and/or holdings sheet missing");
  }
  const holdings = new Map(holdingsArray.map((h) => [h.ticker, h]));
  const holdingsTableHelper =
    makeTableHelper(spreadsheet.getId(), holdingsSheet, HOLDINGS_TABLE_NAME) ??
    (() => {
      throw new Error(
        `${HOLDINGS_SHEET_NAME}:${HOLDINGS_TABLE_NAME} not found`,
      );
    })();
  holdingsTableHelper.ensureRowCount(holdings.size);
  const holdingsRange = holdingsTableHelper.getRange();
  // get column headers please
  const values = holdingsRange.getValues();
  // TODO: Do something for realz! Stomp on Account (needs a lookup), Asset, Shares; and for manual holdings, Price.
  // For everything else just write back what was there, but clear Account and Asset for excess rows (e.g. table is larger than necessary)
  // Write it back
  holdingsRange.setValues(values);
  console.log(Object.entries(classifications).length);
  let lock: GoogleAppsScript.Lock.Lock | undefined;
  try {
    lock = LockService.getScriptLock();
    lock.waitLock(30_000);
    // TODO: updates go here
  } finally {
    lock?.releaseLock();
  }
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

    updateSpreadsheet(holdings, classifications, accounts);
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
