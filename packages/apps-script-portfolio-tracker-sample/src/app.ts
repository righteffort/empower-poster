// Make changes in https://github.com/righteffort/empower-poster/packages/apps-script-portfolio-tracker-sample/src/
import type {
  PostPayload,
  PostResponse,
  HoldingEntry,
  Classifications,
  Account,
} from "@righteffort/empower-poster-types";

// TODO: remove
// @ts-expect-error: Used from Apps Script
function playground() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Sheet1");
  if (sheet == null) {
    throw new Error("Sheet1 not found");
  }
  const table =
    getTable(spreadsheet.getId(), "Sheet1", "Table2") ??
    (() => {
      throw new Error("Sheet1:Table2 not found");
    })();
  const tableRange = getTableGridRange(table);
  console.log(tableRange);
  ensureTableRowCount(sheet, tableRange, 13);
  const newTable =
    getTable(spreadsheet.getId(), "Sheet1", "Table2") ??
    (() => {
      throw new Error("Sheet1:Table2 not found");
    })();
  const newHoldingsTableRange = getTableGridRange(newTable);
  console.log(newHoldingsTableRange);
}

function getTable(
  spreadsheetId: string,
  sheetTitle: string,
  tableName: string,
): Sheets.SafeTable | undefined {
  const gspreadsheet = Sheets.Spreadsheets.get(spreadsheetId, {
    fields:
      "sheets(properties(sheetId,gridProperties(rowCount,columnCount)),tables(name,range,columnProperties))",
  });
  const gsheet = gspreadsheet.sheets?.filter(
    (s) => s.properties?.title === sheetTitle,
  )[0];
  const unsafeTable = gsheet?.tables?.filter((t) => t.name === tableName)[0];
  if (unsafeTable === undefined) {
    return unsafeTable;
  }
  return Object.fromEntries(
    Object.entries(unsafeTable).map(([k, v]) => {
      if (v == null) throw new Error(`${k} is null/undefined`);
      return [k, v];
    }),
  ) as Sheets.SafeTable;
}
function getTableGridRange(table: Sheets.SafeTable): Sheets.SafeGridRange {
  const gridRange = Object.fromEntries(
    Object.entries(table.range).map(([k, v]) => {
      if (v == null) throw new Error(`${k} is null/undefined`);
      return [k, v];
    }),
  ) as Sheets.SafeGridRange;
  // Fail if gridRange has zero data rows or zero data columns
  if (
    gridRange.startRowIndex + 1 >= gridRange.endRowIndex || // +1 for header row
    gridRange.startColumnIndex >= gridRange.endColumnIndex
  ) {
    throw new Error(
      `Table ${table.name} Can't deal with gridRange of ${JSON.stringify(gridRange)}`,
    );
  }
  return gridRange;
}

const HOLDINGS_SHEET_NAME = "Holdings";
const HOLDINGS_TABLE_NAME = "Holdings";
const ASSET_SHEET_NAME = "Asset Setup";
function updateSpreadsheet(
  holdingsArray: HoldingEntry[],
  classifications: Classifications,
  accounts: Account[]
) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const assetsSheet = spreadsheet.getSheetByName(ASSET_SHEET_NAME);
  const holdingsSheet = spreadsheet.getSheetByName(HOLDINGS_SHEET_NAME);

  if (!assetsSheet || !holdingsSheet) {
    throw new Error("classifications and/or holdings sheet missing");
  }
  const holdings = new Map(holdingsArray.map((h) => [h.ticker, h]));
  let holdingsTable =
    getTable(spreadsheet.getId(), HOLDINGS_SHEET_NAME, HOLDINGS_TABLE_NAME) ??
    (() => {
      throw new Error(
        `${HOLDINGS_SHEET_NAME}:${HOLDINGS_TABLE_NAME} not found`,
      );
    })();
  let holdingsTableGridRange = getTableGridRange(holdingsTable);
  ensureTableRowCount(holdingsSheet, holdingsTableGridRange, holdings.size);
  holdingsTable = getTable(
    spreadsheet.getId(),
    HOLDINGS_SHEET_NAME,
    HOLDINGS_TABLE_NAME,
  ) as Sheets.SafeTable;
  holdingsTableGridRange = getTableGridRange(holdingsTable);
  // TODO: factor this out into a function
  const range = holdingsSheet.getRange(
    (holdingsTableGridRange.startRowIndex as number) + 2,
    (holdingsTableGridRange.startColumnIndex as number) + 1,
    (holdingsTableGridRange.endRowIndex as number) -
      (holdingsTableGridRange.startRowIndex as number) -
      1,
    (holdingsTableGridRange.endColumnIndex as number) -
      (holdingsTableGridRange.startColumnIndex as number),
  );
  if (range.getNumRows() < holdings.size) {
    throw new Error("failed to make space for holdings, or else a logic bug");
  }
  // get column headers please
  let values = range.getValues();
  for (let i = 0; i < Math.max(values.length, holdings.size); i++) {
    // do something here
    values = [[7]];
  }
  // Write it back
  // Stomp on Account (needs a lookup), Asset, Shares; and for manual holdings, Price.
  // For everything else just write back what, but clear Account and Asset for excess rows (e.g. table is larger than necessary)
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

/**
 * Make sure the table has at least rowsNeeded rows
 */
function ensureTableRowCount(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  gridRange: Sheets.SafeGridRange,
  rowsNeeded: number,
) {
  // Convert to SpreadsheetApp: 1-indexed and closed-closed ranges
  const holdingsFirstRow = gridRange.startRowIndex + 2; // +1 for header row
  const holdingsLastRow = gridRange.endRowIndex;
  const [holdingsFirstColumn, holdingsLastColumn] = [
    gridRange.startColumnIndex + 1,
    gridRange.endColumnIndex,
  ];
  let holdingsNumRows = holdingsLastRow - holdingsFirstRow + 1;
  const holdingsNumColumns = holdingsLastColumn - holdingsFirstColumn + 1;
  console.log(
    JSON.stringify(
      {
        holdingsFirstRow,
        holdingsLastRow,
        holdingsNumRows,
        holdingsFirstColumn,
        holdingsLastColumn,
        holdingsNumColumns,
      },
      null,
      2,
    ),
  );
  // Extend the table until it will accomodate all the rows
  let totalRowsToAdd = rowsNeeded - holdingsNumRows;
  while (totalRowsToAdd > 0) {
    const rowsToAdd = Math.min(totalRowsToAdd, holdingsNumRows);
    console.log(
      `inserting here: ${JSON.stringify({ holdingsFirstRow, holdingsFirstColumn, rowsToAdd, holdingsNumColumns }, null, 2)}`,
    );
    const range = sheet.getRange(
      holdingsFirstRow,
      holdingsFirstColumn,
      rowsToAdd,
      holdingsNumColumns,
    );
    range.insertCells(SpreadsheetApp.Dimension.ROWS);
    holdingsNumRows += rowsToAdd;
    totalRowsToAdd -= rowsToAdd;
  }
}

// @ts-expect-error: Used from Apps Script
function doPost(event: GoogleAppsScript.Events.DoPost) {
  try {
    const { version: {major, minor}, holdings, classifications, accounts } = JSON.parse(
      event.postData.contents,
    ) as PostPayload;
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
