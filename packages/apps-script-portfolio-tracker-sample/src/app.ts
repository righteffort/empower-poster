// Make changes in https://github.com/righteffort/empower-poster/packages/apps-script-portfolio-tracker-sample/src/
import type {
  PostPayload,
  PostResponse,
  HoldingEntry,
  Classification,
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
  console.log(`range2 before: ${rangeStr(range2)}`);
  helper2.ensureRowCount(13);
  range2 =
    helper2.getColumnRange("g") ??
    (() => {
      throw new Error("Sheet1:Table2:g not found");
    })();
  console.log(`range2 after: ${rangeStr(range2)}`);
}

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

function toLiteralFormula(val: string | boolean | number) {
  if (typeof val === "string") return '="' + val.replace(/"/g, '""') + '"';
  if (typeof val === "boolean") return val ? "=TRUE" : "=FALSE";
  return "=" + val;
}

function isLiteralFormula(formula: string): boolean {
  return /^=(?:"(?:[^"]|"")*"|TRUE|FALSE|-?\d+(?:\.\d+)?)$/.test(formula);
}

const HOLDINGS_SHEET_NAME = "Holdings";
const HOLDINGS_TABLE_NAME = "Holdings";
const ASSET_SHEET_NAME = "Asset Setup"; // TODO rename variable
const ASSETS_TABLE_NAME = "Assets";

class AssetAllocationUpater {
  private readonly holdingsArray: HoldingEntry[];
  private readonly classifications: Record<string, Classification[]>;
  private readonly accountMap: Map<number, string>;
  private readonly spreadsheetId: string;
  private readonly spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  private readonly assetSetupSheet: GoogleAppsScript.Spreadsheet.Sheet;
  private readonly holdingsSheet: GoogleAppsScript.Spreadsheet.Sheet;

  constructor(
    holdingsArray: HoldingEntry[],
    classifications: Classifications,
    accounts: Account[],
  ) {
    this.holdingsArray = holdingsArray;
    this.classifications = classifications;
    this.accountMap = new Map(accounts.map((a) => [a.id, a.name]));
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    this.spreadsheetId = this.spreadsheet.getId();
    const assetSetupSheet = this.spreadsheet.getSheetByName(ASSET_SHEET_NAME);
    const holdingsSheet = this.spreadsheet.getSheetByName(HOLDINGS_SHEET_NAME);
    if (!assetSetupSheet || !holdingsSheet) {
      throw new Error("classifications and/or holdings sheet missing");
    }
    this.assetSetupSheet = assetSetupSheet;
    this.holdingsSheet = holdingsSheet;
  }
  updateSpreadsheet() {
    this.updateAssetSetup();
    this.updateHoldings();
  }
  private updateHoldings() {
    const tableHelper =
      makeTableHelper(
        this.spreadsheetId,
        this.holdingsSheet,
        HOLDINGS_TABLE_NAME,
      ) ??
      (() => {
        throw new Error(
          `${HOLDINGS_SHEET_NAME}:${HOLDINGS_TABLE_NAME} not found`,
        );
      })();
    tableHelper.ensureRowCount(this.holdingsArray.length);
    const getColumnRange = (columnName: string) => {
      const result = tableHelper.getColumnRange(columnName);
      if (result === undefined) {
        throw new Error(`Holdings column ${columnName} not found`);
      }
      return result;
    };
    const padding = tableHelper.getNumRows() - this.holdingsArray.length;
    const accountCol = this.holdingsArray.map(
      (h) =>
        this.accountMap.get(h.userAccountId) ?? `Unknown: ${h.userAccountId}`,
    );
    const assetCol = this.holdingsArray.map((h) => h.ticker);
    const shareCol = this.holdingsArray.map((h) => h.quantity);

    const ACCOUNT_COLUMN_NAME = "Account";
    const ASSET_COLUMN_NAME = "Asset";
    const SHARES_COLUMN_NAME = "Shares";
    // TODO: error checking
    getColumnRange(ACCOUNT_COLUMN_NAME).setValues([
      ...accountCol.map((v) => [formatAccountName(v)]),
      ...Array(padding).fill([""]),
    ]);
    getColumnRange(ASSET_COLUMN_NAME).setValues([
      ...assetCol.map((v) => [v]),
      ...Array(padding).fill([""]),
    ]);
    getColumnRange(SHARES_COLUMN_NAME).setValues([
      ...shareCol.map((v) => [v]),
      ...Array(padding).fill([""]),
    ]);
  }

  private updateAssetSetup() {
    const TICKER_COLUMN_NAME = "Ticker";
    const NAME_COLUMN_NAME = "Name";
    const CLASS_COLUMN_NAME = "Class";
    const CLASS_PCT_COLUMN_NAME = "Class Pct";
    const PRICE_COLUMN_NAME = "Price";
    const assetRows = Object.entries(this.classifications).flatMap(
      ([ticker, v]) =>
        v.map((classification) => ({
          ticker,
          class: classification.classes[1] || classification.classes[0],
          fraction: classification.fraction,
        })),
    );
    const tableHelper =
      makeTableHelper(
        this.spreadsheetId,
        this.assetSetupSheet,
        ASSETS_TABLE_NAME,
      ) ??
      (() => {
        throw new Error(`${ASSET_SHEET_NAME}:${ASSETS_TABLE_NAME} not found`);
      })();
    const tableRange = tableHelper.getRange();
    const priceColRange = tableHelper.getColumnRange(PRICE_COLUMN_NAME);
    if (!priceColRange) {
      throw new Error(`Assets column ${PRICE_COLUMN_NAME} not found`);
    }
    // delete all existing rows with manual 'literals' in price column
    priceColRange
      .getFormulas()
      .flatMap((r, i) => (isLiteralFormula(r[0] ?? "") ? [i] : []))
      .reverse()
      .forEach((i) => {
        this.assetSetupSheet
          .getRange(
            priceColRange.getRow() + i,
            tableRange.getColumn(),
            1,
            tableRange.getNumColumns(),
          )
          .deleteCells(SpreadsheetApp.Dimension.ROWS);
      });
    // to be safe, delete all existing rows that have genuine literals in the price column
    priceColRange
      .getFormulas()
      .flatMap((r, i) => (r[0] === "" ? [i] : []))
      .reverse()
      .forEach((i) => {
        this.assetSetupSheet
          .getRange(
            priceColRange.getRow() + i,
            tableRange.getColumn(),
            1,
            tableRange.getNumColumns(),
          )
          .deleteCells(SpreadsheetApp.Dimension.ROWS);
      });
    // make sure the table is big enough
    tableHelper.ensureRowCount(assetRows.length);
    // fill in ticker, class, pct (straightforward); and name and pct (mix of formulas -- no change -- or actual value, when there is no cusip).
    const getColumnRange = (columnName: string) => {
      const result = tableHelper.getColumnRange(columnName);
      if (result === undefined) {
        throw new Error(`Assets column ${columnName} not found`);
      }
      return result;
    };
    const padding = tableHelper.getNumRows() - assetRows.length;
    const holdings = new Map<string, HoldingEntry>(
      this.holdingsArray.map((h) => [h.ticker, h]),
    );

    // TODO: would be nice not to hardcode the defaults
    const PRICE_FORMULA = "=ASSETPRICE(Assets[Ticker])";
    const NAME_FORMULA = "=ASSETNAME(Assets[Ticker])";

    const tickerCol = assetRows.map((h) => h.ticker);
    const classCol = assetRows.map((h) => h.class);
    const classPctCol = assetRows.map((h) => h.fraction);
    const priceValues = assetRows.map((h) =>
      holdings.get(h.ticker)?.cusip ? null : holdings.get(h.ticker)?.price,
    ); // TODO: this is too permissive
    const priceFormulas = priceValues.map((p) =>
      p == null ? PRICE_FORMULA : toLiteralFormula(p),
    );
    const nameValues = assetRows.map((h) =>
      holdings.get(h.ticker)?.cusip ? null : h.ticker,
    );
    const nameFormulas = nameValues.map((n) =>
      n == null ? NAME_FORMULA : toLiteralFormula(n),
    );
    // TODO: error checking
    getColumnRange(TICKER_COLUMN_NAME).setValues([
      ...tickerCol,
      ...Array(padding).fill([""]),
    ]);
    getColumnRange(NAME_COLUMN_NAME).setValues([
      ...nameFormulas,
      ...Array(padding).fill([PRICE_FORMULA]),
    ]);
    getColumnRange(CLASS_COLUMN_NAME).setValues([
      ...classCol,
      ...Array(padding).fill([""]),
    ]);
    getColumnRange(CLASS_PCT_COLUMN_NAME).setValues([
      ...classPctCol,
      ...Array(padding).fill([""]),
    ]);
    getColumnRange(PRICE_COLUMN_NAME).setValues([
      ...priceFormulas,
      ...Array(padding).fill([PRICE_FORMULA]),
    ]);
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

    let lock: GoogleAppsScript.Lock.Lock | undefined;
    try {
      lock = LockService.getScriptLock();
      lock.waitLock(30_000);
      new AssetAllocationUpater(
        holdings,
        classifications,
        accounts,
      ).updateSpreadsheet();
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
