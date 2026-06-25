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

const ASSET_SETUP_SHEET_NAME = "Asset Setup";
const ASSETS_TABLE_NAME = "Assets";
const ASSET_CLASSES_TABLE_NAME = "Asset Classes";
const CLASS_CATEGORIES_TABLE_NAME = "Class Categories";
const HOLDINGS_SHEET_NAME = "Holdings";
const HOLDINGS_TABLE_NAME = "Holdings";

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
    const assetSetupSheet = this.spreadsheet.getSheetByName(
      ASSET_SETUP_SHEET_NAME,
    );
    const holdingsSheet = this.spreadsheet.getSheetByName(HOLDINGS_SHEET_NAME);
    if (!assetSetupSheet || !holdingsSheet) {
      throw new Error("classifications and/or holdings sheet missing");
    }
    this.assetSetupSheet = assetSetupSheet;
    this.holdingsSheet = holdingsSheet;
  }
  updateSpreadsheet() {
    this.updateAssetsAndCategories();
    this.updateHoldings();
  }
  private updateHoldings() {
    const helper =
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
    helper.ensureRowCount(this.holdingsArray.length);
    const getColumnRange = (columnName: string) => {
      const result = helper.getColumnRange(columnName);
      if (result === undefined) {
        throw new Error(`Holdings column ${columnName} not found`);
      }
      return result;
    };
    const padding = helper.getNumRows() - this.holdingsArray.length;
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

  private updateAssetsAndCategories() {
    const TICKER_COLUMN_NAME = "Ticker";
    const NAME_COLUMN_NAME = "Name";
    const CLASS_COLUMN_NAME = "Class";
    const CLASS_PCT_COLUMN_NAME = "Class Pct";
    const PRICE_COLUMN_NAME = "Price";
    const flatAssets = Object.entries(this.classifications).flatMap(
      ([ticker, v]) =>
        v.map((c) => ({
          ticker,
          classes: [c.classes[0], c.classes[1] || c.classes[0]],
          fraction: c.fraction,
        })),
    );
    const assetClasses = Array.from(
      new Set(flatAssets.map((a) => a.classes.join("\0"))),
    )
      .sort()
      .map((v) => v.split("\0"));
    const classCategories = Array.from(
      new Set(assetClasses.map((v) => v[0] || "")),
    ).sort();
    this.updateAssetClasses(assetClasses);
    this.updateClassCategories(classCategories);
    const assetRows = flatAssets.map((a) => ({
      ticker: a.ticker,
      class: a.classes[1],
      fraction: a.fraction,
    }));
    const helper =
      makeTableHelper(
        this.spreadsheetId,
        this.assetSetupSheet,
        ASSETS_TABLE_NAME,
      ) ??
      (() => {
        throw new Error(
          `${ASSET_SETUP_SHEET_NAME}:${ASSETS_TABLE_NAME} not found`,
        );
      })();
    const priceColRange = helper.getColumnRange(PRICE_COLUMN_NAME);
    if (!priceColRange) {
      throw new Error(`Assets column ${PRICE_COLUMN_NAME} not found`);
    }
    const nameColRange = helper.getColumnRange(NAME_COLUMN_NAME);
    if (!nameColRange) {
      throw new Error(`Assets column ${NAME_COLUMN_NAME} not found`);
    }

    const priceFormula = helper.getColumnDefaultFormula(PRICE_COLUMN_NAME);
    const nameFormula = helper.getColumnDefaultFormula(NAME_COLUMN_NAME);
    if (!priceFormula || !nameFormula) {
      throw new Error(
        `Default formula not found for ${PRICE_COLUMN_NAME} and/or ${NAME_COLUMN_NAME}`,
      );
    }

    // make sure the table is big enough
    helper.ensureRowCount(assetRows.length);
    // fill in ticker, class, pct (straightforward); and name and pct (mix of formulas -- no change -- or actual value, when there is no cusip).
    const getColumnRange = (columnName: string) => {
      const result = helper.getColumnRange(columnName);
      if (result === undefined) {
        throw new Error(`Assets column ${columnName} not found`);
      }
      return result;
    };
    const padding = helper.getNumRows() - assetRows.length;
    const holdings = new Map<string, HoldingEntry>(
      this.holdingsArray.map((h) => [h.ticker, h]),
    );

    const tickerCol = assetRows.map((r) => [r.ticker]);
    const classCol = assetRows.map((r) => [r.class]);
    const classPctCol = assetRows.map((r) => [r.fraction]);
    const priceValues = assetRows.map((r) => {
      const h = holdings.get(r.ticker);
      return h?.cusip ? null : h?.price;
    });
    const priceFormulas = priceValues.map((p) => [
      p == null ? priceFormula : toLiteralFormula(p),
    ]);
    const nameValues = assetRows.map((r) =>
      holdings.get(r.ticker)?.cusip ? null : r.ticker,
    );
    const nameFormulas = nameValues.map((n) => [
      n == null ? nameFormula : toLiteralFormula(n),
    ]);
    // TODO: error checking
    getColumnRange(TICKER_COLUMN_NAME).setValues([
      ...tickerCol,
      ...Array(padding).fill([""]),
    ]);
    getColumnRange(NAME_COLUMN_NAME).setValues([
      ...nameFormulas,
      ...Array(padding).fill([priceFormula]),
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
      ...Array(padding).fill([priceFormula]),
    ]);
  }

  private updateAssetClasses(assetClasses: string[][]) {
    const helper = makeTableHelper(
      this.spreadsheetId,
      this.assetSetupSheet,
      ASSET_CLASSES_TABLE_NAME,
    );
    if (!helper) {
      throw new Error(`${ASSET_CLASSES_TABLE_NAME} not found`);
    }
    const range = helper.getRange();
    range.clear();
    helper.ensureRowCount(assetClasses.length);
    helper.getRange().setValues(assetClasses.map((cs) => [cs[1], cs[0]]));
  }

  private updateClassCategories(classCategories: string[]) {
    const helper = makeTableHelper(
      this.spreadsheetId,
      this.assetSetupSheet,
      CLASS_CATEGORIES_TABLE_NAME,
    );
    if (!helper) {
      throw new Error(`${CLASS_CATEGORIES_TABLE_NAME} not found`);
    }
    const NAME_COLUMN_NAME = "Name";
    const range = helper.getColumnRange(NAME_COLUMN_NAME);
    if (range == null) {
      throw new Error(
        `${CLASS_CATEGORIES_TABLE_NAME}:${NAME_COLUMN_NAME} not found`,
      );
    }
    const existingCategoryNames = range.getValues().map((r) => r[0]);
    const missingCategoryNames = classCategories.filter(
      (c) => !existingCategoryNames.includes(c),
    );
    if (missingCategoryNames.length === 0) {
      return;
    }
    helper.ensureRowCount(
      existingCategoryNames.length + missingCategoryNames.length,
    );
    const missingRange = this.assetSetupSheet.getRange(
      range.getRow() + existingCategoryNames.length,
      range.getColumn(),
      missingCategoryNames.length,
      range.getNumColumns(),
    );
    missingRange.setValues(missingCategoryNames.map((c) => [c]));
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
    const supported = { major: 0, minor: 5 };
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
