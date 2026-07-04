// Make changes in https://github.com/righteffort/empower-poster/packages/apps-script-portfolio-tracker/src/
import type {
  PostPayload,
  PostResponse,
  HoldingEntry,
  Classification,
  Classifications,
  Account,
} from "@righteffort/empower-poster-types";

import { makeTableHelper } from "./sheet-utils.js";

function fail(message = "Internal error"): never {
  throw new Error(message);
}

/**
 * Add a leading apostrophe to purely numeric account names and types.
 */
function formatAccountNameOrType(accountName: string): string {
  if (accountName.search(/^[0-9]+$/) === 0) {
    return `'${accountName}`;
  }
  return accountName;
}

function toLiteralFormula(val: string | boolean | number) {
  if (typeof val === "string") return '="' + val.replace(/"/g, '""') + '"';
  if (typeof val === "boolean") return val ? "=TRUE" : "=FALSE";
  return "=" + val;
}

const ACCOUNT_SETUP_SHEET_NAME = "Account Setup";
const INSTITUTIONS_TABLE_NAME = "Institutions";
const ACCOUNTS_TABLE_NAME = "Accounts";
const ASSET_SETUP_SHEET_NAME = "Asset Setup";
const ASSETS_TABLE_NAME = "Assets";
const ASSET_CLASSES_TABLE_NAME = "Asset Classes";
const CLASS_CATEGORIES_TABLE_NAME = "Class Categories";
const HOLDINGS_SHEET_NAME = "Holdings";
const HOLDINGS_TABLE_NAME = "Holdings";

class AssetAllocationUpater {
  private readonly holdingsArray: HoldingEntry[];
  private readonly classifications: Record<string, Classification[]>;
  private readonly accountMap: Map<number, Account>;
  private readonly spreadsheetId: string;
  private readonly spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  private readonly accountSetupSheet: GoogleAppsScript.Spreadsheet.Sheet;
  private readonly assetSetupSheet: GoogleAppsScript.Spreadsheet.Sheet;
  private readonly holdingsSheet: GoogleAppsScript.Spreadsheet.Sheet;

  constructor(
    holdingsArray: HoldingEntry[],
    classifications: Classifications,
    accounts: Account[],
  ) {
    const allAccounts = new Map(accounts.map((a) => [a.id, a]));
    const accountIds = [...new Set(holdingsArray.map((h) => h.userAccountId))];
    this.accountMap = new Map(
      accountIds.map((id) => [
        id,
        allAccounts.get(id) || fail("Unreachable code"),
      ]),
    );
    this.holdingsArray = holdingsArray;
    this.classifications = classifications;
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    this.spreadsheetId = this.spreadsheet.getId();
    this.accountSetupSheet =
      this.spreadsheet.getSheetByName(ACCOUNT_SETUP_SHEET_NAME) ||
      fail(`'${ACCOUNT_SETUP_SHEET_NAME}' sheet missing`);
    this.assetSetupSheet =
      this.spreadsheet.getSheetByName(ASSET_SETUP_SHEET_NAME) ||
      fail(`'${ASSET_SETUP_SHEET_NAME}' sheet missing`);
    this.holdingsSheet =
      this.spreadsheet.getSheetByName(HOLDINGS_SHEET_NAME) ||
      fail(`'${HOLDINGS_SHEET_NAME}' sheet missing`);
  }
  updateSpreadsheet() {
    this.updateAccountSetup();
    this.updateAssetSetup();
    this.updateHoldings();
  }
  private updateAccountSetup() {
    this.updateInstitutions();
    this.updateAccounts(true);
  }
  private updateInstitutions() {
    const helper =
      makeTableHelper(
        this.spreadsheetId,
        this.accountSetupSheet,
        INSTITUTIONS_TABLE_NAME,
      ) ?? fail(`${INSTITUTIONS_TABLE_NAME} table not found`);
    const NAME_COLUMN_NAME = "Name";
    const institutions = [
      ...new Set([...this.accountMap.values()].map((a) => a.firmName)),
    ].sort();
    helper.updateRowCount(institutions.length);
    const range =
      helper.getColumnRange(NAME_COLUMN_NAME) ??
      fail(`${HOLDINGS_TABLE_NAME}:${NAME_COLUMN_NAME} not found`);
    range.clearContent().setValues(institutions.map((i) => [i]));
  }
  private updateAccounts(pruneRows: boolean) {
    // only show accounts from empower
    // unconditionally show institution from empower (firmName)
    // don't mess with owner *at all*
    // preserve type if it is already there; if not leave blank
    // clear extras
    // this is probably the wrong place, but could guess at type->tax for user. seems silly though.
    const helper =
      makeTableHelper(
        this.spreadsheetId,
        this.accountSetupSheet,
        ACCOUNTS_TABLE_NAME,
      ) || fail(`${ACCOUNTS_TABLE_NAME} not found`);
    const oldTableRange =
      helper.getRange() ||
      fail(`${ACCOUNTS_TABLE_NAME} table missing or has zero data rows`);
    const NAME_COLUMN_NAME = "Name";
    const TYPE_COLUMN_NAME = "Type";
    const INSTITUTION_COLUMN_NAME = "Institution";
    const OWNER_COLUMN_NAME = "Owner";
    const checkPosition = (columnName: string, offset: number) => {
      if (
        (helper.getColumnRange(columnName)?.getColumn() ||
          fail(`${ACCOUNTS_TABLE_NAME}:${columnName} missing`)) -
          oldTableRange.getColumn() !=
        offset
      ) {
        throw new Error(
          `${ACCOUNTS_TABLE_NAME}:${columnName} in wrong position`,
        );
      }
    };
    checkPosition(NAME_COLUMN_NAME, 0);
    checkPosition(TYPE_COLUMN_NAME, 1);
    checkPosition(INSTITUTION_COLUMN_NAME, 2);
    checkPosition(OWNER_COLUMN_NAME, 3);
    const oldRange = this.accountSetupSheet.getRange(
      oldTableRange.getRow(),
      oldTableRange.getColumn(),
      oldTableRange.getNumRows(),
      4,
    );
    const currentValues = new Map(
      oldRange.getValues().map((r) => [
        r[0],
        {
          type: r[1],
          owner: r[3],
        },
      ]),
    );
    const newValues = [...this.accountMap.values()]
      .sort((a: Account, b: Account) => a.name.localeCompare(b.name))
      .map((a) => [
        formatAccountNameOrType(a.name),
        formatAccountNameOrType(currentValues.get(a.name)?.type ?? ""),
        a.firmName,
        currentValues.get(a.name)?.owner ?? "",
      ]);
    helper.updateRowCount(newValues.length, pruneRows);
    const newTableRange = helper.getRange() || fail();
    const newRange = this.accountSetupSheet.getRange(
      newTableRange.getRow(),
      newTableRange.getColumn(),
      newTableRange.getNumRows(),
      4,
    );
    newRange.setValues(newValues);
  }
  private updateHoldings() {
    const holdingRows = this.holdingsArray
      .map((h) => ({
        accountName: this.accountMap.get(h.userAccountId)?.name ?? "Unknown",
        ticker: h.ticker,
        quantity: h.quantity,
      }))
      .sort(
        (a, b) =>
          a.accountName.localeCompare(b.accountName) ||
          a.ticker.localeCompare(b.ticker),
      );
    const helper =
      makeTableHelper(
        this.spreadsheetId,
        this.holdingsSheet,
        HOLDINGS_TABLE_NAME,
      ) ?? fail(`${HOLDINGS_SHEET_NAME}:${HOLDINGS_TABLE_NAME} not found`);
    helper.updateRowCount(holdingRows.length, true);
    const getColumnRange = (columnName: string) => {
      const result = helper.getColumnRange(columnName);
      if (result === undefined) {
        throw new Error(`Holdings column ${columnName} not found`);
      }
      return result;
    };
    const accountCol = holdingRows.map((h) => h.accountName);
    const assetCol = holdingRows.map((h) => h.ticker);
    const shareCol = holdingRows.map((h) => h.quantity);

    const ACCOUNT_COLUMN_NAME = "Account";
    const ASSET_COLUMN_NAME = "Asset";
    const SHARES_COLUMN_NAME = "Shares";
    getColumnRange(ACCOUNT_COLUMN_NAME).setValues(
      accountCol.map((v) => [formatAccountNameOrType(v)]),
    );
    getColumnRange(ASSET_COLUMN_NAME).setValues(assetCol.map((v) => [v]));
    getColumnRange(SHARES_COLUMN_NAME).setValues(shareCol.map((v) => [v]));
  }

  private adjustClasses(classes: [string, string]): [string, string] {
    const [parent, child] = classes;
    if (!child) {
      return [parent, parent];
    }
    const m = parent.match(/^(Intl|U.S.) /);
    if (m) {
      return [parent, `${m[1]} ${child}`];
    }
    return [parent, child];
  }
  private updateAssetSetup() {
    const flatAssets = Object.entries(this.classifications).flatMap(
      ([ticker, v]) =>
        v.map((c) => ({
          ticker,
          classes: this.adjustClasses(c.classes),
          fraction: c.fraction,
        })),
    );
    const classCategories = [
      ...new Set(flatAssets.map((a) => a.classes[0] ?? "")),
    ];
    const classCategoryParentMap = this.updateClassCategories(classCategories);
    const assetClasses = Array.from(
      new Set(flatAssets.map((a) => a.classes.join("\0"))),
    )
      .map((v) => v.split("\0") as [string, string])
      .sort(
        (a: [string, string], b: [string, string]) =>
          (classCategoryParentMap.get(a[0]) ?? "").localeCompare(
            classCategoryParentMap.get(b[0]) ?? "",
          ) ||
          a[0].localeCompare(b[0]) ||
          a[1].localeCompare(b[1]),
      );
    this.updateAssetClasses(assetClasses);
    this.updateAssets(flatAssets, true);
  }
  updateAssets(
    flatAssets: {
      ticker: string;
      classes: [string, string];
      fraction: number;
    }[],
    pruneRows: boolean,
  ) {
    const TICKER_COLUMN_NAME = "Ticker";
    const NAME_COLUMN_NAME = "Name";
    const CLASS_COLUMN_NAME = "Class";
    const CLASS_PCT_COLUMN_NAME = "Class Pct";
    const PRICE_COLUMN_NAME = "Price";
    const assetRows = flatAssets
      .map((a) => ({
        ticker: a.ticker,
        class: a.classes[1],
        fraction: a.fraction,
      }))
      .sort((a, b) => a.ticker.localeCompare(b.ticker));
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

    helper.updateRowCount(assetRows.length, pruneRows);
    // fill in ticker, class, pct (straightforward); and name and pct (mix of formulas -- no change -- or actual value, when there is no cusip).
    const getColumnRange = (columnName: string) => {
      const result = helper.getColumnRange(columnName);
      if (result === undefined) {
        throw new Error(`Assets column ${columnName} not found`);
      }
      return result;
    };
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
    getColumnRange(TICKER_COLUMN_NAME).setValues(tickerCol);
    getColumnRange(NAME_COLUMN_NAME).setValues(nameFormulas);
    getColumnRange(CLASS_COLUMN_NAME).setValues(classCol);
    getColumnRange(CLASS_PCT_COLUMN_NAME).setValues(classPctCol);
    getColumnRange(PRICE_COLUMN_NAME).setValues(priceFormulas);
  }

  private updateAssetClasses(assetClasses: string[][]) {
    const helper = makeTableHelper(
      this.spreadsheetId,
      this.assetSetupSheet,
      ASSET_CLASSES_TABLE_NAME,
    );
    if (assetClasses.length === 0) {
      return;
    }
    if (!helper) {
      throw new Error(`${ASSET_CLASSES_TABLE_NAME} not found`);
    }
    helper.updateRowCount(assetClasses.length);
    const range = helper.getRange();
    if (range == null) {
      throw new Error(
        `Logic error, ${ASSET_CLASSES_TABLE_NAME} has zero data rows`,
      );
    }

    range.clearContent().setValues(assetClasses.map((cs) => [cs[1], cs[0]]));
  }

  private updateClassCategories(incomingClassCategories: string[]) {
    const helper = makeTableHelper(
      this.spreadsheetId,
      this.assetSetupSheet,
      CLASS_CATEGORIES_TABLE_NAME,
    );
    if (!helper) {
      throw new Error(`${CLASS_CATEGORIES_TABLE_NAME} not found`);
    }
    const NAME_COLUMN_NAME = "Name";
    const table_range = helper.getRange();
    if (table_range == null) {
      throw new Error(`${CLASS_CATEGORIES_TABLE_NAME} not found`);
    }
    const name_range = helper.getColumnRange(NAME_COLUMN_NAME);
    if (name_range == null) {
      throw new Error(
        `${CLASS_CATEGORIES_TABLE_NAME}:${NAME_COLUMN_NAME} not found`,
      );
    }
    if (
      table_range.getNumColumns() !== 2 ||
      name_range.getColumn() !== table_range.getColumn()
    ) {
      throw new Error(
        `${CLASS_CATEGORIES_TABLE_NAME}:${NAME_COLUMN_NAME} has unexpected shape`,
      );
    }
    const existingNameParents = new Map(
      (table_range.getValues() as [string, string][]).filter((nameParent) =>
        incomingClassCategories.includes(nameParent[0]),
      ),
    );
    const allNameParents: [string, string][] = incomingClassCategories
      .map((cc): [string, string] => [cc, existingNameParents.get(cc) ?? ""])
      .sort((a, b) => a[1].localeCompare(b[1]) || a[0].localeCompare(b[0]));
    helper.updateRowCount(allNameParents.length);
    (helper.getRange() || fail()).setValues(allNameParents);
    return existingNameParents;
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
    Logger.log(`API version: ${major}.${minor}`);
    const supported = { major: 0, minor: 6 };
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
