// Make changes in https://github.com/righteffort/empower-poster/packages/apps-script-portfolio-tracker/src/
import type {
  PostPayload,
  PostResponse,
  HoldingEntry,
  Classification,
  Classifications,
  Account,
} from "@righteffort/empower-poster-types";

import { TableHelper } from "./sheet-utils.js";

function fail(message = "Internal error"): never {
  throw new Error(message);
}

/**
 * Add a leading apostrophe to purely numeric account names and types.
 */
function formatAsString(accountName: string): string {
  if (accountName.search(/^[0-9]+$/) === 0) {
    return `'${accountName}`;
  }
  return accountName;
}
/*
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
 */
/*
export function NOPE_EMPOWER_VALUE(
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
    return JSON.parse(cached)[rowNumber - 1] ?? "";
  }
  const values = computeCacheEntry(tableName, columnName);
  const serialized = JSON.stringify(values);
  if (serialized.length > 100 * 1000) {
    throw new Error(`Too much data for ${tableName}:${columnName} to cache`);
  }
  cache.put(key, serialized);
  return values[rowNumber - 1] ?? "";
}

function computeCacheEntry(tableName: string, columnName: string) {
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
  return sheet
    .getRange(1, columnNumber, sheet.getLastRow(), 1)
    .getValues()
    .map((r) => r[0]);
}
 */
class AssetAllocationUpater {
  private readonly holdingsArray: HoldingEntry[];
  private readonly classifications: Record<string, Classification[]>;
  private readonly accountMap: Map<number, Account>;
  private sheetsService: MyGoogleAppsScript.Sheets;
  private readonly spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  private readonly spreadsheetId: string;
  private readonly accountSetupSheet: GoogleAppsScript.Spreadsheet.Sheet;
  private readonly assetSetupSheet: GoogleAppsScript.Spreadsheet.Sheet;
  private readonly holdingsSheet: GoogleAppsScript.Spreadsheet.Sheet;
  private readonly empowerAccountsSheet: GoogleAppsScript.Spreadsheet.Sheet;
  private readonly empowerAssetsSheet: GoogleAppsScript.Spreadsheet.Sheet;
  private readonly empowerHoldingsSheet: GoogleAppsScript.Spreadsheet.Sheet;

  constructor(
    holdingsArray: HoldingEntry[],
    classifications: Classifications,
    accounts: Account[],
  ) {
    const ACCOUNT_SETUP_SHEET_NAME = "Account Setup";
    const ASSET_SETUP_SHEET_NAME = "Asset Setup";
    const HOLDINGS_SHEET_NAME = "Holdings";
    const EMPOWER_ACCOUNTS_SHEET_NAME = "Empower Accounts";
    const EMPOWER_ASSETS_SHEET_NAME = "Empower Assets";
    const EMPOWER_HOLDINGS_SHEET_NAME = "Empower Holdings";
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
    if (typeof Sheets === "undefined") {
      throw new Error("Must enable Sheets service");
    }
    this.sheetsService = Sheets as MyGoogleAppsScript.Sheets;
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
    this.empowerAccountsSheet =
      this.spreadsheet.getSheetByName(EMPOWER_ACCOUNTS_SHEET_NAME) ||
      fail(`'${EMPOWER_ACCOUNTS_SHEET_NAME}' sheet missing`);
    this.empowerAssetsSheet =
      this.spreadsheet.getSheetByName(EMPOWER_ASSETS_SHEET_NAME) ||
      fail(`'${EMPOWER_ASSETS_SHEET_NAME}' sheet missing`);
    this.empowerHoldingsSheet =
      this.spreadsheet.getSheetByName(EMPOWER_HOLDINGS_SHEET_NAME) ||
      fail(`'${EMPOWER_HOLDINGS_SHEET_NAME}' sheet missing`);
  }

  updateSpreadsheet() {
    Logger.log("Update accounts");
    this.updateEmpowerAccountsSheet();
    Logger.log("Update assets");
    this.updateEmpowerAssetsSheet();
    Logger.log("Update holdings");
    this.updateEmpowerHoldingsSheet();
    Logger.log("Done");
  }

  private updateEmpowerAccountsSheet() {
    const sheet = this.empowerAccountsSheet;
    sheet.clearContents();
    this.updateAccounts(1);
    this.updateInstitutions(sheet.getLastColumn() + 2);
  }

  private updateInstitutions(column: number) {
    const institutions = [
      ...new Set([...this.accountMap.values()].map((a) => a.firmName)),
    ].sort();
    const sheet = this.empowerAccountsSheet;
    const institutionsRange = sheet.getRange(
      1,
      column,
      institutions.length + 1,
      1,
    );
    const INSTITUTIONS_TABLE_NAME = "Institutions";
    institutionsRange.setValues(
      [`${INSTITUTIONS_TABLE_NAME}:Name`, ...institutions].map((c) => [
        formatAsString(c),
      ]),
    );
    const helper = new TableHelper(
      this.sheetsService,
      this.spreadsheetId,
      this.accountSetupSheet,
      INSTITUTIONS_TABLE_NAME,
    );
    helper.updateRowCount(institutions.length);
  }

  private updateAccounts(column: number) {
    // Capture existing owner and type values
    const ACCOUNTS_TABLE_NAME = "Accounts";
    const helper = new TableHelper(
      this.sheetsService,
      this.spreadsheetId,
      this.accountSetupSheet,
      ACCOUNTS_TABLE_NAME,
    );
    const NAME_COLUMN_NAME = "Name";
    const TYPE_COLUMN_NAME = "Type";
    const OWNER_COLUMN_NAME = "Owner";
    const columnIndices = helper.getColumnIndexByNameMap();
    const existingAccountData = (helper.getRange() ?? fail()).getValues();

    const existingAccounts = new Map(
      existingAccountData.map((r) => [
        r[columnIndices.get(NAME_COLUMN_NAME) ?? fail()],
        {
          type: r[columnIndices.get(TYPE_COLUMN_NAME) ?? fail()],
          owner: r[columnIndices.get(OWNER_COLUMN_NAME) ?? fail()],
        },
      ]),
    );

    // Write to the Empower Sheet
    const accounts: [string, string][] = [...this.accountMap.values()]
      .sort((a: Account, b: Account) => a.name.localeCompare(b.name))
      .map((a) => [formatAsString(a.name), formatAsString(a.firmName)]);
    const headers = ["Name", "Institution"].map(
      (c) => `${ACCOUNTS_TABLE_NAME}:${c}`,
    );
    const sheet = this.empowerAccountsSheet;
    const accountsRange = sheet.getRange(1, column, accounts.length + 1, 2);
    accountsRange.setValues([headers, ...accounts]);

    // Adjust table size and restore owner and type values
    helper.updateRowCount(accounts.length);
    (helper.getColumnRange(TYPE_COLUMN_NAME) ?? fail()).setValues(
      accounts.map((a) => [existingAccounts.get(a[0])?.type ?? ""]),
    );
    (helper.getColumnRange(OWNER_COLUMN_NAME) ?? fail()).setValues(
      accounts.map((a) => [existingAccounts.get(a[0])?.owner ?? ""]),
    );
  }

  private updateEmpowerHoldingsSheet() {
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
      )
      .map((r) => [formatAsString(r.accountName), r.ticker, r.quantity]);
    const HOLDINGS_TABLE_NAME = "Holdings";
    const sheet = this.empowerHoldingsSheet;
    sheet.clearContents();
    const headers = ["Account", "Asset", "Shares"].map(
      (c) => `${HOLDINGS_TABLE_NAME}:${c}`,
    );
    sheet
      .getRange(1, 1, holdingRows.length + 1, 3)
      .setValues([headers, ...holdingRows]);
    const helper = new TableHelper(
      this.sheetsService,
      this.spreadsheetId,
      this.holdingsSheet,
      HOLDINGS_TABLE_NAME,
    );
    helper.updateRowCount(holdingRows.length);
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

  private updateEmpowerAssetsSheet() {
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
    const sheet = this.empowerAssetsSheet;
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
    sheet.clearContents();
    this.updateAssetClasses(assetClasses, 1);
    this.updateAssets(
      flatAssets.sort(
        (a, b) =>
          a.ticker.localeCompare(b.ticker) ||
          (classCategoryParentMap.get(a.classes[0]) ?? "").localeCompare(
            classCategoryParentMap.get(b.classes[0]) ?? "",
          ) ||
          a.classes[0].localeCompare(b.classes[0]) ||
          a.classes[1].localeCompare(b.classes[1]),
      ),
      sheet.getLastColumn() + 2,
    );
  }

  updateAssets(
    flatAssets: {
      ticker: string;
      classes: [string, string];
      fraction: number;
    }[],
    column: number,
  ) {
    const holdings = new Map<string, HoldingEntry>(
      this.holdingsArray.map((h) => [h.ticker, h]),
    );
    const assetRows = flatAssets
      .sort((a, b) => a.ticker.localeCompare(b.ticker))
      .map((a) => {
        const h = holdings.get(a.ticker);
        return [
          a.ticker,
          formatAsString(h?.cusip ?? ""),
          h?.cusip ? "" : a.ticker,
          a.classes[1],
          a.fraction,
          h?.cusip ? "" : h?.price,
        ];
      });
    const ASSETS_TABLE_NAME = "Assets";
    const headers = [
      "Ticker",
      "Cusip",
      "Name",
      "Class",
      "Class Pct",
      "Price",
    ].map((c) => `${ASSETS_TABLE_NAME}:${c}`);
    this.empowerAssetsSheet
      .getRange(1, column, assetRows.length + 1, headers.length)
      .setValues([headers, ...assetRows]);
    const helper = new TableHelper(
      this.sheetsService,
      this.spreadsheetId,
      this.assetSetupSheet,
      ASSETS_TABLE_NAME,
    );
    helper.updateRowCount(assetRows.length);
  }

  private updateAssetClasses(assetClasses: string[][], column: number) {
    const ASSET_CLASSES_TABLE_NAME = "Asset Classes";
    const headers = ["Name", "Category"].map(
      (c) => `${ASSET_CLASSES_TABLE_NAME}:${c}`,
    );
    this.empowerAssetsSheet
      .getRange(1, column, assetClasses.length + 1, headers.length)
      .setValues([headers, ...assetClasses.map((cs) => [cs[1], cs[0]])]);
    const helper = new TableHelper(
      this.sheetsService,
      this.spreadsheetId,
      this.assetSetupSheet,
      ASSET_CLASSES_TABLE_NAME,
    );
    helper.updateRowCount(assetClasses.length);
  }

  private updateClassCategories(incomingClassCategories: string[]) {
    const CLASS_CATEGORIES_TABLE_NAME = "Class Categories";
    const NAME_COLUMN_NAME = "Name";
    const helper = new TableHelper(
      this.sheetsService,
      this.spreadsheetId,
      this.assetSetupSheet,
      CLASS_CATEGORIES_TABLE_NAME,
    );
    const table_range =
      helper.getRange() || fail(`${CLASS_CATEGORIES_TABLE_NAME} not found`);
    const name_range =
      helper.getColumnRange(NAME_COLUMN_NAME) ||
      fail(`${CLASS_CATEGORIES_TABLE_NAME}:${NAME_COLUMN_NAME} not found`);
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
    const newNameParents: [string, string][] = incomingClassCategories
      .map((cc): [string, string] => [cc, existingNameParents.get(cc) ?? ""])
      .sort((a, b) => a[1].localeCompare(b[1]) || a[0].localeCompare(b[0]));
    const newCategoriesSorted = incomingClassCategories.sort();
    const existingCategoriesSorted = [...existingNameParents.keys()].sort();
    if (
      newCategoriesSorted.length !== existingCategoriesSorted.length ||
      !newCategoriesSorted.every((c, i) => existingCategoriesSorted[i] === c)
    ) {
      helper.updateRowCount(newNameParents.length);
      (helper.getRange() || fail()).setValues(newNameParents);
    }
    return new Map(newNameParents);
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
    Logger.log(`${holdings.length} holdings`);
    Logger.log(`${Object.entries(classifications).length} classifications`);
    Logger.log(`${accounts.length} accounts`);
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
      // invalidateCache();
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
