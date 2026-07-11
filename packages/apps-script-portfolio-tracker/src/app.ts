// Make changes in https://github.com/righteffort/empower-poster/packages/apps-script-portfolio-tracker/src/
import type {
  PostPayload,
  PostResponse,
  HoldingEntry,
  Classification,
  Classifications,
  Account,
} from "@righteffort/empower-poster-types";

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

const EMPOWER_ACCOUNTS_SHEET_NAME = "Empower Accounts";
const EMPOWER_ASSETS_SHEET_NAME = "Empower Assets";
const EMPOWER_HOLDINGS_SHEET_NAME = "Empower Holdings";

class AssetAllocationUpater {
  private readonly holdingsArray: HoldingEntry[];
  private readonly classifications: Record<string, Classification[]>;
  private readonly accountMap: Map<number, Account>;
  private readonly spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  private readonly empowerAccountsSheet: GoogleAppsScript.Spreadsheet.Sheet;
  private readonly empowerAssetsSheet: GoogleAppsScript.Spreadsheet.Sheet;
  private readonly empowerHoldingsSheet: GoogleAppsScript.Spreadsheet.Sheet;

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
    this.updateEmpowerAccountsSheet();
    this.updateEmpowerAssetsSheet();
    this.updateEmpowerHoldingsSheet();
  }

  private updateEmpowerAccountsSheet() {
    const sheet = this.empowerAccountsSheet;
    sheet.clearContents();
    this.updateInstitutions(0);
    this.updateAccounts(sheet.getLastColumn() + 2);
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
    institutionsRange.setValues(
      ["Institutions:Name", ...institutions].map((c) => [formatAsString(c)]),
    );
  }

  private updateAccounts(column: number) {
    // TODO: We also have to massage the contents of the Accounts table in Account Setup!
    const accounts: [string, string][] = [...this.accountMap.values()]
      .sort((a: Account, b: Account) => a.name.localeCompare(b.name))
      .map((a) => [formatAsString(a.name), formatAsString(a.firmName)]);
    const ACCOUNTS_TABLE_NAME = "Accounts";
    const headers = ["Name", "Institution"].map(
      (c) => `${ACCOUNTS_TABLE_NAME}:${c}`,
    );
    const sheet = this.empowerAccountsSheet;
    const accountsRange = sheet.getRange(1, column, accounts.length + 1, 2);
    accountsRange.setValues([headers, ...accounts]);
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
      .map((r) => [r.accountName, r.ticker, r.quantity]);
    const HOLDINGS_TABLE_NAME = "Holdings";
    const sheet = this.empowerHoldingsSheet;
    sheet.clearContents();
    const headers = ["Account", "Asset", "Shares"].map(
      (c) => `${HOLDINGS_TABLE_NAME}:${c}`,
    );
    sheet
      .getRange(1, 1, holdingRows.length + 1, 3)
      .setValues([headers, ...holdingRows]);
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
    // const classCategories = [
    //   ...new Set(flatAssets.map((a) => a.classes[0] ?? "")),
    // ];
    const assetClasses = Array.from(
      new Set(flatAssets.map((a) => a.classes.join("\0"))),
    )
      .map((v) => v.split("\0") as [string, string])
      .sort(
        (a: [string, string], b: [string, string]) =>
          a[0].localeCompare(b[0]) || a[1].localeCompare(b[1]),
      );
    const sheet = this.empowerAssetsSheet;
    sheet.clearContents();
    this.updateAssets(flatAssets, 0);
    this.updateAssetClasses(assetClasses, sheet.getLastColumn() + 2);
    // this.updateClassCategories(classCategories, sheet.getLastColumn() + 2);
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
        // TODO: Check that this interoperates with our "column formulas" for name and price
        return [
          a.ticker,
          h?.cusip ?? "",
          h?.cusip ? "" : a.ticker,
          a.classes[1],
          a.fraction,
          h?.cusip ? "" : h?.price,
        ];
      });
    const headers = [
      "Ticker",
      "Cusip",
      "Name",
      "Class",
      "Class Pct",
      "Price",
    ].map((c) => `Assets:${c}`);
    this.empowerAssetsSheet
      .getRange(1, column, assetRows.length + 1, headers.length)
      .setValues([headers, ...assetRows]);
  }

  private updateAssetClasses(assetClasses: string[][], column: number) {
    const headers = ["Name", "Category"].map((c) => `Asset Classes:${c}`);
    this.empowerAssetsSheet
      .getRange(1, column, assetClasses.length, headers.length)
      .setValues([headers, ...assetClasses.map((cs) => [cs[1], cs[0]])]);
  }

  // private updateClassCategories(incomingClassCategories: string[], column: number) {
  //   // TODO: write nothing at all (current), or preserve existing mappings.
  // }
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
  } finally {
    invalidateCache();
  }
}
export function placeholder(): boolean {
  return true;
}
