// Make changes in https://github.com/righteffort/empower-poster/packages/apps-script-sample/src/

import type {
  PostPayload,
  PostResponse,
  HoldingEntry,
  Classifications,
  Account,
} from "@righteffort/empower-poster-types";

function writeHoldings(
  holdingsSheet: GoogleAppsScript.Spreadsheet.Sheet,
  holdings: HoldingEntry[],
) {
  holdingsSheet.clear();
  const headers = [
    "userAccountId",
    "ticker",
    "price",
    "quantity",
    "value",
    "cusip",
    "fundFees",
  ];
  holdingsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (holdings.length > 0) {
    const holdingsData = holdings.map((holding) => [
      holding.userAccountId,
      holding.ticker,
      holding.price,
      holding.quantity,
      holding.value,
      holding.cusip,
      holding.fundFees ?? "",
    ]);
    holdingsSheet
      .getRange(2, 1, holdingsData.length, headers.length)
      .setValues(holdingsData);
  }
}

function writeClassifications(
  classificationsSheet: GoogleAppsScript.Spreadsheet.Sheet,
  classifications: Classifications,
) {
  classificationsSheet.clear();
  const headers = ["ticker", "class", "subclass", "fraction"];
  classificationsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const classificationsData: (string | number)[][] = [];
  for (const [ticker, classificationArray] of Object.entries(classifications)) {
    for (const classification of classificationArray) {
      classificationsData.push([
        ticker,
        classification.classes[0],
        classification.classes[1],
        classification.fraction,
      ]);
    }
  }

  if (classificationsData.length > 0) {
    classificationsSheet
      .getRange(2, 1, classificationsData.length, headers.length)
      .setValues(classificationsData);
  }
}

function writeAccounts(
  accountsSheet: GoogleAppsScript.Spreadsheet.Sheet,
  accounts: Account[],
) {
  accountsSheet.clear();
  const headers = [
    "id",
    "name",
    "accountType",
    "advisoryFeePercentage",
    "balance",
    "firmName",
    "fundFees",
    "isTaxDeferredOrNonTaxable",
  ];
  accountsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (accounts.length > 0) {
    const accountsData = accounts.map((account) => [
      account.id,
      account.name,
      account.accountType,
      account.advisoryFeePercentage,
      account.balance,
      account.firmName,
      account.fundFees ?? "",
      account.isTaxDeferredOrNonTaxable,
    ]);
    accountsSheet
      .getRange(2, 1, accountsData.length, headers.length)
      .setValues(accountsData);
  }
}

export function doPost(e: GoogleAppsScript.Events.DoPost) {
  try {
    const {
      version: { major, minor },
      holdings,
      classifications,
      accounts,
    } = JSON.parse(e.postData.contents) as PostPayload;
    console.log(`API version: ${major}.${minor}`);
    const supported = { major: 0, minor: 6 };
    if (major !== supported.major || minor < supported.minor) {
      throw new Error(
        `data version ${major}.${minor} not supported, expected at least ${supported.major}.${supported.minor}`,
      );
    }
    console.log(`Received ${holdings.length} holdings`);
    console.log(
      `Classifications for ${Object.keys(classifications).length} tickers`,
    );
    console.log(`${accounts.length} accounts`);
    // Write data to sheets
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const classificationsSheet = spreadsheet.getSheetByName("classifications");
    const holdingsSheet = spreadsheet.getSheetByName("holdings");
    const accountsSheet = spreadsheet.getSheetByName("accounts");
    if (!classificationsSheet || !holdingsSheet || !accountsSheet) {
      throw new Error(
        "One or more of classifications, holdings, or accounts sheets missing",
      );
    }
    let lock: GoogleAppsScript.Lock.Lock | undefined;
    try {
      lock = LockService.getScriptLock();
      lock.waitLock(30_000);
      writeHoldings(holdingsSheet, holdings);
      writeClassifications(classificationsSheet, classifications);
      writeAccounts(accountsSheet, accounts);
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

// Placeholder to demonstrate unit testing.
export function placeholder(): boolean {
  return true;
}
