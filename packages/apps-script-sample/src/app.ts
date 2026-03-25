//! Make changes in https://github.com/righteffort/empower-poster/packages/apps-script-sample/src/

import type {
  PostPayload,
  PostResponse,
  HoldingEntry,
  Classifications,
} from "@righteffort/empower-poster-types";

function writeHoldings(
  holdingsSheet: GoogleAppsScript.Spreadsheet.Sheet,
  holdings: HoldingEntry[],
) {
  holdingsSheet.clear();
  const headers = ["userAccountId", "ticker", "price", "quantity", "value"];
  holdingsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (holdings.length > 0) {
    const holdingsData = holdings.map((holding) => [
      holding.userAccountId,
      holding.ticker,
      holding.price,
      holding.quantity,
      holding.value,
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

function doPost(event: GoogleAppsScript.Events.DoPost) {
  try {
    const { version, holdings, classifications } = JSON.parse(
      event.postData.contents,
    ) as PostPayload;

    console.log(`Received ${holdings.length} holdings`);
    console.log(
      `Classifications for ${Object.keys(classifications).length} tickers`,
    );
    console.log(`API version: ${version}`);

    const supportedVersion = "0.3";
    if (version !== supportedVersion) {
      throw new Error(
        `data version ${version} not supported, expected ${supportedVersion}`,
      );
    }

    // Write data to sheets
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const classificationsSheet = spreadsheet.getSheetByName("classifications");
    const holdingsSheet = spreadsheet.getSheetByName("holdings");
    if (!classificationsSheet || !holdingsSheet) {
      throw new Error("classifications and/or holdings sheet missing");
    }
    writeHoldings(holdingsSheet, holdings);
    writeClassifications(classificationsSheet, classifications);

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

// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).doPost = doPost;
