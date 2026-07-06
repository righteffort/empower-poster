export function EMPOWER_VALUE(
  tableName: string,
  columnName: string,
  rowNumber: number,
) {
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
  return sheet.getRange(rowNumber, columnNumber).getValue();
}

function fail(message = "Internal error"): never {
  throw new Error(message);
}

class FormulaInstaller {
  private sheetsService: MyGoogleAppsScript.Sheets;
  private spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  private spreadsheetId: string;
  constructor(
    sheetsService: MyGoogleAppsScript.Sheets,
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
  ) {
    this.sheetsService = sheetsService;
    this.spreadsheet = spreadsheet;
    this.spreadsheetId = this.spreadsheet.getId();
  }
  private setTableEndRowIndex(
    gtable: MyGoogleAppsScript.Sheets.Schema.Table,
    endRowIndex: number,
  ) {
    const utr = this.sheetsService.newUpdateTableRequest();
    utr.fields = "range";
    utr.table = this.sheetsService.newTable();
    utr.table.tableId = gtable.tableId;
    const oldRange = gtable.range || fail();
    console.log(`oldRange: ${JSON.stringify(oldRange, null, 2)}`);
    utr.table.range = this.sheetsService.newGridRange();
    utr.table.range.endColumnIndex = oldRange.endColumnIndex;
    utr.table.range.sheetId = oldRange.sheetId;
    utr.table.range.startColumnIndex = oldRange.startColumnIndex;
    utr.table.range.startRowIndex = oldRange.startRowIndex;
    utr.table.range.endRowIndex = endRowIndex;
    const req = this.sheetsService.newRequest();
    req.updateTable = utr;
    const busr = this.sheetsService.newBatchUpdateSpreadsheetRequest();
    busr.requests = [req];
    console.log(`Time to update table range: ${JSON.stringify(busr, null, 2)}`);
    // TODO this.sheetsService.Spreadsheets.batchUpdate(busr, this.spreadsheetId);
  }
  private installTableColumnFormulas(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    gtable: MyGoogleAppsScript.Sheets.Schema.Table,
    tableName: string,
    columnNames: string[] | undefined,
    customColumns: { name: string; formula: string }[] | undefined,
  ) {
    const makeStandardColumnFormula = (tableName: string, columnName: string) =>
      `=EMPOWER_VALUE("${tableName}", "${columnName}", ROW(Single(${tableName.replace(/ /g, "_")}[${columnName}])) - 2)`;
    const tableStartColumnIndex = gtable?.range?.startColumnIndex ?? 0;
    const tableEndColumnIndex = gtable?.range?.endColumnIndex ?? 0;
    const tableStartRowIndex = gtable?.range?.startRowIndex ?? 0;
    const tableEndRowIndex = gtable?.range?.endRowIndex ?? 0;
    const columnIndicesByName = new Map(
      (gtable.columnProperties ?? []).map((cp) => [
        cp.columnName ?? fail(),
        tableStartColumnIndex + (cp.columnIndex ?? 0),
      ]),
    );
    const newFormulasByColumnName: Map<string, string> = new Map<
      string,
      string
    >();
    (columnNames ?? []).forEach((columnName: string) => {
      newFormulasByColumnName.set(
        columnName,
        makeStandardColumnFormula(tableName, columnName),
      );
    });
    (customColumns ?? []).forEach((customColumn) => {
      newFormulasByColumnName.set(customColumn.name, customColumn.formula);
    });
    const lastRowRange = sheet.getRange(
      tableEndRowIndex,
      tableStartColumnIndex + 1,
      1,
      tableEndColumnIndex - tableStartColumnIndex,
    );
    const oldFormulas = lastRowRange.getFormulas()[0];
    const columnNamesToUpdate: string[] = [];
    columnIndicesByName.forEach((index: number, name: string) => {
      const newFormula: string | undefined = newFormulasByColumnName.get(name);
      if (newFormula) {
        if (oldFormulas?.[index] !== newFormula) {
          columnNamesToUpdate.push(name);
        } else {
          console.log(`Skipping ${tableName} ${name}`);
        }
      }
    });
    if (!columnNamesToUpdate) {
      return;
    }
    // Work around https://issuetracker.google.com/issues/525219695 by overwriting all but the last row.
    // Fail if there are fewer than 2 data rows.
    const numRowsToWrite = tableEndRowIndex - tableStartRowIndex - 2; // -1 for header row, -1 to preserve last row
    if (numRowsToWrite <= 0) {
      throw new Error(
        `Can't deal with table ${tableName} with fewer than 2 data rows`,
      );
    }
    columnNamesToUpdate.forEach((columnName) => {
      const columnRange = sheet.getRange(
        tableStartRowIndex + 2, // +1 for 1-offset, +1 for header row
        tableStartColumnIndex +
          (columnIndicesByName.get(columnName) ?? fail()) +
          1,
        numRowsToWrite,
        1,
      );
      console.log(
        `Time to fill '${columnRange.getSheet().getSheetName()}'!${columnRange.getA1Notation()} with ${newFormulasByColumnName.get(columnName)}`,
      );
      /*
      TODO:
      columnRange.setFormulas(
        new Array(columnRange.getNumRows()).fill([
          newFormulasByColumnName.get(columnName),
        ]),
      );
       */
    });
    this.setTableEndRowIndex(gtable, (gtable.range?.endRowIndex ?? fail()) - 1);
    console.log(
      `Time to deleteCells '${lastRowRange.getSheet().getSheetName()}'!${lastRowRange.getA1Notation()}`,
    );
    // TODO: lastRowRange.deleteCells(SpreadsheetApp.Dimension.ROWS);
    const replacementRowRange = sheet.getRange(
      lastRowRange.getRow() + -1,
      lastRowRange.getColumn(),
      1,
      lastRowRange.getNumColumns(),
    );
    console.log(
      `Time to insertCells '${replacementRowRange.getSheet().getSheetName()}'!${replacementRowRange.getA1Notation()}`,
    );
    // TODO: lastRowRange.insertCells(SpreadsheetApp.Dimension.ROWS);
  }

  private installSheetColumnFormulas(
    gsheet: MyGoogleAppsScript.Sheets.Schema.Sheet,
    sheetName: string,
    tableDescs: {
      name: string;
      columns?: string[];
      customColumns?: { name: string; formula: string }[];
    }[],
  ) {
    const sheet =
      this.spreadsheet.getSheetByName(sheetName) ||
      fail(`Sheet ${sheetName} not found`);
    const tablesByName = new Map(gsheet.tables?.map((t) => [t.name, t]));
    tableDescs.forEach((td) => {
      this.installTableColumnFormulas(
        sheet,
        tablesByName.get(td.name) ?? fail(`Table ${td.name} not found`),
        td.name,
        td.columns,
        td.customColumns,
      );
    });
  }
  installAllColumnFormulas() {
    const formulaDescriptors = [
      {
        sheetName: "Holdings",
        tables: [
          {
            name: "Holdings",
            columns: ["Account", "Asset", "Shares"],
          },
        ],
      },
      {
        sheetName: "Account Setup",
        tables: [
          {
            name: "Accounts",
            customColumns: [
              {
                name: "Institution",
                formula: `=xlookup(SINGLE(Accounts[Name]), 'Empower Accounts'!A2:A, 'Empower Accounts'!B2:B, "")`,
              },
            ],
          },
          {
            name: "Institutions",
            columns: ["Name"],
          },
        ],
      },
      {
        sheetName: "Asset Setup",
        tables: [
          {
            name: "Class Categories",
            columns: ["Name", "Parent"],
          },
          {
            name: "Asset Classes",
            columns: ["Name", "Category"],
          },
          {
            name: "Assets",
            columns: ["Ticker", "Class", "Class Pct"],
            customColumns: [
              {
                name: "Name",
                formula:
                  '=LET(altname, EMPOWER_VALUE("Assets", "Name", ROW(Single(Assets[Name])) - 2), if(NE(altname, ""), altname, ASSETNAME(Assets[Ticker])))',
              },
              {
                name: "Price",
                formula:
                  '=LET(altprice, EMPOWER_VALUE("Assets", "Price", ROW(Single(Assets[Price])) - 2), if(NE(altprice, ""), altprice, ASSETPRICE(Assets[Ticker])))',
              },
            ],
          },
        ],
      },
    ];
    const gspreadsheet = this.sheetsService.Spreadsheets.get(
      this.spreadsheetId,
      {
        fields:
          "sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)),tables(tableId,name,range,columnProperties))",
      },
    );
    const sheetsByTitle = new Map(
      gspreadsheet?.sheets?.map((s) => [s.properties?.title, s]),
    );
    formulaDescriptors.forEach((desc) =>
      this.installSheetColumnFormulas(
        sheetsByTitle?.get(desc.sheetName) ||
          fail(`Sheet ${desc.sheetName} not found`),
        desc.sheetName,
        desc.tables,
      ),
    );
  }
}

export function stuff() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (typeof Sheets === "undefined") {
    throw new Error("Must enable Sheets service");
  }
  new FormulaInstaller(
    Sheets as MyGoogleAppsScript.Sheets,
    spreadsheet,
  ).installAllColumnFormulas();
}

export function placeholder(): boolean {
  return true;
}
