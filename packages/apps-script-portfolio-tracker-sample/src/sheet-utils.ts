// import {sheets_v4} from '@googleapis/sheets';

export class OurError extends Error {}

/**
  Convenience functions for operating on Tables.

  Note that if any table properties change the caller needs to create a new instance.
 */
export function makeTableHelper(
  spreadsheetId: string,
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  tableName: string,
): TableHelper | undefined {
  try {
    return new TableHelper(spreadsheetId, sheet, tableName);
  } catch (e) {
    if (e instanceof OurError) {
      console.log(e.message);
      return;
    }
    throw e;
  }
}

interface GTable {
  columnProperties: {
    columnIndex: number;
    columnName: string;
  }[];
  name: string;
  range: {
    endColumnIndex: number;
    endRowIndex: number;
    sheetId: number;
    startColumnIndex: number;
    startRowIndex: number;
  };
}

interface GSheet {
  properties: {
    sheetId: number;
    gridProperties: {
      rowCount: number;
      columnCount: number;
    };
  };
  tables: GTable[];
}

class TableHelper {
  spreadsheetId: string;
  sheet: GoogleAppsScript.Spreadsheet.Sheet;
  gsheet: GSheet;
  gtable: GTable;
  columnNameToIndex: Map<string, number>;
  constructor(
    spreadsheetId: string,
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    tableName: string,
  ) {
    this.spreadsheetId = spreadsheetId;
    this.sheet = sheet;
    const sheetTitle = sheet.getSheetName();
    const _gspreadsheet = Sheets.Spreadsheets.get(spreadsheetId, {
      fields:
        "sheets(properties(sheetId,gridProperties(rowCount,columnCount)),tables(name,range,columnProperties))",
    });
    const _gsheet = _gspreadsheet.sheets?.filter(
      (s) => s.properties?.title === sheetTitle,
    )[0];
    if (_gsheet === undefined) {
      throw new OurError(`No such sheet ${sheetTitle}`);
    }
    let fail = () => {
      throw new OurError(`Incomplete sheet ${sheetTitle}`);
    };
    const _gproperties = _gsheet?.properties ?? fail();
    this.gsheet = {
      properties: {
        sheetId: _gsheet?.properties?.sheetId ?? fail(),
        gridProperties: {
          rowCount: _gproperties.gridProperties?.rowCount ?? fail(),
          columnCount: _gproperties.gridProperties?.columnCount ?? fail(),
        },
      },
      tables: [],
    };
    const _gtable = _gsheet?.tables?.filter((t) => t.name === tableName)[0];
    if (_gtable === undefined) {
      throw new OurError(`No such table ${sheetTitle}.${tableName}.`);
    }
    fail = () => {
      throw new OurError(`Incomplete table ${sheetTitle}.${tableName}`);
    };
    const columnProperties = _gtable.columnProperties?.map((p) => {
      return {
        columnIndex: p?.columnIndex ?? fail(),
        columnName: p?.columnName ?? fail(),
      };
    });
    this.gtable = {
      columnProperties: columnProperties ?? fail(),
      name: _gtable.name ?? fail(),
      range: {
        endColumnIndex: _gtable.range?.endColumnIndex ?? fail(),
        endRowIndex: _gtable.range?.endRowIndex ?? fail(),
        sheetId: _gtable.range?.sheetId ?? fail(),
        startColumnIndex: _gtable.range?.startColumnIndex ?? fail(),
        startRowIndex: _gtable.range?.startRowIndex ?? fail(),
      },
    };
    this.gsheet.tables = [this.gtable];
    this.columnNameToIndex = new Map<string, number>(
      this.gtable.columnProperties.map((p) => [p.columnName, p.columnIndex]),
    );
  }
  getSheetRangeForTableColumn(
    columnName: string,
  ): GoogleAppsScript.Spreadsheet.Range | undefined {
    const tcolumnIndex = this.columnNameToIndex.get(columnName);
    if (tcolumnIndex == null) {
      return;
    }
    const row = this.gtable.range.startRowIndex + 2;
    const column = this.gtable.range.startColumnIndex + tcolumnIndex + 1;
    const numRows =
      this.gtable.range.endRowIndex - (this.gtable.range.startRowIndex + 1);
    return this.sheet.getRange(row, column, numRows, 1);
  }
}
