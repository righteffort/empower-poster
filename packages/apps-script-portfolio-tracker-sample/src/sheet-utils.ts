export class OurError extends Error {}

/**
  Convenience functions for operating on Tables.
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

interface TableHelperState {
  gsheet: GSheet;
  gtable: GTable;
  columnNameToIndex: Map<string, number>;
}

class TableHelper {
  private spreadsheetId: string;
  private sheet: GoogleAppsScript.Spreadsheet.Sheet;
  private tableName: string;
  private state: TableHelperState;
  constructor(
    spreadsheetId: string,
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    tableName: string,
  ) {
    this.spreadsheetId = spreadsheetId;
    this.sheet = sheet;
    this.tableName = tableName;
    this.state = this.refreshState();
  }
  private static getState(
    spreadsheetId: string,
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    tableName: string,
  ): TableHelperState {
    const _gspreadsheet = Sheets.Spreadsheets.get(spreadsheetId, {
      fields:
        "sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)),tables(name,range,columnProperties))",
    });
    const sheetTitle = sheet.getSheetName();
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
    const gsheet: GSheet = {
      properties: {
        sheetId: _gsheet?.properties?.sheetId ?? 0,
        gridProperties: {
          rowCount: _gproperties.gridProperties?.rowCount ?? 0,
          columnCount: _gproperties.gridProperties?.columnCount ?? 0,
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
        columnIndex: p?.columnIndex ?? 0,
        columnName: p?.columnName ?? fail(),
      };
    });
    const gtable: GTable = {
      columnProperties: columnProperties ?? fail(),
      name: _gtable.name ?? fail(),
      range: {
        endColumnIndex: _gtable.range?.endColumnIndex ?? 0,
        endRowIndex: _gtable.range?.endRowIndex ?? 0,
        startColumnIndex: _gtable.range?.startColumnIndex ?? 0,
        startRowIndex: _gtable.range?.startRowIndex ?? 0,
      },
    };
    if (
      gtable.range.startRowIndex + 1 >= gtable.range.endRowIndex || // +1 for header row
      gtable.range.startColumnIndex >= gtable.range.endColumnIndex
    ) {
      throw new OurError(
        `Table ${sheetTitle}.${tableName} has zero data rows or zero columns`,
      );
    }

    gsheet.tables.push(gtable);
    const columnNameToIndex = new Map<string, number>(
      gtable.columnProperties.map((p) => [p.columnName, p.columnIndex]),
    );
    return { gsheet, gtable, columnNameToIndex };
  }

  /**
   * Make sure the table has at least rowsNeeded rows
   */
  ensureRowCount(rowsNeeded: number) {
    const gridRange = this.state.gtable.range;
    // Convert to SpreadsheetApp: 1-indexed and closed-closed ranges
    const firstRow = gridRange.startRowIndex + 2; // +1 for header row
    const lastRow = gridRange.endRowIndex;
    const [firstColumn, lastColumn] = [
      gridRange.startColumnIndex + 1,
      gridRange.endColumnIndex,
    ];
    let numRows = lastRow - firstRow + 1;
    const numColumns = lastColumn - firstColumn + 1;
    if (numRows < 1 || numColumns < 1) {
      // This is redundant with the check in the constructor.
      throw new Error(
        `Cannot extend table with ${numRows} data rows and ${numColumns} columns`,
      );
    }
    // Extend the table until it will accomodate all the rows
    // Until b/525219695 is fixed, we have to make sure there is always one
    // blank row at the end of the table.
    const rowsActuallyNeeded = rowsNeeded + 1;
    let totalRowsToAdd = rowsActuallyNeeded - numRows;
    if (totalRowsToAdd > 0) {
      // We'll assume a blank first column is a proxy for an empty row (some columns may contain formulas)
      if (this.sheet.getRange(lastRow, firstColumn).getDisplayValue() !== "") {
        throw new OurError(
          "Refusing to extend table with non-blank value in first column of the last row. https://issuetracker.google.com/issues/525219695",
        );
      }
    }
    while (totalRowsToAdd > 0) {
      const rowsToAdd = Math.min(totalRowsToAdd, numRows);
      console.log(
        `inserting here: ${JSON.stringify({ firstRow, firstColumn, rowsToAdd, numColumns }, null, 2)}`,
      );
      const range = this.sheet.getRange(
        firstRow,
        firstColumn,
        rowsToAdd,
        numColumns,
      );
      range.insertCells(SpreadsheetApp.Dimension.ROWS);
      numRows += rowsToAdd;
      totalRowsToAdd -= rowsToAdd;
    }
    // Refresh state after mutation
    this.refreshState();
    // Confirm goal goal was achieved
    const newGridRange = this.state.gtable.range;
    if (
      newGridRange.endRowIndex - (newGridRange.startRowIndex + 1) <
      rowsActuallyNeeded
    ) {
      throw new OurError(
        `Failed to enlarge ${this.sheet.getSheetName()}.${this.tableName} to ${rowsActuallyNeeded} data rows`,
      );
    }
  }

  /** The number of data rows. */
  getNumRows() {
    const gridRange = this.state.gtable.range;
    return gridRange.endRowIndex - 1 - (gridRange.startRowIndex + 1); // -1 for https://issuetracker.google.com/issues/525219695
  }

  /**
   * Returns SpreadsheetApp Range for the data with the given columnName.
   */
  getColumnRange(
    columnName: string,
  ): GoogleAppsScript.Spreadsheet.Range | undefined {
    const tcolumnIndex = this.state.columnNameToIndex.get(columnName);
    if (tcolumnIndex == null) {
      return;
    }
    return this.getRangeForColumns(tcolumnIndex, 1);
  }

  /*
   * Returns SpreadsheetApp Range for the table data
   */
  getRange(): GoogleAppsScript.Spreadsheet.Range {
    const gridRange = this.state.gtable.range;
    return this.getRangeForColumns(
      0,
      gridRange.endColumnIndex - gridRange.startColumnIndex,
    );
  }

  // startColumnIndex is 0-based relative to the table.
  private getRangeForColumns(startColumnIndex: number, numColumns: number) {
    const gridRange = this.state.gtable.range;
    const gridStartDataRowIndex = gridRange.startRowIndex + 1; // For header row
    return this.sheet.getRange(
      gridStartDataRowIndex + 1,
      gridRange.startColumnIndex + startColumnIndex + 1,
      gridRange.endRowIndex - 1 - gridStartDataRowIndex, // -1 for https://issuetracker.google.com/issues/525219695
      numColumns,
    );
  }

  refreshState(): TableHelperState {
    this.state = TableHelper.getState(
      this.spreadsheetId,
      this.sheet,
      this.tableName,
    );
    return this.state;
  }
}
