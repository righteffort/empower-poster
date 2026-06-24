export class OurError extends Error {}

/**
  Convenience functions for operating on Tables in Portfolio Tracker Spreadsheet.
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
  tableId: string;
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
  private isFormulaTable: boolean;
  private lastRowAdjustment: 0 | 1; // 1 for formula tables when https://issuetracker.google.com/issues/525219695 applies
  private isUniqueTableWithLowestRow: boolean;
  constructor(
    spreadsheetId: string,
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    tableName: string,
  ) {
    this.spreadsheetId = spreadsheetId;
    this.sheet = sheet;
    this.tableName = tableName;
    this.state = this.refreshState();
    this.isFormulaTable = this.getIsFormulaTable();
    this.lastRowAdjustment = this.isFormulaTable ? 1 : 0;
    this.isUniqueTableWithLowestRow = this.getIsUniqueTableWithLowestRow();
  }
  private static getState(
    spreadsheetId: string,
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    tableName: string,
  ): TableHelperState {
    const _gspreadsheet = Sheets.Spreadsheets.get(spreadsheetId, {
      fields:
        "sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)),tables(tableId,name,range,columnProperties))",
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
      tableId: _gtable.tableId ?? fail(),
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

  private getIsFormulaTable() {
    const gridRange = this.state.gtable.range;
    const lastRowFormulas = this.sheet
      .getRange(
        gridRange.endRowIndex,
        gridRange.startColumnIndex + 1,
        1,
        gridRange.endColumnIndex - gridRange.startColumnIndex,
      )
      .getFormulas()[0];
    if (lastRowFormulas == null) {
      throw new Error("Logic bug in getLastRowAdjustment");
    }
    return !lastRowFormulas.every((f) => f === "");
  }

  private getIsUniqueTableWithLowestRow() {
    return false; // TODO
  }

  /**
   * Delete a set of rows. Zero-index relative to data.
   */
  deleteRows(rows: Set<number>) {
    const gridRange = this.state.gtable.range;
    const gridStartDataRowIndex = gridRange.startRowIndex + 1; // For header row
    const rowsToDelete = Array.from(rows).sort((a, b) => b - a); // Reverse so row numbers remain correct
    const rangesToDelete = rowsToDelete.map((i) =>
      this.sheet.getRange(
        gridStartDataRowIndex + 1 + i,
        gridRange.startColumnIndex + 1,
        1,
        gridRange.endColumnIndex - gridRange.startColumnIndex,
      ),
    );
    console.log(rangesToDelete.map((r) => r.getA1Notation()));
    rowsToDelete.forEach((i) => {
      this.sheet
        .getRange(
          gridStartDataRowIndex + 1 + i,
          gridRange.startColumnIndex + 1,
          1,
          gridRange.endColumnIndex - gridRange.startColumnIndex,
        )
        .deleteCells(SpreadsheetApp.Dimension.ROWS);
    });
    // Refresh state after mutation
    this.refreshState();
  }

  /**
   * Make sure the table has at least rowsNeeded rows
   */
  ensureRowCount(
    rowsNeeded: number,
  ): GoogleAppsScript.Spreadsheet.Range | undefined {
    return this.isFormulaTable
      ? this.ensureFormulaRowCount(rowsNeeded)
      : this.ensureValueRowCount(rowsNeeded);
  }

  private ensureValueRowCount(
    rowsNeeded: number,
  ): GoogleAppsScript.Spreadsheet.Range | undefined {
    // We grandly assume there is nothing underneath us and make it part of the table if we were wrong.
    const { tableId, range } = this.state.gtable;
    const oldRange = { ...range };
    const numRows = this.getNumRows();
    const rowsToAdd = rowsNeeded - numRows;
    if (rowsToAdd <= 0) {
      return undefined;
    }
    range.endRowIndex += rowsToAdd;
    Sheets.Spreadsheets.batchUpdate(
      {
        requests: [
          {
            updateTable: {
              fields: "range",
              table: { tableId, range },
            },
          },
        ],
      },
      this.spreadsheetId,
    );
    this.refreshState();
    if (this.getNumRows() < rowsNeeded) {
      throw new Error(
        `Tried to extend table ${this.state.gtable.name} past end of sheet`,
      );
    }
    return this.sheet.getRange(
      oldRange.endRowIndex + 1,
      oldRange.startColumnIndex + 1,
      rowsToAdd,
      oldRange.endColumnIndex - oldRange.startColumnIndex,
    );
  }

  private ensureFormulaRowCount(
    rowsNeeded: number,
  ): GoogleAppsScript.Spreadsheet.Range | undefined {
    // TODO: simplify!
    const gridRange = this.state.gtable.range;
    // Convert to SpreadsheetApp: 1-indexed and closed-closed ranges
    const firstRow = gridRange.startRowIndex + 2; // +1 for header row
    const lastRow = gridRange.endRowIndex;
    const [firstColumn, lastColumn] = [
      gridRange.startColumnIndex + 1,
      gridRange.endColumnIndex,
    ];
    const numRows = lastRow - firstRow + 1;
    const numColumns = lastColumn - firstColumn + 1;
    if (numRows < 1 || numColumns < 1) {
      // This is redundant with the check in the constructor.
      throw new Error(
        `Cannot extend table with ${numRows} data rows and ${numColumns} columns`,
      );
    }
    // Extend the table until it will accomodate all the rows
    // Until b/525219695 is fixed, we may have to make sure there is always one
    // blank row at the end of the table.
    const rowsActuallyNeeded = rowsNeeded + this.lastRowAdjustment;
    const totalRowsToAdd = rowsActuallyNeeded - numRows;
    if (totalRowsToAdd > 0 && this.lastRowAdjustment != 0) {
      // We'll assume a blank first column is a proxy for an empty row (some columns may contain formulas)
      if (this.sheet.getRange(lastRow, firstColumn).getDisplayValue() !== "") {
        throw new OurError(
          "Refusing to extend table with non-blank value in first column of the last row. https://issuetracker.google.com/issues/525219695",
        );
      }
    }
    let numRowsAvailable = numRows - 1; // We have to insert after the first row to avoid breaking named ranges.
    if (totalRowsToAdd <= 0) {
      return;
    }
    if (!this.isUniqueTableWithLowestRow) {
      throw new OurError(
        `Cannot extend formula table ${this.state.gtable.name} unless there is no data outside the table on or below its last row`,
      );
    }
    let remainingRowsToAdd = totalRowsToAdd;
    while (remainingRowsToAdd > 0) {
      const rowsToAdd = Math.min(remainingRowsToAdd, numRowsAvailable);
      if (rowsToAdd <= 0) {
        throw new Error("Adding new rows would break named ranges");
      }
      console.log(
        `inserting here: ${JSON.stringify({ firstRow, firstColumn, rowsToAdd, numColumns }, null, 2)}`,
      );
      const range = this.sheet.getRange(
        firstRow + 1, // Insert after the first row.
        firstColumn,
        rowsToAdd,
        numColumns,
      );
      range.insertCells(SpreadsheetApp.Dimension.ROWS);
      numRowsAvailable += rowsToAdd;
      remainingRowsToAdd -= rowsToAdd;
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
    return this.sheet.getRange(
      firstRow + 1, // Inserted after the first row
      firstColumn,
      totalRowsToAdd,
      numColumns,
    );
  }

  /** The number of data rows. */
  getNumRows() {
    const gridRange = this.state.gtable.range;
    return (
      gridRange.endRowIndex -
      this.lastRowAdjustment -
      (gridRange.startRowIndex + 1)
    );
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
      gridRange.endRowIndex - this.lastRowAdjustment - gridStartDataRowIndex,
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
