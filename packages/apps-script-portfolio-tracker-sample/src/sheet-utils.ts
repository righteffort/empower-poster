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
    const sheetsService = MySheets;
    if (sheetsService == null) {
      throw new OurError("Must enable Sheets service");
    }
    return new TableHelper(sheetsService, spreadsheetId, sheet, tableName);
  } catch (e) {
    if (e instanceof OurError) {
      console.log(e.message);
      return;
    }
    throw e;
  }
}

// Type-safe subsets of Sheets.Table and Sheets.Sheet: no undefined fields
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
}

interface TableHelperState {
  gsheet: GSheet;
  gtable: GTable;
  columnNameToIndex: Map<string, number>;
  hasUniqueLargestEndRowIndex: boolean;
}

class TableHelper {
  private sheetsService: MyGoogleAppsScript.Sheets;
  private spreadsheetId: string;
  private sheet: GoogleAppsScript.Spreadsheet.Sheet;
  private tableName: string;
  private state: TableHelperState;
  private isFormulaTable: boolean;
  private lastRowAdjustment: 0 | 1; // 1 for formula tables, to preserve default last row
  constructor(
    sheetsService: MyGoogleAppsScript.Sheets,
    spreadsheetId: string,
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    tableName: string,
  ) {
    this.sheetsService = sheetsService;
    this.spreadsheetId = spreadsheetId;
    this.sheet = sheet;
    this.tableName = tableName;
    this.state = this.refreshState();
    this.isFormulaTable = this.getIsFormulaTable();
    this.lastRowAdjustment = this.isFormulaTable ? 1 : 0;
  }
  private getState(
    spreadsheetId: string,
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    tableName: string,
  ): TableHelperState {
    const _gspreadsheet = this.sheetsService.Spreadsheets.get(spreadsheetId, {
      fields:
        "sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)),tables(tableId,name,range,columnProperties))",
    });
    const sheetTitle = sheet.getSheetName();
    const _gsheet = _gspreadsheet?.sheets?.filter(
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
    };
    const _gtables = _gsheet.tables;
    const endRowIndices = (_gtables ?? []).map(
      (t) => t.range?.endRowIndex ?? 0,
    );
    const largestEndRowIndex = Math.max(...endRowIndices);
    const uniqueLargestEndRowIndexExists =
      endRowIndices.filter((i) => i === largestEndRowIndex).length === 1;

    const _gtable = _gtables?.filter((t) => t.name === tableName)[0];
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
        sheetId: _gtable.range?.sheetId ?? 0,
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
    const hasUniqueLargestEndRowIndex =
      uniqueLargestEndRowIndexExists &&
      _gtable?.range?.endRowIndex === largestEndRowIndex;
    const columnNameToIndex = new Map<string, number>(
      gtable.columnProperties.map((p) => [p.columnName, p.columnIndex]),
    );
    return { gsheet, gtable, columnNameToIndex, hasUniqueLargestEndRowIndex };
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
      throw new Error("Logic bug in getIsFormulaTable");
    }
    return !lastRowFormulas.every((f) => f === "");
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
    // TODO: Check that we won't bump into any tables.
    const { tableId, range } = this.state.gtable;
    const oldRange = { ...range };
    const numRows = this.getNumRows();
    const rowsToAdd = rowsNeeded - numRows;
    if (rowsToAdd <= 0) {
      return undefined;
    }
    range.endRowIndex += rowsToAdd;
    this.sheetsService.Spreadsheets.batchUpdate(
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
    // Did it work?
    if (this.getNumRows() < rowsNeeded) {
      throw new Error(
        `Failed to extend table ${this.state.gtable.name}, probably reached end of sheet`,
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
    // Insert rows above the last row in the table. They will acquire that row's formulas.
    const gridRange = this.state.gtable.range;
    // Convert to SpreadsheetApp: 1-indexed and closed-closed ranges
    const lastRow = gridRange.endRowIndex;
    const [firstColumn, lastColumn] = [
      gridRange.startColumnIndex + 1,
      gridRange.endColumnIndex,
    ];
    const numRows = gridRange.endRowIndex - gridRange.startRowIndex - 1; // +1 for header row
    const numColumns = lastColumn - firstColumn + 1;
    if (numRows < 1 || numColumns < 1) {
      // This is redundant with the check in the constructor.
      throw new Error(
        `Cannot extend table with ${numRows} data rows and ${numColumns} columns`,
      );
    }
    // Extend the table until it will accomodate all the rows
    // We have to make sure there is always one default row at the end of the table
    // so that we can simply insert rows above it.
    const rowsActuallyNeeded = rowsNeeded + 1; // Preserve last row
    const totalRowsToAdd = rowsActuallyNeeded - numRows;
    if (totalRowsToAdd <= 0) {
      return;
    }
    // We'll assume a blank first column is a proxy for an empty row (some columns will contain default formulas)
    if (this.sheet.getRange(lastRow, firstColumn).getDisplayValue() !== "") {
      throw new OurError(
        "Refusing to extend table with non-blank value in first column of the last row.",
      );
    }
    if (!this.state.hasUniqueLargestEndRowIndex) {
      throw new OurError(
        `Cannot extend formula table ${this.state.gtable.name} unless there is no data outside the table on or below its last row`,
      );
    }
    this.sheet.insertRowsBefore(lastRow, totalRowsToAdd);
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
    return; // There's no use case where the caller needs the range of added rows.
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

  getColumnDefaultFormula(columnName: string): string | undefined {
    if (!this.isFormulaTable) {
      throw new Error(`${this.state.gtable.name} is not a formula table`);
    }
    const tcolumnIndex = this.state.columnNameToIndex.get(columnName);
    if (tcolumnIndex == null) {
      return;
    }
    const gridRange = this.state.gtable.range;
    return this.sheet
      .getRange(
        gridRange.endRowIndex,
        gridRange.startColumnIndex + tcolumnIndex + 1,
      )
      .getFormula();
  }

  /*
   * Returns SpreadsheetApp Range for the table data
   */
  getRange(): GoogleAppsScript.Spreadsheet.Range | undefined {
    const gridRange = this.state.gtable.range;
    return this.getRangeForColumns(
      0,
      gridRange.endColumnIndex - gridRange.startColumnIndex,
    );
  }

  // startColumnIndex is 0-based relative to the table.
  private getRangeForColumns(startColumnIndex: number, numColumns: number) {
    const gridRange = this.state.gtable.range;
    if (this.getNumRows() == 0) {
      return;
    }
    const gridStartDataRowIndex = gridRange.startRowIndex + 1; // For header row
    return this.sheet.getRange(
      gridStartDataRowIndex + 1,
      gridRange.startColumnIndex + startColumnIndex + 1,
      gridRange.endRowIndex - this.lastRowAdjustment - gridStartDataRowIndex,
      numColumns,
    );
  }

  refreshState(): TableHelperState {
    this.state = this.getState(this.spreadsheetId, this.sheet, this.tableName);
    return this.state;
  }
}
