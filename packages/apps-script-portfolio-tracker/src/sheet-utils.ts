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
    if (typeof Sheets === "undefined") {
      throw new OurError("Must enable Sheets service");
    }
    const sheetsService = Sheets as MyGoogleAppsScript.Sheets;
    return new TableHelper(sheetsService, spreadsheetId, sheet, tableName);
  } catch (e) {
    if (e instanceof OurError) {
      Logger.log(e.message);
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
  lastTableRowInSheet: number;
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
    const lastTableRowInSheet = Math.max(
      ...(_gtables ?? []).map((t) => t.range?.endRowIndex ?? 0),
    );
    const columnNameToIndex = new Map<string, number>(
      gtable.columnProperties.map((p) => [p.columnName, p.columnIndex]),
    );
    return { gsheet, gtable, columnNameToIndex, lastTableRowInSheet };
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

  updateRowCount(newRowCount: number, pruneRows = false) {
    if (newRowCount < this.getNumRows()) {
      this.shrinkRowCount(newRowCount);
    } else if (newRowCount >= this.getNumRows()) {
      this.expandRowCount(newRowCount);
    }
    if (pruneRows) {
      const lastRow = this.sheet.getMaxRows();
      const deleteCount = Math.max(
        lastRow - this.state.lastTableRowInSheet - 1,
        0,
      );
      if (deleteCount) {
        this.sheet.deleteRows(lastRow - deleteCount + 1, deleteCount);
      }
    }
  }

  private shrinkRowCount(newRowCount: number) {
    if (newRowCount + this.lastRowAdjustment < 2) {
      throw new Error(
        "Must have at least two table rows including empty row for formula table",
      );
    }
    const r = this.getRange();
    if (r == null) {
      throw new Error("Failed to get existing range");
    }
    this.sheet
      .getRange(
        r.getRow(),
        r.getColumn(),
        r.getNumRows() - newRowCount,
        r.getNumColumns(),
      )
      .deleteCells(SpreadsheetApp.Dimension.ROWS);

    // Refresh state after mutation
    this.refreshState();
  }

  private expandRowCount(newRowCount: number) {
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
    const rowsActuallyNeeded = newRowCount + this.lastRowAdjustment;
    const totalRowsToAdd = rowsActuallyNeeded - numRows;
    let numRowsAvailable = numRows - 1; // We have to insert after the first row to avoid breaking named ranges.
    let remainingRowsToAdd = totalRowsToAdd;
    while (remainingRowsToAdd > 0) {
      const rowsToAdd = Math.min(remainingRowsToAdd, numRowsAvailable);
      if (rowsToAdd <= 0) {
        throw new Error("Adding new rows would break named ranges");
      }
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
    if (this.getNumRows() === 0) {
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

  private refreshState(): TableHelperState {
    this.state = this.getState(this.spreadsheetId, this.sheet, this.tableName);
    return this.state;
  }
}
