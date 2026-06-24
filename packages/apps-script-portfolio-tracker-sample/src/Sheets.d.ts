declare namespace Sheets {
  // Convenience types
  type Spreadsheet = import("@googleapis/sheets").sheets_v4.Schema$Spreadsheet;
  type Sheet = import("@googleapis/sheets").sheets_v4.Schema$Sheet;
  type Table = import("@googleapis/sheets").sheets_v4.Schema$Table;
  type SafeTable = {
    [K in keyof Table]-?: NonNullable<X[Table]>;
  };
  type GridRange = import("@googleapis/sheets").sheets_v4.Schema$GridRange;
  type SafeGridRange = {
    [K in keyof GridRange]-?: NonNullable<X[GridRange]>;
  };
  namespace Spreadsheets {
    type BatchUpdateSpreadsheetRequest =
      import("@googleapis/sheets").sheets_v4.Schema$BatchUpdateSpreadsheetRequest;
    type BatchUpdateSpreadsheetResponse =
      import("@googleapis/sheets").sheets_v4.Schema$BatchUpdateSpreadsheetResponse;
    function batchUpdate(
      resource: BatchUpdateSpreadsheetRequest,
      spreadsheetId: string,
    ): BatchUpdateSpreadsheetResponse;
    type GetOptionalArgs = Omit<
      import("@googleapis/sheets").sheets_v4.Params$Resource$Spreadsheets$Get,
      "spreadsheetId"
    >;
    function get(
      spreadsheetId: string,
      optionalArgs: Spreadsheets.GetOptionalArgs? = null,
    ): Spreadsheet;
  }
}
