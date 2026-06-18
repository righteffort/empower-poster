declare namespace Sheets {
  // Convenience types
  type Spreadsheet = import("@googleapis/sheets").sheets_v4.Schema$Spreadsheet;
  type Table = import("@googleapis/sheets").sheets_v4.Schema$Table;
  type SafeTable = {
    [K in keyof Table]-?: NonNullable<X[Table]>;
  };
  type GridRange = import("@googleapis/sheets").sheets_v4.Schema$GridRange;
  type ValueRange = import("@googleapis/sheets").sheets_v4.Schema$ValueRange;
  type SafeGridRange = {
    [K in keyof GridRange]-?: NonNullable<X[GridRange]>;
  };
  namespace Spreadsheets {
    type GetOptionalArgs = Omit<
      import("@googleapis/sheets").sheets_v4.Params$Resource$Spreadsheets$Get,
      "spreadsheetId"
    >;
    function get(
      spreadsheetId: string,
      optionalArgs: SpreadsheetsGetOptionalArgs? = null,
    ): Spreadsheet;
    namespace Values {
      type GetOptionalArgs = Omit<
        import("@googleapis/sheets").sheets_v4.Params$Resource$Spreadsheets$Values$Update,
        "range",
        "spreadsheetId"
      >;
      type UpdateValuesResponse =
        import("googleapis/sheets").sheets_v4.Schema$UpdateValuesResponse;
      function update(
        valueRange: ValueRange,
        spreadsheetId: string,
        range: string,
        options: GetOptionalArgs,
      ): UpdateValuesResponse;
    }
  }
}
