- ditch vendor/src/sheets_v4.d.ts as soon as possible
- expand/shrink tables to fit data!

maybe simplify
- this depends on https://issuetracker.google.com/issues/525219695 being fixed
- 1-time setup (but check each time)
  - add 3 empower sheets, with filter for Empower Holdings sheet
  - define 2 named functions
  - add / replace formulas for multiple columns (see my [manual spreadsheet](https://docs.google.com/spreadsheets/d/1bRJIR8yysVWtk6TS3frWLeK4ki04d_NJ9hlkGcPrI0Q/edit?gid=1843166815#gid=1843166815)
    - but not Accounts:Name & Accounts:Institution
  - adjust % formatting on Class Pct (no idea why this is needed)
  - adjust == 1 check on sum of Class Pct (should be upstream)
- each time
  - replace all data in the 3 empower sheets with lightly massaged empower data (e.g. blank name & price if cusip is present)
  - keep existing handling of 'Accounts' table to preserve manual columns (type, owner)
  - shrink/expand tables and tidy up sheets as before
- document limitations

```javascript
// standard formula descriptors
// ugh what about that final row that we're not allowed to write to???
// standard columns produce `=EMPOWER_VALUE("${tableName}", "${columnName}", ROW(Single(${tableName.replaceAll(' ', '_')}[${columnName}])) - 2)`,
[
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
    ]
  }
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
              formula: '=LET(altname, EMPOWER_VALUE("Assets", "Name", ROW(Single(Assets[Name])) - 2), if(NE(altname, ""), altname, ASSETNAME(Assets[Ticker])))',
            },
            {
              name: "Price",
              formula: '=LET(altprice, EMPOWER_VALUE("Assets", "Price", ROW(Single(Assets[Price])) - 2), if(NE(altprice, ""), altprice, ASSETPRICE(Assets[Ticker])))',
            },
         ],
       },
    ],
  }
]
```

`EMPOwER_SHEET(tablename)`
```
=ifs(eq(tablename, "Holdings"), "Empower Holdings",
     countif({"Assets"; "Asset Classes"; "Class Categories"},tablename) > 0, "Empower Assets",
     countif({"Accounts"; "Institutions"}, tablename) > 0, "Empower Accounts",
     true, na()
 )
 ```
 
 `EMPOWER_VALUE(tablename, columnname, rownumber)`
 ```
 =let(sheetname, EMPOWER_SHEET(tablename),
      headers, INDIRECT("'" & sheetname & "'!$A$1:$1"),
      columnindex, MATCH(tablename & ":" & columnname, headers, 0),
      reference, "'" & sheetname & "'!R"&rownumber & "C" & columnindex,
      INDIRECT(reference, FALSE)
 )
 ```
 
