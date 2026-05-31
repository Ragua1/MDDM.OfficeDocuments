# OfficeDocuments.Excel

Date: 2026-05-31

This guide describes the current public Excel API as implemented in `src/OfficeDocuments.Excel/Interfaces/*` and exercised by the current test suite.

## Scope

`OfficeDocuments.Excel` is the primary and more mature module in this repository.

The library is designed for:

- creating and opening `.xlsx` workbooks from files or streams
- writing business-oriented spreadsheet content with a small object model
- reading values back as strings or typed values
- applying styles and light workbook automation features without forcing direct OpenXml usage

The default object model is:

`ISpreadsheet -> IWorksheet -> IRange / IRow -> ICell`

## What the library supports today

### Workbook workflows

- Create a workbook in memory with `Spreadsheet.CreateDocument(Stream)`
- Open a workbook from a stream with `Spreadsheet.OpenDocument(Stream, bool isEditable = true)`
- Create or open a workbook on disk with `new Spreadsheet(string filePath, bool createNew)`
- Add, rename, move, copy, hide, and remove worksheets
- Create reusable styles with `ISpreadsheet.CreateStyle(...)`
- Create, look up, rename, resize, enumerate, and remove structured tables
- Add named ranges
- Protect workbook structure

### Worksheet workflows

- Add rows sequentially or by explicit row index
- Add cells sequentially or by explicit row and column coordinates
- Get cells by coordinates or A1 reference
- Work with rectangular ranges through `GetRange(...)` and `TryGetRange(...)`
- Bulk insert rows from nested collections or object collections
- Set column widths, auto-fit columns, freeze panes, and clear frozen panes
- Protect worksheets
- Embed worksheet images from streams or files

### Range workflows

- Read values row-by-row through `GetValues()`
- Write rectangular value sets with `SetValues(...)`
- Apply styles to the whole range
- Merge a range
- Apply auto-filter
- Sort by a relative column index
- Add data validation
- Add conditional formatting

### Cell workflows

- Write values, formulas, hyperlinks, and comments
- Read values back with typed getters and `TryGetValue(...)`
- Detect value or formula presence

## Main API surface

### `ISpreadsheet`

- `AddWorksheet(...)`
- `GetWorksheet(...)`
- `GetWorksheetsName()`
- `RenameWorksheet(...)`
- `MoveWorksheet(...)`
- `CopyWorksheet(...)`
- `SetWorksheetHidden(...)`
- `RemoveWorksheet(...)`
- `CreateStyle(...)`
- `AddTable(...)`
- `GetTable(...)`
- `GetTables(...)`
- `RenameTable(...)`
- `ResizeTable(...)`
- `RemoveTable(...)`
- `AddNamedRange(...)`
- `ProtectWorkbook(...)`
- `Close()`

### `IWorksheet`

- `AddRow(...)`
- `AddCell(...)`
- `AddCellWithFormula(...)`
- `AddCellOnRange(...)`
- `GetRange(...)`
- `TryGetRange(...)`
- `GetRow(...)`
- `GetCell(...)`
- `GetCellByReference(...)`
- `AddRows(...)`
- `SetColumnWidth(...)`
- `FreezePanes(...)`
- `ClearFrozenPanes()`
- `AutoFitColumns()`
- `Protect(...)`
- `AddImage(...)`

### `IRange`

- `GetCell(...)`
- `GetValues()`
- `SetValues(...)`
- `ApplyStyle(...)`
- `Merge()`
- `ApplyAutoFilter()`
- `SortByColumn(...)`
- `AddValidation(...)`
- `AddConditionalFormatting(...)`

### `IRow`

- `AddCell(...)`
- `AddCellWithFormula(...)`
- `AddCellOnRange(...)`
- `GetCell(...)`
- `GetCellByReference(...)`

### `ICell`

- `SetValue(...)`
- `SetFormula(...)`
- `SetHyperlink(...)`
- `GetHyperlink()`
- `SetComment(...)`
- `GetComment()`
- `GetFormula()`
- `GetStringValue()`
- `GetBoolValue()`
- `GetIntValue()`
- `GetLongValue()`
- `GetDoubleValue()`
- `GetDecimalValue()`
- `GetDateValue(...)`
- `TryGetValue(...)`
- `HasValue()`
- `HasFormula()`

## Usage examples

### Create a workbook in memory

```csharp
using System.IO;
using OfficeDocuments.Excel;

var stream = new MemoryStream();

using var spreadsheet = Spreadsheet.CreateDocument(stream);
var worksheet = spreadsheet.AddWorksheet("Summary");

worksheet.AddRow().AddCell("Hello");
worksheet.GetRow()?.AddCell("World");

spreadsheet.Close();
```

### Create or open a workbook on disk

```csharp
using OfficeDocuments.Excel;

const string filePath = "report.xlsx";

using (var spreadsheet = new Spreadsheet(filePath, createNew: true))
{
    spreadsheet.AddWorksheet("Sheet1");
    spreadsheet.Close();
}

using (var spreadsheet = new Spreadsheet(filePath, createNew: false))
{
    var worksheet = spreadsheet.GetWorksheet("Sheet1");
    var worksheetNames = spreadsheet.GetWorksheetsName();

    spreadsheet.Close();
}
```

### Create and apply styles

```csharp
using OfficeDocuments.Excel;
using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Styles;
using Color = System.Drawing.Color;

using var spreadsheet = new Spreadsheet("styled.xlsx", createNew: true);
var worksheet = spreadsheet.AddWorksheet("Sheet1");

var style = spreadsheet.CreateStyle(
    font: new Font
    {
        Bold = true,
        FontSize = 12,
        FontName = FontNameValues.Calibri,
        Color = Color.DarkBlue
    },
    fill: new Fill(Color.LightYellow),
    border: new Border
    {
        Top = BorderStyleValues.Thin,
        Right = BorderStyleValues.Thin,
        Bottom = BorderStyleValues.Thin,
        Left = BorderStyleValues.Thin
    },
    alignment: new Alignment
    {
        Horizontal = HorizontalAlignmentValues.Center,
        Vertical = VerticalAlignmentValues.Center,
        WrapText = true
    },
    numberFormat: new NumberingFormat("#,##0.00")
);

worksheet.AddRow(style).AddCell(1234.56m);
spreadsheet.Close();
```

### Merge styles

```csharp
using OfficeDocuments.Excel;
using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Styles;
using Color = System.Drawing.Color;

using var spreadsheet = new Spreadsheet("merged-style.xlsx", createNew: true);
var worksheet = spreadsheet.AddWorksheet("Sheet1");

var fontStyle = spreadsheet.CreateStyle(font: new Font { Bold = true, Color = Color.DarkGreen });
var fillStyle = spreadsheet.CreateStyle(fill: new Fill(Color.Honeydew));
var borderStyle = spreadsheet.CreateStyle(border: new Border(BorderStyleValues.Medium));

var headerStyle = fontStyle.CreateMergedStyle(fillStyle).CreateMergedStyle(borderStyle);

worksheet.AddRow(headerStyle).AddCell("Merged style example");
spreadsheet.Close();
```

### Work with ranges and bulk insert

```csharp
using OfficeDocuments.Excel;
using OfficeDocuments.Excel.Enums;

using var spreadsheet = new Spreadsheet("range-ops.xlsx", createNew: true);
var worksheet = spreadsheet.AddWorksheet("Sheet1");

var insertedRange = worksheet.AddRows(
[
    ["Name", "Score"],
    ["Alice", 95],
    ["Bob", 88],
    ["Cara", 99]
]);

var dataRange = worksheet.GetRange("A1:B4");
dataRange.ApplyAutoFilter();
dataRange.SortByColumn(2, SortDirection.Descending, hasHeader: true);

spreadsheet.Close();
```

### Add formulas

```csharp
using OfficeDocuments.Excel;

using var spreadsheet = new Spreadsheet("formulas.xlsx", createNew: true);
var worksheet = spreadsheet.AddWorksheet("Sheet1");

worksheet.AddRow().AddCell(100).AddCell(200);
worksheet.AddRow().AddCell(300).AddCell(400);

var totalRow = worksheet.AddRow();
totalRow.AddCell("Total");
totalRow.AddCellWithFormula("SUM(B1:B2)");

spreadsheet.Close();
```

### Read values back

```csharp
using OfficeDocuments.Excel;

using var spreadsheet = new Spreadsheet("existing.xlsx", createNew: false);
var worksheet = spreadsheet.GetWorksheet("Sheet1");

var cell = worksheet?.GetCellByReference("B2");
var rawValue = cell?.Value;
var stringValue = cell?.GetStringValue();

if (cell?.TryGetValue(out decimal amount) == true)
{
    // Use amount.
}

if (cell?.TryGetValue(out DateTime dateValue) == true)
{
    // Use dateValue.
}

spreadsheet.Close();
```

### Add validation, formatting, hyperlinks, comments, named ranges, and protection

```csharp
using OfficeDocuments.Excel;
using OfficeDocuments.Excel.Options;
using OfficeDocuments.Excel.Styles;
using Color = System.Drawing.Color;

using var spreadsheet = new Spreadsheet("advanced.xlsx", createNew: true);
var worksheet = spreadsheet.AddWorksheet("Sheet1");
var style = spreadsheet.CreateStyle(fill: new Fill(Color.LightGoldenrodYellow));
var range = worksheet.GetRange("A1:A3");

range.SetValues(
[
    ["A"],
    ["B"],
    ["C"]
]);

range.AddValidation(DataValidationOptions.List(["A", "B", "C"]));
range.AddConditionalFormatting(ConditionalFormattingOptions.ContainsText("A", style));

var cell = worksheet.GetCell(1, 1)!;
cell.SetHyperlink("https://example.com", "Docs");
cell.SetComment("Validated value");

spreadsheet.AddNamedRange("Codes", range);
worksheet.Protect("secret");
spreadsheet.ProtectWorkbook("secret");
spreadsheet.Close();
```

### Create and manage structured tables

```csharp
using OfficeDocuments.Excel;
using OfficeDocuments.Excel.Options;

using var spreadsheet = new Spreadsheet("table.xlsx", createNew: true);
var worksheet = spreadsheet.AddWorksheet("Scores");

worksheet.AddRows(
[
    ["Name", "Score"],
    ["Alice", 95],
    ["Bob", 88]
]);

var range = worksheet.GetRange("A1:B3");
var tableInfo = spreadsheet.AddTable(range, ["Name", "Score"], new TableCreateOptions
{
    TableName = "ScoresTable",
    Style = new TableStyleOptions
    {
        StyleName = "TableStyleMedium2",
        ShowBandedRows = true
    }
});

var found = spreadsheet.GetTable("Scores", "ScoresTable");
var worksheetTables = spreadsheet.GetTables("Scores");
var allTables = spreadsheet.GetTables();

spreadsheet.RenameTable("Scores", "ScoresTable", "PlayerScores");
spreadsheet.ResizeTable("Scores", "PlayerScores", worksheet.GetRange("A1:B5"));
spreadsheet.RemoveTable("Scores", "PlayerScores");

spreadsheet.Close();
```

### Embed images in a worksheet

```csharp
using OfficeDocuments.Excel;
using OfficeDocuments.Excel.Enums;

using var spreadsheet = new Spreadsheet("report.xlsx", createNew: true);
var worksheet = spreadsheet.AddWorksheet("Sheet1");

worksheet.AddImage("logo.png", fromColumn: 1, fromRow: 1, toColumn: 4, toRow: 5);

using var imageStream = File.OpenRead("chart.jpeg");
worksheet.AddImage(imageStream, ImageType.Jpeg, fromColumn: 6, fromRow: 1, toColumn: 10, toRow: 5);

spreadsheet.Close();
```

## Consumer notes

- Rows and columns are 1-based.
- `AddCell(...)` is the preferred value-writing API.
- `AddCellWithValue(...)` remains available only for compatibility and is obsolete.
- `GetRange(...)` and `AddRows(...)` are the preferred high-level entry points for multi-cell work.
- Formula support writes formulas to the workbook. It does not provide an Excel calculation engine.
- The document is persisted when you call `Close()` or dispose the spreadsheet.
- Some raw OpenXml-oriented members still exist for compatibility. They should be treated as non-preferred interop surfaces rather than the default API.

## Related documents

- [README.md](README.md)
- [word-library.md](word-library.md)
- [terminology.md](terminology.md)
- [library-benchmark-report.md](library-benchmark-report.md)
