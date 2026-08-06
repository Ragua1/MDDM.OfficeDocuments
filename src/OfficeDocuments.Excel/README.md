# OfficeDocuments.Excel

A .NET library for creating and reading Excel (`.xlsx`) documents via the Open XML SDK,
with a clean, fluent interface that keeps OpenXml internals out of consumer code.

## Install

```sh
dotnet add package OfficeDocuments.Excel
```

## Quick start

```csharp
using OfficeDocuments.Excel;

// Create
using (var spreadsheet = new Spreadsheet("report.xlsx", createNew: true))
{
    var sheet = spreadsheet.AddWorksheet("Summary");

    sheet.AddRows(
    [
        ["Product", "Amount"],
        ["Widget", 1250.50m],
        ["Gadget", 890.00m],
    ]);

    sheet.AddCellWithFormula(2, 4, "SUM(B2:B3)");
}

// Read
using (var spreadsheet = new Spreadsheet("report.xlsx", createNew: false))
{
    var sheet = spreadsheet.GetWorksheet("Summary");
    var product = sheet?.GetCellByReference("A2")?.GetStringValue();

    if (sheet?.GetCell(2, 2)?.TryGetValue(out decimal amount) == true)
    {
        // Use amount.
    }
}
```

Streams work the same way: `Spreadsheet.CreateDocument(stream)` and
`Spreadsheet.OpenDocument(stream, isEditable)`. Rows and columns are **1-based**. The workbook is
written when you call `Close()` or dispose it.

## What it covers

Workbooks and worksheets, rows, cells, and rectangular ranges; typed reads; formulas, hyperlinks,
and comments; reusable styles with merging; bulk insert from collections or objects; sorting,
auto-filter, data validation, conditional formatting, freeze panes, auto-fit; worksheet lifecycle
operations; named ranges, protection, structured tables, and worksheet images.

## Targets

`net8.0` · `net9.0` · `net10.0`

## Links

- [Full documentation](https://github.com/Ragua1/MDDM.OfficeDocuments/blob/master/.docs/excel-library.md)
- [Repository](https://github.com/Ragua1/MDDM.OfficeDocuments)
- [Changelog / releases](https://github.com/Ragua1/MDDM.OfficeDocuments/releases)
