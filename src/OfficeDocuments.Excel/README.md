# OfficeDocuments.Excel

A .NET library for creating and reading Excel (`.xlsx`) documents via the OpenXml SDK,
with a clean, fluent interface that keeps OpenXml internals out of consumer code.

## Install

```
dotnet add package OfficeDocuments.Excel
```

## Quick start

```csharp
using OfficeDocuments.Excel;

// Create
using var doc = Spreadsheet.Create("report.xlsx");
var sheet = doc.AddWorksheet("Sheet1");
var row = sheet.AddRow(1);
row.AddCell(1).SetValue("Hello");
row.AddCell(2).SetValue(42);
doc.Save();

// Read
using var doc = Spreadsheet.Open("report.xlsx");
var sheet = doc.GetWorksheet("Sheet1");
var value = sheet.GetRow(1)?.GetCell(1)?.GetValue<string>();
```

## Targets

`net8.0` · `net9.0` · `net10.0`

## Links

- [Full documentation](.doc/excel-library.md)
- [Repository](https://github.com/Ragua1/MDDM.OfficeDocuments)
- [Changelog / releases](https://github.com/Ragua1/MDDM.OfficeDocuments/releases)
