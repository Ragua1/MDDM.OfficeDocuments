# Migrating OfficeDocuments.Excel from v3 to v4

Date: 2026-08-06

Version 4 is a major release. It carries two kinds of breaking change:

1. **A package split** — the advanced Excel features moved out of `OfficeDocuments.Excel` into a new
   optional package, `OfficeDocuments.Excel.Advanced` (EXCEL-010 Tier 2).
2. **Stricter, more correct behaviour** — several previously-accepted-but-wrong inputs now throw, and
   a few values are now written in their OOXML-conformant form.

If you only use cells, rows, ranges, and styles, the package split does not affect you; skip to
[Other v4 breaking changes](#other-v4-breaking-changes).

## 1. The advanced-features package split

### What moved

These APIs are no longer on the core `ISpreadsheet` / `IWorksheet` interfaces. They now live in
`OfficeDocuments.Excel.Advanced` as **extension methods** over the same interfaces:

| Feature | Members | Was on | Now |
| --- | --- | --- | --- |
| Structured tables | `AddTable`, `GetTable`, `GetTables`, `RenameTable`, `ResizeTable`, `RemoveTable` | `ISpreadsheet` | extension on `ISpreadsheet` |
| Named ranges | `AddNamedRange` | `ISpreadsheet` | extension on `ISpreadsheet` |
| Workbook protection | `ProtectWorkbook` | `ISpreadsheet` | extension on `ISpreadsheet` |
| Worksheet protection | `Protect` | `IWorksheet` | extension on `IWorksheet` |
| Image embedding | `AddImage` | `IWorksheet` | extension on `IWorksheet` |

These supporting types also moved into the `OfficeDocuments.Excel.Advanced` namespace:

| Type | Old namespace | New namespace |
| --- | --- | --- |
| `ITableInfo` | `OfficeDocuments.Excel.Interfaces` | `OfficeDocuments.Excel.Advanced` |
| `TableCreateOptions` | `OfficeDocuments.Excel.Options` | `OfficeDocuments.Excel.Advanced` |
| `TableStyleOptions` | `OfficeDocuments.Excel.Options` | `OfficeDocuments.Excel.Advanced` |
| `ImageType` | `OfficeDocuments.Excel.Enums` | `OfficeDocuments.Excel.Advanced` |

### What did **not** move

Data validation, conditional formatting, hyperlinks, and comments stay in the core package — they
are on `IRange` (`AddValidation`, `AddConditionalFormatting`) and `ICell` (`SetHyperlink`,
`GetHyperlink`, `SetComment`, `GetComment`), and are unchanged.

### How to migrate

1. Add the package:

   ```
   dotnet add package OfficeDocuments.Excel.Advanced
   ```

2. Add one `using` to each file that uses an advanced feature. It brings both the extension methods
   and the moved types into scope:

   ```csharp
   using OfficeDocuments.Excel.Advanced;
   ```

3. The call sites do not change. This code compiles unchanged once the `using` is present:

   ```csharp
   using OfficeDocuments.Excel;
   using OfficeDocuments.Excel.Advanced;

   using var spreadsheet = Spreadsheet.CreateDocument(stream);
   var sheet = spreadsheet.AddWorksheet("Data");
   var range = sheet.GetRange("A1:C10");

   spreadsheet.AddTable(range, ["Id", "Name", "Value"]); // extension method
   spreadsheet.AddNamedRange("MyData", range);           // extension method
   sheet.AddImage("logo.png", 1, 1, 3, 5);               // extension method
   sheet.Protect("secret");                              // extension method
   spreadsheet.ProtectWorkbook("secret");                // extension method
   ```

If you drop an advanced type's old namespace import (e.g. `using OfficeDocuments.Excel.Enums;` used
only for `ImageType`), replace it with `using OfficeDocuments.Excel.Advanced;`.

### Why it changed

The two central classes were decomposed into a small core plus focused collaborators (EXCEL-010
Tier 1). Tier 2 finishes that by moving the heavier, less-common features into a separate package so
a consumer who only needs the minimal core does not carry structured-table, named-range, protection,
and image code. The advanced layer drives the same internal document state through the same
collaborators; it is an add-on, not a fork. See
[tasks/core/excel/EXCEL-010-god-class-decomposition.md](tasks/core/excel/EXCEL-010-god-class-decomposition.md)
and [architecture/target-package-boundaries-and-instantiation.md](architecture/target-package-boundaries-and-instantiation.md).

### Note on custom `ISpreadsheet` / `IWorksheet` implementations

The advanced extensions require the built-in `Spreadsheet` / `Worksheet` types — they reach the
underlying OpenXml parts that only those types own. Calling an advanced extension on a foreign
implementation throws `ArgumentException`. In practice every instance you get from
`Spreadsheet.CreateDocument` / `OpenDocument` / `AddWorksheet` is the built-in type, so normal code
is unaffected.

## Other v4 breaking changes

These landed on the 4.0.0 line independently of the package split, as part of the correctness
hardening pass. Full rationale is in
[excel-state-verdict.md](excel-state-verdict.md); the highlights:

- **`ICell.GetFormulaValue()` now returns `double`** (was `int`). It is also row-aware, matches
  function names exactly (`SUMIF` ≠ `SUM`), throws `NotSupportedException` for unknown functions and
  `InvalidOperationException` when the cell has no formula, and no longer truncates `MEDIAN`.
- **Boolean cells serialize as `1` / `0`** (the OOXML-conformant form) instead of `True` / `False`.
  Reading stays tolerant of the old form.
- **Stricter input validation now throws** where v3 silently produced a schema-invalid or
  wrong-reading file: non-finite numbers (`NaN`, `Infinity`), XML control characters in cell text,
  invalid worksheet names (length/characters), and malformed ARGB hex colours.
- **Integer and long cell reads are culture-invariant**, so a non-US current culture no longer
  changes what a numeric cell parses to.
- **Dates before 1900-03-01** are written with the serial number Excel expects (v3 was one day off
  for that range).
- **`AddCellOnRange(...)` returns a non-nullable `ICell` and throws on an inverted range.** Every
  overload on `IWorksheet` and `IRow` is affected. In v3 an out-of-order range (`endColumn` before
  `beginColumn`, `endRow` before `beginRow`) returned `null` while an out-of-bounds index threw, so
  the same method reported two failure classes two different ways and forced a null check on every
  call site. All of them now throw `ArgumentException` naming the offending parameter. A range that
  covers exactly one cell is valid, returns that cell, and writes no `mergeCell` element — v3 either
  returned `null` for it (the `IRow` and 3-argument `IWorksheet` overloads) or wrote a degenerate
  one-cell merge (the 5-argument overload).

  ```csharp
  // v3 — the null was the only signal, and only for some invalid inputs
  var cell = worksheet.AddCellOnRange(5, 4, 1);
  if (cell is null) { /* ... */ }

  // v4 — invalid input throws, valid input always returns a cell
  var cell = worksheet.AddCellOnRange(2, 4, 1);
  cell.SetValue("Header");
  ```

## Version alignment

`OfficeDocuments.Excel` and `OfficeDocuments.Excel.Advanced` ship on the same version line. Install
matching versions.
