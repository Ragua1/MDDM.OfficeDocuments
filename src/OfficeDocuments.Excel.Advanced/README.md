# OfficeDocuments.Excel.Advanced

Optional advanced layer for [`OfficeDocuments.Excel`](https://www.nuget.org/packages/OfficeDocuments.Excel).
It adds the heavier, less-common Excel features on top of the minimal core so a consumer who only
needs cells, rows, ranges, and styles does not carry them.

The advanced features are exposed as **extension methods** over the core `ISpreadsheet` and
`IWorksheet` interfaces. Add the package, add a `using`, and the calls light up on the objects you
already have.

## What is in here

| Feature | API |
| --- | --- |
| Structured tables | `ISpreadsheet.AddTable` / `GetTable` / `GetTables` / `RenameTable` / `ResizeTable` / `RemoveTable` |
| Named ranges | `ISpreadsheet.AddNamedRange` |
| Workbook protection | `ISpreadsheet.ProtectWorkbook` |
| Worksheet protection | `IWorksheet.Protect` |
| Image embedding | `IWorksheet.AddImage` |

Data validation, conditional formatting, hyperlinks, and comments are **not** here — they live on
`IRange` / `ICell` in the core package.

## Usage

```csharp
using OfficeDocuments.Excel;
using OfficeDocuments.Excel.Advanced; // brings the advanced extensions into scope

using var spreadsheet = Spreadsheet.CreateDocument(stream);
var sheet = spreadsheet.AddWorksheet("Data");
// ... write cells ...

var range = sheet.GetRange("A1:C10");
spreadsheet.AddTable(range, ["Id", "Name", "Value"]);
spreadsheet.AddNamedRange("MyData", range);
sheet.AddImage("logo.png", 1, 1, 3, 5);
spreadsheet.ProtectWorkbook("secret");
```

The extensions require the built-in `Spreadsheet` / `Worksheet` implementation from the core
package; they throw `ArgumentException` if handed a foreign `ISpreadsheet` / `IWorksheet`.

## Versioning

Ships in lockstep with `OfficeDocuments.Excel` on the same version line. See the
[v3 → v4 migration guide](https://github.com/Ragua1/MDDM.OfficeDocuments/blob/master/.doc/migration-v3-to-v4.md)
for what moved here in v4 and how to update your code.
