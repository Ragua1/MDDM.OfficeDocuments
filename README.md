# OfficeDocuments

[![Excel build](https://github.com/Ragua1/MDDM.OfficeDocuments/actions/workflows/github-build-excel.yml/badge.svg)](https://github.com/Ragua1/MDDM.OfficeDocuments/actions/workflows/github-build-excel.yml)
[![Word build](https://github.com/Ragua1/MDDM.OfficeDocuments/actions/workflows/github-build-word.yml/badge.svg)](https://github.com/Ragua1/MDDM.OfficeDocuments/actions/workflows/github-build-word.yml)
[![NuGet Excel](https://img.shields.io/nuget/v/OfficeDocuments.Excel.svg?label=OfficeDocuments.Excel)](https://www.nuget.org/packages/OfficeDocuments.Excel/)
[![NuGet Excel.Advanced](https://img.shields.io/nuget/v/OfficeDocuments.Excel.Advanced.svg?label=OfficeDocuments.Excel.Advanced)](https://www.nuget.org/packages/OfficeDocuments.Excel.Advanced/)
[![NuGet Word](https://img.shields.io/nuget/v/OfficeDocuments.Word.svg?label=OfficeDocuments.Word)](https://www.nuget.org/packages/OfficeDocuments.Word/)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE.md)

Generate and read `.xlsx` and `.docx` from .NET without learning the Open XML SDK.

`DocumentFormat.OpenXml` is complete, fast, and faithful to the file format — which is exactly the
problem when all you need is a report with a header row and a total. `OfficeDocuments` is a smaller,
task-oriented API over it: a workbook is worksheets, rows, cells and ranges; a document is
paragraphs, tables and images. The Open XML types stay behind the interface.

```csharp
using OfficeDocuments.Excel;

record Sale(string Region, decimal Revenue);

var sales = new[] { new Sale("North", 1_240_000m), new Sale("South", 980_000m) };

using var spreadsheet = new Spreadsheet("report.xlsx", createNew: true);
var sheet = spreadsheet.AddWorksheet("Summary");

sheet.AddRows(sales, includeHeader: true);
sheet.AutoFitColumns();

spreadsheet.Close();
```

## Is this the right library for you?

**A good fit when** you generate business documents server-side — reports, exports, invoices,
protocols, filled templates — and you want a small API, no native dependencies, and files that open
in Excel and Word without a repair prompt.

**Look elsewhere when** you need to *calculate* formulas rather than write them (this library writes
formulas; it does not ship a full calculation engine), or you need legacy binary `.xls` / `.doc`,
which is permanently out of scope.

**Honest positioning.** This is a wrapper, and so are most libraries in this space — including the
one you are probably comparing it to. [ClosedXML](https://github.com/ClosedXML/ClosedXML) and
[NPOI](https://github.com/nissl-lab/npoi) are larger, older, and have more contributors; if you want
the broadest possible feature surface, use them. What this project offers instead is a deliberately
narrow API, both formats behind one consistent design, and a correctness bar described below. Pick
on that basis, not on a feature-count table.

## Install

The Excel and Word packages are independent — neither drags in the other. Install only what you need.
Excel's heavier, less-common features live in an optional add-on, `OfficeDocuments.Excel.Advanced`.

```sh
dotnet add package OfficeDocuments.Excel
dotnet add package OfficeDocuments.Excel.Advanced   # optional: tables, named ranges, protection, images
dotnet add package OfficeDocuments.Word
```

Targets `net8.0`, `net9.0`, and `net10.0`. The only runtime dependency is `DocumentFormat.OpenXml`.
Upgrading from v3? See the [migration guide](.doc/migration-v3-to-v4.md).

## Excel

```csharp
using OfficeDocuments.Excel;
using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Styles;
using Color = System.Drawing.Color;

using var spreadsheet = new Spreadsheet("report.xlsx", createNew: true);
var sheet = spreadsheet.AddWorksheet("Summary");

// Styles are values: create once, reuse everywhere, compose by merging.
var header = spreadsheet.CreateStyle(
    font: new Font { Bold = true, FontSize = 12, Color = Color.DarkBlue },
    fill: new Fill(Color.LightYellow));
var boxed = header.CreateMergedStyle(
    spreadsheet.CreateStyle(border: new Border(BorderStyleValues.Thin)));

var headerRow = sheet.AddRow(boxed);
headerRow.AddCell("Region");
headerRow.AddCell("Revenue");

var dataRow = sheet.AddRow();
dataRow.AddCell("North");
dataRow.AddCell(1_240_000m);

sheet.GetRange("A1:B1").ApplyAutoFilter();
sheet.AutoFitColumns();

spreadsheet.Close();
```

Server-side generation works over streams, with no file system involved:

```csharp
var stream = new MemoryStream();
using var spreadsheet = Spreadsheet.CreateDocument(stream);
// ...
spreadsheet.Close();
```

Beyond the basics, in the core package: ranges, bulk insert from object collections, typed reads with
`TryGetValue`, formulas, hyperlinks, comments, sorting, auto-filter, data validation, conditional
formatting, and freeze panes. In the optional `OfficeDocuments.Excel.Advanced` package: structured
tables, named ranges, worksheet images, and worksheet/workbook protection — added as extension
methods over the same objects, so a `using OfficeDocuments.Excel.Advanced;` is all it takes.
Full guide: [.doc/excel-library.md](.doc/excel-library.md).

## Word

```csharp
using OfficeDocuments.Word;
using OfficeDocuments.Word.Enums;
using OfficeDocuments.Word.Formatting;

using var document = new Wordprocessing("report.docx", createNew: true);

// Formatting is immutable records: null means inherit, false means an active override.
var bodyText = new TextFormat { FontName = "Calibri", FontSize = 11 };
var justified = new ParagraphFormat { Alignment = ParagraphAlignment.Justify, SpacingAfter = 6 };

var body = document.GetBody();
body.AddParagraph("Quarterly report", new ParagraphFormat { StyleId = WordStyleIds.Title });
body.AddHeading("Summary", 1);
body.AddParagraph("Revenue grew steadily.", justified, bodyText);

body.AddParagraph(justified)
    .AddText("Total: ", bodyText)
    .AddText("1 240 000 CZK", bodyText with { Bold = true });
```

The body, a header, a footer, and a table cell all implement the same `IBlockContainer`, so
paragraphs, headings, lists and tables behave identically in all four.

Beyond the basics: run and paragraph formatting, built-in styles and headings, bullet and numbered
lists, tables with repeating header rows and nested content, hyperlinks, inline images sized from the
image itself, headers and footers, page setup, document metadata, paragraph search, and text
replacement that survives the run boundaries Word inserts mid-word.
Full guide: [.doc/word-library.md](.doc/word-library.md).

Deliberately out of scope for now: multiple sections with differing page setups, footnotes,
bookmarks, comments, tracked changes, and a generated table of contents.

## Correctness

A generated file that round-trips through this library proves only that the library agrees with
itself. Two rules exist because of that:

- **Every test that produces a complete document ends with the Open XML schema validator.** Not a
  read-back assertion — the validator.
- **Element order is treated as a correctness invariant.** OOXML fixes the order of child elements.
  Emit them out of order and the file still round-trips perfectly and still reads back correctly; it
  only fails when Excel or Word opens it. That class of bug has hit this repository three times.

Even that is not sufficient, and the repository documents where. `<v>NaN</v>` in a numeric cell
passes schema validation, because `v` is declared `xsd:string`. `ToOADate`/`FromOADate` are exact
inverses, so a date serial can be wrong by one day in Excel's reckoning and still pass any
round-trip test. Both were real defects here, found by probing the running library rather than by
any gate the suite had.

Tests are split into tiers with explicit entry criteria — pure unit tests with no document, behaviour
tests through the public API over `MemoryStream`, whole-document verification including reopening and
reading foreign files, and performance guards. No performance test asserts on a duration; a
millisecond threshold measures the CI runner, not the code, so the guards assert growth ratios and
allocation counts instead. See [test/README.md](test/README.md).

## Documentation

| Document | Contents |
| --- | --- |
| [.doc/excel-library.md](.doc/excel-library.md) | Excel API guide, semantics, worked examples |
| [.doc/word-library.md](.doc/word-library.md) | Word API guide, semantics, worked examples |
| [.doc/library-benchmark-report.md](.doc/library-benchmark-report.md) | Capability comparison against ClosedXML, EPPlus, NPOI, openpyxl, python-docx |
| [.doc/excel-performance-baseline.md](.doc/excel-performance-baseline.md) | Measured baselines and the known hot spots |
| [.doc/tasks/roadmap-overview.md](.doc/tasks/roadmap-overview.md) | Planned work |
| [.doc/README.md](.doc/README.md) | Documentation index |

### Known performance characteristics

Documented rather than hidden: creating many *distinct* styles allocates quadratically. Reusing a
small set of styles instead of creating one per cell is dramatically faster and lighter — the
baseline document quantifies it. If you style a large sheet, create your styles once and reuse them.

## Building from source

```powershell
dotnet build OfficeDocuments.slnx
dotnet test  OfficeDocuments.slnx

dotnet test  OfficeDocuments.Excel.slnx    # one module at a time
dotnet test  OfficeDocuments.Word.slnx
```

The two modules are independent by design: neither references the other, and each per-module solution
omits the other, so a cross-reference fails the build instead of passing review. Package versions are
centralized in `Directory.Packages.props`; compilation settings in `Directory.Build.props`.

## Contributing

Contributions are welcome. Start with [AGENTS.md](AGENTS.md) — it states the working rules that apply
to every change, human or AI, and routes to the detailed guidance in
[.doc/ai-instructions/](.doc/ai-instructions/README.md). The essentials:

- Keep diffs minimal and aimed at the root cause; no drive-by refactors.
- The interface layer is the public API. Do not leak Open XML types across it.
- Every test producing a complete document must pass schema validation.
- Documentation is English, and snippets must match the real API.

Please open an issue before starting anything large, so you do not build something that does not fit
the roadmap.

## Support

Best-effort, no SLA. See [SUPPORT.md](SUPPORT.md) for what that means in practice and how to file a
report that can actually be acted on.

## License

[MIT](LICENSE.md). Copyright © MDDM / Martin Domanský.
