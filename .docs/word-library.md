# OfficeDocuments.Word

Date: 2026-07-28

This guide describes the current public Word API as implemented in `src/OfficeDocuments.Word/Interfaces/*`
and `src/OfficeDocuments.Word/Formatting/*`, and exercised by `test/OfficeDocuments.Word.Tests`.
Package version `4.0.0`.

## Scope

`OfficeDocuments.Word` covers authoring and reading the content of a `.docx` with the structure and
formatting that ordinary business documents need: reports, letters, protocols, invoices, and records.

The object model is:

```text
IWordprocessing
├── IBody ─────────────┐
├── IHeaderFooter ─────┤ all are IBlockContainer
└── ITable             │
    └── ITableRow      │
        └── ITableCell ┘
                │
                └── IParagraph → IRun → IText
```

`IBlockContainer` is the shared contract: the body, a header, a footer, and a table cell all hold
block content on identical terms, so paragraphs, headings, lists, and tables work the same in all four.

## What the library supports today

| Area | Features |
| --- | --- |
| Documents | Create or open from a file or stream, read-only open, idempotent close |
| Text | Paragraphs, runs, breaks, multi-line text, whitespace-faithful reading |
| Run formatting | Bold, italic, underline, strikethrough, all-caps, small caps, highlight, superscript and subscript, font, size, colour, character styles |
| Paragraph formatting | Alignment, spacing, line spacing, indentation, hanging indent, page-break-before, keep-with-next, keep-lines, named styles |
| Styles | Built-in `Title`, `Subtitle`, `Heading1`–`Heading6`, and `Hyperlink`, defined on first use |
| Lists | Bullet and numbered lists, nine nesting levels |
| Tables | Create by size or from data, rows and cells, header rows that repeat, width, alignment, borders, cell padding, shading, vertical alignment, column spanning, nested tables |
| Hyperlinks | External links with separate display text |
| Images | Inline images from a stream or a file, intrinsic or explicit sizing, alternative text |
| Page setup | Paper size, orientation, margins, header and footer distance |
| Headers and footers | Default, first-page, and even-page, each holding full block content |
| Metadata | Title, subject, author, keywords, description, category, last modified by, dates |
| Search | Find paragraphs by text, walk every paragraph including table content |
| Update | Replace text across runs, per paragraph, per container, or document-wide; set a paragraph's text; remove paragraphs, tables, and table rows |

Not supported, and deliberately out of scope for now: multiple sections with different page setups,
footnotes and endnotes, bookmarks and internal links, comments, tracked changes, and a generated table
of contents.

## Units

The public API measures everything in **points**. WordprocessingML does not: font size is in
half-points, spacing and indentation in twentieths of a point, borders in eighths of a point, and
drawings in English Metric Units. Those conversions happen inside the library, so `FontSize = 10.5`
and `MarginLeft = 56` mean what they say.

Two exceptions, both because the underlying value is not a length:

- `LineSpacing` is a multiple of single spacing, so `1.5` is one-and-a-half lines.
- `WidthPercent` on tables and cells is a percentage of the available width.

## Main API surface

### `IWordprocessing`

- `new Wordprocessing(string filePath, bool createNew, bool isEditable = true)`
- `new Wordprocessing(Stream stream, bool createNew, bool isEditable = true)`
- `GetBody()` — returns the same instance on every call
- `AddHeader(HeaderFooterKind kind = Default)` / `AddFooter(...)` — idempotent per kind
- `HeadersAndFooters` — read from the document, so an opened document reports the ones it already had
- `ReplaceText(string oldValue, string newValue, StringComparison = Ordinal)` — body, tables, headers,
  and footers
- `PageSetup` / `ApplyPageSetup(PageSetup setup)`
- `Metadata` / `SetMetadata(DocumentMetadata metadata)`
- `Close(bool saveDocument = true)` — idempotent; `false` genuinely discards
- `Dispose()`

### `IBlockContainer`

Implemented by `IBody`, `IHeaderFooter`, and `ITableCell`.

- `Paragraphs`, `Tables`
- `AddParagraph()`, `AddParagraph(ParagraphFormat?)`, `AddParagraph(string)`,
  `AddParagraph(string, ParagraphFormat?, TextFormat?)`
- `AddHeading(string text, int level)` — level 1 to 6
- `AddListItem(string text, ListStyle style = Bullet, int level = 0)`
- `AddTable(int rowCount, int columnCount, TableFormat? format = null)`
- `AddTable(IEnumerable<IEnumerable<string>> rows, TableFormat? format = null)`
- `GetAllParagraphs()` — document order, descending into tables at any depth
- `FindParagraphs(string text, StringComparison = Ordinal)`
- `ReplaceText(string oldValue, string newValue, StringComparison = Ordinal)`
- `Remove(IParagraph)`, `Remove(ITable)` — `false` if it is not a child of this container
- `GetAllTexts()`

### `IParagraph`

- `Runs`, `Format`, `ApplyFormat(ParagraphFormat)`
- `AddText(string)`, `AddText(string, TextFormat?)`, `AddRun(string, TextFormat? = null)`
- `AddBreak(BreakType)`
- `AddHyperlink(string text, string url, TextFormat? format = null)`
- `AddImage(Stream content, ImageSize? size = null, string? description = null)`
- `AddImage(Stream content, ImageType imageType, ImageSize? size = null, string? description = null)`
- `AddImage(string filePath, ImageSize? size = null, string? description = null)`
- `SetText(string text, TextFormat? format = null)` — replaces the content, keeps the formatting
- `ReplaceText(string oldValue, string newValue, StringComparison = Ordinal)`
- `GetTextElements()`, `GetTexts()`

### `IRun`

- `Text` — get and set
- `Format`, `ApplyFormat(TextFormat)`

### `ITable`, `ITableRow`, `ITableCell`

- `ITable`: `Rows`, `ColumnCount`, `Format`, `ApplyFormat(TableFormat)`, `AddRow()`,
  `AddRow(params string[])`, `Remove(ITableRow)`, `GetCell(int rowIndex, int columnIndex)`,
  `GetAllTexts()`
- `ITableRow`: `Cells`, `AddCell(string? text = null, TableCellFormat? format = null)`,
  `RepeatAsHeader(bool = true)`, `IsHeader`, `GetAllTexts()`
- `ITableCell`: everything from `IBlockContainer`, plus `Format`, `ApplyFormat(TableCellFormat)`, and
  `SetText(string, TextFormat? = null)`

### Formatting records

All in `OfficeDocuments.Word.Formatting`. Each has optional properties, `IsEmpty`, and
`Merge(overrides)`.

| Record | Properties |
| --- | --- |
| `TextFormat` | `StyleId`, `Bold`, `Italic`, `Underline`, `Strikethrough`, `AllCaps`, `SmallCaps`, `Highlight`, `VerticalPosition`, `FontName`, `FontSize`, `Color` |
| `ParagraphFormat` | `StyleId`, `Alignment`, `SpacingBefore`, `SpacingAfter`, `LineSpacing`, `IndentLeft`, `IndentRight`, `IndentFirstLine`, `PageBreakBefore`, `KeepWithNext`, `KeepLines`, `ListStyle`, `ListLevel` |
| `TableFormat` | `WidthPercent`, `Alignment`, `Borders`, `BorderColor`, `BorderWidth`, `CellPadding`, `StyleId` |
| `TableCellFormat` | `WidthPercent`, `BackgroundColor`, `VerticalAlignment`, `ColumnSpan` |
| `PageSetup` | `PaperSize`, `PageWidth`, `PageHeight`, `Orientation`, `MarginTop`, `MarginBottom`, `MarginLeft`, `MarginRight`, `HeaderDistance`, `FooterDistance`, plus `WithUniformMargins(double)` |
| `DocumentMetadata` | `Title`, `Subject`, `Author`, `Keywords`, `Description`, `Category`, `LastModifiedBy`, `Created`, `Modified` |
| `ImageSize` | `ImageSize.Intrinsic`, `ImageSize.Exact(w, h)`, `ImageSize.FromWidth(w)`, `ImageSize.FromHeight(h)` |

### `WordStyleIds`

Constants for the built-in styles: `Normal`, `Title`, `Subtitle`, `Heading1` to `Heading6`,
`Hyperlink`, and `Heading(int level)`.

## How formatting works

Every property of every format record is optional, and this matters more than it looks.

`null` means "leave this alone", not "turn this off". In WordprocessingML `<w:b/>` switches bold on,
`<w:b w:val="0"/>` switches it off, and writing neither lets the paragraph's style decide. So
`Bold = false` is an active override of a bold style, while leaving `Bold` unset inherits it.

`ApplyFormat` follows the same rule: it writes the properties the format sets and leaves everything
else — including properties this library does not model — as it was. It is not a replacement.

Because the formats are records, one base format can carry many variations:

```csharp
var bodyText = new TextFormat { FontName = "Calibri", FontSize = 11 };

paragraph.AddText("normal", bodyText);
paragraph.AddText("emphasis", bodyText with { Bold = true });
paragraph.AddText("aside", bodyText.Merge(new TextFormat { Italic = true, Color = "5A5A5A" }));
```

`Merge` layers a second format on top of a first, with the argument winning. `with` is the shorter
form when the variation is known at the call site.

## Colours

Colours are 6 hex digits (`RRGGBB`), optionally `#`-prefixed, or the literal `auto`. They are
validated and normalized when assigned, so a bad colour throws where it was written rather than
producing a document Word has to repair.

WordprocessingML colours have no alpha channel. An 8-digit ARGB value — the form the Excel module
accepts — is rejected with a message saying so, rather than silently losing the alpha.

Highlighting is separate from colour and comes from a fixed palette, which is why `Highlight` takes
the `HighlightColor` enum rather than a hex string.

## Definitions, not just references

Three features are only a pointer in the paragraph, with the appearance defined elsewhere in the
package. Writing the pointer alone produces a document that looks like nothing happened, so the
library adds the definition on first use:

- **Named styles.** `w:pStyle` needs a style definition. Using `WordStyleIds.Heading1` or `AddHeading`
  defines it, along with the `Normal` style it is based on and the outline level that puts headings in
  Word's navigation pane and in a generated table of contents.
- **Lists.** `w:numPr` needs a numbering definition. `AddListItem` creates one per list style, shared
  by every item of that style, with all nine levels declared.
- **First-page and even-page headers.** These need `w:titlePg` on the section and
  `w:evenAndOddHeaders` in the document settings respectively. Without them the header is valid but
  never appears; `AddHeader` sets them.

An identifier the library does not know is written through untouched. That is deliberate: a document
created from a template legitimately references styles defined in that template, and inventing a
definition would overwrite the template's look. A document that uses no named styles gets no styles
part at all.

## Text fidelity

`GetTexts()` and `GetAllTexts()` return what the document contains, without trimming:

- leading and trailing spaces survive, because the library writes `xml:space="preserve"` when a run
  needs it — without that attribute an XML processor is free to collapse the whitespace and
  `"Total: " + "42"` closes up to `"Total:42"`
- a newline in the text passed to `AddText` becomes a real line break (`w:br`), because a newline
  inside `w:t` is only whitespace to Word
- breaks read back as `\n` and tabs as `\t`
- `GetAllTexts()` reads paragraphs and tables in document order, joins blocks with `\n`, and keeps
  empty paragraphs, so blank lines the author put in the document are still visible when reading back
- table rows are joined with `\n` and cells within a row with `\t`
- `IParagraph.Runs` includes runs nested inside a hyperlink, so a link's text is never skipped

## Searching and replacing text

A run boundary in a `.docx` carries no meaning. Word starts a new run wherever spell-check state,
revision identifiers, or editing history change, so a placeholder that a person typed as `{{customer}}`
and sees as one word is very often stored as three runs:

```xml
<w:r><w:t>Dear {{</w:t></w:r>
<w:r><w:proofErr w:type="spellStart"/><w:t>customer</w:t></w:r>
<w:r><w:t>}}, thank you.</w:t></w:r>
```

Any search that looks inside individual runs finds nothing here, and a template fill silently does
nothing at all. This is the single most common way a `.docx` find-and-replace goes wrong.

`ReplaceText` therefore works on the paragraph's text as a whole. What that means in practice:

- **A match is found regardless of how the text is split.** The runs are only where the characters
  happen to be stored.
- **A match never crosses a paragraph boundary.** Two paragraphs are two texts, and replacing a phrase
  spanning them would have to merge or delete a paragraph.
- **The replacement takes the formatting of the run where the match starts.** Replacing `{{amount}}`
  that began in a bold red run produces bold red text.
- **A run the replacement empties is removed**, so filling a template repeatedly does not accumulate
  content-free runs.
- **Line breaks and tabs are part of the text**, reading as `\n` and `\t`. A match may run through one,
  and a `\n` in the replacement becomes a real line break.
- **The return value is the number of occurrences replaced**, which is what a template fill should
  assert on: zero means the placeholder was not there.

`ReplaceText` exists at three levels, each reaching further:

| Called on | Covers |
| --- | --- |
| `IParagraph` | that paragraph |
| `IBlockContainer` | its paragraphs and every table cell inside it, at any nesting depth |
| `IWordprocessing` | the body, plus every header and footer |

The document-level one is the template entry point. A date or a customer name in a running header is
exactly what a body-only pass leaves behind.

Comparison defaults to `StringComparison.Ordinal`. The ordinal and `OrdinalIgnoreCase` options are the
useful ones; the culture-sensitive options work, and are handled correctly for the case where the
matched span is a different length than the search text — under a culture-sensitive comparison the
single character `ﬁ` matches the two characters `fi`.

## Usage examples

### A formatted report

```csharp
using OfficeDocuments.Word;
using OfficeDocuments.Word.Enums;
using OfficeDocuments.Word.Formatting;

using var document = new Wordprocessing("report.docx", createNew: true);

var bodyText = new TextFormat { FontName = "Calibri", FontSize = 11 };
var justified = new ParagraphFormat { Alignment = ParagraphAlignment.Justify, SpacingAfter = 6 };

var body = document.GetBody();
body.AddParagraph("Quarterly report", new ParagraphFormat { StyleId = WordStyleIds.Title });
body.AddHeading("Summary", 1);
body.AddParagraph("Revenue grew steadily.", justified, bodyText);

body.AddParagraph(justified)
    .AddText("Total: ", bodyText)
    .AddText("1 240 000 CZK", bodyText with { Bold = true })
    .AddText(" (unaudited)", bodyText with { Italic = true });
```

### Page setup and metadata

```csharp
document
    .ApplyPageSetup(new PageSetup { PaperSize = PaperSize.A4, Orientation = PageOrientation.Portrait }
        .WithUniformMargins(56))
    .SetMetadata(new DocumentMetadata
    {
        Title = "Quarterly report",
        Author = "Finance",
        Keywords = "report;finance;Q3",
    });
```

### Headers and footers

```csharp
// A logo on the first page only, and a page footer everywhere.
using var logo = File.OpenRead("logo.png");
document.AddHeader(HeaderFooterKind.First)
    .AddParagraph(new ParagraphFormat { Alignment = ParagraphAlignment.Right })
    .AddImage(logo, ImageSize.FromWidth(80));

document.AddFooter()
    .AddParagraph("Confidential",
        new ParagraphFormat { Alignment = ParagraphAlignment.Center },
        new TextFormat { FontSize = 8, Color = "808080" });
```

Calling `AddHeader` again with the same kind returns the header that already exists, so a document can
be added to without duplicating it.

### Lists

```csharp
body.AddListItem("First point");
body.AddListItem("Second point");
body.AddListItem("A sub-point", ListStyle.Bullet, level: 1);

body.AddListItem("Step one", ListStyle.Number);
body.AddListItem("Step two", ListStyle.Number);
```

### Tables

```csharp
// From data, sized to the longest row.
var table = body.AddTable(
    [
        ["Item", "Quantity", "Price"],
        ["Widget", "2", "19.90"],
        ["Gadget", "1", "45.00"],
    ],
    new TableFormat
    {
        WidthPercent = 100,
        Borders = TableBorders.All,
        BorderColor = "#4472C4",
        CellPadding = 3,
    });

// A header row that repeats on every page the table spans.
table.Rows[0].RepeatAsHeader();
table.Rows[0].Cells[0].ApplyFormat(new TableCellFormat { BackgroundColor = "D9E2F3" });

// A cell is a block container, so it takes anything the body does.
table.GetCell(1, 0).SetText("Widget", new TextFormat { Bold = true });
table.GetCell(2, 2).AddListItem("plus delivery");
```

```csharp
// Or an empty grid to fill in afterwards.
var grid = body.AddTable(rowCount: 3, columnCount: 4);
grid.GetCell(0, 0).SetText("A1");
```

A cell spanning two columns replaces the cells it covers, so that row holds one fewer cell than the
table has columns:

```csharp
var row = table.AddRow();
row.Cells[0].ApplyFormat(new TableCellFormat { ColumnSpan = 2 });
```

### Hyperlinks

```csharp
body.AddParagraph()
    .AddText("See ")
    .AddHyperlink("our documentation", "https://example.com/docs")
    .AddText(" or write to ")
    .AddHyperlink("support", "mailto:support@example.com")
    .AddText(" for help.");
```

The link text gets the built-in `Hyperlink` character style, so it is blue and underlined. A format
passed to `AddHyperlink` layers on top of that rather than replacing it.

### Images

```csharp
// From a file, at the image's own size.
body.AddParagraph().AddImage("chart.png");

// From a stream, scaled to a width with the height following the aspect ratio.
using var photo = File.OpenRead("photo.jpg");
body.AddParagraph().AddImage(photo, ImageSize.FromWidth(300), description: "Site photograph");

// At an exact size, which may change the proportions.
body.AddParagraph().AddImage("logo.png", ImageSize.Exact(120, 40));
```

Without a size the library reads the image's own dimensions and resolution from its header, for PNG,
JPEG, GIF, and BMP. For any other format, or for a stream that cannot seek, pass the `ImageType` and
an `ImageSize.Exact(...)`.

### Read a document without modifying it

```csharp
using OfficeDocuments.Word;

using var document = new Wordprocessing("report.docx", createNew: false, isEditable: false);

var body = document.GetBody();
var allText = body.GetAllTexts();

foreach (var paragraph in body.Paragraphs)
{
    var style = paragraph.Format.StyleId;
    var isListItem = paragraph.Format.ListStyle is not null;

    foreach (var run in paragraph.Runs)
    {
        var isBold = run.Format.Bold == true;
    }
}

foreach (var table in body.Tables)
{
    foreach (var row in table.Rows)
    {
        var cells = row.Cells.Select(cell => cell.GetAllTexts()).ToArray();
    }
}
```

Opening with `isEditable: false` also stops `Close()` and `Dispose()` from writing the file back, so
reading a document leaves its bytes untouched.

### Navigate an existing document

```csharp
using var document = new Wordprocessing("report.docx", createNew: false, isEditable: false);

var body = document.GetBody();

// Every paragraph in document order, table content included, at any nesting depth.
foreach (var paragraph in body.GetAllParagraphs())
{
    Console.WriteLine(paragraph.GetTexts());
}

// Or only the ones that matter.
var invoiceLines = body.FindParagraphs("Invoice", StringComparison.OrdinalIgnoreCase);

// Headers and footers the document already contained are reported too.
foreach (var container in document.HeadersAndFooters)
{
    Console.WriteLine($"{(container.IsHeader ? "header" : "footer")} {container.Kind}: {container.GetAllTexts()}");
}
```

`Paragraphs` and `GetAllParagraphs()` answer different questions: the first is this container's own
paragraphs, the second is everything below it. Use `Paragraphs` to edit the body's own structure and
`GetAllParagraphs()` to sweep the text.

### Fill a template

```csharp
using var document = new Wordprocessing("template.docx", createNew: false);

// One call covers the body, every table cell, and every header and footer.
var filled = document.ReplaceText("{{customer}}", "Acme s.r.o.");
if (filled == 0)
{
    throw new InvalidOperationException("The template does not contain a {{customer}} placeholder.");
}

document.ReplaceText("{{date}}", DateTime.Today.ToString("yyyy-MM-dd"));

// A section this run does not need can go away entirely.
var body = document.GetBody();
foreach (var heading in body.FindParagraphs("{{optional}}").ToList())
{
    body.Remove(heading);
}
```

Materialize the result of `FindParagraphs` with `ToList()` before removing anything: the collections are
read from the document on each access, so enumerating one while editing it walks a moving target.

### Edit an existing document

```csharp
using var document = new Wordprocessing("report.docx", createNew: false);

var body = document.GetBody();
body.AddParagraph("Appended after opening.");

// Existing runs can be rewritten in place, keeping their formatting.
body.Paragraphs[0].Runs[0].Text = "Replaced title";

// Or a whole paragraph, keeping its style and alignment.
body.Paragraphs[1].SetText("Rewritten, still a heading");

// Existing headers, styles, and list numbering are reused rather than duplicated.
document.AddHeader().AddParagraph("Added to the existing header");

// Structure can shrink as well as grow.
body.Tables[0].Remove(body.Tables[0].Rows[^1]);
```

## Consumer notes

- The fluent flow is the intended authoring model; `AddRun` exists for when the run itself is needed.
- Collections are projected from the document on every access, so content added or removed through the
  API is visible immediately and always in document order. Hold the list in a local when indexing it
  inside a loop, and materialize it before editing what you are iterating over.
- `GetBody()` returns the same instance every time it is called, as does `AddHeader`/`AddFooter` for a
  given kind — and so does `HeadersAndFooters` for a header this instance already handed out.
- `Close()` is idempotent, so `using` combined with an explicit `Close()` is safe.
- The document is persisted when you call `Close()` or dispose the document instance.
  `Close(saveDocument: false)` discards the changes instead, and opening with `isEditable: false`
  discards them whatever you pass.

## Related documents

- [README.md](README.md)
- [excel-library.md](excel-library.md)
- [terminology.md](terminology.md)
- [tasks/README.md](tasks/README.md)
- [tasks/core/word/](tasks/core/word/) — the Word backlog and its progress logs
- [architecture/word-002-readiness-audit.md](architecture/word-002-readiness-audit.md)
