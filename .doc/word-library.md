# OfficeDocuments.Word

Date: 2026-05-31

This guide describes the current public Word API as implemented in `src/OfficeDocuments.Word/Interfaces/*` and exercised by the current test suite.

## Scope

`OfficeDocuments.Word` is intentionally smaller than the Excel module.

The library is designed for:

- creating and opening `.docx` documents from files or streams
- building simple document content through a fluent paragraph API
- reading paragraph text back from existing documents

The default object model is:

`IWordprocessing -> IBody -> IParagraph -> IText`

## What the library supports today

- Create or open `.docx` documents from files or streams
- Get the document body
- Add paragraphs
- Add text runs in a fluent style
- Add page, column, and text-wrapping breaks
- Read paragraph text from an existing document

## Main API surface

### `IWordprocessing`

- `new Wordprocessing(string filePath, bool createNew)`
- `new Wordprocessing(Stream stream, bool createNew)`
- `GetBody()`
- `Close()`

### `IBody`

- `AddParagraph()`
- `Paragraphs`
- `GetAllTexts()`

### `IParagraph`

- `AddText(...)`
- `AddBreak(...)`
- `GetTextElements()`
- `GetTexts()`

### `IText`

- `TextValue`

## Usage examples

### Create a Word document

```csharp
using OfficeDocuments.Word;
using OfficeDocuments.Word.Enums;

using var word = new Wordprocessing("sample.docx", createNew: true);

var body = word.GetBody();
body.AddParagraph()
    .AddText("First page")
    .AddBreak(BreakType.Page);

body.AddParagraph()
    .AddText("Second page");

word.Close();
```

### Open a Word document and read text

```csharp
using OfficeDocuments.Word;

using var word = new Wordprocessing("sample.docx", createNew: false);

var body = word.GetBody();
var allText = body.GetAllTexts();

foreach (var paragraph in body.Paragraphs)
{
    var paragraphText = paragraph.GetTexts();
}

word.Close();
```

## Consumer notes

- The Word module currently exposes a deliberately small API surface.
- The fluent paragraph flow is the intended authoring model.
- Advanced Word features such as tables, styling, headers, or images are backlog items, not current public features.
- The document is persisted when you call `Close()` or dispose the document instance.

## Related documents

- [README.md](README.md)
- [excel-library.md](excel-library.md)
- [terminology.md](terminology.md)
- [tasks/README.md](tasks/README.md)
