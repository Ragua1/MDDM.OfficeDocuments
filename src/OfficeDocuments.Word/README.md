# OfficeDocuments.Word

A .NET library for creating and reading Word (`.docx`) documents via the Open XML SDK, with a fluent
interface for paragraphs, runs, tables, images, and page structure.

## Install

```sh
dotnet add package OfficeDocuments.Word
```

## Quick start

```csharp
using OfficeDocuments.Word;
using OfficeDocuments.Word.Enums;
using OfficeDocuments.Word.Formatting;

// Create
using (var document = new Wordprocessing("report.docx", createNew: true))
{
    document.SetMetadata(new DocumentMetadata { Title = "Quarterly report", Author = "Finance" });
    document.AddHeader().AddParagraph("Contoso — internal");

    var body = document.GetBody();
    var bodyText = new TextFormat { FontName = "Calibri", FontSize = 11 };

    body.AddParagraph("Quarterly report", new ParagraphFormat { StyleId = WordStyleIds.Title });
    body.AddHeading("Summary", 1);

    body.AddParagraph(new ParagraphFormat { Alignment = ParagraphAlignment.Justify })
        .AddText("Total: ", bodyText)
        .AddText("1 240 000 CZK", bodyText with { Bold = true });

    body.AddTable(
    [
        ["Product", "Amount"],
        ["Widget", "1 250.50"],
    ]);
}

// Read
using (var document = new Wordprocessing("report.docx", createNew: false, isEditable: false))
{
    var text = document.GetBody().GetAllTexts();
}

// Fill a template — one call covers the body, every table cell, and every header and footer
using (var document = new Wordprocessing("template.docx", createNew: false))
{
    document.ReplaceText("{{customer}}", "Acme s.r.o.");
}
```

Font sizes, spacing, indentation, margins, and image dimensions are in **points**; the library
converts them to the half-points, twips, and EMUs WordprocessingML actually stores. The document is
written when you call `Close()` or dispose it.

## What it covers

Paragraphs, runs, breaks, headings, and built-in styles; run and paragraph formatting; bullet and
numbered lists; tables with header rows, borders, shading, and column spans; hyperlinks; inline
images sized from the file itself or explicitly; headers and footers; page size, orientation, and
margins; document metadata; text reading from existing documents; paragraph search and navigation;
and text replacement that works across the run boundaries Word inserts mid-word — the reason a naive
`.docx` find-and-replace usually finds nothing.

## Targets

`net8.0` · `net9.0` · `net10.0`

## Links

- [Full documentation](https://github.com/Ragua1/MDDM.OfficeDocuments/blob/master/.doc/word-library.md)
- [Repository](https://github.com/Ragua1/MDDM.OfficeDocuments)
- [Changelog / releases](https://github.com/Ragua1/MDDM.OfficeDocuments/releases)
