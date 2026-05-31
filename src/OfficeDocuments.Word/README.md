# OfficeDocuments.Word

A .NET library for creating and reading Word (`.docx`) documents via the OpenXml SDK,
with a fluent interface for building paragraphs, text runs, and document content.

## Install

```
dotnet add package OfficeDocuments.Word
```

## Quick start

```csharp
using OfficeDocuments.Word;

// Create
using var doc = Wordprocessing.Create("document.docx");
doc.GetBody()
   .AddParagraph()
   .AddText("Hello, World!");
doc.Save();

// Read
using var doc = Wordprocessing.Open("document.docx");
var text = doc.GetBody().GetParagraphs().First().GetText();
```

## Targets

`net8.0` · `net9.0` · `net10.0`

## Links

- [Full documentation](.doc/word-library.md)
- [Repository](https://github.com/Ragua1/MDDM.OfficeDocuments)
- [Changelog / releases](https://github.com/Ragua1/MDDM.OfficeDocuments/releases)
