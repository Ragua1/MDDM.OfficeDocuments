# OfficeDocuments Library Benchmark Report

Date: 2026-05-31

## Purpose

This report evaluates the current state of `MDDM.OfficeDocuments`, compares it with established .NET and Python libraries for Office document work, and identifies the most important remaining feature and usability gaps.

The goal is not to copy other libraries wholesale. The goal is to identify the smallest set of high-value capabilities and design patterns that would most improve this library.

## Scope reviewed

Internal evidence reviewed:

- `README.md`
- `.doc/excel-library.md`
- `.doc/word-library.md`
- `src/OfficeDocuments.Excel/Interfaces/*`
- `src/OfficeDocuments.Excel/Spreadsheet.cs`
- `src/OfficeDocuments.Excel/DataClasses/*`
- `src/OfficeDocuments.Word/Interfaces/*`
- `src/OfficeDocuments.Word/DataClasses/*`
- `src/OfficeDocuments.Word/Wordprocessing.cs`
- `test/OfficeDocuments.Excel.Tests/*`
- `test/OfficeDocuments.Word.Tests/*`

External reference libraries reviewed:

- ClosedXML
- EPPlus
- NPOI
- openpyxl
- XlsxWriter
- python-docx
- DocX

## Executive summary

`OfficeDocuments.Excel` is already practical for a broad set of report-style spreadsheet workflows. It now supports ranges, bulk insert, worksheet lifecycle operations, sorting, auto-filter, validation, conditional formatting, hyperlinks, comments, named ranges, protection, structured tables, and worksheet images in addition to the original workbook, row, cell, formula, and style workflows.

`OfficeDocuments.Word` remains intentionally small. It covers document creation, paragraph authoring, text runs, breaks, and paragraph-text reading, but it still lacks the richer authoring primitives expected from a broadly usable Word library.

Compared with established libraries, the biggest differentiator is still API scope rather than raw file support. `OfficeDocuments` is a focused, approachable wrapper with stronger Excel ergonomics than before, while Word is still at an earlier maturity stage.

## Current state of the library

### Excel: current capabilities

Verified in API, documentation, and tests:

- Create and open workbooks from file paths and streams.
- Use a clear object hierarchy: `ISpreadsheet -> IWorksheet -> IRange / IRow -> ICell`.
- Add worksheets and retrieve worksheet names.
- Rename, move, copy, hide, and remove worksheets.
- Add rows and cells by position or sequentially.
- Bulk insert nested rows and object collections.
- Write strings, booleans, numbers, decimals, dates, hyperlinks, comments, and formulas.
- Read values back via typed getters and `TryGetValue(...)` overloads.
- Create styles for font, fill, border, alignment, and number format.
- Merge styles with `CreateMergedStyle(...)`.
- Work with ranges for reading, writing, styling, merging, sorting, filtering, validation, and conditional formatting.
- Freeze panes and auto-fit columns.
- Add named ranges.
- Protect worksheets and workbook structure.
- Create, query, rename, resize, enumerate, and remove Excel tables.
- Embed worksheet images from streams or files.

### Excel: usability strengths

- The public workflow is materially richer than a plain workbook-row-cell wrapper.
- The range abstraction improves several common spreadsheet tasks.
- Stream support keeps server-side generation straightforward.
- Style composition is much cleaner than working directly with raw OpenXml elements.
- Tests now cover a broader advanced feature set, including ranges, validation, sorting, tables, protection, and images.

### Excel: limitations and design debt

- The public API still exposes some compatibility-oriented OpenXml members that are marked obsolete but not yet removed from the default surface.
- Formula support writes formulas but does not provide a calculation engine.
- Higher-level import and export helpers for common external data sources are still limited.
- Charts, pivot tables, and broader template-oriented workflows remain out of scope today.
- `ReaderTest` is still effectively empty, so dedicated Excel read-path validation is thinner than write-path validation.

### Word: current capabilities

Verified in API and tests:

- Create or open `.docx` documents from files and streams.
- Get the document body.
- Add paragraphs.
- Add text runs in a fluent style.
- Add page, column, and text-wrapping breaks.
- Read paragraph text elements from an existing document body.

### Word: usability strengths

- The fluent `GetBody().AddParagraph().AddText().AddBreak()` flow is easy to understand.
- The API stays intentionally small and approachable for the current feature set.

### Word: limitations and design debt

- The feature surface is still very small compared with mainstream Word libraries.
- There is no public support for:
  - run formatting
  - paragraph formatting
  - headings and styles
  - tables
  - images
  - hyperlinks
  - headers and footers
  - sections and page setup
  - bookmarks
  - comments
  - document properties
  - find and replace
  - table-of-contents groundwork
- Word tests are still relatively sparse.
- One reading test remains skipped because it depends on an external resource file.

## Comparison with public libraries

### Excel-focused libraries

| Library | Platform | API style | Strong points | Gap vs `OfficeDocuments.Excel` |
| --- | --- | --- | --- | --- |
| ClosedXML | .NET | Very high-level workbook/worksheet/range/cell API | Bulk insert, tables, styles, sorting, autofilter, pivot tables, protection, formulas, themes | `OfficeDocuments.Excel` is still smaller and lighter, but now overlaps with more day-to-day worksheet automation features |
| EPPlus | .NET | Strongly typed Excel object model close to the Excel/VBA mental model | Broad feature set: formula engine, import/export helpers, validation, conditional formatting, charts, comments, images, protection, performance focus | `OfficeDocuments.Excel` covers several common report workflows but remains behind on enterprise-scale breadth |
| NPOI | .NET | Lower-level but broad Office support | Reads and writes `xls`, `xlsx`, `docx`; broad Excel coverage | `OfficeDocuments.Excel` remains narrower but more approachable for modern XML-only scenarios |
| openpyxl | Python | Workbook/worksheet/cell with iteration helpers | Read/write, worksheet copy, iteration, read-only mode, image support, broad worksheet manipulation | `OfficeDocuments.Excel` now covers more worksheet operations but still lacks the broader read/edit ecosystem |
| XlsxWriter | Python | Write-focused workbook/worksheet API | Charts, tables, comments, images, conditional formatting, validation, memory optimization | `OfficeDocuments.Excel` now has meaningful overlap for writing workflows, though not the same output breadth |

### Word-focused libraries

| Library | Platform | API style | Strong points | Gap vs `OfficeDocuments.Word` |
| --- | --- | --- | --- | --- |
| python-docx | Python | High-level document/paragraph/run/table/section model | Tables, styles, sections, headers/footers, comments, shapes, hyperlinks, rich text | `OfficeDocuments.Word` currently covers only a small subset of this surface |
| DocX | .NET | High-level document authoring API | Paragraphs, lists, tables, images, bookmarks, hyperlinks, TOC, sections, protection, find/replace, properties | `OfficeDocuments.Word` is still far behind on authoring depth and document structure |
| NPOI | .NET | Lower-level multi-format Office API | Supports `docx` in addition to Excel formats | Broader file-format support, but not necessarily a cleaner high-level authoring API |

## What works well elsewhere

Patterns worth reusing as inspiration:

- Range-first modeling for spreadsheets.
- One-line bulk operations for common import/export scenarios.
- Explicit task-oriented helpers instead of raw structural manipulation.
- High-level worksheet operations treated as normal API features.
- Word document structure modeled as first-class public concepts when the surface grows.
- Recipe-style documentation for common business scenarios instead of method lists only.

## Gap analysis

### Highest-value Excel gaps

- Finish removal or isolation of raw OpenXml-oriented compatibility members.
- Add a clearer documented strategy for formula calculation.
- Add higher-level import/export helpers and templates after the current core stabilizes.

### Highest-value Word gaps

- Run formatting
- Paragraph formatting
- Tables
- Images
- Hyperlinks
- Headers, footers, and sections

Without these, the Word module is still too small for most production document scenarios.

### Structural gaps

- The public Excel API is cleaner than before but still carries some compatibility members that do not fit the preferred minimal-core surface.
- Read/edit workflows remain less mature than write workflows, especially for Word.
- The Word test surface is not yet strong enough to support rapid feature growth as safely as the Excel module.

## Recommended roadmap

### Priority 0

- Finish the Excel public-surface cleanup.
- Strengthen Word into a more complete document-authoring surface.
- Add stronger Word round-trip and read/edit tests.

### Priority 1

- Improve higher-level Excel import/export ergonomics.
- Expand Word navigation and basic edit workflows.

### Priority 2

- Revisit heavier Excel output features such as charts, pivot tables, and templates.
- Consider optional advanced Word capabilities only after the core authoring model is broader and stable.

## Recommended positioning

The strongest current positioning is not "full Office automation suite".

The strongest positioning is:

- a simple, modern, server-friendly Office document library
- optimized for common business spreadsheet generation
- intentionally smaller and more predictable than large general-purpose libraries
- strongest today in Excel, emerging in Word

That positioning is credible today for Excel and only partially credible for Word. The next major credibility gain comes from richer Word authoring and continued cleanup of the remaining Excel compatibility surface.
