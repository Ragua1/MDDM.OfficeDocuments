# OfficeDocuments Library Benchmark Report

Baseline: 2026-05-31 · Last reconciled against the code: 2026-07-27

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
- `test/OfficeDocuments.Excel.*Tests/*`
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

`OfficeDocuments.Word` closed most of its baseline gap during the `WORD-001` to `WORD-003` sequence. It now covers run and paragraph formatting, built-in styles, lists, tables, hyperlinks, inline images, headers and footers, page setup, and document metadata, on top of the original paragraph and text-reading workflows. What remains missing is the second tier of document features — sections, footnotes, bookmarks, comments, tracked changes, and a generated table of contents.

Compared with established libraries, the biggest differentiator is still API scope rather than raw file support. `OfficeDocuments` is a focused, approachable wrapper; Excel is the broader surface and Word is now credible for the common business-document case rather than only for plain text.

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
- Read-path validation is still thinner than write-path validation, although the verification tier now covers round-trips and foreign input workbooks.

### Word: current capabilities

Verified in API and tests:

- Create or open `.docx` documents from files and streams, including read-only open.
- Author block content through one shared `IBlockContainer` contract, so the body, headers, footers, and table cells behave identically.
- Add paragraphs, headings, runs, and page, column, and text-wrapping breaks.
- Apply run formatting: bold, italic, underline, strikethrough, caps, highlight, super/subscript, font, size, colour, and character styles.
- Apply paragraph formatting: alignment, spacing, line spacing, indentation, page-break-before, and keep-together control.
- Use built-in styles and headings, defined in the document on first use.
- Create bullet and numbered lists with nested levels.
- Create tables by size or from data, with repeating header rows, borders, shading, cell padding, column spans, and nested content.
- Add hyperlinks and inline images, with intrinsic sizing read from PNG, JPEG, GIF, and BMP headers.
- Add headers and footers, including first-page and even-page variants.
- Set paper size, orientation, margins, and document metadata.
- Read paragraph, run, and table text back with whitespace and line structure preserved.

### Word: usability strengths

- The fluent `GetBody().AddParagraph().AddText()` flow is still easy to understand at the entry point.
- `IBlockContainer` means a feature added once works in every place block content is allowed, instead of the body-only special case most small libraries end up with.
- Immutable formatting records make one base format reusable across many runs and paragraphs.
- Every generated document passes a schema-validation gate in the test suite.

### Word: limitations and design debt

- No public support yet for: multiple sections with different page setups, footnotes and endnotes, bookmarks and internal links, comments, tracked changes, find and replace, or a generated table of contents.
- `IParagraph.Runs` lists direct children only, so a run nested in a hyperlink is reachable through `GetTexts()` but not through `Runs`.
- Read and edit workflows are still thinner than authoring workflows.
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
| python-docx | Python | High-level document/paragraph/run/table/section model | Tables, styles, sections, headers/footers, comments, shapes, hyperlinks, rich text | `OfficeDocuments.Word` now overlaps on most of this; sections and comments are the remaining gaps |
| DocX | .NET | High-level document authoring API | Paragraphs, lists, tables, images, bookmarks, hyperlinks, TOC, sections, protection, find/replace, properties | `OfficeDocuments.Word` matches the common authoring set; bookmarks, TOC, find/replace, and protection are still missing |
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

- Read, navigate, and edit helpers — find and replace, and locating content without walking the tree by hand
- Multiple sections with independent page setups
- Bookmarks and internal links, which a generated table of contents also depends on

The authoring baseline is no longer the gap; the remaining work is on reading, editing, and document structure beyond one section.

### Structural gaps

- The public Excel API is cleaner than before but still carries some compatibility members that do not fit the preferred minimal-core surface.
- Excel read/edit workflows remain less mature than its write workflows. Word closed that gap on 2026-07-27 with `WORD-004`: navigation, search, cross-run text replacement, structural removal, and the template scenario end to end.
- The Word test suite is still one tier, whereas Excel is split into unit, integration, and verification.

## Recommended roadmap

### Priority 0

- Finish the Excel public-surface cleanup. This is now the only remaining `P0`.

### Priority 1

- Improve higher-level Excel import/export ergonomics.
- Split the Word test suite into tiers once the surface justifies it.

### Priority 2

- Revisit heavier Excel output features such as charts, pivot tables, and templates.
- Consider optional advanced Word capabilities only after the core authoring model is broader and stable.

## Recommended positioning

The strongest current positioning is not "full Office automation suite".

The strongest positioning is:

- a simple, modern, server-friendly Office document library
- optimized for common business document and spreadsheet generation
- intentionally smaller and more predictable than large general-purpose libraries
- broadest today in Excel, with Word now covering the common business-document case

That positioning is credible today for both modules on the authoring side. The next credibility gain comes from read and edit workflows, and from continued cleanup of the remaining Excel compatibility surface.
