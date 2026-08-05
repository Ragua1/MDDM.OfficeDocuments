# OfficeDocuments Terminology

Date: 2026-05-31

This glossary defines the preferred project terminology and abbreviations used across the repository documentation.

## Product and architecture terms

| Term | Meaning | Preferred usage |
| --- | --- | --- |
| `OfficeDocuments` | Repository and product name covering both Excel and Word modules | Use for the overall library family |
| `OfficeDocuments.Excel` | Excel-focused module for `.xlsx` work | Use when referring to the spreadsheet library specifically |
| `OfficeDocuments.Word` | Word-focused module for `.docx` work | Use when referring to the Word library specifically |
| minimal core | Small, default consumer-facing feature set that should remain lightweight and easy to learn | Prefer over broader phrases such as "full office platform" |
| advanced layer | Optional future layer for broader or heavier feature slices | Use for roadmap and architecture planning only; it is not a current shipped package |
| interop surface | Public API that exposes raw OpenXml concepts or types | Use when discussing compatibility-oriented APIs that are not the preferred consumer surface |

## Excel terms

| Term | Meaning | Preferred usage |
| --- | --- | --- |
| spreadsheet | Workbook-level object represented by `ISpreadsheet` or `Spreadsheet` | Use for the workbook entry point |
| workbook | Excel file as a whole | Use for the document concept; "spreadsheet" is the API entry point |
| worksheet | A single sheet inside the workbook | Use instead of the looser word "sheet" in formal documentation |
| row | 1-based row inside a worksheet | State explicitly that indexes are 1-based when relevant |
| cell | Single worksheet value container identified by row and column | Use for `ICell` instances and Excel cells generally |
| range | Rectangular worksheet area represented by `IRange` | Prefer over phrases such as "cell block" |
| A1 reference | Excel address notation such as `B2` or `A1:C4` | Use when documenting coordinate-based APIs |
| structured table | Excel table object with a name, reference, and style options | Use instead of just "table" when ambiguity is possible |
| named range | Workbook or worksheet-scoped symbolic name over a range | Use when documenting formulas, interoperability, or workbook metadata |
| style | Reusable Excel formatting bundle created through `CreateStyle(...)` | Use for the public formatting concept |
| style merge | Combining two styles through `IStyle.CreateMergedStyle(...)` | Use for the supported composition workflow |
| worksheet image | Image embedded in a worksheet and anchored to a rectangular area | Prefer over vague phrases such as "picture support" |

## Word terms

| Term | Meaning | Preferred usage |
| --- | --- | --- |
| document | Word `.docx` file as a whole | Use for the conceptual file |
| block container | Anything that holds block-level content, exposed by `IBlockContainer` | Use when a rule applies to the body, headers, footers, and table cells alike |
| block content | Paragraphs and tables — the content a block container holds | Prefer over "top-level content", which is no longer accurate |
| body | Main document content area exposed by `IBody` | Use for the primary block container; it is one of several |
| header, footer | Page furniture exposed by `IHeaderFooter`, in default, first-page, and even-page variants | Use `header kind` for the variant, not "type" |
| paragraph | Word paragraph exposed by `IParagraph` | Use as the main authoring building block |
| run | Contiguous text with one character format, exposed by `IRun` | Use for character-level formatting discussions |
| text element | `IText` value returned when reading paragraph content | Use when documenting read workflows |
| format record | Immutable options record such as `TextFormat` or `ParagraphFormat` | Prefer over "style object"; a Word *style* is a named definition in the package |
| Word table | Table exposed by `ITable`, `ITableRow`, and `ITableCell` | Qualify as "Word table" when Excel structured tables are also in scope |
| fluent API | Chained authoring style such as `GetBody().AddParagraph().AddText(...)` | Use for the current Word authoring model |
| run splitting | Word's habit of starting a new run for reasons unrelated to the text — spell-check state, revision identifiers, editing history | Use when explaining why a search or replacement has to work on the whole paragraph |
| template fill | Replacing placeholders in an existing document through `ReplaceText` | Prefer over "mail merge", which means something more specific in Word |

## Technology and process terms

| Term | Meaning | Preferred usage |
| --- | --- | --- |
| OpenXml | Short name for the Open XML SDK or raw OpenXml document model | Capitalize consistently as `OpenXml` in prose |
| `DocumentFormat.OpenXml` | The .NET package and namespace used as the primary implementation foundation | Use the full name when precision matters |
| coding agent | Any AI coding tool working in this repository | Preferred term in new task documents; the working rules live in [../AGENTS.md](../AGENTS.md) |
| GHC | Retired abbreviation for GitHub Copilot | Do not use. It was removed from the task documents on 2026-07-27; use `coding agent`, or name the tool if the tool actually matters |
| benchmark report | Comparative analysis against other libraries | Use for `library-benchmark-report.md` |
| backlog | Planned but not yet delivered features or engineering tasks | Use for roadmap and task documentation |
| round-trip test | Test that creates, saves, reopens, and verifies a document | Prefer over generic phrases such as "integration test" when that exact behavior is meant |
| projected collection | Collection read from the document on every access rather than stored, so it cannot go stale | Use for `Paragraphs`, `Runs`, `Rows`, `Cells`, and `Tables`; see `AGENTS.md` rule 10 |
| foreign document | Input written by a producer other than this library | Use for real Excel files and for the Word markup `ForeignDocuments` builds; the point is that the library did not choose its structure |
| inherited defect | Schema-validation error that arrived with a foreign input document | Use for `AssertValid`'s `inheritedDefects` parameter; never for a defect this library produced |

## Writing rules

- Prefer `Excel` and `Word` as module names only when the surrounding context already makes `OfficeDocuments` clear.
- Prefer `worksheet`, `range`, `paragraph`, and `body` over generic synonyms.
- Use `OpenXml` for the concept and `DocumentFormat.OpenXml` for the concrete package or namespace.
- Keep consumer documentation focused on public APIs and workflow terms; reserve architecture terms such as `advanced layer` or `interop surface` for planning documents.
