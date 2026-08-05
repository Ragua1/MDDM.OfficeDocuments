# OfficeDocuments Feature Gap Backlog

Date: 2026-05-31

This backlog reflects the current repository state after reviewing the public API, README, and test suite.

Legend:

- `P0`: highest-value gap for product credibility
- `P1`: strong usability improvement after the most urgent work
- `P2`: broader or heavier feature slice

## Already delivered since the earlier benchmark baseline

### Excel

- [x] Range abstraction and rectangular range operations
- [x] Bulk row insertion from nested collections and object collections
- [x] Worksheet rename, remove, move, copy, and hide operations
- [x] Auto-filter and range sorting
- [x] Data validation
- [x] Conditional formatting
- [x] Freeze panes and auto-fit columns
- [x] Hyperlinks and comments
- [x] Named ranges
- [x] Worksheet and workbook protection
- [x] Worksheet image embedding
- [x] Structured table lookup and lifecycle operations

### Word

- [x] File and stream create/open workflows, including read-only open
- [x] Fluent paragraph and text authoring
- [x] Paragraph and run text reading, with whitespace and line structure preserved
- [x] Run formatting: bold, italic, underline, strikethrough, caps, highlight, super/subscript, font, size, colour, character styles
- [x] Paragraph formatting: alignment, spacing, line spacing, indentation, page-break and keep-together control
- [x] Built-in paragraph and character styles, headings, and lists, defined in the document on first use
- [x] Tables: header rows, borders, shading, cell padding, column spanning, nested content
- [x] Hyperlinks with relationship handling
- [x] Inline images with intrinsic sizing read from PNG, JPEG, GIF, and BMP headers
- [x] Headers and footers, including first-page and even-page
- [x] Page size, orientation, and margins
- [x] Document core properties
- [x] Schema-validation gate over every generated document
- [x] Paragraph navigation and search, including inside tables
- [x] Text replacement across the run boundaries Word inserts, per paragraph, container, or document
- [x] Structural removal of paragraphs, tables, and table rows

## P0

- [ ] Finish the cleanup of raw OpenXml-oriented compatibility members on the Excel public surface.
  Why: the library is positioned as a simpler API over OpenXml, but several compatibility members still expose implementation details.

- [x] Add Word run formatting: bold, italic, underline, font family, font size, and color.
  Delivered 2026-07-27 by `WORD-001` as the `TextFormat` record.

- [x] Add Word paragraph formatting: alignment, spacing, indentation, and heading-like styles.
  Delivered 2026-07-27 by `WORD-001` as the `ParagraphFormat` record, with built-in style definitions.

- [x] Add Word tables.
  Delivered 2026-07-27 by `WORD-002A`, together with the shared block-container extraction.

- [x] Add Word hyperlinks and images.
  Delivered 2026-07-27 by `WORD-002B` and `WORD-002C`.

- [x] Add Word headers, footers, sections, and document metadata.
  Delivered 2026-07-27 by `WORD-003`. Multiple sections with differing page setups remain out of scope.

- [x] Strengthen Word round-trip and read/edit test coverage.
  Delivered 2026-07-27 by `WORD-004`. The read and update paths turned up three defects the
  authoring-only tests could not reach: stale projected collections after a removal, headers that an
  opened document did not report, and a discard-on-close that always saved.

## P1

- [ ] Add Excel formula evaluation or an explicit documented calculation strategy.
  Why: formulas can be written today, but calculation is still delegated to Excel or another consumer.

- [ ] Add higher-level Excel import and export helpers for common tabular sources.
  Why: the new range and bulk APIs improved ergonomics, but typed import/export recipes are still thin.

- [x] Expand Word read/edit helpers for basic navigation and find/replace scenarios.
  Delivered 2026-07-27 by `WORD-004`: paragraph walking and search, text replacement that works across
  the run boundaries Word inserts, and removal of paragraphs, tables, and table rows.

- [x] Replace external-resource-dependent Word tests with repository-local fixtures where possible.
  Delivered 2026-07-27 by `WORD-004`. `TestKit/ForeignDocuments.cs` builds Word-shaped markup through
  the SDK, so the input is reviewable in a diff instead of being a binary nobody can inspect.

## P2

- [ ] Add Excel charts.
  Why: useful, but materially heavier than the current data-centric core.

- [ ] Add Excel pivot tables.
  Why: valuable for reporting, but expensive to design and validate well.

- [ ] Add richer Excel template-oriented workflows.
  Why: template processing is attractive, but it should follow the current core ergonomics work.

- [ ] Add optional advanced Word capabilities such as bookmarks, comments, or table-of-contents groundwork.
  Why: these are follow-up slices after the core authoring surface is stronger.

## Design principles to keep

- [x] Keep the small workbook -> worksheet -> row/cell/range mental model
- [x] Keep stream-first create/open support
- [x] Keep typed value reads in Excel
- [x] Keep reusable style creation and style merging
- [x] Keep the small fluent Word authoring model as the foundation for future growth

## Suggested delivery order

1. ~~Strengthen Word formatting and tables.~~ Delivered 2026-07-27.
2. ~~Add Word hyperlinks, images, headers, footers, and sections.~~ Delivered 2026-07-27.
3. ~~Harden Word read and edit workflows (`WORD-004`).~~ Delivered 2026-07-27. The Word core backlog is
   now empty.
4. Finish the Excel public-surface cleanup — the only remaining `P0`.
5. Revisit heavier Excel output features only after the current core remains stable.
