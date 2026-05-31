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

- [x] File and stream create/open workflows
- [x] Fluent paragraph and text authoring
- [x] Paragraph text reading

## P0

- [ ] Finish the cleanup of raw OpenXml-oriented compatibility members on the Excel public surface.
  Why: the library is positioned as a simpler API over OpenXml, but several compatibility members still expose implementation details.

- [ ] Add Word run formatting: bold, italic, underline, font family, font size, and color.
  Why: without text styling, the Word module still cannot cover most real business documents.

- [ ] Add Word paragraph formatting: alignment, spacing, indentation, and heading-like styles.
  Why: paragraph layout is a baseline requirement for usable `.docx` generation.

- [ ] Add Word tables.
  Why: tables are a core requirement for invoices, reports, and formal business documents.

- [ ] Add Word hyperlinks and images.
  Why: they remain part of the practical minimum for generated Word documents.

- [ ] Add Word headers, footers, sections, and document metadata.
  Why: branding and basic document structure are still missing from the current Word API.

- [ ] Strengthen Word round-trip and read/edit test coverage.
  Why: the Word module still has a smaller and more fragile test surface than the Excel module.

## P1

- [ ] Add Excel formula evaluation or an explicit documented calculation strategy.
  Why: formulas can be written today, but calculation is still delegated to Excel or another consumer.

- [ ] Add higher-level Excel import and export helpers for common tabular sources.
  Why: the new range and bulk APIs improved ergonomics, but typed import/export recipes are still thin.

- [ ] Expand Word read/edit helpers for basic navigation and find/replace scenarios.
  Why: document updates and template workflows need more than pure authoring.

- [ ] Replace external-resource-dependent Word tests with repository-local fixtures where possible.
  Why: deterministic local test assets improve reliability and maintainability.

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

1. Finish the Excel public-surface cleanup.
2. Strengthen Word formatting and tables.
3. Add Word hyperlinks, images, headers, footers, and sections.
4. Harden Word tests and read/edit workflows.
5. Revisit heavier Excel output features only after the current core remains stable.
