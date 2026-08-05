# WORD-002A Basic Tables

- Module: `OfficeDocuments.Word`
- Priority: `P0`
- Status: `Delivered` (2026-07-27 — see [Progress log](#progress-log))

## Business goal

Tables are required for invoices, summaries, protocols, offer documents, and other business documents where structured data must be readable in Word output.

## Why this belongs in the core backlog

Basic table creation is part of normal `.docx` authoring. It delivers clear business value without requiring the broader advanced-feature positioning that heavier layout or media scenarios might need.

## Functional description

The library should support:

- creating a table from the body
- adding rows and cells
- writing text into cells
- applying basic formatting such as width, borders, and simple alignment

## Technical guidance

### Public API direction

- Prefer a small body-first creation API such as `GetBody().AddTable(...)`.
- Introduce table interfaces only if they improve clarity enough to justify the extra surface.
- Keep the first version focused on creation, fill, and simple formatting.

### Implementation steps

- Model the feature around WordprocessingML table, row, and cell elements.
- Avoid building a full table-style engine in the first iteration.
- Keep the first cell-content workflow aligned with the current paragraph and text model.

### Tests

- Add a test for a multi-row table.
- Add a test for basic table formatting persistence.
- Validate structure first; do not wait for full rendering validation.

### Documentation

- Add a table example to [../../../word-library.md](../../../word-library.md).

## Complexity

- Estimate: `M`

## Risks

- The table surface can expand too quickly if the first version tries to handle every Word table scenario.
- Cell-content rules can become inconsistent if they diverge from the existing paragraph model.

## Dependencies

- benefits from `WORD-001`, but can start with a smaller direct body-level model if needed
- depends on clean paragraph insertion rules for cell content

## Subtasks

- [x] Finalize the first table API.
- [x] Implement create and fill behavior.
- [x] Add basic formatting options.
- [x] Add focused tests.
- [x] Update detailed Word documentation.

## Acceptance criteria

- Consumers can create and populate a basic table without using OpenXml directly.
- The first API stays small, readable, and extensible.
- Tests verify the expected table structure and persisted core settings.

## Progress log

### 2026-07-27 — delivered

API: `IBlockContainer.AddTable(rowCount, columnCount, format?)` and
`AddTable(IEnumerable<IEnumerable<string>> rows, format?)`, with `ITable → ITableRow → ITableCell`.
Formatting through the `TableFormat` and `TableCellFormat` records, matching the pattern `WORD-001`
established.

**A block container was extracted first.** A table cell holds the same block content as the body, and
so does a header, so `IBlockContainer` / `DataClasses/BlockContainer.cs` now implements paragraphs,
headings, lists, and tables once for the body, headers, footers, and cells. Building tables into `Body`
directly would have meant duplicating that API for `WORD-003`. `IBody` is now an empty marker over
`IBlockContainer`.

Deliberate choices:

- `AddTable(rows)` sizes the grid to the **longest** row and pads shorter ones, because a row narrower
  than the grid renders as a broken table rather than as a short row.
- New tables default to full width with all borders. A table with no borders is rarely what a caller
  producing a business document wants, and `TableFormat` overrides it.
- `TableBorders.Outline` writes the inside borders as an explicit "none", so a table style cannot
  reinstate the grid lines the caller asked not to have.
- Cell shading writes an explicit `w:val="clear"` pattern; a fill colour without one has no effect.
- `TableCell.CreateElement` always creates a paragraph, because `CT_Tc` requires block content and Word
  offers to repair a document with an empty cell.
- `GetAllTexts()` now walks blocks in document order, so a table between two paragraphs reads between
  them. Rows join with `\n`, cells with `\t`.

Column spanning is supported through `TableCellFormat.ColumnSpan`. Row height, cell-level borders, and
table styles beyond passing a `StyleId` through are not implemented.

24 tests added.
