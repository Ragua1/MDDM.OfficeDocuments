# EXCEL-006 Worksheet and row lookup indexing

Date: 2026-05-31

## Business goal

Make common read and write operations scale better on sparse or larger worksheets without changing the public API.

## Why core or advanced

Core. Row and cell lookup is part of nearly every worksheet workflow, including write, read, range, and roundtrip scenarios.

## Functional description

The Excel library should maintain fast internal lookup paths for rows and cells while preserving existing ordering rules and public API behavior.

## Technical guidance

Current lookup hotspots are concentrated in:

- `src/OfficeDocuments.Excel/DataClasses/Worksheet.cs`
- `src/OfficeDocuments.Excel/DataClasses/Row.cs`
- `src/OfficeDocuments.Excel/DataClasses/Range.cs`
- `src/OfficeDocuments.Excel/DataClasses/Cell.cs`

Observed issues:

- `Worksheet.GetRow(...)` scans the ordered row list
- `Worksheet.GetCellByReference(...)` flattens worksheet cells and scans by reference
- `Row.GetCell(...)` scans the row cell list
- sparse-cell creation backfills by repeatedly probing existing cells
- range operations multiply these costs through nested loops

Implementation direction:

- keep the ordered `Rows` and `Cells` public surfaces for compatibility, but add internal dictionaries for fast lookup
- maintain a worksheet-level `rowIndex -> row` index
- maintain a worksheet-level `cellReference -> cell` index using ordinal-ignore-case comparison
- maintain a row-level `columnIndex -> cell` index
- update indexes both when opening an existing workbook and when creating rows or cells programmatically
- preserve current OpenXml ordering by continuing to insert new rows and cells into the DOM and ordered lists at the correct position
- validate with focused tests plus a reopen roundtrip test so the indexes work both for new and reopened workbooks

## Complexity

Medium.

## Risks

- indexes can become stale if not updated on every insertion path
- ordering bugs can appear if list maintenance and dictionary maintenance diverge
- roundtrip loading must build the same in-memory state as write-time creation

## Dependencies

- none

## Subtasks

- add worksheet row and cell lookup dictionaries
- add row cell lookup dictionary
- centralize row and cell registration helpers to avoid drift
- update open-from-existing-document constructors to populate indexes
- add focused tests for sparse rows, row lookup, and cell reference lookup
- add a roundtrip test covering reopen and lookup behavior

## Acceptance criteria

- `GetRow(...)` no longer scans the worksheet row list for normal lookup
- `GetCellByReference(...)` no longer materializes and scans all worksheet cells for normal lookup
- `Row.GetCell(...)` no longer scans the row cell list for normal lookup
- current row and cell ordering behavior remains unchanged
- focused Excel row, worksheet, and reader tests pass
