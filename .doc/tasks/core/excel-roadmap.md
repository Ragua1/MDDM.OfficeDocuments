# Excel Core Roadmap

Date: 2026-05-31

This document groups the Excel backlog slices that belong in the minimal core and records their current delivery status.

## EXCEL-001 Range-centric API and public surface cleanup

- Status: Delivered, with follow-up cleanup still open
- Goal: introduce first-class range workflows and reduce direct OpenXml exposure on the default consumer path
- Current state:
  - `IRange` exists
  - range reads, writes, styling, merge, filter, sorting, validation, and conditional formatting exist
  - several raw OpenXml-oriented members remain available only as compatibility surfaces
- Remaining follow-up:
  - continue reducing or isolating the remaining compatibility-oriented OpenXml members

## EXCEL-002 Bulk insert and tabular import workflows

- Status: Delivered
- Goal: improve reporting ergonomics by avoiding repetitive row-by-row code
- Current state:
  - `IWorksheet.AddRows(IEnumerable<IEnumerable<object?>> ...)` exists
  - `IWorksheet.AddRows<T>(IEnumerable<T> items, ...)` exists
- Remaining follow-up:
  - add higher-level import/export helpers only after the current core surface remains stable

## EXCEL-003 Worksheet operations and workbook usability

- Status: Delivered
- Goal: make normal workbook maintenance tasks first-class public features
- Current state:
  - worksheet rename, move, copy, hide, and remove operations exist
  - freeze panes and auto-fit workflows exist
- Remaining follow-up:
  - keep this surface documentation aligned if more workbook ergonomics are added

## EXCEL-004 Validation, formatting, and annotations

- Status: Delivered
- Goal: support richer editable-workbook scenarios without leaving the minimal-core story
- Current state:
  - data validation exists
  - conditional formatting exists
  - hyperlinks and comments exist
  - named ranges exist
  - worksheet and workbook protection exist
  - worksheet images exist
- Remaining follow-up:
  - document calculation behavior clearly where formulas and annotations overlap with consumer expectations

## Next slices

- [EXCEL-005 Style pipeline performance hardening](excel/EXCEL-005-style-pipeline-performance-hardening.md)
- [EXCEL-006 Worksheet and row lookup indexing](excel/EXCEL-006-worksheet-and-row-lookup-indexing.md)
- [EXCEL-007 OpenXml compatibility surface isolation](excel/EXCEL-007-openxml-compatibility-surface-isolation.md)
