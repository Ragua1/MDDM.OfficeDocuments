---
description: "Use when changing the Excel library, Spreadsheet/Worksheet/Row/Cell behavior, style handling, formulas, ranges, tables, or Excel tests in MDDM.OfficeDocuments."
name: "Excel Library Guidance"
applyTo:
  - "src/OfficeDocuments.Excel/**/*.cs"
  - "test/OfficeDocuments.Excel.Tests/**/*.cs"
---
# Excel Guidance

- Treat `src/OfficeDocuments.Excel` as the primary and more mature module in this repository.
- Respect the public boundary in `Interfaces/*`. Keep OpenXml-specific logic inside implementation classes.
- Work through the established hierarchy: `ISpreadsheet -> IWorksheet -> IRow -> ICell`.
- Keep row and column indexing 1-based. Invalid indexes should remain explicit failures, not silent coercions.
- Preserve OpenXml ordering rules. `Row.CreateCell` backfills missing earlier cells intentionally; do not replace this with sparse append behavior.
- Reuse the central style pipeline: `Spreadsheet.CreateStyle(...)`, `IStyle.CreateMergedStyle(...)`, and the helpers in `Utils`.
- When touching ranges, style creation, or XML merging, optimize for efficiency on larger sheets. Avoid nested scans and repeated DOM lookups in loops.
- Do not introduce new public APIs that require callers to understand OpenXml internals.
- If you add or change public Excel behavior, update tests first or alongside the change.
- Keep the root `README.md` high-level and update the detailed Excel documentation in `.doc/excel-library.md`.
- Keep terminology aligned with `.doc/terminology.md`.
