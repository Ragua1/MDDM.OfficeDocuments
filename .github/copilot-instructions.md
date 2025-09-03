# Copilot instructions for MDDM.OfficeDocuments

Purpose: Help AI agents work productively in this .NET library that wraps OpenXml for Excel (mature) and Word (early). Keep edits idiomatic to this repo.

## Big picture
- Layout: `src/OfficeDocuments.Excel` (main), `src/OfficeDocuments.Word` (scaffold), `test/*` (xUnit tests), CI in `.github/workflows`.
- Public Excel API in `Interfaces/*` (ISpreadsheet, IWorksheet, IRow, ICell, IStyle); implementations are internal in `DataClasses/*` and `Styles/*`.
- `Spreadsheet` owns `SpreadsheetDocument`, `WorkbookPart`, and the `Stylesheet`. Styles are centralized; avoid ad‑hoc OpenXml edits in features.

## Excel architecture & patterns
- Abstractions: ISpreadsheet → IWorksheet → IRow → ICell. Use these; don’t pass OpenXml types across public boundaries.
- Styles: create via `Spreadsheet.CreateStyle(...)`; compose with `CreateMergedStyle(...)`. Default style set in `Spreadsheet.InitStylesheet()`; style ID 0 = default.
- `Worksheet` lazily creates `Columns` and `MergeCells`. Creating a cell can backfill missing earlier cells to maintain OpenXml order.
- Style merge: fonts/fills/borders merged by `Utils.Merge*`; number format and alignment are applied but not deeply merged. Numeric/Date setters add common defaults if none.
- Tables: `ISpreadsheet.AddTable(sheetName, startCell, endCell, columns)` creates `TableDefinitionPart` and updates `TableParts`.

## Conventions & gotchas
- Indices are 1-based for rows/columns; invalid (<1) throws `ArgumentException`.
- Prefer `AddCell(...)`; `[Obsolete] AddCellWithValue(...)` is kept for legacy tests.
- `Close()` saves only when the document is editable; always dispose/close to flush and release resources.
- Keep argument validation, naming, and null-handling consistent with existing methods.

## Team standards
- Operate as a senior developer: propose minimal, high-impact changes, explain trade-offs briefly, and account for edge cases.
- Prioritize Microsoft documentation and official .NET guidelines; when unspecified, follow current OSS community conventions.
- Unit tests: include both positive (happy path) and negative (error/edge) cases; name tests using `MethodName_StateUnderTest_ExpectedOutcome`.
 - Match existing naming and nullability patterns; prefer explicit argument validation and consistent exceptions (`ArgumentException`, `ArgumentNullException`).
 - Avoid exposing OpenXml types in public APIs; stick to the `Interfaces/*` layer.
 - Prefer lazy creation patterns already used (e.g., `Columns`, `MergeCells`); avoid redundant DOM traversals in tight loops.
 - When adding large-range operations, ensure OpenXml node order is preserved and backfill logic remains O(range) without quadratic behavior.
 - Update `README.md` snippets when adding new public features; mirror patterns from tests.
 - Add a small example to tests for new API surfaces (kept fast and deterministic).
 - Ensure builds pass with the SDK from `global.json` (9.0.0) and tests run quickly; keep long-running integration tests behind explicit flags.

## Extending safely
- Define/extend interfaces in `Interfaces/*`, implement in `DataClasses/*` (internal). Reuse `Style`/`Utils` to manage stylesheet entries.
- Validate inputs consistently; keep 1-based semantics and exception types aligned with existing methods.
- Add focused tests under `test/OfficeDocuments.Excel.Tests` following `CreationTest.cs` and `UtilsTest.cs` patterns; include positive and negative cases and name them `MethodName_StateUnderTest_ExpectedOutcome`.

## Key files
- `src/OfficeDocuments.Excel/Spreadsheet.cs` (workbook, stylesheet init, sheets, tables)
- `src/OfficeDocuments.Excel/DataClasses/{Worksheet,Row,Cell,Style}.cs`
- `src/OfficeDocuments.Excel/Styles/*`, `src/OfficeDocuments.Excel/Enums/*`
- Tests: `test/OfficeDocuments.Excel.Tests/{CreationTest,UtilsTest,SpreadsheetTestBase}.cs`

If anything is unclear (e.g., alignment merge nuances or conditional formatting scope), flag it so we can refine these rules.
