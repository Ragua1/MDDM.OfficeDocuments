# EXCEL-007 OpenXml compatibility surface isolation

Date: 2026-05-31

## Business goal

Reduce coupling between the public Excel API and raw `DocumentFormat.OpenXml` types so future library evolution stays safer and more predictable.

## Why core or advanced

Core. The default value of the library is a smaller and more consumer-friendly API than raw OpenXml.

## Functional description

The Excel library should keep existing compatibility members available where needed, but all raw OpenXml exposure should be clearly marked as compatibility-oriented and discouraged on the main consumer path.

## Technical guidance

Relevant files include:

- `src/OfficeDocuments.Excel/Interfaces/IRow.cs`
- `src/OfficeDocuments.Excel/Interfaces/IWorksheet.cs`
- `src/OfficeDocuments.Excel/Interfaces/IStyle.cs`
- `src/OfficeDocuments.Excel/Interfaces/IOpenXmlWrapper.cs`
- `src/OfficeDocuments.Excel/Interfaces/ISpreadsheet.cs`
- `src/OfficeDocuments.Excel/Spreadsheet.cs`

Observed issue:

- most raw OpenXml compatibility members are already hidden behind `EditorBrowsable(Never)` and `Obsolete(...)`
- `IRow.Element` still exposes `DocumentFormat.OpenXml.Spreadsheet.Row` without the same compatibility warning layer

Implementation direction:

- mark remaining raw OpenXml members consistently with `EditorBrowsable(EditorBrowsableState.Never)` and targeted `Obsolete(...)` guidance
- preserve binary compatibility and avoid removing members in the current major version
- keep tests focused on public behavior first; only use raw element access where serialization-critical assertions require it
- as follow-up work, review whether more internal test assertions can move from raw element checks to public API checks

## Complexity

Low.

## Risks

- new obsolete warnings can surface in tests or downstream consumers that still use compatibility members
- inconsistent messages across interfaces can confuse migration guidance

## Dependencies

- none

## Subtasks

- audit public interfaces for remaining raw OpenXml members without compatibility annotations
- annotate the remaining members consistently
- keep existing members intact for compatibility
- update or suppress internal test usage only where warnings become too noisy

## Acceptance criteria

- remaining raw OpenXml compatibility members are consistently marked as non-primary API surface
- no public member required for compatibility is removed
- Excel projects still build and existing behavior remains unchanged

## Progress log

### 2026-08-06 — test suite no longer consumes the compatibility surface

Delivered the follow-up subtask ("review whether more internal test assertions can move from raw
element checks to public API checks"). A build of `OfficeDocuments.slnx` reported 180 warnings, all
of them in `Excel.IntegrationTests` and `Excel.VerificationTests` and none in `src/` — 60 unique
sites across three target frameworks. That count was the useful signal: it measured how much of the
deprecated surface the test suite still required, and therefore how much of it could not be removed
in v5 without leaving a behaviour untested.

All 32 `CS0618` sites are gone, and **none of them needed a suppression** — every one had a public
or TestKit replacement already:

- 10 were `Assert.NotNull(style.Element)` on a non-nullable property. Tautologies; deleted.
- 9 reached the stylesheet by hand (`style.Stylesheet.Fonts!...ElementAt(style.FontId)`) to fetch the
  entry a style points at. `StylesheetProbe.Font/Fill` already did exactly this. `StyleAllocationTests`
  simply predated the probe and was never migrated.
- 5 read `style.Element.Alignment`, the one facet `IStyle` exposes no accessor for. Added
  `StylesheetProbe.Alignment(IStyle)` rather than widening `IStyle` — the probe is where the
  suppression is confined by design.
- 1 compared two styles' stylesheets by identity. Added `StylesheetProbe.ShareStylesheet`.
- 3 read `cell.Element.CellFormula.Text`; `ICell.GetFormula()` is literally that expression.
- 4 were `Assert.NotNull(x.Element)` on `IRow` / `IOpenXmlWrapper<Cell>`. Tautologies; deleted.
- 1 used `IWorksheet.Spreadsheet` to reach `CreateStyle` from a private test helper. The caller
  already had the `ISpreadsheet`; it is now a parameter.

Also fixed 28 `CS8602`/`CS8604` in the same files, and the API defect one group of them pointed at —
see the `AddCellOnRange` entry below.

Verified: `dotnet build OfficeDocuments.Excel.slnx` → 0 errors; `dotnet test OfficeDocuments.slnx`
→ 484 Excel + 241 Word tests pass on `net8.0`, `net9.0`, and `net10.0`.

Left deliberately: 3 sites (9 warnings) where `WorksheetPart.Worksheet` is nullable in
`DocumentFormat.OpenXml` itself — `DifferentialFormatTests.cs:40`,
`DocumentStructureTests.cs:107`, `WorkbookRoundtripTests.cs:208`. Not this library's surface.

**Consequence for the P0:** the test suite no longer pins any deprecated member. Removing the
`Element` / `Stylesheet` / `Spreadsheet` compatibility properties in a future major version is now a
question of consumer impact only, not of losing coverage.

### 2026-08-06 — `AddCellOnRange` error contract unified (breaking)

Three `CS8602` sites traced back to `ICell? AddCellOnRange(...)`. The nullable return existed
because the method reported one invalid-input class by returning `null` while reporting the others
with `ArgumentException`, and each of the three `Worksheet` overloads carried its own copy of the
bounds check — which is how they came to disagree: the 3-argument overload rejected
`beginColumn >= endColumn`, the 5-argument one rejected `beginColumn > endColumn`, so a
single-column range was invalid on one path and produced a degenerate `mergeCell ref="A1:A1"` on the
other.

Now, on `IWorksheet` and `IRow` alike: an index below 1 or an end index before its begin index
throws `ArgumentException` naming the parameter; a valid range always returns a non-nullable `ICell`;
a range covering exactly one cell returns that cell and writes no merge element. The 3-argument
`Worksheet` overload delegates to the 5-argument one, so there is one copy of the contract.

Documented in [../../../migration-v3-to-v4.md](../../../migration-v3-to-v4.md). Regression tests:
`WorksheetTest.AddCellOnRange_*` and `RowTest.AddCellOnRange_*`.
