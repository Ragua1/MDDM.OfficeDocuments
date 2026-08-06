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

### 2026-08-06 — the defects the merge coverage was hiding

Verifying the change above meant asking what the produced document actually looked like, rather
than whether the tests passed. Building the pre-change library into a temporary worktree and
running the same generator against both DLLs showed `styles.xml` and every `sheet*.xml`
byte-identical for valid input — but writing that harness surfaced two defects that no test
covered, because **the merge reference for an ordinary horizontal `AddCellOnRange` was never
asserted anywhere.** The only merge reference pinned in the whole suite belonged to
`IRange.Merge()`.

**A parent style was re-applied on every access, not once at creation.** `Row.GetOrCreateCell` and
`Worksheet.GetOrCreateRow` both merged the parent's style into whatever they returned, including a
cell or row that already existed. Anything touching the same cell twice therefore stamped the row
style back over what the caller had set, so for a facet both levels set the *wider* level won —
the inverse of the precedence `StyleInheritanceTests` documents. `Range.Merge()` calls
`EnsureCells()` a second time after `ApplyStyle` has run, so `AddCellOnRange(..., style)` on a
styled row lost the caller's font size on every call. Reproducible without `AddCellOnRange` at all:

```csharp
var range = worksheet.GetRange("B1:D1");
range.ApplyStyle(cellStyle);   // correct on its own
range.Merge();                 // second touch — row style wins, cellStyle's size is gone
```

Fixed by applying a parent style only when the child is created. Two follow-ons fell out of it:
`Row.CreateCell` now styles the cells it backfills, which it never did — they used to sit bare next
to styled neighbours, and only picked the row style up if something addressed them again later —
and the workbook default is skipped there, because an unstyled sheet still hands down a style whose
index is `0` and `s="0"` means what leaving the attribute off means.

**Overlapping merges were written without complaint.** `AppendMergeReference` skipped an exact
duplicate reference and nothing else, so `A1:D1` followed by `C1:F1` produced a file Excel reports
as damaged. It now throws `ArgumentException` naming the range it collides with. The check reads
the existing merges through `GetFirstChild` rather than the `MergeCells` property, because that
property creates the element on access and `CT_MergeCells` requires at least one child — a rejected
merge would otherwise leave an empty and schema-invalid `<mergeCells/>` behind. Complexity is
unchanged: the duplicate check was already a linear scan per call.

Coverage added — the gap list that produced them, in order: `MergedCellTests` (new) pins the
horizontal, block, vertical and single-cell cases, `IRow`/`IWorksheet` equivalence, style
application across the *whole* range rather than only the returned cell, the overlap rejection, and
that a rejected merge leaves no empty element; `StyleInheritanceTests` gains the backfill and
second-access cases; `DocumentStructureTests.MergedRanges_AreSchemaValidAndKeepWorksheetChildrenInOrder`
puts merges through the schema validator and pins `mergeCells` after `sheetData`, which no test did
before.

Verified: `dotnet build OfficeDocuments.slnx` → 0 errors, 9 warnings, all of them
`WorksheetPart.Worksheet` being nullable in `DocumentFormat.OpenXml` itself.
`dotnet test OfficeDocuments.slnx` → 1 774 tests pass on `net8.0`, `net9.0` and `net10.0`. The
old-versus-new document diff is now down to one thing: cells touched by a styleless range operation
no longer carry a redundant `s="0"`.
