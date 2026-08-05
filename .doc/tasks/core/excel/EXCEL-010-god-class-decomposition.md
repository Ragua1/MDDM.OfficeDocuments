# EXCEL-010 God-class decomposition (Worksheet and Spreadsheet)

Date: 2026-07-24

## Status

- Tier 1 Step A (partial-class split): **Delivered 2026-07-24.** Both classes were split into a
  coordinator plus responsibility partials, with no public API change. Verified green across
  `net8.0`/`net9.0`/`net10.0` (174 tests). See the progress log at the end.
- Tier 1 Step B (collaborator extraction): **Delivered 2026-07-27.** Nine internal collaborators;
  the coordinator partials are now pure lazy-field + delegation. Verified green across all TFMs
  (174 tests). See the progress log at the end.
- Tier 2 (interface split → `Excel.Advanced`): Open.

## Business goal

`DataClasses/Worksheet.cs` (~1183 lines) and `Spreadsheet.cs` (~841 lines) each mix many
unrelated responsibilities. This makes them hard to read, risky to change, and a poor base for
the remaining Excel work (style-pipeline and lookup-indexing perf, leakage removal, and any
future feature). The goal is to decompose both into a small coordinator plus focused
collaborators, **without changing the public API first**, so later feature and performance work
lands on smaller, testable units.

## Why core or advanced

Core. This is an internal maintainability refactor of the two central Excel types. It keeps the
public `IWorksheet` / `ISpreadsheet` contracts and the public `Spreadsheet` class unchanged in
its first (non-breaking) tier. The optional second tier (interface splitting) is the seam that
later enables the `Excel.Advanced` package split described in
[../../../architecture/target-package-boundaries-and-instantiation.md](../../../architecture/target-package-boundaries-and-instantiation.md).

## Current-state responsibility map

### `Worksheet` (internal class : Base, IWorksheet)

Shared mutable state (must stay in the coordinator, never duplicated into a collaborator):
`Spreadsheet` back-reference, `WorksheetPart`, `Element` (SheetData), `WorksheetElement`,
`Rows`, `_rowsByIndex`, `_cellsByReference`, `_currentRow`, `_columns`, `_mergeCells`,
static `PropertyCache`, `Style` (from `Base`).

| # | Cluster | Members (approx. lines) | Touches row/cell state? |
| --- | --- | --- | --- |
| 1 | Row/cell store (**core**) | `AddRow`, `AddCell*`, `AddCellOnIndex`, `AddCellOnRange`, `AddCellWithFormula`, `GetOrCreateRow`, `GetRow`, `GetCell`, `GetCellByReference`, `RegisterCell`, `RegisterRow`, `CurrentRow/CurrentCell`, `Next*Index` (103–290, 815–869) | **Yes** — the hot state |
| 2 | Ranges (**core**) | `GetRange`, `TryGetRange` (198–241) | Reads only |
| 3 | Bulk import (**core**) | `AddRows`, `AddRows<T>`, `IsScalarType`, `GetReadableProperties`, `GetDisplayText` (290–395, 871–904) | Via `AddRow`/`AddCell` |
| 4 | Columns (**core**) | `Columns`, `SetColumnWidth`, `AutoFitColumns` (36–52, 397–508) | Reads rows for autofit |
| 5 | Merge (**core**) | `MergeCells`, `AppendMergeReference` (54–70, 529–537) | No |
| 6 | Freeze panes (**core**) | `FreezePanes`, `ClearFrozenPanes` (427–468) | No |
| 7 | AutoFilter (**advanced**) | `SetAutoFilter` (539–549) | No |
| 8 | Data validation (**advanced**) | `AddDataValidation` (551–622) | No |
| 9 | Conditional formatting (**advanced**) | `AddConditionalFormatting`, `GetNextConditionalFormattingPriority`, `EscapeFormulaString` (624–679, 906–913, 1044) | No |
| 10 | Hyperlinks (**advanced**) | `SetCellHyperlink`, `GetCellHyperlink` (681–749) | No |
| 11 | Comments + VML (**advanced**) | `SetCellComment`, `GetCellComment`, `UpdateCommentVml`, `BuildCommentVml` (751–813, 1000–1042) | No |
| 12 | Images / drawing (**advanced**) | `AddImage` ×2, `BuildTwoCellAnchor`, `EnsureDrawingElement`, `ToImagePartType`, `DetectImageType` (1049–1182) | No |
| 13 | Protection (**advanced**) | `Protect` (510–527) | No |
| 14 | Element-order infra (**infra**) | `InsertAfterMergeCells`, `InsertAfterConditionalFormatting`, `InsertAfterDataValidations` (915–998) | No |

Key fact: clusters 7–13 operate only on `WorksheetElement` / `WorksheetPart` and the element-order
infra (14). **None of them touch the row/cell dictionaries or `_currentRow`.** That makes them
cleanly extractable. The three `InsertAfter*` ladders (14) are near-identical and encode the
`CT_Worksheet` child order — they must be centralized first because 8–12 all depend on them.

### `Spreadsheet` (public class : ISpreadsheet)

Shared state: `_worksheets`, `_document`, `_isEditable`, `_defaultStyle`, `_disposed`, and the
derived `WorkbookPartInternal` / `SheetsInternal` / `StylesheetInternal` accessors.

| # | Cluster | Members | Notes |
| --- | --- | --- | --- |
| 1 | Lifecycle (**core, public**) | ctors, `CreateDocument`, `OpenDocument`, `Close`, `Dispose`, finalizer | keep in coordinator |
| 2 | Worksheet CRUD (**core, public**) | `AddWorksheet`, `GetWorksheet`, `GetWorksheetsName`, `RenameWorksheet`, `MoveWorksheet`, `CopyWorksheet`, `SetWorksheetHidden`, `RemoveWorksheet`, `GetSheet*`, `GetWorksheetOrThrow`, `EnsureWorksheetNameAvailable` | operates on `Sheets` + `_worksheets` |
| 3 | Styles + differential formats (**core/advanced**) | `CreateStyle` ×2, `InitStylesheet`, `StylesheetInternal`, `GetStyleElement`, `GetOrCreateDifferentialFormat`, `CreateDifferentialFormat` | natural home for the EXCEL-005 perf work |
| 4 | Tables (**advanced**) | `AddTable` ×2, `GetTable`, `GetTables` ×2, `RenameTable`, `ResizeTable`, `RemoveTable`, `FindTableDefinitionPart` | biggest self-contained cluster |
| 5 | Named ranges (**advanced**) | `AddNamedRange`, `IsValidNamedRange` | operates on `Workbook.DefinedNames` |
| 6 | Protection (**advanced**) | `ProtectWorkbook`, `ComputeProtectionPassword` | small |

## Tier 1 — non-breaking decomposition

Two complementary techniques. Neither changes any public signature, so the whole tier stays
binary- and source-compatible; `IWorksheet`, `ISpreadsheet`, and the public `Spreadsheet` class
are untouched.

### Step A — `partial class` split (mechanical, zero risk)

Move code into multiple files under the same `partial class`. No behavior change, fully shared
state. This is the cheap first cut that makes everything below reviewable.

- `Worksheet` → `Worksheet.cs` (fields, ctor, clusters 1–2), `Worksheet.Columns.cs` (4–6),
  `Worksheet.BulkImport.cs` (3), `Worksheet.Annotations.cs` (7–11),
  `Worksheet.Drawing.cs` (12), `Worksheet.ElementOrder.cs` (13–14).
- `Spreadsheet` → `Spreadsheet.cs` (1–2), `Spreadsheet.Styles.cs` (3),
  `Spreadsheet.Tables.cs` (4), `Spreadsheet.NamedRanges.cs` (5), `Spreadsheet.Protection.cs` (6).

Mark the type `internal sealed partial class Worksheet` / `public sealed partial class Spreadsheet`
(sealing `Spreadsheet` is technically a source-compat change for subclassers; it currently has a
`protected virtual Dispose(bool)` and a finalizer, so if strict compat is required, keep it
unsealed in Tier 1 and seal in Tier 2).

### Step B — extract internal collaborators (the real decoupling)

Each collaborator is an `internal` class that owns one responsibility, receives only the
OpenXml part(s)/accessors it needs, and is delegated to by the (now-thin) coordinator method.
Non-breaking because the collaborators are internal and observable behavior is unchanged.

Extract in this order (dependencies first):

1. **`WorksheetElementOrderer`** — wraps `WorksheetElement` and centralizes the three
   `InsertAfter*` ladders behind one `InsertInOrder(element, afterOneOf: [...])` API plus the
   lazy `Columns` / `MergeCells` / `AutoFilter` accessors. Removes the triplication and becomes
   the single place that knows the `CT_Worksheet` child order. Prerequisite for 3–7 below.
2. **`StylesheetManager` / `DifferentialFormatCache`** (Spreadsheet cluster 3) — owns
   `InitStylesheet`, the default style, `CreateStyle`, and differential-format dedup. This is
   also the home where EXCEL-005 fixes the O(N²) style scans and the `id <= 0` dedup bug.
3. **`CommentWriter`** (Worksheet cluster 11) — comments + VML, takes `WorksheetPart`.
   Self-contained; the natural place to later fix the O(N²) "rebuild all VML on every comment".
4. **`WorksheetImageWriter`** (12) — images/drawing, takes `WorksheetPart` + orderer.
   Mostly static builders already.
5. **`DataValidationWriter`** (8) and **`ConditionalFormattingWriter`** (9) — take
   `WorksheetElement` + orderer (+ the differential-format cache for CF).
6. **`HyperlinkStore`** (10) — takes `WorksheetPart` + orderer.
7. **`TableManager`** (Spreadsheet cluster 4) — the biggest single win; takes `WorkbookPart` +
   a worksheet lookup. **`NamedRangeManager`** (5) and **`WorkbookProtector`** (6) follow.
8. Optional: **`TabularImporter`** (Worksheet cluster 3) — the reflection-based `AddRows<T>`
   with the `PropertyCache`, behind a row-writer callback.

Each extraction is one commit/PR: create the collaborator, move the code, replace the coordinator
body with a one-line delegation, run the full Excel suite (currently 174 tests, many asserting
real OpenXml nodes and element ordering).

Worked example (Comments):

```csharp
// Worksheet.Annotations.cs (coordinator, after extraction)
private CommentWriter? _commentWriter;
private CommentWriter CommentWriter => _commentWriter ??= new CommentWriter(WorksheetPart);

internal void SetCellComment(Cell cell, string text, string? author)
    => CommentWriter.Set(cell.CellReference, text, author);

internal string? GetCellComment(string cellReference)
    => CommentWriter.Get(cellReference);
```

```csharp
// CommentWriter.cs (new internal class)
internal sealed class CommentWriter(WorksheetPart worksheetPart)
{
    public void Set(string cellReference, string text, string? author) { /* moved body */ }
    public string? Get(string cellReference) { /* moved body */ }
    // UpdateVml / BuildVml move here as private members
}
```

### Tier-1 end state

- `Worksheet` and `Spreadsheet` become coordinators of ~200–300 lines each: the row/cell store,
  ranges, lifecycle, and worksheet CRUD, delegating everything else.
- Each advanced feature lives in its own tested collaborator.
- The public API is byte-for-byte the same.

## Tier 2 — breaking-change follow-up (optional, later)

Only after Tier 1 is stable. These change the public surface (major version) and align with the
`Excel.Advanced` direction:

1. **Split the fat interfaces into role interfaces.** `IWorksheet` keeps the core store/range
   members; new `IWorksheetAnnotations` (validation, conditional formatting, hyperlinks,
   comments), `IWorksheetDrawing` (images), and `IWorksheetProtection` carry the advanced
   members. Same for `ISpreadsheet` → core + `IWorkbookTables` + `IWorkbookProtection` +
   `INamedRanges`. The concrete classes still implement all interfaces; consumers can depend on
   the narrow ones. Moving advanced members **off** the core interface is the breaking part.
2. **Relocate the advanced collaborators + their role interfaces into `OfficeDocuments.Excel.Advanced`**,
   surfaced via extension methods on the core types or a thin facade. The Tier-1 collaborators are
   exactly the units that move.
3. **Rationalize construction**: seal `Spreadsheet`, collapse the overlapping
   constructor/`CreateDocument`/`OpenDocument` paths, and remove the obsolete raw-OpenXml members
   once the compatibility window closes.

## Complexity

- Tier 1 Step A (partial split): Low.
- Tier 1 Step B (collaborator extraction): Medium — one collaborator at a time, each small.
- Tier 2 (interface split + relocation): Medium-High; coordinate with the package-boundary work.

## Risks

- **OpenXml child order** is load-bearing. Centralize it in `WorksheetElementOrderer` and rely on
  the existing ordering assertions (e.g. drawing-before-`TableParts`, and the freeze/validation/CF
  round-trip tests) to catch regressions.
- **Shared mutable state.** The row/cell dictionaries, `_currentRow`, and the lazy
  `Columns`/`MergeCells` caches must stay owned by the coordinator; collaborators receive accessors,
  never their own copies.
- **The `Spreadsheet` back-reference for differential formats.** Conditional formatting currently
  calls `Spreadsheet.GetOrCreateDifferentialFormat`; inject the differential-format cache (or a
  delegate) into `ConditionalFormattingWriter` rather than duplicating it.
- **Keep behavior identical during extraction.** Do not fix the known O(N²) comment-VML and style
  hot paths in the same commit as the move — extract first (green tests), then let EXCEL-005/006
  optimize the now-isolated units.
- **Sealing `Spreadsheet`** is a source-compat change for subclassers; defer to Tier 2 if strict
  compatibility is required.

## Dependencies

- Builds on the current green state (174 Excel tests) after the 2026-07-24 hardening pass.
- Enables / de-risks: EXCEL-005 (style pipeline perf), EXCEL-006 (worksheet/row lookup indexing),
  EXCEL-007 (leakage isolation), and the `Excel.Advanced` split.

## Subtasks

1. Tier 1 Step A: partial-class split of `Worksheet` and `Spreadsheet`; build + full test.
2. Extract `WorksheetElementOrderer`; collapse the three `InsertAfter*` ladders.
3. Extract `StylesheetManager` / `DifferentialFormatCache`.
4. Extract `CommentWriter`, then `WorksheetImageWriter`.
5. Extract `DataValidationWriter`, `ConditionalFormattingWriter`, `HyperlinkStore`.
6. Extract `TableManager`, `NamedRangeManager`, `WorkbookProtector`.
7. (Optional) Extract `TabularImporter`.
8. (Tier 2) Split interfaces into role interfaces; then relocate advanced collaborators to
   `Excel.Advanced`.

## Acceptance criteria

- `Worksheet.cs` and `Spreadsheet.cs` coordinators are materially smaller (target: no single file
  over ~350 lines), each advanced responsibility living in its own internal collaborator.
- Public `IWorksheet` / `ISpreadsheet` / `Spreadsheet` surface is unchanged through all of Tier 1.
- The full Excel test suite stays green across `net8.0`, `net9.0`, and `net10.0` after every step.
- No new OpenXml types leak onto the public surface.
- The element-order knowledge lives in exactly one place.

## Progress log

### 2026-07-24 — Tier 1 Step A (partial-class split) delivered

Both god classes were split into a coordinator plus responsibility partials. No public API
changed (`Worksheet` is `internal`; `Spreadsheet` and its interfaces are untouched). Verified
green across `net8.0`/`net9.0`/`net10.0` (174 tests) after each extraction.

`Worksheet` (1183 → 634-line coordinator):

| File | Lines | Responsibility |
| --- | --- | --- |
| `DataClasses/Worksheet.cs` | 634 | core: row/cell store, ranges, bulk import, columns, merge, freeze, autofilter, protect |
| `DataClasses/Worksheet.Annotations.cs` | 224 | data validation, conditional formatting, hyperlinks |
| `DataClasses/Worksheet.Drawing.cs` | 146 | images / drawing |
| `DataClasses/Worksheet.Comments.cs` | 118 | comments + VML |
| `DataClasses/Worksheet.ElementOrder.cs` | 92 | the three `InsertAfter*` ladders (CT_Worksheet child order) |

`Spreadsheet` (841 → 450-line coordinator):

| File | Lines | Responsibility |
| --- | --- | --- |
| `Spreadsheet.cs` | 450 | core: workbook lifecycle, worksheet CRUD, stylesheet init |
| `Spreadsheet.Tables.cs` | 236 | structured-table create/lookup/lifecycle |
| `Spreadsheet.Styles.cs` | 95 | `CreateStyle`, differential formats, style-element access |
| `Spreadsheet.NamedRanges.cs` | 56 | named ranges + validation |
| `Spreadsheet.Protection.cs` | 39 | workbook protection + password hash |

Side benefit: the split surfaced coupling that had hidden in the monolith — several `using`
directives were only needed by a single moved cluster, so the coordinators now import materially
fewer namespaces (e.g. `Worksheet.cs` dropped from 14 usings to 6).

### 2026-07-24 — Tier 1 Step B (collaborator extraction) started

Three internal collaborators extracted; the coordinator partials now hold only a lazy
collaborator field plus thin delegations. Verified green across `net8.0`/`net9.0`/`net10.0`
(174 tests) after each extraction. No public API change.

| Collaborator | Replaces | Notes |
| --- | --- | --- |
| `DataClasses/WorksheetElementOrderer.cs` | the three `InsertAfter*` ladders | Two of the three were byte-identical; the CT_Worksheet order now lives once, behind `InsertConditionalFormatting` / `InsertDataValidations` / `InsertHyperlinks`. |
| `DataClasses/CommentWriter.cs` | `SetCellComment` / `GetCellComment` / VML | Self-contained (takes `WorksheetPart` + worksheet element). Natural home for the pending O(N²) "rebuild all VML on every comment" fix. |
| `DataClasses/WorksheetImageWriter.cs` | `AddImage` ×2 + drawing builders | Self-contained; mostly static builders already. |

### 2026-07-27 — Tier 1 Step B completed

The remaining six collaborators were extracted; every coordinator partial is now a lazy
collaborator field plus one-line delegations (8–37 lines each). Verified green across
`net8.0`/`net9.0`/`net10.0` (174 tests). No public API change.

| Collaborator | Owns | Injected dependencies |
| --- | --- | --- |
| `DataClasses/DataValidationWriter.cs` | `AddDataValidation` | worksheet element, `WorksheetElementOrderer` |
| `DataClasses/ConditionalFormattingWriter.cs` | `AddConditionalFormatting` + priority/escape helpers | worksheet element, orderer, `Func<IStyle,uint>` differential-format resolver |
| `DataClasses/HyperlinkStore.cs` | `Set`/`Get` hyperlink (display text stays with the caller) | worksheet part, worksheet element, orderer |
| `DataClasses/TableManager.cs` | all structured-table create/lookup/lifecycle | workbook part, `Func<string,Worksheet>` lookup, `Func<IEnumerable<Worksheet>>` catalog |
| `DataClasses/NamedRangeManager.cs` | `AddNamedRange` + validation | workbook part, `Func<string,int>` sheet-index resolver |
| `DataClasses/WorkbookProtector.cs` | `ProtectWorkbook` + the legacy password hash | workbook part |

`ComputeProtectionPassword` was shared by both `Spreadsheet.ProtectWorkbook` and
`Worksheet.Protect`; it moved to `WorkbookProtector` as a `public static` helper and both call
sites now use it.

Step B end state: the two coordinators (`Worksheet` ~634 lines, `Spreadsheet` ~450 lines) plus
nine focused, independently constructible collaborators. The advanced writers/managers depend
only on OpenXml parts and small delegate seams — exactly the units that move to
`OfficeDocuments.Excel.Advanced` in Tier 2.

Next: Tier 2 — split the fat `IWorksheet`/`ISpreadsheet` interfaces into role interfaces, then
relocate the advanced collaborators into `OfficeDocuments.Excel.Advanced` (breaking; major version).
