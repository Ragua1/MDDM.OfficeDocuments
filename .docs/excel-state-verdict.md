# OfficeDocuments.Excel — State Verdict and Direction

Date: 2026-07-24

This document is an independent assessment of the **Excel module only**. It reviews the
implementation state, catalogs the real technical debt (verified against the code, not just
the existing planning docs), states what is genuinely good, and compares three possible
directions for the project.

It complements the existing planning material and does not replace it:

- [library-benchmark-report.md](library-benchmark-report.md) — external positioning
- [feature-gap-backlog.md](feature-gap-backlog.md) — feature gaps
- [architecture/target-package-boundaries-and-instantiation.md](architecture/target-package-boundaries-and-instantiation.md) — the core/advanced/interop split direction
- [tasks/roadmap-overview.md](tasks/roadmap-overview.md) — cross-module roadmap

## Executive verdict

`OfficeDocuments.Excel` is **functionally healthy and genuinely useful**. It builds clean,
168 tests pass (net9.0, 0 failed, 0 skipped), and it covers a broad, practical feature set
(ranges, bulk insert, worksheet lifecycle, tables, validation, conditional formatting,
hyperlinks, comments, named ranges, protection, images) with mostly real round-trip OpenXml
assertions rather than no-throw smoke tests. The core value proposition — "a smaller, more
predictable API over the Open XML SDK" — is credible today for Excel.

The debt is **not in "does it work" but in "is it finished and safe at the edges"**. There are
a handful of concrete correctness bugs, one design decision that actively contradicts the stated
scope (a half-implemented formula calculation engine), two god classes, a broken CI wiring, and
documentation that has drifted ahead of / behind the code. None of these block everyday use;
all of them are the kind of thing that will bite a new consumer or a future contributor.

**Bottom line:** the Excel module is at "works well, needs a hardening pass and one scope
decision" — not at "needs a rewrite" and not at "done, ship and forget".

## Implementation state snapshot

| Aspect | State |
| --- | --- |
| Build | Clean (net8.0 / net9.0 / net10.0 multi-target) |
| Tests | 168 pass / 0 fail / 0 skip (net9.0); mostly real round-trip OpenXml assertions |
| Package | Version 4.0.0, full NuGet metadata, central package management, OpenXml 3.5.1 |
| CI | **Broken** — the restore step targets a file that does not exist (see D-1) |
| Activity | Last real commit 2026-06-01; ~7 weeks dormant as of this date |
| History | Started 2019; 118 commits over ~7 years — long-lived, burst-driven side project |
| OpenXml leakage cleanup | ~50% done — interface-level leaks are `[Obsolete]` + `[EditorBrowsable(Never)]`; `Styles.*` and `Utils` leaks are not yet annotated |
| Factory internalization (EXCEL-009) | Effectively done in code (interfaces are `internal`), but now dead and undeleted |

## Technical debt

Severities are this review's own judgment after verifying each item in the source, not a
restatement of the planning backlog. `file:line` references point at the current code.

### A. Correctness bugs (verified)

- **[High] The formula "calculation engine" only works horizontally on the current row.**
  `GetFormulaValue()` dispatches to `FormulaSum` / `CountCellsWithValue` / `CountCellsIf` /
  `GetMedian`, all of which iterate `Worksheet.GetCell(columnIndex)` — the single-argument
  overload, which reads the *current row only* — and destructure `GetExcelCellIndex()` as
  `var (_, col)`, discarding the row entirely. So `SUM(A1:A10)` (a vertical range) collapses to
  one cell and returns the wrong value; only same-row horizontal ranges are correct.
  `DataClasses/Cell.cs:187-295`.
- **[High] `MEDIAN` crashes on an empty range and truncates.** `Median(int[])` indexes
  `data[data.Length/2]` with no empty guard (`IndexOutOfRangeException`), and uses integer
  division for the even-count average. `DataClasses/Cell.cs:297-303`.
- **[High] `AddWorksheet(..., sheetStyle)` silently drops the style on opened workbooks.**
  Line 147 computes `_defaultStyle?.CreateMergedStyle(sheetStyle)`. `_defaultStyle` is only set
  inside `InitStylesheet()`, which is skipped when opening a document that already has a
  `WorkbookStylesPart`. When `_defaultStyle` is null the `?.` short-circuits and `sheetStyle` is
  thrown away with no error. `Spreadsheet.cs:147` (with the init guard at `:78`).
- **[Medium] `IRange.Worksheet` is null through the interface.** `Range` exposes
  `public new IWorksheet Worksheet => _worksheet`, which only shadows on the concrete type; the
  `IBase.Worksheet` slot is filled by `Base.Worksheet`, and `Range` calls `base(null)`, so any
  access through the `IRange`/`IBase` contract returns null. Range methods work because they use
  the private `_worksheet` field, but `Base.OwnerWorksheet` / `OwnerSpreadsheet` (a hard cast of
  the null property) would throw. `DataClasses/Range.cs:14,28` + `DataClasses/Base.cs:9`.
- **[Medium] Boolean cells are written as `"True"/"False"`.** `SetValue(bool)` writes
  `value.ToString()` with `DataType=Boolean`; the OOXML spec requires `"1"/"0"`. It round-trips
  inside this library but is non-conformant for other readers. `DataClasses/Cell.cs:111-114`.
- **[Medium] Culture-inconsistent reads.** Values are *written* invariant, and
  `GetDoubleValue`/`GetDecimalValue`/`TryGetValue(double|decimal)` read invariant — but
  `GetIntValue`/`GetLongValue`/`GetBoolValue` use current-culture `int.Parse`/`long.Parse`.
  Under a non-invariant `CurrentCulture` reads can diverge. `DataClasses/Cell.cs:335-366`.
- **[Medium] `Base.OwnerWorksheet` hard-casts `IWorksheet` to the concrete `Worksheet`.**
  Couples the whole data-class hierarchy to the concrete implementation;
  `InvalidCastException` for any alternate `IWorksheet`. `DataClasses/Base.cs:9`.
- **[Low-Med] `CopyWorksheet` can leave dangling relationships.** It clones the worksheet XML but
  strips only `Hyperlinks`, `LegacyDrawing`, `TableParts`; any other `r:id`-bearing child
  (drawings, comments) keeps references that do not exist in the new part. `Spreadsheet.cs:263-283`.
- **[Low-Med] Workbook child-order risk.** `AddNamedRange` / `ProtectWorkbook` append
  `DefinedNames` / `WorkbookProtection` at the end of `Workbook`, which can violate the
  `CT_Workbook` child sequence and produce an invalid file. `Spreadsheet.cs:423,446`.
- **[Low] No Excel bounds validation in reference parsing.** Columns beyond XFD (16384) and rows
  beyond 1048576 parse and format successfully, so `ZZZ9999999`-style references are accepted.
  `Extensions/CellExtension.cs:84-207`. Still open. The neighbouring input-validation gaps were
  closed by EXCEL-011 phase 6 (2026-07-28) — worksheet-name legality, C0 control characters,
  non-finite numbers and the 1900 date-serial conversion; see
  [`excel-library.md`](excel-library.md#what-the-library-refuses-to-write) for the resulting
  contract. Three of those four were invisible to the schema validator, the round trip, or both.

### B. Design and architecture debt

- **[High] The half-baked formula engine contradicts the stated scope.** The docs say the library
  "writes formulas but does not provide a calculation engine" — yet `Cell.GetFormulaValue()` *is*
  a calculation engine: 4 hard-coded functions matched by `StartsWith` (so `SUM` also matches
  `SUMIF`/`SUMPRODUCT`), returning `int` (lossy), using `-1` as a magic empty sentinel, and
  `throw new NotImplementedException()` for anything else. This needs an explicit decision:
  remove it, or commit to it as a real, documented feature. `DataClasses/Cell.cs:169-303`.
- **[High] `Worksheet` is a god class (~1183 lines).** It mixes the row/cell store, columns,
  merges, freeze panes, autofit, protection, data validation, conditional formatting, hyperlinks,
  comments (incl. raw VML generation), and image/drawing construction, with duplicated
  `InsertAfterX` fallback ladders. `DataClasses/Worksheet.cs`.
- **[Medium] `Spreadsheet` is a second god class (~841 lines)** — workbook lifecycle, worksheet
  CRUD, table CRUD, named ranges, protection, stylesheet init, differential-format dedup, and
  password hashing all in one type. `Spreadsheet.cs`.
- **[Medium] `public InitStylesheet()` is destructive, raw, and not obsoleted.** It assigns a
  brand-new `Stylesheet` (wiping any existing one), returns raw OpenXml, and mutates
  `_defaultStyle` as a side effect — and unlike the other raw members it carries no `[Obsolete]`.
  `Spreadsheet.cs:504-523`.
- **[Medium] The `Styles.*` layer is the least-guarded leak and is write-only.** All five wrappers
  (`Font`, `Fill`, `Border`, `Alignment`, `NumberingFormat`) expose a public non-obsolete
  `Element` of a raw `DocumentFormat.OpenXml.Spreadsheet.*` type, take OpenXml-typed constructors,
  and have **set-only** properties (state cannot be read back). This is on the everyday
  `CreateStyle` path. `Styles/*.cs`.
- **[Medium] The factory layer is dead plumbing.** The four factory interfaces + classes are now
  `internal` with **zero consumers** anywhere in `src/` (only `new X(...)` passthroughs that
  duplicate directly reachable constructors). The architecture doc already flagged them for
  removal; they were internalized but not deleted. `Factory/*`.
- **[Medium] "Getter"-named members with mutation side effects.** `Range.Rows` adds rows to the
  worksheet on read; `Row.CurrentCell` creates a cell on read; `Style.Get*Id` append to the
  shared stylesheet. Reads that mutate shared state are a debugging hazard.
  `Range.cs:22-24`, `Row.cs:14-16`, `Style.cs:156-251`.
- **[Low] `Utils.MergeFonts/MergeFills/MergeBorders` are public and take/return raw OpenXml**,
  guarded only by `EditorBrowsable(Never)` (not `[Obsolete]`). `Utils.cs:27-57`.

### C. Performance debt

**All four were measured on 2026-07-27 and are now pinned by CI guards** — see
[`excel-performance-baseline.md`](excel-performance-baseline.md) for the numbers and
[`test/OfficeDocuments.Excel.PerformanceTests`](../test/README.md) for the guards. The guards stop
them getting worse; they do not fix them. Fixing them is EXCEL-005, and the measurements argue for
this order: style dedup (worst absolute cost — 1 000 distinct styles is 3.1 s and 1.2 GB), then
comments (steepest growth — a clean 16× for 4× the input), then the row backfill, then the sort.

- **[Medium] Style creation is O(N²).** Every `Style` construction linearly scans fonts, fills,
  borders, numbering formats, and cell formats with no cache/index; the dedup guard `id <= 0`
  also treats a legitimate match at index 0 (equal to the default) as "not found" and appends a
  duplicate. `DataClasses/Style.cs:156-251`.
- **[Medium] Comments are O(N²) with repeated I/O.** `SetCellComment` regenerates the entire VML
  for all comments and calls `comments.Save()` on every single call.
  `DataClasses/Worksheet.cs:751-803,1000-1042`.
- **[Medium] `Row.CreateCell` backfills every missing cell up to the requested index** via an
  O(n) `InsertCell` (so O(n²) to fill a row); a single large column index can materialize a huge
  number of cells. `DataClasses/Row.cs:166-208`.
- **[Medium] `Range.SortByColumn` clones the OpenXml subtree of every cell** into snapshots and
  replaces them — O(rows×cols) DOM clones/allocations. `DataClasses/Range.cs:128-290`.

### D. Process / infrastructure debt

- **[High] CI is broken.** `github-build-excel.yml` and `github-build-word.yml` both run
  `dotnet restore OfficeDocuments.sln`, and `copilot-instructions.md` references the same name —
  but the repo tracks only `OfficeDocuments.slnx` (renamed 2026-05-31). Reproduced:
  `dotnet restore OfficeDocuments.sln` → `MSB1009: project file does not exist`. The restore step
  fails, so the whole workflow fails. Fix: point at `OfficeDocuments.slnx` (or drop the argument).
- **[Medium] Documentation has drifted from the code.** The planning docs (dated 2026-05-31) still
  describe the factory layer as public and list it as work to be done, but it is already
  `internal`. The consumer guide and architecture notes should be reconciled with the current
  code before they are trusted.
- **[Medium] Duplicate task IDs.** `tasks/core/excel/EXCEL-005/006/007` name different work than
  `tasks/advanced/excel-roadmap.md` EXCEL-005/006/007. The ID reuse makes the roadmap ambiguous.
- **[Low] Tests depend on the raw members slated for removal.** `CellTest`/`RowTest`/`StyleTest`
  assert against `Element`/`Stylesheet`, producing many `CS0618` obsolete-usage warnings and
  coupling test code to the very surface the project wants to drop. Leakage removal and test
  refactoring must land together.
- **[Low] Nullable warnings (`CS8602`) in the test project** — latent NRE risk in test helpers.
- **Test blind spots:** no coverage for culture-independence, concurrency/thread-safety, or scale
  (near Excel's row/column limits); thin coverage of null-argument contracts and the
  `DataValidationOptions`/`ConditionalFormattingOptions` guard clauses.

## What is genuinely good (keep it)

- **Broad, practical feature set that beats raw OpenXml ergonomics** — ranges, bulk insert,
  worksheet lifecycle, tables, validation, conditional formatting, images, protection.
- **Real round-trip tests.** Most tests write a real `.xlsx` and reopen it (often via raw
  `SpreadsheetDocument.Open`) to assert concrete nodes — far stronger than no-throw smoke tests.
- **The OpenXml-leakage cleanup is well-signposted.** Interface-level raw members are consistently
  `[Obsolete]` + `[EditorBrowsable(Never)]` — a clean, low-risk transitional pattern. Removing
  them later is a mechanical, well-marked step.
- **Style composition (`CreateStyle` + `CreateMergedStyle`) is a real value-add** the base SDK
  makes awkward.
- **Culture handling in the write path and the double/decimal read path is correct**
  (`InvariantCulture` + `OADate`) — a common wrapper bug that this library mostly avoids.
- **Solid engineering hygiene** — multi-targeting, central package management, nullable enabled,
  a dedicated `.docs/` discipline, PR/issue templates, and a genuine roadmap.

## Direction: three paths compared

The three paths are framed as the user posed them: keep the necessary minimum, keep adding
features, or split into core + extensions. They are **not mutually exclusive in time** — see
"How to choose".

### Path A — Maintain a hardened minimal core

*Freeze the Excel feature set; fix the debt; treat Excel as "done + maintained".*

| | |
| --- | --- |
| Scope | Fix the correctness bugs (section A), fix CI (D-1), finish the OpenXml leakage removal, delete the dead factory, and make the scope decision on the formula engine. No new features. |
| Pros | Lowest effort; Excel already covers most report/export scenarios; matches the stated positioning ("smaller and more predictable than ClosedXML/EPPlus"); smallest maintenance surface; frees energy for Word (where the real gap is). |
| Cons | No adoption growth from Excel; the formula-engine decision still has to be made; does nothing for the Word ambition on its own. |
| Cost | Low — roughly 1–2 focused sessions for the whole hardening pass. |
| Choose if | This is primarily a personal/internal tool that already does its job, and your appetite is "keep it correct and clean" rather than "grow it". |

### Path B — Keep adding features in one package

*Stay single-package; work the backlog (mostly Word, plus Excel import/export helpers, later charts).*

| | |
| --- | --- |
| Scope | Continue the feature backlog without a physical split. The P0 backlog is dominated by Word (run/paragraph formatting, tables, images, hyperlinks, headers/footers). |
| Pros | Directly increases usefulness and adoption; Word is where the biggest capability gap is; keeps packaging simple (one artifact, one version). |
| Cons | Ever-growing single surface is exactly what the minimal-core guidance warns against; heavy features (charts, pivots, templates) would bloat the dependency for consumers who only export tabular data; the god classes get worse unless refactored first. |
| Cost | Ongoing; each backlog item is M–L. |
| Choose if | You want this to become a broadly useful general-purpose library and are willing to invest continuously — and you keep the truly heavy features out until a split is justified. |

### Path C — Split into core + advanced (+ interop)

*Extract `OfficeDocuments.Excel` (minimal core), `OfficeDocuments.Excel.Advanced`, and optionally `OfficeDocuments.Excel.Interop`.*

| | |
| --- | --- |
| Scope | The split the repo's own architecture docs already recommend. Core keeps workbook/worksheet/row/cell/range/style; Advanced holds heavier features; Interop holds the raw-OpenXml compat surface so the core can finally drop the obsolete members. |
| Pros | Preserves a small, predictable core per the positioning; lets heavy features grow without bloating core; gives the raw-OpenXml compatibility members a clean home so the core surface can shrink; aligns with the existing roadmap. |
| Cons | Highest one-time cost (project split, per-package versioning, packaging, CI × N, docs); **premature today** — the current "advanced" features (images, validation, formatting, tables) are already cleanly typed with no OpenXml leakage (per the interface audit), so a split buys less separation than it looks like until there is real heavy-feature volume (charts, pivots, templates). |
| Cost | Significant one-time; then per-feature. |
| Choose if | You already have — or are committed to building — enough heavy Excel features (charts/pivots/templates) to justify a second package. The architecture doc itself says the split is "a guide, not an immediate physical split". |

### How to choose

- **Do the Path A hardening regardless.** The correctness bugs, the broken CI, the dead factory,
  and the formula-engine decision are debt under *every* path — none of them depend on the
  strategic choice. This is the cheapest, highest-value work available and it unblocks everything
  else.
- **The real strategic fork is B vs C, and it is a function of feature volume.** A split (C) with
  only two or three advanced features is packaging overhead for its own sake; a single package (B)
  that accumulates charts + pivots + templates becomes the bloated general-purpose library the
  positioning explicitly rejects.
- **This review's leaning (an opinion, not a directive):** harden now (A), keep the Excel feature
  scope frozen as the minimal core, and point new-feature energy at **Word** (that is where the
  backlog's P0 gap actually is). Defer the physical core/advanced split (C) until there are
  ≥2–3 heavy Excel features that genuinely justify a second package. In short: **A now → B on
  Word → C only when Excel advanced volume demands it.**

## Suggested near-term backlog (path-independent)

Ordered by value-to-effort. Everything here is worth doing under any of the three paths.

1. **Fix CI** — point the workflows and `copilot-instructions.md` at `OfficeDocuments.slnx`. (D-1)
2. **Decide the formula engine** — remove `GetFormulaValue()` and friends, or promote them to a
   real, documented, `double`-returning feature with row-aware ranges. As-is it is a correctness
   trap that contradicts the docs. (A + B-formula)
3. **Fix the verified correctness bugs** — row-aware formula ranges (if kept), empty-`MEDIAN`
   guard, `AddWorksheet` style drop, boolean `"1"/"0"`, culture-consistent int/long/bool reads.
4. **Delete the dead factory layer.** (B-factory)
5. **Reconcile the docs with the code** — factory status, dated planning notes, duplicate task IDs.
6. **Finish the OpenXml leakage removal** — annotate/hide the `Styles.*` `Element` getters and
   `Utils.Merge*`, and refactor the tests off the raw members in the same change.
7. **Break up `Worksheet` / `Spreadsheet`** before adding more features to them (partial classes
   at minimum; extracted collaborators ideally).
8. **Add the missing test dimensions** — culture, null-argument contracts, and the option guards.

## Progress log

### 2026-07-24 — hardening pass (Path A)

Delivered in this pass, all verified against the Excel test suite (174 passing, up from 168):

- **CI fixed** (D-1): both workflows and `copilot-instructions.md` now reference `OfficeDocuments.slnx`; `dotnet restore` verified to succeed.
- **Formula engine reworked** (B-formula; chosen direction: keep and fix properly): `ICell.GetFormulaValue()` now returns `double`, evaluates the range in two dimensions (fixing the row-blind `SUM`/`COUNT`/`COUNTIF`/`MEDIAN` bug), matches the exact function name (so `SUMIF` is no longer mistaken for `SUM`), guards empty `MEDIAN`, computes the median without integer truncation, throws `NotSupportedException` for unknown functions instead of `NotImplementedException`, and throws `InvalidOperationException` on a cell without a formula instead of returning the `-1` sentinel. Five regression tests were added and `excel-library.md` was updated.
- **`AddWorksheet` style drop fixed** (A): a `sheetStyle` is now applied even when the default style is not initialized (opened workbooks), via `?? sheetStyle`.
- **Boolean conformance fixed** (A): boolean cells now write `"1"`/`"0"` per OOXML; reads accept both `"1"`/`"0"` and legacy `"True"`/`"False"`.
- **Culture-consistent reads** (A): `GetIntValue`/`GetLongValue` and `TryGetValue(int|long)` now parse with `InvariantCulture`, matching the invariant write path.
- **`IRange.Worksheet` fixed** (A): `Range` now sets the base worksheet, so the property is non-null through the `IRange`/`IBase` interface and `OwnerWorksheet` no longer risks an invalid cast; regression test added.
- **Dead factory layer removed** (B-factory): the entire unused `Factory/*` layer was deleted and the planning docs were reconciled.

Still open from the backlog: reconcile the remaining dated planning notes and the duplicate core/advanced `EXCEL-005/006/007` task IDs; finish the `Styles.*` / `Utils.Merge*` leakage removal; break up the `Worksheet` / `Spreadsheet` god classes; and add culture/concurrency/scale test dimensions.

### 2026-07-27 — schema-validation gate (EXCEL-011 phase 1)

A test-suite analysis and restructuring proposal was written to
[tasks/core/excel/EXCEL-011-test-suite-restructuring.md](tasks/core/excel/EXCEL-011-test-suite-restructuring.md)
(unit / integration / verification / performance tiers, entry criteria per tier, a style-testing
deep-dive, and a 15-item blind-spot catalogue drawn from defects other OOXML libraries shipped).

Its phase 1 — an `OpenXmlValidator` gate on every test that produces a complete document — was
delivered and **immediately caught three real correctness bugs**, all now fixed: merged styles
violated the CT_Font/CT_Border child sequence (`Utils.MergeElements` appended instead of inserting
in order, masked by order-insensitive dedup); `Font.ArgbHexColor` / `Fill(string, …)` wrote
un-normalized colour strings into the `hexBinary` `rgb` attribute; and `workbookProtection` was
appended after `sheets` — **the workbook child-order bug predicted in section A above**, now fixed
via a new `WorkbookElementOrderer` that also closes the same latent defect in `AddNamedRange`.
Suite 174 → 187 tests, green on all three TFMs. Details and follow-ups in the EXCEL-011 progress log.
