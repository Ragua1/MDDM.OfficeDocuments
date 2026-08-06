# EXCEL-011 Test suite restructuring (unit / integration / verification / performance)

Date: 2026-07-27

## Status

- Phase 1 (schema-validation gate): **Delivered 2026-07-27.** Found and fixed three real
  correctness bugs. See the progress log at the end.
- Phase 2 (`TestKit` + `UnitTests`): **Delivered 2026-07-27.** 92 new unit tests; the tier runs in
  ~55 ms.
- Phase 3 (integration / verification split): **Delivered 2026-07-27.** Three tiers now exist;
  the integration tier no longer touches disk and got ~10× faster.
- Phase 4 (style deep-dive): **Delivered 2026-07-27.** 73 new style tests; found and fixed two
  more correctness bugs.
- Phase 5 (performance): **Delivered 2026-07-27.** Benchmarks project + 14 CI guards; the four
  known hot spots are quantified for the first time in
  [`excel-performance-baseline.md`](../../../excel-performance-baseline.md).
- Phase 6 (blind spots B-2..B-6): **Delivered 2026-07-28.** Four confirmed defects fixed, three of
  them invisible to every gate the suite had. B-7..B-15 remain open.

## Business goal

`OfficeDocuments.Excel.Tests` is a single project holding 174 tests of four genuinely different
kinds, all bound to the same base class and all writing real files to disk. The mixture hides
three problems: there is no fast pure-logic tier to run on every build, there is no tier that
verifies the *document* (only tiers that verify API results), and there is no tier at all for
performance. The goal is to split the suite along the axis of **what a failure means**, define
entry criteria per tier so future tests land in the right place, and close the coverage gaps that
the current single-tier design structurally cannot express.

This is a direct follow-up to two items left open by
[../../../excel-state-verdict.md](../../../excel-state-verdict.md): *"add the missing test
dimensions — culture, null-argument contracts, and the option guards"* and the four unaddressed
O(N²) hot spots in section C, none of which currently has a test that would notice a regression.

## Current state

### Inventory

174 tests execute (153 declared methods; `CellExtensionTest` expands via 29 `[InlineData]`),
across 12 files, ~3,700 lines. Runtime: **1 second** for the whole suite on `net9.0`, 0 failed,
0 skipped. Stack: xUnit 2.9.3 + `coverlet.collector`, no other test dependencies.

| File | Tests | What it actually is | Disk I/O |
| --- | --- | --- | --- |
| `CellExtensionTest.cs` | 8 (29 cases) | **True unit tests** — A1↔index math, range parsing | none |
| `UtilsTest.cs` | 4 | Unit-ish, but asserts on raw OpenXml `Merge*` helpers slated for removal | none |
| `CellTest.cs` | 50 | Cell value/formula behaviour — but every test creates a real `.xlsx` | 50 files |
| `RowTest.cs` | 15 | Row/cell creation behaviour | 15 files |
| `WorksheetTest.cs` | 8 | Worksheet cell/range lookups | 8 files |
| `StyleTest.cs` | 12 | Style id allocation and merging | 14 files |
| `WriterTest.cs` | 9 | Document lifecycle (create/open/stream) | 7 files |
| `RangeAndAdvancedFeaturesTest.cs` | 31 | Ranges, tables, validation, CF, images — mixed API + raw XML asserts | 33 files |
| `ReaderTest.cs` | 8 | **Round-trip verification** — write, close, reopen, assert | 8 files |
| `CreationTest.cs` | 8 | Large realistic workbooks — but almost assertion-free | 7 files |

### Findings

**F-1 — There is no tier boundary, only a naming coincidence.** `SpreadsheetTestBase` gives every
test class a temp workspace and every test a `.xlsx` on disk. 143 `GetFilepath()` calls against 12
`MemoryStream` uses. `CellTest` is named and written like a unit test but is an integration test
by construction: it cannot fail without the OpenXml package layer being involved. The suite is
fast today (1 s) only because every file is tiny — this is not a speed problem yet, it is an
*intent* problem: nothing tells a contributor where a new test belongs.

**F-2 — Zero schema validation.** `OpenXmlValidator` appears nowhere in the suite. The library can
emit a workbook that Excel refuses to open and every test would still pass. This is not
hypothetical for this codebase: the verdict document already flags *"`AddNamedRange` /
`ProtectWorkbook` append `DefinedNames` / `WorkbookProtection` at the end of `Workbook`, which can
violate the `CT_Workbook` child sequence"* and *"`CopyWorksheet` can leave dangling
relationships"*. Both are exactly the class of defect a validator pass catches for free, and
neither has a test. The Open XML SDK itself shipped a bug producing invalid files past 26 columns
([dotnet/Open-XML-SDK#440](https://github.com/OfficeDev/Open-XML-SDK/issues/440)) — schema-order
bugs are the normal failure mode of this domain, not an exotic one.

**F-3 — No golden/reference files.** ClosedXML keeps a `Resource` directory of reference `.xlsx`
files and compares generated output against them, updating a reference only after inspecting the
diff visually in Excel and as XML
([CONTRIBUTING.md](https://github.com/ClosedXML/ClosedXML/blob/develop/CONTRIBUTING.md)). This
repo has the machinery — `Extensions/XElementExtensions.CompareXml` does namespace- and
order-normalised XML equality — but it is used in exactly **one** test
(`MergeStylesAcrossWorkbooks_UsesTargetStylesheetIds`).

**F-4 — No performance tests of any kind.** No BenchmarkDotNet, no timing assertions, no
allocation assertions. The four known quadratic hot spots (style dedup, comment VML regeneration
with a `Save()` per call, `Row.CreateCell` backfill, `Range.SortByColumn` DOM cloning) have no
guard, so neither a regression nor a *fix* can be demonstrated.

**F-5 — Assertion-free tests inflate the count.** `CreationTest.BasicFile` has an empty `using`
body (literally `;`). `CreationTest.CustomFile1/2/3` build 150-line realistic workbooks and assert
nothing about them. `WriterTest.CreateDocumentToStream` asserts nothing. These are "does not
throw" smoke tests counted as coverage — and ironically they are the scenarios best suited to
become the verification tier, because they already build realistic documents.

**F-6 — Heavy duplication where parameterisation belongs.** `CellTest` is 1,125 lines for 50
tests, largely because `SetInteger` / `SetDouble` / `SetLong` / `SetDecimal` / `SetDate` each
exist twice (with and without style) as near-identical copies, and `CellSetAndGet*Value` repeats
the same shape six more times. Only `CellExtensionTest` uses `[Theory]`. A typed matrix would cut
this file by well over half while *increasing* the number of cases.

**F-7 — Weak style assertions.** Every `StyleTest` case asserts `Assert.True(s.FontId > 0)` — that
*an* id was allocated, not that the produced font is correct. Nothing asserts the rendered
`styles.xml` fragment, so a wrong colour, a swapped bold/italic flag, or a mis-encoded ARGB would
pass. See "Style testing" below.

**F-8 — The read path only ever reads files this library wrote.** `ReaderTest` is eight
write-then-reopen round-trips. `Resources/Example_1.xlsx` is used only as bytes to materialise a
file which is then opened; `Resources/Example_2.xlsx` is tracked in git but **referenced by
nothing**. There is no test against a workbook produced by Excel, LibreOffice, or Google Sheets,
and no test against malformed input.

**F-9 — Hygiene.** `TestResults/` is not in `.gitignore` (currently untracked, but it will
accumulate). The test project produces `CS0618` obsolete-usage warnings because tests assert
against `Element` / `Stylesheet` — the very surface the library wants to drop.

## Proposed tier model

The split axis is **what a failing test tells you**, which in turn fixes the dependency each tier
is allowed to touch. Four tiers, plus a shared kit.

| | Unit | Integration | Verification | Performance |
| --- | --- | --- | --- | --- |
| Question answered | "is this function correct?" | "do these types behave correctly together through the public API?" | "is the produced file a valid, correct Excel document?" | "is it fast enough, and still the right complexity class?" |
| May touch | pure types only | `SpreadsheetDocument` over `MemoryStream` | a complete `.xlsx` byte array | either |
| Must not touch | any OpenXml package, any stream, the file system | the file system | — | — |
| Assertion target | return values, exceptions | object model + specific OpenXml DOM nodes | validator verdict, package integrity, golden XML, third-party reader | duration ratios, allocated bytes |
| Test style | `[Theory]`-heavy, table-driven | one feature per test, targeted node asserts | realistic whole-document scenarios | N vs k·N scaling pairs |
| Budget | tier < 200 ms; each test < 10 ms | tier < 10 s; each test < 100 ms | tier < 60 s | out of band |
| Runs | every build, pre-commit | every build | every CI push | nightly + on demand |
| Failure means | a rule is wrong | a collaboration is wrong | the file is wrong | a promise is broken |

### Why this axis and not "fast/slow"

Tiering by speed is tempting but useless here — the whole suite is 1 second. Tiering by *blast
radius of a failure* is what makes the tiers actionable: a unit failure points at one function, a
verification failure means "users get a corrupt file", and those two deserve different urgency,
different reviewers, and different CI gates. This also matches the standard .NET guidance that
integration tests stay in a separate project precisely so the unit project carries no
infrastructure dependency
([SSW](https://www.ssw.com.au/rules/follow-naming-conventions-for-tests-and-test-projects),
[Microsoft](https://learn.microsoft.com/en-us/dotnet/core/testing/unit-testing-best-practices)).

### Entry criteria (the rule a contributor applies)

**Unit** — the test compiles without referencing `DocumentFormat.OpenXml.Packaging`. If you need
a `Spreadsheet` to write the test, it is not a unit test.

Eligible surface today, and it is larger than it looks because EXCEL-010 just created nine
internal collaborators:

- `Extensions/CellExtension` — column name↔index, cell reference parse/format, range parsing,
  and the missing Excel bounds checks (XFD / 1048576).
- `Utils.ColorConverter` — `System.Drawing.Color` → ARGB hex.
- `Styles/NumberingFormat.DefaultNumberFormats` — built-in id lookup, custom-format allocation.
- `Options/*` guard clauses — `DataValidationOptions`, `ConditionalFormattingOptions`,
  `TableCreateOptions`, `TableStyleOptions`. Currently untested per the verdict doc.
- `WorksheetElementOrderer` — the `CT_Worksheet` child-order ladder. This is pure DOM logic over a
  `Worksheet` element and needs no package; it is the single highest-value new unit target
  because every advanced feature depends on it.
- `Enums/*` → OOXML value mapping, including the round-trip of every enum member.

Requires `[assembly: InternalsVisibleTo("OfficeDocuments.Excel.UnitTests")]`.

**Integration** — exercises two or more collaborating types through the **public** API, over a
`MemoryStream`. Assert on the object model and on the specific DOM nodes you care about — never on
a whole-document snapshot, so the test does not break when an unrelated part of the output
changes. No `GetFilepath`, no temp directory.

**Verification** — produces a complete document and then interrogates it as a *document*. Every
verification scenario must pass four gates:

1. **Schema** — `new OpenXmlValidator(FileFormatVersions.Office2021).Validate(doc)` yields zero
   errors. Non-negotiable, and the gate that closes F-2.
2. **Package integrity** — every `r:id` resolves to an existing part; every part has a
   `[Content_Types].xml` entry; no orphan parts. Closes the `CopyWorksheet` dangling-relationship
   risk.
3. **Round-trip** — reopen through the library and read every value back.
4. **Golden** — normalised XML of the interesting parts compared against a committed reference,
   with ClosedXML's discipline: a reference is only regenerated after the diff has been inspected
   and the file opened in Excel.

  Gate 4 needs **output determinism** first: `docProps/core.xml` timestamps and any GUID-ish
  relationship-id churn must be normalised out, or golden comparison is unusable. Worth resolving
  as part of this work — determinism also means "open + save with no changes produces the same
  bytes", which is itself a valuable regression test.

  Optional fifth gate, cheap and very strong: **read the generated file with a different
  implementation** (reference ClosedXML or NPOI in the verification project only). If a foreign
  reader agrees with the expected values, the file is genuinely well-formed rather than merely
  self-consistent. This directly closes F-8.

**Performance** — see below; two sub-kinds that must not be conflated.

## Target layout

```
test/
  OfficeDocuments.Excel.TestKit/            # shared, no tests
    XElementExtensions.cs                   # moved from Tests/Extensions
    OpenXmlValidation.cs                    # AssertValid(doc), AssertPackageIntact(doc)
    WorkbookFactory.cs                      # in-memory builders, replaces SpreadsheetTestBase
    TempWorkspace.cs                        # today's TestSettings, only for tiers that need disk
    Scenarios/                              # the realistic documents, shared by verification + perf
  OfficeDocuments.Excel.UnitTests/
  OfficeDocuments.Excel.IntegrationTests/
  OfficeDocuments.Excel.VerificationTests/
    Golden/                                 # committed reference .xlsx + expected XML fragments
    Inputs/                                 # foreign-producer workbooks (Excel/LibreOffice/Sheets)
  OfficeDocuments.Excel.Benchmarks/         # BenchmarkDotNet console app, not a test project
```

Naming follows the `[Project].UnitTests` / `[Project].IntegrationTests` convention rather than the
current bare `.Tests`, so `dotnet test --filter` is unnecessary — CI selects tiers by project.
`Word` should get the same shape later; keeping the names symmetrical now avoids a second rename.

An alternative to four projects is one project with xUnit `[Trait("Tier", "...")]` filters. It is
rejected: traits do not prevent a "unit" test from taking a dependency on the packaging layer, and
the dependency *is* the point of the split.

### Migration mapping

| Current | Goes to | Notes |
| --- | --- | --- |
| `CellExtensionTest` | Unit | Moves as-is. Extend with bounds cases (F-2 adjacent). |
| `UtilsTest` | Unit | Only the `ColorConverter` case survives long-term; the `Merge*` cases die with EXCEL-007. |
| `CellTest` | Integration | Collapse the per-type duplication into typed `[Theory]` matrices (F-6). |
| `RowTest`, `WorksheetTest` | Integration | Switch `GetFilepath` → `MemoryStream`. |
| `StyleTest` | Integration + Unit | Id-allocation and merge-truth-table logic → unit where possible; rendered-XML asserts → integration. |
| `WriterTest` | Integration, except 2 | `NotCreateFileOnNonExistDirectory` and the file-path lifecycle cases genuinely need disk — keep those on `TempWorkspace`. |
| `RangeAndAdvancedFeaturesTest` | Split | Behaviour asserts → integration; the eight `SpreadsheetDocument.Open` raw-XML cases → verification. 637 lines / 31 tests is already too big for one file; split by feature. |
| `ReaderTest` | Verification | Already the right shape; add the validator gate. |
| `CreationTest` | Verification | The realistic workbooks become the golden scenarios and gain the assertions they never had (F-5). `BasicFile` is deleted. |

## Style testing

Called out separately because the current coverage is the weakest part of the suite (F-7) and
because styles are where this library claims its main value-add over the raw SDK.

What is missing, in rough priority order:

1. **Rendered-output assertions.** For each style facet, assert the normalised `styles.xml`
   fragment via `CompareXml`, not the allocated id. A test that only checks `FontId > 0` cannot
   distinguish `Color.Blue` from `Color.Red`.
2. **The dedup identity matrix.** Creating the same style twice must yield the same `StyleIndex`;
   changing exactly one attribute must yield a different one; creating N distinct styles must
   produce exactly N + defaults entries. This is also the regression test for the verdict doc's
   *"the dedup guard `id <= 0` treats a legitimate match at index 0 as not-found and appends a
   duplicate"* — a style equal to the default is the specific missing case.
3. **A merge truth table.** Four ad-hoc merge tests exist today. Replace with a `[Theory]` over
   {font, fill, border, numberFormat, alignment} × {base only, overlay only, both, neither},
   asserting which side wins per facet. `CreateMergedStyle` is the API's signature feature and its
   semantics are currently pinned by example rather than by rule.
4. **The inheritance chain.** Sheet style → row style → cell style precedence, including partial
   overrides (row sets a font, cell sets a fill, both must survive) and explicit-null-vs-unset.
   Two tests cover this today and neither tests a conflict.
5. **Colour encoding.** `System.Drawing.Color` → ARGB round-trip, alpha handling, the
   `ArgbHexColor` string path with and without a leading `#`, invalid input, and `Color.Empty`.
   Theme colours and indexed colours (`<color theme="1"/>`, `<bgColor indexed="64"/>`) are a known
   source of confusion in every OOXML implementation
   ([officeopenxml.com](http://officeopenxml.com/SSstyles.php)) — at minimum the library's
   position on them should be pinned by a test.
6. **Number formats.** Built-in id mapping (`"@"` → 49) vs custom allocation from 170, per-workbook
   independence (already tested — keep), format codes containing quotes/brackets/escapes, and
   culture-dependent codes.
7. **Borders per edge.** Each edge independently, per-edge style, per-edge colour, diagonal.
   Today: one all-edges case and one two-edge case.
8. **Differential formats (`dxf`).** Conditional formatting allocates dxf entries and the
   `Spreadsheet` deduplicates them; that dedup path has no test at all.
9. **Cross-workbook merge** — already covered by one good test; extend it to fills/borders/
   alignment (currently font + fill only).

## Performance testing

Two distinct kinds. Conflating them is the usual mistake — one is a measurement instrument, the
other is a CI gate.

### Benchmarks — `OfficeDocuments.Excel.Benchmarks`

A BenchmarkDotNet console app, Release-only, **not** run by `dotnet test`. It answers "did this
change make things faster or slower" with ns/op, allocated bytes, and gen0/1/2 counts. This is the
standard shape — Microsoft's own MSAL.NET perf project is a BenchmarkDotNet console app for
exactly this reason
([docs](https://learn.microsoft.com/ko-kr/entra/msal/dotnet/advanced/performance-testing)).
Export JSON, keep a baseline, diff in a nightly job
([approach](https://amarozka.dev/extending-benchmarkdotnet-exporters-metrics-ci-cd/)).

Benchmark targets, taken straight from the verdict document's section C so that the fixes become
demonstrable:

- `CreateStyle` with 100 / 1,000 / 5,000 distinct styles (the O(N²) linear scan).
- `SetCellComment` with 10 / 100 / 500 comments (full VML regeneration + `Save()` per call).
- `Row.CreateCell` at a far column index (the O(n²) backfill).
- `Range.SortByColumn` over 1k / 10k rows (DOM subtree cloning).
- `AddRows<T>` bulk import, 10k / 100k rows.
- Whole-document write, 100k rows × 20 columns, and the reopen of the same.

### Perf guards — in `VerificationTests` or their own project

A small number of xUnit tests that **do** run in CI, with deliberately generous ceilings. Two
assertion strategies, both machine-independent, which wall-clock alone is not:

- **Complexity assertions.** Measure at N and 4N and assert the ratio, e.g.
  `t(4N) / t(N) < 8` for something claimed to be near-linear. This catches "someone reintroduced a
  nested scan" without depending on how fast the CI runner is. This is the right guard for all
  four quadratic hot spots.
- **Allocation ceilings.** `GC.GetAllocatedBytesForCurrentThread()` around the operation is
  deterministic and machine-independent — far more stable than time as a CI gate.

Plus one **scale ceiling** test that documents the practical limit rather than asserting speed:
ClosedXML throws `OutOfMemoryException` around 400k rows/sheet and EPPlus around 1M
([comparison](https://hackernoon.com/c-excel-library-in-depth-comparison-tested-for-2026),
[ClosedXML#818](https://github.com/ClosedXML/ClosedXML/issues/818)). This library is DOM-based over
the OpenXml SDK and will have a similar ceiling. A test that pins "100k rows × 20 columns
completes within X MB" turns an unknown into a documented contract — and tells you when a future
change moves the ceiling.

## Blind spots — gaps this suite does not currently express

Compiled from the failure modes that other OOXML implementations actually shipped. Ordered by
expected value.

**B-1 through B-6 are closed** (phases 1 and 6); the entries are kept because they record what the
defect was and how it was found. B-7 through B-15 are still open — B-7 (culture) and B-9 (foreign
producers) are the two with the most left in them.

**B-1 — Schema validity.** Covered above (F-2). Highest value single change in this document: one
helper, called from every verification scenario.

**B-2 — Sheet-name legality.** Verified absent from the source: `AddWorksheet` checks uniqueness
via `EnsureWorksheetNameAvailable` but performs **no** length or character validation. Excel
requires ≤ 31 characters and forbids `: \ / ? * [ ]`. Every major library has shipped a bug here —
[exceljs#1474](https://github.com/exceljs/exceljs/issues/1474) (`/` and `:` silently produce
"Sheet 1"), [openxlsx#211](https://github.com/ycphs/openxlsx/issues/211),
[ImportExcel#362](https://github.com/dfinke/ImportExcel/issues/362). Currently this library will
happily write a 40-character sheet name and produce a file Excel repairs or rejects.

**Closed in phase 6.** Confirmed exactly as described: a 40-character name and every forbidden character produced a schema-invalid workbook. `WorksheetNameValidator` now applies the rules on create and on rename.

**B-3 — XML escaping.** `&`, `<`, `>`, `"`, `'` in cell values, sheet names, defined names, table
names, comment text, and hyperlink tooltips. openxlsx corrupted files outright because sheet names
were written into `workbook.xml` unescaped
([openxlsx#518](https://github.com/ycphs/openxlsx/issues/518)). The SDK escapes for you on the
typed paths, but anywhere this library builds XML as a string — and `CommentWriter`'s VML
generation does exactly that — the guarantee is gone. `Worksheet.EscapeFormulaString` exists,
which implies the author already met this problem once.

**Closed in phase 6 — and it was already correct.** The SDK escapes every typed path, and the VML `CommentWriter` builds by hand turns out to carry only numeric anchors, never caller text. Nothing to fix; tests added so it cannot regress silently.

**B-4 — Illegal XML control characters.** `0x00`–`0x08`, `0x0B`, `0x0C`, `0x0E`–`0x1F` are not
representable in XML 1.0 and must be stripped or `_x####_`-escaped. Real libraries produce
`Removed Part: /xl/sharedStrings.xml` errors over this
([libxlsxwriter#276](https://github.com/jmcnamara/libxlsxwriter/issues/276)). No test, no guard.

**Closed in phase 6.** The SDK did refuse, but only when serializing, so the whole document was lost to an exception naming neither sheet nor cell. `XmlText` checks at the point of assignment instead.

**B-5 — Numeric edge cases.** `double.NaN` and `±Infinity` are **not valid** OOXML numeric cell
values; verified there is no `NaN`/`IsInfinity` check anywhere in the source, so they will be
written and produce a corrupt file. Also: round-trip precision (`"R"`/G17 formatting),
`decimal`→`double` precision loss, `double.MaxValue`, negative zero, and values exceeding the
32,767-character cell text limit
([Excel limits](https://support.microsoft.com/en-us/excel/excel-specifications-and-limits)).

**Closed in phase 6 for the non-finite half**, which was the corrupting one — and which the schema validator does not catch, because `v` is declared as a string. Round-trip precision, `decimal`→`double` loss and the 32,767-character limit remain open.

**B-6 — Date edge cases.** The 1900 leap-year bug (serial 60 = the non-existent 29 Feb 1900) means
every date before 1 Mar 1900 is ambiguous between producers
([Eric White](https://www.ericwhite.com/blog/dates-in-spreadsheetml/)). The library uses `OADate`,
which agrees with Excel above that boundary and diverges below it. Also untested: `DateTime.Min/
MaxValue`, `DateTimeKind`, time-only values, negative serials, and the 1904 date system (which the
library appears not to support — that is fine, but it should be a documented, tested position).

**Closed in phase 6.** Confirmed: every date before 1 March 1900 was written one day late, invisibly, because `ToOADate` and `FromOADate` are exact inverses. `ExcelSerialDate` replaces both. The 1904 system stays unsupported, now as a documented position.

**B-7 — Culture.** Partially known already, but two specific cases are worth naming.
`CurrentCulture = de-DE` exercises decimal-separator divergence on the read path.
`CurrentCulture = tr-TR` is the more dangerous one: the dotless-i breaks
`ToUpper()`/`ToLower()`-based comparisons, and this library does case-insensitive sheet-name
lookups and exact function-name matching in `GetFormulaValue`. Best implemented as a fixture that
swaps the culture around the *existing* integration tier, not as a handful of bespoke tests.

**B-8 — Excel limits enforcement.** The verdict doc notes reference parsing accepts columns beyond
XFD and rows beyond 1,048,576. Add: cell text > 32,767 chars, defined-name length, and sheet-count
limits. The library should either reject or document; today it silently produces an invalid file.

**B-9 — Foreign-producer input.** The read path has never seen a file it did not write (F-8).
Commit small workbooks saved by Excel, LibreOffice Calc, and Google Sheets into
`VerificationTests/Inputs/` and read them. This is where "we only support our own dialect" bugs
surface — inline strings vs shared strings, `sheetData` without `dimension`, styles referencing
themes, `r` attributes omitted on cells.

**B-10 — Malformed input robustness.** Truncated zip, valid zip with no `workbook.xml`, `.xls`
renamed to `.xlsx`, password-protected file, and a part declared in `[Content_Types].xml` but
missing. Expected behaviour is a clear typed exception, not an `IndexOutOfRange` or a hang.

**B-11 — Lifecycle and resource release.** Double `Dispose`, use-after-`Dispose`, `Dispose` without
`Close`, and — Windows-specific and easy to get wrong — that the file handle is actually released
so the file can be deleted or reopened immediately afterwards. `Spreadsheet` has a finalizer per
EXCEL-010; none of this is tested.

**B-12 — Concurrency.** `Worksheet` holds a `static ConcurrentDictionary<Type, PropertyInfo[]>
PropertyCache` — shared mutable static state across all instances. The type is thread-safe, so the
realistic risk is low, but the *contract* is undocumented and untested. The valuable test is not
"is `Spreadsheet` thread-safe" (it is not, and should not claim to be) but "two `Spreadsheet`
instances used concurrently on different threads do not interfere" — that pins the supported
usage. Pair it with a documented statement in the consumer guide.

**B-13 — Idempotent re-save.** Open a document, save with no changes, and assert the output is
schema-equal to the input. Catches the family of bugs where each open/save cycle duplicates a
part, appends an empty element, or grows the stylesheet. Also a prerequisite for golden files.

**B-14 — Formula surface.** Now that `GetFormulaValue` is a real feature: cross-sheet references
(`Sheet2!A1`), sheet names requiring quotes (`'My Sheet'!A1`), shared and array formulas,
`#REF!`/`#DIV/0!` error values, and circular references. Currently only same-workbook,
same-worksheet ranges are exercised.

**B-15 — Null-argument contracts.** Flagged in the verdict doc, unchanged: the option objects'
guard clauses and the public API's null handling are thinly covered. Cheap, mechanical, and
belongs entirely in the unit tier.

## Suggested phasing

Each phase is independently valuable and leaves the suite green.

1. **Add the validator gate to the existing project.** No restructuring. Write
   `OpenXmlValidation.AssertValid`, call it at the end of the eight `ReaderTest` round-trips and
   the four `CreationTest` workbooks. This alone closes B-1 and will likely find the two
   child-order bugs the verdict doc predicts. Highest value-to-effort in this document.
2. **Create `TestKit` + `UnitTests`.** Move `CellExtensionTest` and `XElementExtensions`, add
   `InternalsVisibleTo`, and write the new unit coverage for `WorksheetElementOrderer`, the option
   guards, and B-15. This is where the EXCEL-010 collaborators finally pay off.
3. **Split integration from verification.** Move the projects, switch integration to
   `MemoryStream`, give `CreationTest`'s workbooks real assertions, delete `BasicFile`, add
   `Example_2.xlsx` to a scenario or delete it.
4. **Style deep-dive.** Sections 1–4 of "Style testing" (rendered output, dedup matrix, merge truth
   table, inheritance chain). Do this before EXCEL-005's style-pipeline performance work, so the
   optimisation has a correctness net under it.
5. **Performance.** Benchmarks project first (measurement), then the perf guards (gating) once a
   baseline exists.
6. **Blind spots by priority.** B-2 through B-6 are all small, self-contained, and each is a real
   defect today rather than a hypothetical.

Phases 1, 2 and 4 are worth doing under any of the strategic paths in the verdict document. Phase
3 is the prerequisite for keeping `Word` tests symmetrical when Word grows.

## Open decisions

- **Golden files: yes or no?** They are the strongest verification gate and the most maintenance.
  They require output determinism (timestamps, relationship ids) to be solved first. A defensible
  middle position is: golden files for `styles.xml` and `workbook.xml` only, targeted XML asserts
  everywhere else.
- **Third-party reader in the verification tier.** Referencing ClosedXML from a test project is a
  test-only dependency and a very strong signal, but it is a competitor's library in this repo's
  own benchmark report. Cheap enough to be worth it; the objection is aesthetic.
- ~~**Perf guards in CI, or nightly only?**~~ **Resolved in phase 5: per-push CI.** Nothing
  asserts on a duration, so the runner's speed cancels out of every comparison. Measured
  reproducibility on the development machine is ±3% run to run.
- ~~**Four projects or three?**~~ **Resolved in phase 5: four.** Perf guards got their own
  project rather than living in `VerificationTests`, for three reasons that only became clear once
  they existed. They need `TieredCompilation=false` and xUnit parallelism off, neither of which
  should be forced on a correctness tier. They run for about 40 s, which would triple the
  verification tier's budget. And they are single-TFM, where the correctness tiers multi-target.

## Progress log

### 2026-07-27 — Phase 1: schema-validation gate (closes B-1 / F-2)

`test/OfficeDocuments.Excel.Tests/Validation/OpenXmlValidation.cs` validates a finished workbook
against `FileFormatVersions.Office2021` and fails with a readable per-error report (part URI +
XPath). Wired into all 8 `ReaderTest` round-trips (twice in the one that reopens and modifies) and
all 8 `CreationTest` workbooks. Suite: 174 → **187 tests, green on net8.0/net9.0/net10.0**.

The gate immediately failed 5 tests, exposing **three real correctness bugs** plus one fixture
defect. All three are fixed:

- **[High] Merged styles violated the CT_Font / CT_Border child sequence.** `Utils.MergeElements`
  appended a merged-in child at the end of the base element. Whenever the overlay contributed a
  child that must sort *before* one the base already had, the result was schema-invalid — e.g.
  merging a bold onto a size-only font produced `<sz/><b/>`, which Excel repairs or rejects.
  Fixed by sorting merged children into the declared sequence (`Utils.ApplySchemaChildOrder`, with
  explicit order tables for CT_Font, CT_Border and CT_PatternFill).
  *Why this survived 174 tests:* `Utils.OpenXmlElementsEqual` compares children **order-insensitively**,
  so style dedup silently substituted an existing correctly-ordered font whenever one existed. The
  defect only escaped to disk when the mis-ordered element was the first of its combination — which
  is why it showed up on realistic multi-style workbooks and not on focused style tests.
- **[High] `Font.ArgbHexColor` and `Fill(string, …)` wrote the user's string verbatim** into the
  `rgb` attribute, which is typed `hexBinary`. The documented-looking `new Font { ArgbHexColor =
  "#2A66FF" }` produced `rgb="#2A66FF"` — an invalid file. Fixed with `Utils.NormalizeArgbHex`:
  strips `#`, upper-cases, expands 6-digit RGB to 8-digit ARGB, and throws `ArgumentException`
  on anything else instead of emitting a corrupt document.
- **[Medium] `workbookProtection` was appended at the end of `Workbook`** but CT_Workbook requires
  it before `sheets`. This is the bug predicted in
  [../../../excel-state-verdict.md](../../../excel-state-verdict.md) section A ("workbook
  child-order risk"). Fixed with a new `WorkbookElementOrderer` (same pattern as the existing
  `WorksheetElementOrderer`), which inserts before the first child that must follow. The same
  helper also fixes `NamedRangeManager`, whose `definedNames` append would have landed after
  `calcPr` on any workbook opened from Excel — latent, and now closed.

Not a library bug: `Resources/Example_1.xlsx` itself carries `pageSetup/@verticalDpi="0"`, which
the schema forbids. Real-world Excel files are not always schema-clean, so `AssertValid` takes an
`inheritedDefects` parameter to tolerate named defects that arrived with a foreign input rather
than switching the gate off for that test.

Regression tests added (13): merged-font child order, ARGB normalization (4 valid + 5 invalid
cases across `Font` and `Fill`), and workbook child order for protection + named ranges.

Follow-ups this surfaced, recorded but not actioned:

- `Utils.OpenXmlElementsEqual` being order-insensitive is right for dedup but means the stylesheet
  can hold two elements this library considers equal and the schema does not. Worth an explicit
  test when the style deep-dive (phase 4) lands.
- `Spreadsheet.Close()` saves before checking `_disposed`, so a second `Close()` throws on a
  disposed document. Harmless today, but it belongs in the B-11 lifecycle tests.

### 2026-07-27 — Phase 2: `TestKit` + `UnitTests` (closes F-1 for the unit tier)

Two new projects, registered in `OfficeDocuments.slnx` and in the Excel CI workflow (which now
runs the unit tier first, since its failures point at a single function):

- **`test/OfficeDocuments.Excel.TestKit`** — shared helpers, not a test project (no test SDK, no
  `[Fact]`, `IsTestProject=false`). Holds `Validation/OpenXmlValidation` (from phase 1),
  `XElementExtensions.CompareXml`, and `TempWorkspace` (the former `TestSettings`, renamed to say
  what it is). Referenced by the integration tier only.
- **`test/OfficeDocuments.Excel.UnitTests`** — the unit tier. `InternalsVisibleTo` was added to the
  library so the EXCEL-010 collaborators can be tested directly. The project **deliberately does
  not reference `TestKit`**, so the temp-file workspace and the validator are unreachable from it —
  that is the structural half of the entry rule.

`test/README.md` states the entry criteria per tier, so a contributor no longer has to infer where
a new test belongs (finding F-1).

Coverage added — 92 new tests, tier total **121 tests in ~55 ms** on net8.0/net9.0/net10.0
(budget was < 200 ms). The integration tier is 158 (174 + 13 from phase 1, minus the 29 moved).

| Target | Tests | Why it matters |
| --- | --- | --- |
| `WorksheetElementOrderer` | 8 | Pure CT_Worksheet ordering logic that **every** advanced worksheet feature depends on, and previously had no direct test. Includes "insert in reverse schema order, still comes out in schema order". |
| `WorkbookElementOrderer` | 9 | Regression cover for the phase 1 workbook-order bug plus its latent `definedNames`/`calcPr` twin. |
| `DataValidationOptions` | 26 | Guard clauses were listed as untested in the verdict doc. Covers quote escaping, blank filtering, the between-operator formula2 requirement. |
| `ConditionalFormattingOptions` | 17 | Same; uses a minimal `IStyle` stub, since the factories only store the reference. |
| `NumberingFormat` | 20 | Built-in id lookup (incl. its case-sensitivity), custom-id allocation from 170, `General` fallback. |
| `Utils` colour handling | 12 | Pure-function half of the phase 1 ARGB fix, with the invalid inputs that used to reach the `rgb` attribute. |

`CellExtensionTest` moved across unchanged as `CellExtensionTests` (it was already a true unit
test — the only one in the old project).

One test failure during development was mine, not the library's: `Color.Transparent` is
ARGB(0, 255, 255, 255), not a zeroed colour. The corrected assertion now documents that trap.

### 2026-07-27 — Phase 3: integration / verification split (closes F-1, F-5, F-8; part of B-9)

The Excel suite is now three tiers. `OfficeDocuments.Excel.Tests` was renamed to
`OfficeDocuments.Excel.IntegrationTests` (the name finally matches the content) and
`OfficeDocuments.Excel.VerificationTests` was created. Both are in `OfficeDocuments.slnx` and in
the Excel CI workflow, which runs unit → integration → verification.

| Tier | Tests | Runtime |
| --- | --- | --- |
| Unit | 121 | ~60 ms |
| Integration | 139 | **~300 ms** (was ~3 s) |
| Verification | 24 | ~2 s |

**The integration tier no longer touches disk.** 109 of the 112 disk-bound tests were converted:
92 to `CreateInMemorySpreadsheet()` and 17 to an explicit `MemoryStream` where the test reopens the
workbook. Three keep a real path because they assert `File.Exists`, and two image tests keep one
because the API under test takes a file path. That is a ~10× speed-up, but the point is the
entry rule: a tier that *cannot* reach the file system cannot silently drift into a verification
tier.

**What moved to verification**, by the rule "interrogates the finished document as a document":

- `ReaderTest` → `WorkbookRoundtripTests` (8) — write, reopen, read back, validate.
- `CreationTest` → `RealisticWorkbookTests` (8) — the large multi-sheet scenarios.
- Three tests out of `RangeAndAdvancedFeaturesTest` → `DocumentStructureTests` — the two
  child-order assertions and the annotation-persistence round-trip. The other six raw-XML tests
  stayed in integration: they assert "feature X produced node Y", which is an integration concern
  regardless of how the assertion is made.

**New coverage (+5):**

- `ForeignWorkbookTests` (4) — reads `Example_2.xlsx`, a genuine Microsoft Excel file that was
  tracked in git but **referenced by nothing** (finding F-8). It stores text in
  `sharedStrings.xml` rather than as inline strings, so it is the only coverage the shared-string
  read path has; every other test round-trips through our own writer and therefore only ever
  proves the reader understands our own dialect. Also covers non-ASCII text and extending a
  foreign workbook without invalidating it. This is a first slice of blind spot B-9.
- `DocumentStructureTests.ReopenAndSaveWithoutChanges_KeepsTheDocumentValid` — blind spot B-13.
  Guards the family of defects where every open/save cycle duplicates a part, and is the
  precondition for any future golden-file work.

**The realistic scenarios finally assert something (finding F-5).** `CustomFile1/2/3` built
150-line workbooks and checked nothing but the phase 1 validator. They are now named for what they
verify (`MultiSheetStyledReport_IsValidAndReadable`, `LargeStyledSheet_IsValidAndReadable`,
`SheetWithTable_IsValidAndReadable`) and each reopens the document to assert sheet names, known
cell values at known references, formula text, and table metadata.

Two deviations from the plan above, both deliberate:

- **`BasicFile` was not deleted.** Phase 1 had already turned it from an empty `using` body into a
  validated minimal workbook, so it now earns its place as `MinimalWorkbook_IsValidAndHasOneSheet`.
- **`Example_2.xlsx` was not deleted** — see above; it turned out to be the most valuable fixture
  in the repository.

Incidental fix: four Czech string literals in the scenario data had been destroyed by a legacy
encoding round-trip (`Mno�stvo`, `p.�.`, `Id m�sta`, `Mat�j Z�bsk�`).
The original bytes were long gone — they were already U+FFFD replacement characters — so the
intended text was restored and is now asserted on, which also gives the suite its only non-ASCII
write-path coverage.

Also moved into `TestKit`: `SpreadsheetTestBase` (shared by both tiers, with
`CreateInMemorySpreadsheet()` as the documented default), `WorkbookParts` (sheet-name → part
resolution and workbook child names, previously duplicated in two test classes), and
`TestImages.MinimalPng()`.

Not addressed here, and not a regression from this work: `src/OfficeDocuments.Word` currently does
not compile (`Body.cs:25`, `Paragraph.cs:30`) due to in-flight WORD work. All three Excel tiers
build and pass on net8.0/net9.0/net10.0 independently of it.

### 2026-07-27 — Phase 4: style deep-dive (closes F-7; style-testing items 1–4 and 7–9)

73 new tests in `IntegrationTests/Styles/`, replacing id-only assertions with assertions on what
`styles.xml` actually contains. Excel totals: **121 unit / 212 integration / 24 verification**,
green on net8.0/net9.0/net10.0. The tier found and fixed **two more correctness bugs**.

| File | Tests | Covers |
| --- | --- | --- |
| `StyleRenderingTests` | 30 | Every facet's rendered fragment: bold/italic/underline variants, colour as 8-digit ARGB, fill patterns, every border style value, per-edge borders, alignment on the `cellXfs` entry, built-in vs custom number formats — plus CT_Font / CT_Border / CT_PatternFill child order. |
| `StyleDedupTests` | 11 | Same input reuses an entry, one differing attribute allocates a new one, N distinct styles produce exactly N entries, 50 repeats grow nothing. |
| `StyleMergeTests` | 21 | The merge truth table per facet across base-only / overlay-only / both / neither, non-commutativity, chained merges, and cross-workbook copying of every facet. |
| `StyleInheritanceTests` | 11 | Sheet → row → cell precedence, same-facet conflicts, different-facet composition, `AddStyle` layering. |
| `DifferentialFormatTests` | 7 | Conditional formatting `dxfs`: allocation, dedup across rules and worksheets, what a dxf carries, and that a colour scale needs none. |

**[Medium] Dedup rejected a legitimate match at index 0.** `GetFontId` / `GetFillId` / `GetBorderId`
appended whenever `FindElementIndex` returned `<= 0`, but `-1` means "not found" while `0` means
"matched the default entry". Asking for an empty border — reachable as `new Border()` — therefore
appended a duplicate of the default and returned index 1. This is the bug predicted in
[../../../excel-state-verdict.md](../../../excel-state-verdict.md) section C. Guard corrected to
`< 0`; `Style.cs`.

**[Medium] Cross-workbook merge silently dropped a facet.** `CreateMergedStyle` skipped a facet
when `fontId == style.FontId`, which is a valid shortcut inside one stylesheet and meaningless
across two — index 1 in the source workbook has nothing to do with index 1 in the target. Merging a
bold font from workbook B onto a 9pt font from workbook A dropped the bold entirely, because both
happened to sit at index 1. Fixed by applying the shortcut only when the two stylesheets are the
same instance; `Style.cs`. The pre-existing cross-workbook test passed only because its indexes
happened to differ.

Two behaviours turned out to be correct and are now pinned explicitly rather than "fixed", because
both surprise on first contact:

- **Merging folds in the workbook default font.** Any merge involving a level with no font of its
  own runs against the default entry, so asking only for bold yields
  `<b/><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/>`. My first draft of the
  inheritance tests asserted the narrow fragment and failed; the library was right.
- **A style reaching a cell through the sheet → row → cell chain is a new stylesheet entry.**
  Comparing `StyleIndex` against the style handed to `AddRow` does not work, so the inheritance
  tests assert on the resolved font instead.

New TestKit helpers: `StylesheetProbe` (reads the entry a style points at, and confines the
obsolete raw-stylesheet access to one place) and `OoxmlAssert.RendersAs` / `ChildOrder`. Writing
the latter surfaced a defect in `XElementExtensions.CompareXml`: it treated namespace declarations
as content, so the same element written with a default namespace and with an `x:` prefix compared
unequal. Namespace declarations are now excluded from the comparison.

### 2026-07-27 — Opt-in artifact capture

Requested alongside phase 4: keep the produced workbooks on disk when a human wants to open them.
`TestArtifacts` reads `OFFICEDOCS_TEST_OUTPUT` — unset means the previous behaviour (nothing is
kept), `1`/`true` writes under `%TEMP%/MDDM.OfficeDocuments.Tests/Output`, any other value is taken
as the target directory. With capture on, `TempWorkspace` roots itself there and stops deleting
itself, so all 19 verification workbooks survive automatically in plain per-class folders; the
scenario files were renamed from `doc2.xlsx`-style placeholders to names that say what they are.
`SpreadsheetTestBase.SaveArtifact(stream, name)` covers in-memory tests and is a no-op when capture
is off.

### 2026-07-27 — Phase 5: performance (closes the performance half of the tier model)

Two projects, resolving the "four or three" open decision in favour of four.
`test/OfficeDocuments.Excel.Benchmarks` is a BenchmarkDotNet console app that measures and never
fails; `test/OfficeDocuments.Excel.PerformanceTests` is 14 xUnit guards that run on every push.
Both are single-TFM (`net10.0`) — performance is not a per-runtime property and multi-targeting
would triple both the wall clock and the number of chances to flake. Full numbers:
[`excel-performance-baseline.md`](../../../excel-performance-baseline.md).

**The four known hot spots are now quantified.** They had been named in the verdict document since
2026-07-24 but never measured:

| Path | Growth for 4× input | Worst measured |
| --- | --- | --- |
| `CreateStyle`, distinct styles | 4× allocation per 2× input — quadratic | 1 000 styles → 3.1 s, 1.2 GB |
| `SetComment` | 16× — the steepest | 200 comments → 166 ms, 59 MB |
| `Row.CreateCell` backfill | ~15× | one cell at column 8 000 → 500 ms |
| `Range.SortByColumn` | 2.0× allocation over building the range | 2 000 rows → +67 MB |

Reusing eight styles instead of allocating a thousand distinct ones is 87× faster and allocates
84× less, which makes the documented workaround worth stating in the user-facing docs.

**No guard asserts on a duration.** Only growth ratios and allocation counts, both of which cancel
out the runner's speed. Four things had to be true before those numbers meant anything, and three
of them were discovered by getting them wrong first:

- **Tiered compilation had to be disabled.** .NET jits with a fast non-optimizing compiler and
  only replaces the code after enough calls, so a guard that ran early in the process measured
  tier-0 and one that ran late measured tier-1. The same 500-row round trip measured 97 ms and
  23 ms depending only on test order. With `TieredCompilation=false` the suite reproduces to ±3%,
  and allocation figures become deterministic too (12 KB per row here against 19 KB under
  BenchmarkDotNet, where tiering is on).
- **A ratio must be taken over one path, not a pipeline.** The first read-path guard timed
  write + close + reopen + read together and measured 10.3×. Raising the ceiling would have made
  it useless: reading is the minority share, so a read path that turned quadratic would still land
  near 8× and pass. Split so the workbooks are built outside the measured region, it measures 4.4–4.7×
  and would actually catch the regression.
- **Timing cannot classify the distinct-style path.** It allocates so heavily that GC dominates:
  33× for 4× input, between quadratic (16×) and cubic (64×), which identifies neither. Its
  complexity guard is therefore an allocation ratio, which reads a clean 4.0× for 2× input.
- **A measurement floor is needed.** Below ~4 ms a ratio is noise, so `Measure` fails with "raise
  the base size in this test" rather than passing by luck. It fired three times during calibration.

The split between `LinearScalingGuards` and `KnownHotSpotGuards` is the load-bearing part of the
design. The first states promises; the second pins behaviour that is already quadratic so it
cannot get worse before EXCEL-005. They look alike and mean opposite things, so they are separate
files, and moving a test from the second to the first is the definition of done for a fix.

Also fixed: BenchmarkDotNet writes its generated runner project into the benchmark project's `bin`
folder, so it inherited the root `Directory.Build.props` and its `TargetFrameworks`, which
overrode the single `TargetFramework` BDN sets for itself and failed the restore with NU1201. The
`TargetFrameworks` property is now conditioned on the project name.

### 2026-07-28 — Phase 6: blind spots B-2 through B-6

Each of the five was probed against the running library before anything was written, because a
test that pins wrong behaviour is worse than no test. Four turned out to be real defects; one was
already correct and only lacked cover. 61 new unit tests, 40 integration, 3 verification.

| Blind spot | What the probe found | Outcome |
| --- | --- | --- |
| B-2 sheet names | No length or character validation at all — a 40-character name or one holding `/` produced a **schema-invalid** workbook | `WorksheetNameValidator`, applied on create and rename |
| B-3 XML escaping | Already correct: the SDK escapes markup on every typed path, and the VML the comment writer builds by hand contains only numeric anchors | Tests only |
| B-4 control characters | Accepted silently by `SetValue`; the SDK threw during `Close()`, losing the whole document with a message naming neither sheet nor cell | `XmlText`, checked at the point of assignment |
| B-5 `NaN` / `Infinity` | Written verbatim as `<v>NaN</v>` into a cell marked `t="n"` | Rejected in `Cell.SetNumberValue` |
| B-6 dates | Every date before 1 March 1900 written one day late | `ExcelSerialDate` replaces `ToOADate` on both sides |

**Three of the four were invisible to every gate this suite has**, which is the finding worth
keeping:

- **`NaN` passes schema validation.** `v` is declared as a string, and "this must be a number"
  comes from the cell's `t` attribute — a semantic rule, not a grammatical one, so the validator
  never evaluates it. Phase 1 installed that validator as the principal safety net; this is the
  class of defect it cannot catch by construction.
- **The date bug passes any round trip.** `ToOADate` and `FromOADate` are exact inverses, so the
  library's own files always read back correctly while Excel read them a day late. `AGENTS.md`
  rule 4 says a round trip proves self-consistency and nothing more; this is that rule biting
  somewhere other than element order.
- **The control-character failure was in the right place but at the wrong time.** The SDK does
  refuse, but only when the part is serialized. Moving the check to the assignment turned "the
  document is gone and here is an exception about `0x01`" into "cell B3 cannot hold this".

The sheet-name defect was the only one an existing gate would have caught — the schema constrains
the `name` attribute directly — but no test had ever written an illegal name.

Two things deliberately **not** changed, and pinned instead:

- A carriage return does not survive a round trip. XML requires a parser to normalize `\r\n` to
  `\n` before the application sees it, and Excel agrees that an in-cell break is `\n`. Documented
  in the consumer guide so callers passing Windows line endings are not surprised.
- Reading stays more permissive than writing. Serial 60 is Excel's phantom 29 February 1900; a
  foreign file containing it resolves to 1 March rather than throwing, because refusing to open a
  file is a worse answer than resolving an impossible date.

**This implies a major version bump for `OfficeDocuments.Excel`.** Four inputs that were previously
accepted now throw, and the serial written for pre-March-1900 dates changed. Per the SemVer rules
in `build-and-packaging.md` that is major, and phases 1–4 already carried unversioned behaviour
changes of the same kind (`ArgbHexColor` began throwing on malformed input). The bump is one
decision covering all of it and has been left for the commit rather than made piecemeal here.

## Sources

- [ClosedXML CONTRIBUTING.md](https://github.com/ClosedXML/ClosedXML/blob/develop/CONTRIBUTING.md) — reference-file discipline
- [OpenXmlValidator.Validate](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.validation.openxmlvalidator.validate?view=openxml-3.0.1) and [Open-XML-SDK validator tests](https://github.com/OfficeDev/Open-XML-SDK/blob/master/test/DocumentFormat.OpenXml.Tests/ofapiTest/OpenXmlValidatorTest.cs)
- [Validating OpenXml generated documents](https://tech.trailmax.info/2014/04/validating-of-openxml-generated-documents-or-the-file-cannot-be-opened-because-there-are-problems-with-contents/)
- [dotnet/Open-XML-SDK#440 — invalid file past 26 columns](https://github.com/OfficeDev/Open-XML-SDK/issues/440)
- [exceljs#1474](https://github.com/exceljs/exceljs/issues/1474), [openxlsx#211](https://github.com/ycphs/openxlsx/issues/211), [openxlsx#518](https://github.com/ycphs/openxlsx/issues/518), [ImportExcel#362](https://github.com/dfinke/ImportExcel/issues/362) — sheet-name legality and escaping
- [libxlsxwriter#276](https://github.com/jmcnamara/libxlsxwriter/issues/276) — control characters in strings
- [Dates in SpreadsheetML](https://www.ericwhite.com/blog/dates-in-spreadsheetml/), [Excel 1900 leap year bug](https://learn.microsoft.com/en-us/answers/questions/416681/excel-1900-leap-year-bug)
- [Excel specifications and limits](https://support.microsoft.com/en-us/excel/excel-specifications-and-limits)
- [OOXML styles overview](http://officeopenxml.com/SSstyles.php)
- [BenchmarkDotNet](https://github.com/dotnet/BenchmarkDotNet), [MSAL.NET performance testing](https://learn.microsoft.com/ko-kr/entra/msal/dotnet/advanced/performance-testing), [BenchmarkDotNet in CI/CD](https://amarozka.dev/extending-benchmarkdotnet-exporters-metrics-ci-cd/)
- [ClosedXML#818 — large-file memory](https://github.com/ClosedXML/ClosedXML/issues/818), [C# Excel library comparison 2026](https://hackernoon.com/c-excel-library-in-depth-comparison-tested-for-2026)
- [Verify snapshot testing](https://github.com/verifytests/verify), [Verify.ClosedXml](https://github.com/VerifyTests/Verify.ClosedXml)
- [SSW — test project naming](https://www.ssw.com.au/rules/follow-naming-conventions-for-tests-and-test-projects), [Microsoft — unit testing best practices](https://learn.microsoft.com/en-us/dotnet/core/testing/unit-testing-best-practices)
