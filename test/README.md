# Test projects

Tests are split by **what a failure tells you**, not by how fast they run. That axis decides which
dependencies each tier is allowed to take, which is what keeps the split from eroding.

The full rationale, the migration plan, and the open blind-spot catalogue live in
[../.doc/tasks/core/excel/EXCEL-011-test-suite-restructuring.md](../.doc/tasks/core/excel/EXCEL-011-test-suite-restructuring.md).

## Where does my new test go?

| Project | Answers | May touch | Must not touch | Budget |
| --- | --- | --- | --- | --- |
| `OfficeDocuments.Excel.UnitTests` | "is this function correct?" | pure types only | any `SpreadsheetDocument`, any stream, the file system | tier < 200 ms |
| `OfficeDocuments.Excel.IntegrationTests` | "do these types behave correctly together through the public API?" | a workbook over `MemoryStream` | the file system | tier < 10 s |
| `OfficeDocuments.Excel.VerificationTests` | "is the finished file a valid, correct Excel document?" | complete documents, disk, foreign input files | — | tier < 60 s |
| `OfficeDocuments.Excel.PerformanceTests` | "did the cost of this change shape?" | large workbooks, timing, allocation counters | any assertion on an absolute duration | tier < 120 s |
| `OfficeDocuments.Word.Tests` | "does the Word surface behave correctly through the public API?" | a document, ideally over `MemoryStream` | — | tier < 10 s |

**The unit-tier rule, in one line:** if you need a `Spreadsheet` to write the test, it is not a
unit test. The project deliberately does not reference `OfficeDocuments.Excel.TestKit`, so the
temp-file workspace and the schema validator are simply not reachable from it.

Note the limit of that guarantee: the library itself depends on `DocumentFormat.OpenXml`, so the
packaging types remain *technically* reachable from the unit project. The rule is enforced by
convention and review, not by the compiler.

**Integration vs verification**, since this is the boundary that is easy to blur: a test that says
*"feature X produced node Y"* is integration. A test that says *"the document is schema-valid"*,
*"it survives a close and reopen"*, *"its children are in CT_Workbook order"*, or *"a file Excel
wrote reads correctly"* is verification. The integration tier stays in memory; only verification
touches disk, and only where the path is part of what is being checked.

`OfficeDocuments.Word.Tests` is still a single tier; it gets the same split when the Word surface
grows enough to need it.

## Performance: `OfficeDocuments.Excel.PerformanceTests` and `OfficeDocuments.Excel.Benchmarks`

Two projects, because measuring and gating are different jobs. **Benchmarks** produce numbers a
person reads and never fail a build; they are not run by `dotnet test` and not run in CI.
**PerformanceTests** are ordinary xUnit tests that run on every push, and every threshold in them
is traceable to a benchmark number.

**No test here may assert on an absolute duration.** A millisecond threshold measures the machine,
so one tuned on a workstation is either meaningless or flaky on a shared runner. Two kinds of
assertion survive the move to CI:

- **A growth ratio.** Measure at N and at 4N in the same run and compare. Linear work costs about
  4x, quadratic about 16x, cubic about 64x — wide bands that hardware speed cancels out of.
- **An allocation count.** `GC.GetAllocatedBytesForCurrentThread()` is counted, not sampled, so it
  does not care what else the machine is doing. Where a defect shows up in allocation at all,
  guard it there; those are the tests that never flake.

Three files, and the split between the first two carries the meaning:

- `LinearScalingGuards` — paths promised to stay linear. A failure is a regression.
- `KnownHotSpotGuards` — paths that are *already* quadratic, pinned so they cannot get worse
  before `EXCEL-005` fixes them. A failure here does not mean "this is slow now", it means "this
  changed complexity class". When one is fixed, its test moves to `LinearScalingGuards`.
- `AllocationGuards` — the deterministic half of both of the above.
- `ScaleCeilingTests` — correctness at a size no other tier reaches: 25 000 rows survive a round
  trip, column 16 384 works.

Guards report their measurement whether or not they pass, so the log is a cheap trend line:

```sh
dotnet test test/OfficeDocuments.Excel.PerformanceTests/… --logger "console;verbosity=detailed"
```

Two settings in that project exist to make a handful of runs mean something, and removing either
brings the noise back: the assembly disables xUnit parallelism, because a timing measurement taken
while other tests run on other cores is not a measurement; and it disables tiered compilation,
because .NET jits with a fast non-optimizing compiler first and only replaces it after enough
calls — the same 500-row round trip measured 97 ms and 23 ms depending purely on which test ran
first in the process.

To measure rather than gate:

```sh
dotnet run -c Release --project test/OfficeDocuments.Excel.Benchmarks -- --filter '*'
dotnet run -c Release --project test/OfficeDocuments.Excel.Benchmarks -- --filter '*Comment*'
```

Current baseline and the four known hot spots:
[../.doc/excel-performance-baseline.md](../.doc/excel-performance-baseline.md).

## `OfficeDocuments.Excel.TestKit`

Shared helpers for the tiers that legitimately touch packaging or disk. Not a test project — it
declares no test SDK and holds no `[Fact]`.

- `Validation/OpenXmlValidation` — `AssertValid(...)` runs `OpenXmlValidator` against
  `FileFormatVersions.Office2021` and reports each error with its part URI and XPath.
  **Every test that produces a complete document should end with this call.** It is what catches
  schema-order and relationship defects that a round-trip through this library cannot see, because
  a round-trip only proves self-consistency. The `inheritedDefects` parameter tolerates named
  defects that arrived with a foreign input document — real Excel files are not always schema-clean.
- `XElementExtensions.CompareXml` — namespace- and order-normalized XML equality, for asserting on
  generated OOXML fragments without depending on attribute or sibling ordering.
- `TempWorkspace` — per-test-class temporary directory, for the tests that genuinely need a file on
  disk. Prefer a `MemoryStream` when you do not.
- `SpreadsheetTestBase` — workbook factory methods shared by the integration and verification
  tiers. `CreateInMemorySpreadsheet()` is the default; `GetFilepath(...)` is for the handful of
  tests where the path itself is the subject.
- `WorkbookParts` — resolves a `WorksheetPart` by sheet name and lists the workbook's child
  element names, for assertions made against the raw package.
- `TestImages.MinimalPng()` — a 1×1 PNG, so image tests need no binary fixture.
- `StylesheetProbe` — reads the `font`/`fill`/`border`/`numFmt`/`cellXfs`/`alignment` entry a style
  actually points at, plus the entry counts and whether two styles share a stylesheet. Assert on
  these, not on the allocated id: an id assertion cannot tell a blue font from a red one. It is also
  the **only** place a test may touch `IStyle.Element` or `IStyle.Stylesheet` — the obsolete raw
  access lives behind one documented suppression here so no test project carries its own.
- `OoxmlAssert.RendersAs(...)` — compares a produced element against an expected fragment written
  without namespace declarations. `OoxmlAssert.ChildOrder(...)` pins a schema sequence.
- `TestArtifacts` — see below.

## Inspecting the produced workbooks

Tests leave nothing on disk by default. To keep the generated files and open them in Excel, set
`OFFICEDOCS_TEST_OUTPUT`:

```sh
OFFICEDOCS_TEST_OUTPUT=1              dotnet test …   # writes under %TEMP%/MDDM.OfficeDocuments.Tests/Output
OFFICEDOCS_TEST_OUTPUT=C:/tmp/xlsx    dotnet test …   # writes there instead
```

With capture on, `TempWorkspace` roots itself under that directory and stops deleting itself, so
every verification test leaves its workbook behind automatically, in a plain per-class folder. For
an in-memory test whose output is worth eyeballing, call `SaveArtifact(stream, "name.xlsx")` — it
is a no-op when capture is off, so it is safe to leave in place.

## `OfficeDocuments.Word.TestKit`

The same two helpers for the Word module, against `WordprocessingDocument`, plus two input builders.
Not a test project.

- `Validation/OpenXmlValidation` — as above. It earns its place here for a Word-specific reason: the
  child-order rules in WordprocessingML are strict and easy to break by accident. `w:sectPr` has to be
  the last child of `w:body`, `w:pPr` the first child of `w:p`, and the children of `w:rPr`, `w:pPr`,
  and `w:style` each follow a fixed sequence. A document that violates one of these still round-trips
  through this library and still reads back correctly — it only fails in Word.
- `TempWorkspace` — per-test-class temporary directory under a `Word` root.
- `TestImages` — minimal PNGs, including one built to a requested size and resolution, for the image
  tests. No binary fixtures.
- `ForeignDocuments` — documents as a producer other than this library writes them. **Use this for
  anything that reads or updates.** A document this library wrote and then read back proves only that
  it agrees with itself; what breaks read paths in practice is that Word starts a new run wherever
  spell-check state or revision tracking changes, so a placeholder typed as one word arrives as three
  runs. `ForeignDocuments` reproduces that splitting through the SDK, so the input is reviewable in a
  diff and deterministic on every platform, and includes the `w:sectPr` every real document carries.

`OfficeDocuments.Word.Tests` inherits `WordTestBase`, whose `WriteAndValidate(...)` authors a
document into a `MemoryStream` and validates it, so the gate applies even to tests that were written
to check something else. Assert on values through `ReadDocumentElement(...)` rather than by matching
against `ReadMainDocumentXml(...)`: both `w:val="0"` and `w:val="false"` mean the same thing to Word,
so a string assertion pins the SDK's formatting instead of the library's behaviour. String matching
is still the right tool for structure, such as the presence of `xml:space="preserve"`.

How to write a Word test once you know it belongs here — including the negative-assertion rule that
three Word defects got past — is in
[../.doc/ai-instructions/testing.md](../.doc/ai-instructions/testing.md).

## Running

```sh
dotnet test test/OfficeDocuments.Excel.UnitTests/OfficeDocuments.Excel.UnitTests.csproj   # fast
dotnet test OfficeDocuments.Excel.slnx                                                    # Excel, all tiers
dotnet test OfficeDocuments.Word.slnx                                                     # Word
dotnet test OfficeDocuments.slnx                                                          # everything
```

The per-module solutions are the normal entry point. Besides loading faster, they are what keeps
the two modules from growing a dependency on each other: neither contains the other's projects, so
a cross-module reference fails to build instead of quietly appearing in a diff.
