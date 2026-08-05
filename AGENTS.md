# AGENTS.md

Primary instruction file for AI coding agents working in `MDDM.OfficeDocuments`. Keep it short;
depth lives in [`.doc/ai-instructions/`](.doc/ai-instructions/README.md).

## The project

A .NET library that wraps `DocumentFormat.OpenXml` in a smaller, task-oriented API for creating and
reading `.xlsx` and `.docx`. Legacy binary `.xls` / `.doc` is out of scope, permanently.

- SDK `10.0.300` (pinned in `global.json`), C# at `default` language level, `net8.0;net9.0;net10.0`.
- Single runtime dependency: `DocumentFormat.OpenXml`. Tests are xUnit.
- `OfficeDocuments.Excel` (`4.0.0`) and `OfficeDocuments.Word` (`4.0.0`) version independently.
  Word is the smaller surface, but its core backlog is complete as of 2026-07-27 — further Word work
  is a choice, not a dependency. The remaining `P0` is the Excel public-surface cleanup.

## Repository map

| Path | What it is |
| --- | --- |
| `src/OfficeDocuments.Excel/` | Excel library. Public contract in `Interfaces/`; coordinators `Spreadsheet`/`Worksheet` delegate to internal collaborators in `DataClasses/` |
| `src/OfficeDocuments.Word/` | Word library. Public contract in `Interfaces/`; formatting records in `Formatting/` |
| `test/OfficeDocuments.Excel.UnitTests/` | Fast tier — pure types, no `SpreadsheetDocument`, no I/O |
| `test/OfficeDocuments.Excel.IntegrationTests/` | Excel behaviour through the public API, over `MemoryStream` |
| `test/OfficeDocuments.Excel.VerificationTests/` | Whole documents — schema validity, reopen, foreign input files |
| `test/OfficeDocuments.Excel.PerformanceTests/` | Growth-ratio and allocation guards. Never asserts on a duration |
| `test/OfficeDocuments.Excel.Benchmarks/` | BenchmarkDotNet console app. Measures; never fails a build |
| `test/OfficeDocuments.Word.Tests/` | Word behaviour through the public API |
| `test/OfficeDocuments.*.TestKit/` | Shared helpers (schema validator, temp workspace). Not test projects |
| `.doc/` | All documentation. Index: [`.doc/README.md`](.doc/README.md) |
| `OfficeDocuments.slnx` | Everything. There is no `.sln` |
| `OfficeDocuments.Excel.slnx` | Excel module only — contains no Word project, by design |
| `OfficeDocuments.Word.slnx` | Word module only — contains no Excel project, by design |

Which tier a new test belongs in: [`test/README.md`](test/README.md).

## Rules that always apply

1. **Minimal, root-cause diffs.** No drive-by refactors, no reformatting bundled into a behaviour
   change, no compatibility shims — change the code instead of wrapping it.
2. **The interface layer is the public API.** Never add a public member that requires the caller to
   understand OpenXml internals, and do not widen the leakage that already exists.
3. **Element order is a correctness invariant, not a style preference.** OOXML fixes the order of
   child elements. A wrong order still round-trips through this library and still reads back
   correctly — it only fails when Excel or Word opens the file. This bug class has hit the repo three
   times. Use the orderers (Excel) and the SDK's typed property setters (Word).
4. **Every test that produces a complete document ends with the schema validator.** A round-trip
   proves self-consistency, nothing more.
5. **Central Package Management.** No `Version` attribute on any `PackageReference`.
6. **Documentation is English**, the root `README.md` stays product-oriented, detail goes in `.doc/`,
   and snippets must match the real API.
7. **Verify before reporting done.** Run the focused tier, widen for shared code. If a check did not
   run, say which one and why.
8. **The two modules are independent.** Nothing under `*.Excel*` may reference `*.Word*`, or the
   reverse — not a `ProjectReference`, not a `using`, not a shared helper. They ship as separate
   packages on separate version lines, and a user who installs one must not drag in the other.
   Anything genuinely common belongs in a third project, not in whichever module wrote it first.
   `OfficeDocuments.Excel.slnx` and `OfficeDocuments.Word.slnx` are how this is enforced: each omits
   the other module, so a cross-reference fails the build rather than passing review.
9. **No performance test asserts on a duration.** A millisecond threshold measures the CI runner,
   not the code. Assert a growth ratio between t(N) and t(4N), or an allocation count — both cancel
   the hardware out. Thresholds are traceable to
   [`.doc/excel-performance-baseline.md`](.doc/excel-performance-baseline.md).
10. **Never keep a second copy of the document's structure.** The package is the single source of
    truth; anything a wrapper exposes about it — child order, which part owns a relationship, which
    headers exist — is derived, not stored. This bug class has now cost three defects, and the second
    one is the instructive one: `ElementWrapperList` cached the list and hand-synchronized *additions*,
    which held until something removed an element. **Caching part of a derived value is still
    duplication** — it only makes the drift rarer and harder to find. Where a lookup genuinely has to
    be cached for performance, it must be invalidated by construction, never by remembering to call
    something.

## Commands

Work on one module through that module's solution. It is faster, and it is the check that keeps
the two modules independent — see rule 8.

```powershell
dotnet build OfficeDocuments.Excel.slnx
dotnet test  OfficeDocuments.Excel.slnx                                                     # Excel, all tiers
dotnet build OfficeDocuments.Word.slnx
dotnet test  OfficeDocuments.Word.slnx                                                      # Word
dotnet test  OfficeDocuments.slnx                                                           # everything

dotnet test test/OfficeDocuments.Excel.UnitTests/OfficeDocuments.Excel.UnitTests.csproj     # fast tier
dotnet test test/OfficeDocuments.Excel.IntegrationTests/OfficeDocuments.Excel.IntegrationTests.csproj
dotnet test test/OfficeDocuments.Excel.VerificationTests/OfficeDocuments.Excel.VerificationTests.csproj
dotnet test test/OfficeDocuments.Excel.PerformanceTests/OfficeDocuments.Excel.PerformanceTests.csproj

# Measure rather than gate. Release only; not run by `dotnet test`.
dotnet run -c Release --project test/OfficeDocuments.Excel.Benchmarks -- --filter '*'
```

## Detailed instructions

Read the file that matches what you are touching. Do not load all of them.

| Working on | Read |
| --- | --- |
| Any non-trivial task | [`.doc/ai-instructions/workflow.md`](.doc/ai-instructions/workflow.md) |
| Any `.cs` file | [`.doc/ai-instructions/csharp.md`](.doc/ai-instructions/csharp.md) |
| Excel module | [`.doc/ai-instructions/excel.md`](.doc/ai-instructions/excel.md) |
| Word module | [`.doc/ai-instructions/word.md`](.doc/ai-instructions/word.md) |
| Tests | [`.doc/ai-instructions/testing.md`](.doc/ai-instructions/testing.md) + [`test/README.md`](test/README.md) |
| Anything performance | [`.doc/excel-performance-baseline.md`](.doc/excel-performance-baseline.md) |
| Projects, props, CI | [`.doc/ai-instructions/build-and-packaging.md`](.doc/ai-instructions/build-and-packaging.md) |
| Any `.md` | [`.doc/ai-instructions/documentation.md`](.doc/ai-instructions/documentation.md) |
| Deciding what to build | [`.doc/tasks/roadmap-overview.md`](.doc/tasks/roadmap-overview.md) |
| Deciding core vs advanced | [`.doc/architecture/minimal-core-pr-guidelines.md`](.doc/architecture/minimal-core-pr-guidelines.md) |

## Precedence

The user's explicit request wins. Then the most specific instruction file — module beats language
beats this file. Then the conventions of the file you are editing. Then Microsoft .NET guidance and
common OSS practice. When you deviate, keep it localized and say so.
