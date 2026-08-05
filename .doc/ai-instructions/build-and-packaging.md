---
name: Build and packaging
description: MSBuild layout, central package management, target frameworks, versioning, and CI expectations.
applyTo:
  - "**/*.csproj"
  - "**/*.props"
  - "**/*.slnx"
  - ".github/workflows/**"
---

# Build and packaging

## Where settings live

| File | Owns |
| --- | --- |
| [`global.json`](../../global.json) | SDK pin: `10.0.300`, `rollForward: latestFeature`, no prerelease |
| [`Directory.Build.props`](../../Directory.Build.props) | `LangVersion`, `Nullable`, `ImplicitUsings`, `TargetFrameworks`, shared NuGet metadata |
| [`Directory.Packages.props`](../../Directory.Packages.props) | Every `PackageVersion` — Central Package Management is on |
| `src/*/OfficeDocuments.*.csproj` | Only project-specific metadata: `Version`, `Title`, `Description`, `PackageTags`, package content |

Rules that follow from that split:

- A `PackageReference` must **never** carry a `Version` attribute. Add the version to
  `Directory.Packages.props` instead.
- Do not copy `TargetFrameworks`, `Nullable`, or `LangVersion` into an individual project.
- Do not add `GeneratePackageOnBuild`. Packing is an explicit `dotnet pack` step.
- Do not add `Newtonsoft.Json` or `System.IO.Packaging`. They are absent on purpose.
- Every new dependency must be justified against the cost it adds to a minimal-core library. See
  [../architecture/minimal-core-pr-guidelines.md](../architecture/minimal-core-pr-guidelines.md).

## Target frameworks

`net8.0;net9.0;net10.0` for every project. A change is not verified until it builds on all three —
see the multi-target trap in [csharp.md](csharp.md).

## Versioning

Packages version independently: `OfficeDocuments.Excel` is `4.0.0`, `OfficeDocuments.Word` is
`4.0.0`. Follow SemVer against the *public* surface:

- Removing or changing the signature, behaviour, or exception type of a public member is **major**.
  WORD-001 took Word from `1.0.0` to `2.0.0` for exactly this reason.
- **Adding a member to a public interface is also major**, because it breaks every external
  implementer. This is why Word has gone up a major on each of its four core tasks despite being
  almost entirely additive — the surface is interface-first, so "additive" and "non-breaking" are not
  the same thing here. Adding a member to a public *class* is minor.
- Internal refactoring with no surface change is patch.
- Obsoleting is not removing. `AddCellWithValue` stays until a deliberate major.

Word's line so far: `1.0.0` → `2.0.0` (WORD-001, formatting) → `3.0.0` (WORD-002A/B/C and WORD-003,
tables through metadata) → `4.0.0` (WORD-004, search and update).

## Solution files

Three, all `.slnx` — the older `.sln` no longer exists, and CI once broke by referencing it.

| File | Holds |
| --- | --- |
| [`OfficeDocuments.slnx`](../../OfficeDocuments.slnx) | Everything |
| [`OfficeDocuments.Excel.slnx`](../../OfficeDocuments.Excel.slnx) | Excel and its tests, with no Word project |
| [`OfficeDocuments.Word.slnx`](../../OfficeDocuments.Word.slnx) | Word and its tests, with no Excel project |

The per-module files are not just a convenience. The two modules ship as separate packages and must
not depend on each other (`AGENTS.md` rule 8), and omitting the other module from each solution is
what turns that rule into a build failure instead of a review comment. Each CI workflow restores
only its own solution, so work in flight in one module cannot break the other's build.

A new project is not part of the build until it is added to the solutions it belongs in **and** to
the relevant workflow.

## CI

- [`github-build-excel.yml`](../../.github/workflows/github-build-excel.yml) — restores and builds
  `OfficeDocuments.Excel.slnx`, then runs the tiers in widening order (unit, integration,
  verification, performance) so a failure points at the smallest possible scope. Uploads `.trx`
  results and Cobertura coverage. The performance step runs with detailed console logging, because
  those tests report their measurement whether or not they pass.
- [`github-build-word.yml`](../../.github/workflows/github-build-word.yml) — the same shape for
  `OfficeDocuments.Word.slnx`, with a single test step because Word is still one tier. When it is
  split, this workflow gains the widening order too.
- Both pin the SDK from `global.json`. Do not hardcode a version in a workflow.
- Known broken: [`publish.yml`](../../.github/workflows/publish.yml) still points at the
  pre-rename path `OpenXmlApi\OfficeDocumentsApi.Excel\...` and uses a long-abandoned action. Treat
  NuGet publishing as manual until it is rewritten.

## Local validation

Build and test commands are in [AGENTS.md](../../AGENTS.md). Additionally, for changes in this area:

```powershell
dotnet pack src/OfficeDocuments.Excel/OfficeDocuments.Excel.csproj -c Release
dotnet list package --vulnerable --include-transitive   # when dependencies change
```

A build or packaging change is verified against the **whole solution**, not one project — that is
the class of change most likely to break a target framework you did not think about.
