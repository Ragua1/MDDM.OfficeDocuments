# Proposal: a guide web app (Blazor + MudBlazor)

Date: 2026-08-06
Status: **Proposal — approved in outline, nothing implemented**

A Blazor web application that presents the `OfficeDocuments.Excel` and `OfficeDocuments.Word` APIs
as a readable, navigable guide: prose plus real, compiled C# for each topic.

The application **presents** the guide. It does not invoke either library at runtime.

## Decisions already taken

| Decision | Value | Where it lands |
| --- | --- | --- |
| Location | New top-level `samples/` | Repository map in [AGENTS.md](../../AGENTS.md) |
| Name | `OfficeDocuments.Guide.Web` | Folder, project, namespace, solution |
| Single-targeting | Scoped `samples/Directory.Build.props`, verified below | New file |
| Runtime behaviour | **Presentation only — no library execution in the web app** | Drives the whole architecture |
| Content model | **A** — guided tutorial content, *derived from* the root README and `.doc/`, which stay the normative base | [Content model](#content-model) |

## Why this, and why now

The two consumer guides ([excel-library.md](../excel-library.md),
[word-library.md](../word-library.md)) are good reference material and a poor first contact. They are
alphabetical-ish lists of members followed by examples, which is what you want on your third day and
not on your first. A guided path — object model, then writing, then reading, then the sharp edges —
is a different document with a different shape.

It matters more now than it did: the repository is going public, and a public OSS library is judged
in the first ninety seconds.

### Goals

1. **A guided reading order**, not a reference index.
2. **The displayed code is compiled code.** See [Snippet rot](#snippet-rot-is-still-the-design-problem).
3. **Teach the sharp edges.** The refusal rules, the 1900 leap-year bug, `null` vs `false` in Word
   formatting, and run-boundary text replacement are what this library gets right and its
   competitors get subtly wrong. They are the reason to choose it, and a feature list cannot say so.
4. **Stay a leaf.** Consumed by nothing.

### Non-goals

- **The web app never calls into `OfficeDocuments.Excel` or `OfficeDocuments.Word`.** No documents
  are produced, none are downloaded, nothing is validated at runtime. This is a decision, not a
  simplification to be walked back later without a rethink — see [Consequences](#consequences-of-presentation-only).
- No live code editing or Roslyn compilation.
- No in-browser rendering of `.xlsx` / `.docx`.
- No deployment, no CI workflow, no hosting in the first version. It runs with `dotnet run`.
- No multi-targeting. `net10.0` only.
- No persistence, accounts, or user uploads.

### Consequences of presentation-only

Worth stating plainly, because two of them are losses:

- **Lost:** a reader cannot press a button and get the `.xlsx`. The pages that teach the refusal
  rules and the 1900 date bug now *describe* the exception rather than raising it in front of the
  reader, which is less vivid.
- **Lost:** the read-back panel, which would have demonstrated the read API on every page for free.
  The read API now needs its own pages, which the content plan below accounts for.
- **Gained:** no server-side execution of anything, no temp files, no per-request cost, no attack
  surface beyond serving static content.
- **Gained:** the app has no dependency on `DocumentFormat.OpenXml`, which is what makes the
  WebAssembly / static-hosting option in [Hosting](#hosting-and-render-model) real.

## Snippet rot is still the design problem

Everything else here is ordinary web work. This is the part worth designing, and dropping runtime
execution makes it *more* important rather than less.

[documentation.md](../ai-instructions/documentation.md) sets the bar: *"Every snippet must match the
real public API… the project's own docs once demonstrated calling `Close()` inside a `using`, which
threw `ObjectDisposedException`."* A guide app has far more sample code than the Markdown guides do,
and a polished site implies a stronger promise of correctness than a text file does.

`EXCEL-007` and `EXCEL-009` are open and **will** change public signatures. Snippet drift is not
hypothetical here; it is scheduled.

**The rule: a snippet is compiled code, and the site displays that code's own source text.** The
snippets live in their own project, which references both libraries and compiles against them. The
web app reads the `.cs` files as text. Nothing executes.

| Property | Achieved by |
| --- | --- |
| The snippet compiles | It is ordinary project code; drift is a build error |
| The snippet matches the API | There is only one copy of it |
| The snippet *works* | **Not covered in v1.** See [Verification](#verification) |

That last row is the honest cost of presentation-only, and it is called out again below rather than
buried here.

## Project shape

Two projects. The split is what turns "the web app does not use the libraries" from a promise into
something the compiler enforces.

```text
samples/
  Directory.Build.props                        # single-target override, scoped to this folder

  OfficeDocuments.Guide.Snippets/               # compiles against both libraries; never executed
    OfficeDocuments.Guide.Snippets.csproj       #   ProjectReference: Excel, Excel.Advanced, Word
    Excel/
      E01_HelloWorkbook.cs
      E02_RowsAndCells.cs
      ...
    Word/
      W01_HelloDocument.cs
      ...

  OfficeDocuments.Guide.Web/                    # renderer only; no library reference
    OfficeDocuments.Guide.Web.csproj            #   MudBlazor; snippets included as *text*
    Program.cs
    Content/                                    # prose, one file per topic
    Guide/
      GuideTopic.cs                             # id, title, summary, module, category, snippet ids
      GuideIndex.cs                             # explicit ordered list — drives the nav
      SnippetSource.cs                          # loads + extracts a region from embedded text
    Components/
      Layout/MainLayout.razor, NavMenu.razor
      Pages/Home.razor, Topic.razor
      Shared/CodeBlock.razor, SnippetView.razor
    wwwroot/
```

### How the web app gets the snippet text without referencing the libraries

```xml
<!-- OfficeDocuments.Guide.Web.csproj -->

<!--
  The snippet sources are content, not code: they are embedded as text and displayed. This project
  deliberately has no reference to OfficeDocuments.* and no DocumentFormat.OpenXml dependency —
  the guide presents the API, it does not call it.
-->
<ItemGroup>
  <EmbeddedResource Include="../OfficeDocuments.Guide.Snippets/**/*.cs"
                    Exclude="../OfficeDocuments.Guide.Snippets/{bin,obj}/**"
                    LogicalName="snippets/%(RecursiveDir)%(Filename)%(Extension)" />
</ItemGroup>

<!--
  ReferenceOutputAssembly=false: build ordering and a visible dependency in the solution graph,
  without an assembly reference. If the snippets stop compiling, this build fails too.
-->
<ItemGroup>
  <ProjectReference Include="../OfficeDocuments.Guide.Snippets/OfficeDocuments.Guide.Snippets.csproj"
                    ReferenceOutputAssembly="false" />
</ItemGroup>
```

The `ReferenceOutputAssembly="false"` line is the important one: without it the guide would happily
render snippets that no longer compile.

### Region extraction

Each snippet file marks the interesting part:

```csharp
public static class E03_Styles
{
    public static void Run()
    {
        #region snippet
        using var spreadsheet = new Spreadsheet("styled.xlsx", createNew: true);
        var worksheet = spreadsheet.AddWorksheet("Sheet1");
        // ...
        #endregion
    }
}
```

The site shows the region, de-indented; the file stays ordinary compilable code. A file with no
region markers is shown whole — a degraded display, not a crash. Multiple named regions per file
(`#region snippet:create`, `#region snippet:apply`) let one topic show two steps.

Because nothing runs, the snippets can use the file-path overloads where those read better —
`new Spreadsheet("report.xlsx", createNew: true)` is clearer than a `MemoryStream` for a first
example. A separate topic covers the stream API, which is what server-side consumers actually need.

### Where it sits in the solutions

`OfficeDocuments.Guide.Snippets` references both modules. That is allowed —
[AGENTS.md](../../AGENTS.md) rule 8 governs `*.Excel*` and `*.Word*` projects, and this is neither.
It is a leaf consumer, exactly like an end user's application.

Two guard rails:

- **Both projects go into [OfficeDocuments.slnx](../../OfficeDocuments.slnx) only.** Never into
  `OfficeDocuments.Excel.slnx` or `OfficeDocuments.Word.slnx`. Those exist to turn rule 8 into a
  build failure, and a both-modules project in either would defeat that. The solution entry should
  carry a comment saying so, because "add it to the others too" is the natural next move.
- **No shared abstraction over the two modules inside the snippets project.** `Excel/` and `Word/`
  are two folders that share nothing but a namespace root. The moment something there wants to be "a
  generic document builder over both", that is rule 8's spirit leaking somewhere the compiler cannot
  check.

`ProjectReference`, not `PackageReference`: the snippets must track the working tree. A NuGet package
would demonstrate an API the repository no longer has, which is the failure this design exists to
prevent.

## Build-system changes

Four, all small. These are the ones most likely to be got wrong, so they are spelled out.

### 1. Single-targeting — `samples/Directory.Build.props`

[Directory.Build.props](../../Directory.Build.props) applies `net8.0;net9.0;net10.0` to every
project except the BenchmarkDotNet runner. A `Microsoft.NET.Sdk.Web` project cannot multi-target,
and **`TargetFrameworks` wins over `TargetFramework`** — so setting `<TargetFramework>net10.0</TargetFramework>`
in the `.csproj` does nothing at all and the project builds three times.

Scoped to the folder rather than by growing a name-exclusion list in the root file:

```xml
<!-- samples/Directory.Build.props -->
<Project>

  <!--
    Import the repository-wide settings explicitly: MSBuild uses the *nearest* Directory.Build.props
    and stops, so without this the samples would lose Nullable and ImplicitUsings.
  -->
  <Import Project="$([MSBuild]::GetPathOfFileAbove('Directory.Build.props', '$(MSBuildThisFileDirectory)../'))" />

  <!--
    Samples are applications, not packages. Single-target because a web SDK project cannot
    multi-target, and there is nothing to gain: the libraries' net8.0/net9.0 builds are proven by the
    test suite, not here. TargetFrameworks must be cleared, not just overridden — a non-empty
    TargetFrameworks makes MSBuild run an outer multi-targeting build and ignore TargetFramework.
  -->
  <PropertyGroup>
    <TargetFrameworks />
    <TargetFramework>net10.0</TargetFramework>
    <IsPackable>false</IsPackable>
    <IsTestProject>false</IsTestProject>
  </PropertyGroup>

</Project>
```

**Verified 2026-08-06**, not assumed: a throwaway project under this layout, using the repository's
real root `Directory.Build.props` and `global.json`, produced a single `bin/Debug/net10.0/` and
inherited `Nullable=enable`, `ImplicitUsings=enable`, and the shared NuGet metadata. Omitting the
`<Import>` loses all three; leaving `TargetFrameworks` unset instead of cleared produces the
three-fold build.

The SDK pin in `global.json` (`10.0.302`) is already sufficient — no change there.

Note the consequence for the snippets project: it compiles against the libraries' `net10.0` build
only. That is enough for signature checking, which is what it is for.

### 2. Central Package Management

`MudBlazor` gets a `PackageVersion` in [Directory.Packages.props](../../Directory.Packages.props) in
a new `Samples` group, and a bare `PackageReference` in the web `.csproj` with **no `Version`
attribute** (rule 5).

```xml
<ItemGroup Label="Samples">
  <PackageVersion Include="MudBlazor" Version="…" />
</ItemGroup>
```

Pin the exact version `dotnet add package MudBlazor` resolves against the .NET 10 SDK at
implementation time; this document deliberately does not guess it.

[build-and-packaging.md](../ai-instructions/build-and-packaging.md) requires every new dependency to
be justified against the cost it adds to a minimal-core library. **MudBlazor adds exactly zero cost
to any shipped package** — it is referenced by a non-packable application that no library
references. Worth saying in the PR rather than leaving a reviewer to derive it.

The snippets project takes **no** package reference at all, only the three `ProjectReference`s.

### 3. Solution membership

Both projects into `OfficeDocuments.slnx` under a new `/samples/` folder. Nowhere else.

### 4. CI

Nothing in the first version. The Excel and Word workflows restore only their own module solution,
so projects absent from both are invisible to them — which is correct: a broken guide must not fail
a library build.

**This is a real gap, not an omission:** until a workflow exists, nothing outside a developer's
machine compiles the snippets, and the anti-rot guarantee is only as good as whoever last ran a
full-solution build. Options, in increasing cost: run `dotnet build OfficeDocuments.slnx` before any
public-surface change; add a build step for it to a full-solution job; give the guide its own
workflow. Recommendation: accept the gap for v1 and close it when the Excel surface cleanup starts,
since that is the change most likely to break the snippets.

## Hosting and render model

Presentation-only changes this materially — the earlier reasoning assumed OOXML code had to execute
on the server.

| Option | Verdict |
| --- | --- |
| **Blazor Web App, `InteractiveServer`** | **Recommended for v1.** Prerendered HTML on first paint (so the content is indexable and fast), MudBlazor's interactive components work, `dotnet run` and it works. Needs a host if it is ever published |
| Blazor WebAssembly, standalone | **Now genuinely viable**, and it was not before: with no OOXML in the browser there is no compatibility question left. `dotnet publish` produces static files, so GitHub Pages hosts a public guide for free with no server to maintain. Costs a multi-megabyte runtime download on first visit and poor search-engine indexing without prerendering |
| Static SSR only, no interactivity | Cheapest, but most MudBlazor components need an interactive render mode; the drawer, tabs, and copy button would all need replacing |
| Razor Pages / MVC | No reason to give up the component model |

Recommendation: **build for `InteractiveServer`, keep the WASM door open.** The app is a renderer
over embedded text, so it has no server-only dependency — porting cost is the render mode and the
hosting, not the code. Revisit the moment public deployment is actually decided
([open decision 1](#open-decisions)); if the answer is "GitHub Pages", WASM is the better answer and
switching early is cheaper than switching late.

## Content model

**Decided: A.** The alternatives are kept below because the reason for rejecting them is the same
reason approach A has to be policed. Prose has to come from somewhere, and the repository already
has 1 005 lines of it in `.doc/excel-library.md` and `.doc/word-library.md`.

| Approach | What it means | Cost | Risk |
| --- | --- | --- | --- |
| **A. Derived tutorial content** | New Markdown under `Content/`, restructuring the base into a guided path. The root README and `.doc/*-library.md` stay normative | High | Two sets of prose about the same API. Handled by [the source-of-truth rule](#the-source-of-truth-rule), not eliminated |
| **B. Render the existing `.doc/` files** | The app is a Markdown viewer over the current guides | Very low | Adds little over reading them on GitHub; snippets stay uncompiled strings, so the anti-rot design gets no use |
| **C. Guide becomes the source** | Tutorial content in `Content/`, and `.doc/*-library.md` are reduced to short pointers at it | High + a docs restructure | Reference material stops being readable in the repository and in the NuGet package readmes |

### The source-of-truth rule

The Markdown already in the repository — the root [README.md](../../README.md) and everything under
`.doc/` — **is the base. The guide only interprets it.** No behavioural fact originates in the
guide. When a base document changes, the guide topics derived from it have to be updated too; that
coupling is accepted deliberately rather than engineered away.

Stated as rules, because "accepted deliberately" decays into "forgotten" without them:

- **The base is normative.** `.doc/excel-library.md`, `.doc/word-library.md`, and the root README
  hold the authoritative wording for every rule, guarantee, and limitation. They stay complete, stay
  readable in the repository, and stay shippable as package readmes.
- **The guide is a re-presentation, not a second author.** It restructures the base into a reading
  order, adds connective narrative, and may omit what would distract a beginner. It does not
  introduce a rule the base does not state, and it does not paraphrase a normative sentence into
  something subtly different.
- **Every guide topic declares what it derives from**, in front matter:

  ```yaml
  ---
  id: excel/refusals
  title: What the library refuses to write
  source:
    - .doc/excel-library.md#what-the-library-refuses-to-write
    - .doc/excel-library.md#dates-and-the-1900-leap-year-bug
  ---
  ```

  This is the mechanism that makes the coupling actionable. Changing a section of
  `excel-library.md` becomes a greppable question — `rg "excel-library.md#what-the-library"
  samples/` — instead of a memory test.

- **Links point one way**, guide → base. The base gains no links back, so the reference does not
  turn into a table of contents for the guide.
- **The rule belongs in [documentation.md](../ai-instructions/documentation.md)** when this ships,
  next to "Keeping the index honest". Otherwise it will be followed for two months and then quietly
  not.

### The snippets are the one place the guide is authoritative

An honest exception, and a useful one. The base documents' code examples are Markdown strings that
nothing compiles. The guide's are compiled code. Where the two disagree, **the compiled one is
right and the base document has the bug.**

A predictable consequence worth planning for rather than discovering: writing
`OfficeDocuments.Guide.Snippets` is the first time the examples in `excel-library.md` and
`word-library.md` will be handed to a compiler. Some of them will not build — the repository has
already shipped one example that threw `ObjectDisposedException`. Fixing those base documents is
part of steps 3 and 4, not a distraction from them.

The long-term resolution is the reverse of today's arrangement: the base documents eventually pull
their examples from the compiled snippets, leaving one copy. That is out of scope here and worth
revisiting once the snippets project exists and has proven itself.

### Format

Prose in **Markdown** files embedded as resources, rendered with a small Markdown library, with YAML
front matter for the topic metadata above. One more package, and it buys a much better authoring and
review experience than prose escaped into Razor markup — which matters more now that the guide's
content is a derivation someone will diff against the base. Razor-only is the fallback if the extra
dependency is unwelcome, but it costs the front matter and the greppability with it.

## UI structure

MudBlazor close to defaults. The library is the subject; the site should not be.

- `MudLayout` with a persistent `MudDrawer` on desktop, temporary on mobile.
- `MudNavMenu`: Concepts, Excel, Excel.Advanced, Word, each with `MudNavGroup` per category,
  generated from `GuideIndex`.
- Topic page: title, summary, then interleaved prose and snippets — code belongs *next to* the
  sentence explaining it, not in a separate tab. (The tabbed layout in the earlier draft existed to
  separate code from a live result; with no result, tabs only add clicks.)
- Prev / next links at the foot of every topic. A guided path needs a path.
- Home: what the library is, install commands, the two-package split, and the first three topics.

**Syntax highlighting.** MudBlazor has no code component. Self-host `highlight.js` (C#-only build)
in `wwwroot` — no CDN, so no third-party origin and it works offline. `CodeBlock.razor` wraps
`<pre><code>` and adds a copy button. Prism is an equivalent alternative. Neither goes in
`Directory.Packages.props`: they are static assets, not NuGet packages.

## Content plan

Ordered as a reader would go through it. "Základ z chování knihovny", as scoped.

The **Base** column is the `source:` front matter each topic will carry — the derivation record, and
the thing to grep when a base document changes. `excel` and `word` abbreviate
`.doc/excel-library.md` and `.doc/word-library.md`; `README` is the repository root readme.

### Concepts

| Topic | Content | Base |
| --- | --- | --- |
| What this is | Positioning, the two Excel packages, install, when to look elsewhere | `README` §Is this the right library for you?, §Install |
| The object models | `ISpreadsheet → IWorksheet → IRange/IRow → ICell`, and the Word tree with `IBlockContainer` as the shared contract | `excel` §Scope, `word` §Scope |
| Correctness | Why a round trip proves nothing, element order as an invariant, what the schema validator does and does not catch | `README` §Correctness |

### Excel — core

| # | Topic | Teaches | Base |
| --- | --- | --- | --- |
| 1 | Hello workbook | `new Spreadsheet(path, createNew)`, `AddWorksheet`, `Close()`; rows and columns are 1-based | `excel` §Create or open a workbook on disk, §Consumer notes |
| 2 | Files and streams | `CreateDocument(Stream)` / `OpenDocument(Stream, isEditable)`; why server-side code wants the stream form | `excel` §Create a workbook in memory; `README` §Excel |
| 3 | Rows and cells | `AddRow` / `AddCell`, `CurrentRow`, and the trap that `AddCell` returns the *cell*, not the row | `excel` §Add formulas (where the trap is noted), §Consumer notes |
| 4 | Styles | `CreateStyle` with font, fill, border, alignment, number format | `excel` §Create and apply styles |
| 5 | Merging styles | `CreateMergedStyle`, and why reuse matters | `excel` §Merge styles; `README` §Known performance characteristics |
| 6 | Bulk insert | `AddRows` from nested collections and from a record collection with `includeHeader: true` | `excel` §Work with ranges and bulk insert |
| 7 | Ranges | `GetRange`, `SetValues`, `GetValues`, `Merge`, `ApplyStyle` | `excel` §Main API surface → `IRange` |
| 8 | Sorting and auto-filter | `ApplyAutoFilter`, `SortByColumn(…, hasHeader: true)` | `excel` §Work with ranges and bulk insert |
| 9 | Formulas | `AddCellWithFormula`; that this writes formulas and does not calculate them; the four functions `GetFormulaValue()` actually evaluates | `excel` §Add formulas, §Consumer notes |
| 10 | Reading values | `TryGetValue<T>`, typed getters, `HasValue` / `HasFormula` | `excel` §Read values back |
| 11 | Columns and panes | `SetColumnWidth`, `AutoFitColumns`, `FreezePanes` | `excel` §Worksheet workflows |
| 12 | **What it refuses to write** | The five refusal rules, the exception each throws, and why failing at the call beats a workbook Excel offers to repair | `excel` §What the library refuses to write, §Line endings |
| 13 | **Dates and 1900** | The phantom 29 February 1900, why `ToOADate` is not used, why reading is more permissive than writing | `excel` §Dates and the 1900 leap-year bug |

12 and 13 are the differentiating topics. They should be written first, not last.

### Excel — Advanced

| # | Topic | Teaches | Base |
| --- | --- | --- | --- |
| 14 | Structured tables | `AddTable`, `TableCreateOptions`, `TableStyleOptions`, and the `using OfficeDocuments.Excel.Advanced;` that unlocks them | `excel` §Create and manage structured tables |
| 15 | Named ranges and protection | `AddNamedRange`, `Protect`, `ProtectWorkbook` | `excel` §Add validation, formatting, hyperlinks, comments, named ranges, and protection |
| 16 | Images | `AddImage` from a file and from a stream with an explicit `ImageType` | `excel` §Embed images in a worksheet |

### Word

| # | Topic | Teaches | Base |
| --- | --- | --- | --- |
| 17 | Hello document | `new Wordprocessing(path, createNew)`, `GetBody`, `AddParagraph` | `word` §A formatted report, §Consumer notes |
| 18 | **Formatting: `null` vs `false`** | Inherit versus active override, `with`, `Merge`. The most misunderstood thing in the Word API | `word` §How formatting works, §Formatting records |
| 19 | Units | Everything in points, and the two exceptions that are not lengths | `word` §Units |
| 20 | Headings and styles | `AddHeading`, `WordStyleIds`, and that a style definition is written on first use | `word` §WordStyleIds, §Definitions, not just references |
| 21 | Lists | Bullet and numbered, nesting levels | `word` §Lists, §Definitions, not just references |
| 22 | Tables | From data, `RepeatAsHeader`, cell formatting, a cell as a block container | `word` §Tables |
| 23 | Hyperlinks and images | `AddHyperlink`, `AddImage` with `ImageSize.FromWidth` | `word` §Hyperlinks, §Images |
| 24 | Page setup, headers, metadata | `ApplyPageSetup`, `AddHeader(HeaderFooterKind.First)`, `SetMetadata` | `word` §Page setup and metadata, §Headers and footers |
| 25 | Reading a document | `GetAllParagraphs`, `FindParagraphs`, `isEditable: false`, `Paragraphs` vs `GetAllParagraphs()` | `word` §Read a document without modifying it, §Navigate an existing document |
| 26 | **Template fill** | `ReplaceText` across run boundaries — the three-run XML that defeats a naive find-and-replace, the three levels of `ReplaceText`, and why the return count is what a template fill should assert on | `word` §Searching and replacing text, §Fill a template |

18 and 26 are Word's differentiating topics.

Topic 2 and topic 19 are new relative to the earlier draft: with no runnable examples, the
stream-versus-file distinction and the units rule have to be taught explicitly rather than absorbed
from watching code run.

### Deferred

Colour validation; text fidelity and `xml:space="preserve"`; conditional formatting and data
validation; Excel hyperlinks and comments; worksheet lifecycle (move, copy, hide); nested tables;
column spanning; the performance topic.

## Verification

Compilation is the gate. `dotnet build OfficeDocuments.slnx` builds the snippets project, so a
signature change that breaks a snippet breaks the build.

**What that does not cover: a snippet that compiles and does not work.** Argument order, a wrong
range reference, a call sequence that throws at runtime — the compiler sees none of it. The earlier
draft caught these by executing every sample; presentation-only gives that up. Two ways to get it
back, both cheap, neither in v1:

- `samples/OfficeDocuments.Guide.Snippets.Tests` — xUnit, one test per snippet: call it, schema-validate
  what it produced with `OpenXmlValidator` against `FileFormatVersions.Office2021`, the same gate the
  test kits use. This keeps execution out of the web app while keeping the guarantee. **Recommended
  as the first follow-up**, and the natural pair to the CI workflow.
- Failing that, snippets are reviewed against the corresponding test in the existing suite, which is
  a human gate and should be written down as such rather than assumed.

Manual gate before merging v1: full-solution build, every topic page opened, snippets spot-checked
against [excel-library.md](../excel-library.md) and [word-library.md](../word-library.md).

## Risks

| Risk | Severity | Mitigation |
| --- | --- | --- |
| A snippet compiles but does not work | **High — this is v1's real weakness** | Stated above; the snippets test project closes it. Until then, no snippet should be written that is not derived from an existing test or documented example |
| Guide prose drifts from the base documents it derives from | **High — inherent to the chosen model** | The base is normative, the guide declares its `source:` in front matter so the coupling is greppable, and the rule goes into `documentation.md`. Not eliminated, made visible |
| Guide rots when the Excel surface changes | Medium | Compile-time reference plus `ReferenceOutputAssembly="false"`; but nothing outside a dev machine builds it in v1 — see [CI](#4-ci) |
| A both-modules project weakens rule 8 | Medium | `OfficeDocuments.slnx` only; no shared Excel+Word abstraction in the snippets project |
| `TargetFrameworks` override done wrong | Low, confusing when hit | Scoped props file with the comment; verified |
| MudBlazor version drift against a new SDK | Low | Pinned centrally; affects no shipped package |

## Open decisions

1. **Public deployment?** Only matters because "GitHub Pages" makes WebAssembly the better render
   model, and switching early is cheaper than switching late. Not blocking for steps 1 and 2.
2. **Czech localization?** Recommendation: English only, consistent with the rest of the repository
   and with a public audience.
3. **Which Markdown package?** Decided in principle (Markdown, not Razor); the specific package is a
   step-2 detail, pinned in `Directory.Packages.props` under the same `Samples` group as MudBlazor.

## Delivery plan

Four increments, each independently mergeable, each leaving the repository buildable.

| Step | Contents | Rough size |
| --- | --- | --- |
| 1 | `samples/Directory.Build.props`; both projects; solution entries; MudBlazor wired up; layout and nav shell; one Excel snippet rendered end to end from embedded source | S |
| 2 | The machinery: `GuideIndex`, `GuideTopic`, region extraction, Markdown rendering, `CodeBlock` with copy, prev/next | S–M |
| 3 | Excel content — topics 1 to 16, the two differentiating ones first. Includes fixing any example in `excel-library.md` that turns out not to compile | M–L |
| 4 | Word content — topics 17 to 26, plus the three concept topics. Same for `word-library.md` | M |

Step 1 de-risks the rest: it proves the single-target override, the CPM entry, and the
embed-and-extract mechanism. If something in this proposal is wrong, it is most likely there.

Recommended step 5, outside this scope: the snippets test project and a CI workflow, together.

## Related documents

- [../excel-library.md](../excel-library.md) — the Excel reference this guide narrates
- [../word-library.md](../word-library.md) — the same for Word
- [../ai-instructions/build-and-packaging.md](../ai-instructions/build-and-packaging.md) — the build
  rules this has to satisfy
- [../ai-instructions/documentation.md](../ai-instructions/documentation.md) — the snippet accuracy
  bar that drives the architecture
- [minimal-core-pr-guidelines.md](minimal-core-pr-guidelines.md) — dependency justification
- [../tasks/roadmap-overview.md](../tasks/roadmap-overview.md) — where this slots in
