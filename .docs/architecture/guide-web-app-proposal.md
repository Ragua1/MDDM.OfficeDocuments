# Design: the public documentation site

Date: 2026-08-06
Status: **Agreed. Not yet implemented.**

A public documentation site for `OfficeDocuments`: the reference documentation, a guided path
through it, and a generated API reference. Built with DocFX, published to GitHub Pages.

## Decisions

| Decision | Value |
| --- | --- |
| Hosting | GitHub Pages at `ragua1.github.io`, deployed from a GitHub Actions artifact |
| Generator | DocFX, a `dotnet` tool — no application code |
| Content source | `README.md` and `.docs/**` stay in place and stay normative |
| Guided path | Derived pages in `.docs/guide/`, maintained by the [`update-guide` skill](../../.claude/skills/update-guide/SKILL.md) |
| API reference | Generated from the assemblies and their XML comments |
| Language | English only |

Deploying from an Actions artifact means Jekyll never runs, so none of its conventions apply — no
`.nojekyll`, no dot-prefix problem with `.docs/`, no per-page front matter.

The GitHub **wiki** was considered and rejected: it lives in a separate repository, so its content
escapes pull-request review, and this repository enforces its documentation rules there.

## Why DocFX

- **The API reference cannot drift.** The public interfaces carry roughly 200 `<summary>` blocks
  (verified 2026-08-06). DocFX generates the reference from the compiled assemblies, so it is
  derived mechanically from the code — which matters because `EXCEL-007` and `EXCEL-009` will change
  public signatures.
- **Static HTML.** Fast first paint, indexable, no runtime download.
- **A `dotnet` tool.** No Node, no Ruby. Pin it in a tool manifest (`.config/dotnet-tools.json`) so
  it restores like any other dependency.
- **It supplies what would otherwise be hand-built:** the page tree (`toc.yml`), cross-document
  Markdown link resolution, full-text search, and code highlighting.

### Two prerequisites

**`GenerateDocumentationFile` is not set anywhere in this repository**, so no XML documentation file
is produced today and DocFX would have nothing to read. Enabling it also switches on **CS1591 for
every undocumented public member**. Expect a wave of warnings; decide deliberately whether to fix
them, scope the property to the library projects, or suppress CS1591 with a stated reason.

**The library projects multi-target `net8.0;net9.0;net10.0`.** DocFX analyses a project through
Roslyn and needs to be told which target framework to use; left unset it either fails or picks one
silently.

```json
{
  "metadata": [
    {
      "src": [
        { "files": ["src/OfficeDocuments.Excel/*.csproj",
                    "src/OfficeDocuments.Excel.Advanced/*.csproj",
                    "src/OfficeDocuments.Word/*.csproj"] }
      ],
      "dest": "api",
      "properties": { "TargetFramework": "net10.0" }
    }
  ]
}
```

The three library projects only — not the tests, not the benchmarks.

## Content model

| Layer | Files | Rule |
| --- | --- | --- |
| **Source** | `README.md`, `.docs/*.md` | Normative. Authoritative wording for every rule, limit, guarantee, and exception. Unchanged by this work |
| **Derived** | `.docs/guide/*.md` | A guided path. Restructures and narrates. Introduces no fact the source does not state |
| **Generated** | `api/` | From the assemblies. Nobody writes it, nobody edits it, it is not committed |

**The guide need not be 1:1 with the source.** Order, structure, framing, and how much connective
narrative to add are free. Anything normative is not: thresholds, exception types, defaults, names,
and the direction of a rule carry across exactly.

Derived pages are ordinary Markdown, so they stay readable in the repository and go through
pull-request review like everything else.

### Keeping the derivation honest

Each guide page records what it derives from and the commit it was last derived against:

```yaml
---
id: excel/refusals
title: What the library refuses to write
source:
  - .docs/excel-library.md#what-the-library-refuses-to-write
  - .docs/excel-library.md#line-endings
source-revision: 0ddda02
---
```

`source-revision` makes staleness detectable rather than remembered:
`git log <revision>..HEAD -- .docs/excel-library.md` answers "has this page's source moved" with a
command. DocFX ignores unknown front-matter keys, so both fields cost nothing at build time.

The [`update-guide` skill](../../.claude/skills/update-guide/SKILL.md) maintains these pages. It
scopes work from `source-revision`, re-derives affected pages, checks code examples against the real
interfaces rather than against the source documents' uncompiled examples, and reports what needs a
decision.

That check matters because the API reference covers signatures but **not worked examples**, and the
examples in `excel-library.md` and `word-library.md` are Markdown strings that nothing compiles. The
repository has shipped a broken one before — the guides once demonstrated `Close()` inside a
`using`, which throws `ObjectDisposedException`.

## Page tree

Curated in `toc.yml`, never derived from the folder structure.

| Section | Pages |
| --- | --- |
| Getting started | `README.md` |
| Guide | `.docs/guide/*.md` |
| Reference | `.docs/excel-library.md`, `.docs/word-library.md`, `.docs/migration-v3-to-v4.md`, `.docs/terminology.md` |
| API | Generated |
| Comparison | `.docs/library-benchmark-report.md` |
| Performance | `.docs/excel-performance-baseline.md` |
| Contributing | `AGENTS.md`, `test/README.md`, `SUPPORT.md` |

Not published: `.docs/ai-instructions/**` (working rules, decided 2026-08-06), `.docs/tasks/**` (a
backlog that churns and reads as a promise), `.docs/feature-gap-backlog.md` and
`.docs/excel-state-verdict.md` (internal analysis), `.docs/architecture/**` including this file,
`src/OfficeDocuments.*/README.md` (near-duplicates of pages already selected), `CLAUDE.md` and
`LICENSE.md`.

### The guided path, first version

Around ten pages: getting started per module, the object models, styles and formatting, reading —
and the four topics that differentiate this library, written **first**:

- Excel: what the library refuses to write; dates under the 1900 leap-year bug
- Word: `null` versus `false` in formatting; `ReplaceText` across run boundaries

Those four are what this library gets right and its competitors get subtly wrong, and a feature list
cannot communicate them.

## Delivery

| Step | Contents | Size |
| --- | --- | --- |
| 1 | `GenerateDocumentationFile` on the three library projects; count and resolve the CS1591 wave | S |
| 2 | Tool manifest, `docfx.json`, `toc.yml`, local build; verify cross-document links and the API tree | S–M |
| 3 | First guided pages via the `update-guide` skill; fix any source example found to be wrong | M |
| 4 | Publish workflow (`upload-pages-artifact` + `deploy-pages`); theme and branding pass | S |

Step 1 first: it is the only step touching the shipping projects, and the CS1591 count is the one
number that could change the plan. DocFX emits relative links throughout, so serving from
`/MDDM.OfficeDocuments/` needs no base-path handling.

## Risks

| Risk | Severity | Mitigation |
| --- | --- | --- |
| `GenerateDocumentationFile` floods the build with CS1591 | Medium, immediate | Step 1 exists to find the number first |
| DocFX metadata picks the wrong target framework | Medium | `properties.TargetFramework` pinned in `docfx.json` |
| Guide prose drifts from its source | Medium — inherent to a derived guide | `source-revision` plus the `update-guide` skill make it detectable |
| Worked examples in the source docs are wrong | Medium — pre-existing, amplified by publication | Beyond the API reference's reach; covered by the skill's example check |
| Publishing a backlog or internal audit by accident | Low | `toc.yml` is curated explicitly |

## Related documents

- [`.claude/skills/update-guide/SKILL.md`](../../.claude/skills/update-guide/SKILL.md) — how derived
  guide pages are kept in sync with their sources
- [../excel-library.md](../excel-library.md), [../word-library.md](../word-library.md) — the primary
  source content
- [../ai-instructions/documentation.md](../ai-instructions/documentation.md) — where documentation
  lives and the accuracy bar it is held to
- [../ai-instructions/build-and-packaging.md](../ai-instructions/build-and-packaging.md) — the build
  rules step 1 has to satisfy
