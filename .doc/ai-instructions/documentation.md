---
name: Documentation guidance
description: Where each kind of documentation lives, language and terminology rules, and the accuracy bar for code snippets.
applyTo: "**/*.md"
---

# Documentation guidance

## Where it goes

| Content | Location |
| --- | --- |
| Short product overview, what the library is, how to get started | Root [`README.md`](../../README.md) |
| Consumer API guides and examples | [`../excel-library.md`](../excel-library.md), [`../word-library.md`](../word-library.md) |
| NuGet package readme | `src/OfficeDocuments.*/README.md` (shipped in the package) |
| Architecture decisions and audits | [`../architecture/`](../architecture/) |
| Backlog, task specs, roadmap | [`../tasks/`](../tasks/README.md) |
| Test-tier rules | [`../../test/README.md`](../../test/README.md) |
| Rules for AI agents | [`./`](README.md) |

The root `README.md` stays short and product-oriented. Depth belongs in `.doc/`. When you add a
public API, update the detailed guide; touch the root README only if the high-level product story
actually changed.

## Language and terminology

- All documentation is written in **English**, in clear, technical, unambiguous prose.
- Use the terms in [`../terminology.md`](../terminology.md). If you need a term that is not there,
  add it there rather than inventing a synonym in one document.
- Prefer relative links between files inside `.doc/`, so the documentation stays portable.
- Dates are absolute (`2026-07-27`), never relative.

## Accuracy bar

- **Every snippet must match the real public API.** Verify against the interface, the implementation,
  and a test before copying a pattern forward. A snippet that no longer compiles is a bug report
  waiting to happen — the project's own docs once demonstrated calling `Close()` inside a `using`,
  which threw `ObjectDisposedException`.
- Prefer short, copy-pastable examples that create, use, and dispose a document correctly.
- Consumer documentation should cover the whole workflow, not just creation: reading values, applying
  styles, working with ranges, configuring columns, adding tables, and correct disposal.

## Keeping the index honest

When you add a file under `.doc/`, add it to [`../README.md`](../README.md). An unlinked document
will not be found — by a human or by an agent — and will quietly rot.
