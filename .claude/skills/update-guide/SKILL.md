---
name: update-guide
description: Re-derive the guide pages under .docs/guide/ from their source documentation. Use when README.md or a file under .docs/ has changed and the guide may now be stale, when adding a guide page for new source material, or when asked to "update the guide", "sync the guide", "refresh the docs site content", or "check whether the guide is out of date". Also use after any change to a public interface, because guide code examples are derived from the real API and nothing else compiles them.
---

# Update the guide

The guide pages under `.docs/guide/` are **derived**. Their source is the root `README.md` and the
reference documentation under `.docs/`. This skill re-derives them when the source moves.

The derivation is not a copy. Content need not be 1:1 with the source — the guide restructures,
reorders, and narrates. What must survive is the **information**, and in particular every normative
statement the source makes.

## The contract

| Role | Files | Rule |
| --- | --- | --- |
| **Source** | `README.md`, `.docs/*.md` (excluding `.docs/guide/`) | Normative. Holds the authoritative wording for every rule, guarantee, limit, and exception |
| **Derived** | `.docs/guide/*.md` | Re-presentation. Introduces no fact the source does not state |
| **Ground truth for code** | `src/**/Interfaces/*.cs`, the implementations, the test suite | Beats both of the above. See [Code examples](#code-examples) |

Links point one way: guide → source. Never add a link from a source document into the guide.

## Guide page front matter

Every file in `.docs/guide/` carries this. It is the derivation record and this skill depends on it.

```yaml
---
id: excel/refusals
title: What the library refuses to write
section: Excel
order: 120
source:
  - README.md#correctness
  - .docs/excel-library.md#what-the-library-refuses-to-write
  - .docs/excel-library.md#line-endings
source-revision: 0ddda02
---
```

- `source` — every section this page derives from, as `path#anchor`. Anchors use GitHub's slug rules.
- `source-revision` — the commit the page was last derived against. This is what makes staleness
  detectable rather than remembered.

## Procedure

### 1. Scope the work

With no arguments, find every stale page:

```bash
# For each guide page, what changed in its sources since it was last derived?
git log --oneline <source-revision>..HEAD -- README.md .docs/
```

Run this per distinct `source-revision` found in the front matter. A page is stale when any file
named in its `source` list appears in that range. With an argument (`/update-guide
.docs/word-library.md`), restrict to pages whose `source` names that file. With `--all`, re-derive
everything.

Also detect, and report rather than silently handle:

- **A source file with no guide coverage** — new documentation nobody has surfaced.
- **An anchor in `source:` that no longer exists** — a heading was renamed or removed. This is
  either a rename to follow or a deletion to reflect; both need a decision, not a guess.
- **A guide page whose sources are all gone** — the page may need deleting.

### 2. Read before writing

For each stale page, read the **current** source sections in full — not the diff. A diff tells you
what changed; it does not tell you whether the surrounding paragraph still makes the page's framing
correct. Read the guide page too, so the rewrite preserves what is already good.

### 3. Re-derive

Rewrite the affected parts of the page. The derivation rules:

**Free to change.** Order, structure, headings, the amount of connective narrative, what to omit as
distracting for a first read, splitting one source section across two guide pages or merging several
into one, worked scenarios that string several source sections together.

**Not free to change — carry these exactly:**

- **Thresholds and limits.** `1–31 characters`, `1 January 1900`, `nine nesting levels`, `levels 1
  to 6`. Never round, never approximate, never write "about".
- **Exception types.** `ArgumentException` and `ArgumentOutOfRangeException` are not
  interchangeable, and a reader writes a `catch` against what the guide says.
- **Defaults.** `StringComparison.Ordinal`, `createNew`, `isEditable: true`,
  `saveDocument: true`, `includeHeader: false`.
- **Names.** Type, member, parameter, enum, and package names, spelled as the code spells them.
- **Direction of a rule.** "Reading is more permissive than writing" must not become "reading and
  writing differ". `null` means inherit and `false` means an active override — do not soften that
  into "optional".
- **What is out of scope.** If the source says footnotes are not supported, the guide does not imply
  they might be.

**Never introduce.** A behaviour, a guarantee, a limit, or a recommendation the source does not
state. If the guide needs a fact the source lacks, the fix is to add it to the source document
first — and that is a separate, deliberate edit to normative documentation, not a side effect of
this skill.

### 4. Code examples

Guide examples are held to a harder standard than the prose, because the source documents' own
examples are Markdown strings that nothing compiles. **A source example is evidence, not authority.**

For every example in a page you touch:

- Verify each call against the real signature in `src/OfficeDocuments.Excel/Interfaces/*.cs`,
  `src/OfficeDocuments.Excel.Advanced/*.cs`, or `src/OfficeDocuments.Word/Interfaces/*.cs`. Check
  argument order, optional parameters, and the return type — `AddCell` returns the cell, not the row.
- Prefer patterns that appear in the test suite. A pattern exercised by a test is known to work.
- Never call `Close()` inside a `using` block on the same instance. This repository has shipped that
  example before and it throws `ObjectDisposedException`.
- Show the `using` directives when a type comes from a namespace the reader would not guess,
  especially `OfficeDocuments.Excel.Advanced`.
- If a source document's example turns out to be wrong, **fix the source document too** and say so
  in the report. Leaving a known-broken example in normative documentation is not acceptable just
  because the guide worked around it.

### 5. Update the front matter

Set `source-revision` to the current `HEAD` short sha on every page you re-derived. Add or remove
`source:` entries if the derivation now draws on different sections.

Do not touch `source-revision` on a page you did not actually re-read and re-derive. A stamped
revision is a claim that the page was checked against that commit.

### 6. Report

State, in this order:

1. Pages re-derived, and for each, what changed in substance — not "updated wording", but "the
   worksheet name limit changed from 31 to 32 characters".
2. Pages checked and left alone, with the reason (source change was editorial).
3. Source documents fixed, if any, and why.
4. Anything needing a decision: uncovered source files, dead anchors, pages whose source has gone,
   facts the guide wanted that the source does not state.
5. Verification that ran, and anything that did not.

## Verification

- Every `source:` path exists and every anchor resolves to a real heading in that file.
- Every code example checked against the interface, as above.
- Internal links in touched pages resolve.
- English throughout. Absolute dates (`2026-08-06`), never relative.
- If the documentation site is set up, build it and confirm the touched pages render and their
  links resolve.

## What this skill does not do

- It does not write new source documentation. Adding a fact means editing `.docs/` deliberately.
- It does not restructure the guide's page tree. Adding, splitting, or deleting a page is a decision
  to raise in the report, not to take.
- It does not touch the documentation site's code, only its content.

## Related

- [`.docs/ai-instructions/documentation.md`](../../../.docs/ai-instructions/documentation.md) — where
  documentation lives, the language rules, and the snippet accuracy bar
- [`.docs/architecture/guide-web-app-proposal.md`](../../../.docs/architecture/guide-web-app-proposal.md)
  — the site this content feeds, and why the source stays normative
