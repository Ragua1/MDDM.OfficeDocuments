---
name: Testing guidance
description: xUnit design rules, the schema-validation gate, and assertion strategy. Tier entry criteria live in test/README.md.
applyTo: "test/**/*.cs"
---

# Testing guidance

**Which project does my test go in?** [../../test/README.md](../../test/README.md) is the authority.
It owns the tier table, the entry criteria, and the TestKit helper catalogue. This file owns how a
test is written once you know where it lives.

## The validation gate

Every test that produces a complete document ends with the schema validator:

```csharp
OpenXmlValidation.AssertValid(stream);   // Excel: TestKit; Word: WordTestBase.WriteAndValidate
```

This is not ceremony. A round-trip through this library only proves self-consistency — it cannot see
a schema-order or relationship defect, because the same wrong assumption reads the file back. The
validator caught three real bugs on the day it was introduced, including one that focused unit tests
were structurally incapable of finding.

`AssertValid` runs against `FileFormatVersions.Office2021` and reports each error with its part URI
and XPath. The `inheritedDefects` parameter tolerates named defects that arrived with a foreign input
document; real-world Excel files are not always schema-clean.

## Design rules

- Name tests `MethodName_StateUnderTest_ExpectedOutcome`.
- AAA structure, one primary behaviour per test. Do not assert things unrelated to the test name.
- Use `[Theory]` with inline or member data for the same behaviour across several inputs.
- Keep setup local unless sharing genuinely removes duplication without hiding intent.
- Deterministic only: no wall-clock dependence, no machine-local paths, no locale assumptions, no
  sleeps, no ordering dependence between tests.
- Prefer a `MemoryStream` over a temp file. Reach for `TempWorkspace` only when the test genuinely
  needs a file on disk.

## What to assert

- Assert through the public API first. Testing private implementation detail is an anti-pattern here
  even though `InternalsVisibleTo` makes it possible — that hatch exists for the extracted
  collaborators, not for reaching into coordinators.
- For document behaviour, verify both in memory and after reopen.
- **Do not assert on raw XML strings for values.** `w:val="0"` and `w:val="false"` mean the same
  thing to Word, so a string assertion pins the SDK's formatting rather than this library's
  behaviour. Use `ReadDocumentElement(...)`. String matching is still the right tool for
  *structure*, such as the presence of `xml:space="preserve"`.
- Assert the exact exception type, and assert the constraint that made it throw.
- Keep row/column index assertions explicitly 1-based.
- **Assert that the things that should not happen did not.** This is the blind spot with the worst
  record in this repository. Three Word defects survived every existing test because all of them wrote
  a document and then read it back — the case where saving, appending, and finding content is exactly
  what you want. Nothing asserted that a read-only open leaves the file byte-identical, that
  `Close(saveDocument: false)` discards, or that a one-paragraph table cell reports *one* paragraph.
  All three were real, and one of them made a documented API parameter do nothing at all. For any
  operation with an "and otherwise nothing changes" clause, write that clause down as an assertion.

## Test the input you do not control

A document this library wrote and then read back proves only that it agrees with itself: the same wrong
assumption produces the file and reads it. What breaks read paths in practice is another producer's
markup — Word splits a paragraph's text across runs wherever spell-check state or revision tracking
changes, so a placeholder typed as one word arrives as three runs.

Build that input rather than checking in a binary: `ForeignDocuments` (Word) constructs it through the
SDK, so the fixture is reviewable in a diff and deterministic on every platform, and the Excel
verification tier accepts genuine foreign files with `AssertValid`'s `inheritedDefects`. A permanently
skipped test that depends on a file nobody has is worth less than no test — one sat in the Word suite
from the day it was written until WORD-004 replaced it.

## Overlap strategy

For every critical capability, write two complementary tests: a focused one for the smallest
behaviour and its edge cases, and a scenario test that exercises the same behaviour after
save-and-reopen, alongside neighbouring features. Single-layer coverage produces false confidence —
the merged-style child-order bug was invisible to focused tests and only appeared in a realistic
multi-style workbook.

Critical capabilities: cell value types and formulas; range operations; worksheet lifecycle; table
operations; metadata persistence (hyperlinks, comments, protection, named ranges); Word paragraph and
run formatting; and the Word read and update paths — opening an existing document, replacing text
across runs, and removing structure.

## Coverage

- Cover happy path, edge cases, and invalid arguments.
- Every bug fix gets a regression test that reproduces the original failure.
- Treat coverage percentage as a signal, not a target. Prioritize branch-heavy, behaviour-critical
  code.
