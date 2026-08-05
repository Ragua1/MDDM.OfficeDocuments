# WORD-004 Search, Navigation, and Test Hardening

- Module: `OfficeDocuments.Word`
- Priority: `P1`
- Status: `Delivered`

## Business goal

Once the Word authoring surface grows, consumers will need safer read and update workflows. This task strengthens the ability to inspect, navigate, and evolve existing `.docx` documents while also improving confidence in the module through stronger tests.

## Why this belongs in the core backlog

Reliable read and update behavior is necessary if the Word module is expected to support more than one-shot document creation. Hardening tests and navigation APIs supports everyday use, not just advanced extensions.

## Functional description

The library should support:

- finding and navigating key document content more predictably
- safer updates to existing document structures
- stronger test coverage for both write and read paths

## Technical guidance

### Public API direction

- Prefer small search and navigation helpers over a large query framework.
- Keep the first iteration aligned with the existing body, paragraph, and text model.
- Strengthen read and update scenarios only where the public API can stay understandable.

### Implementation steps

- Expand the current read-path coverage in the Word tests.
- Identify the smallest useful navigation seams for existing documents.
- Avoid overdesigning a full document-query abstraction.

### Tests

- Add coverage for realistic open-read-update scenarios.
- Add regression tests for the most important document structures introduced by earlier Word tasks.
- Prefer deterministic XML-structure validation over environment-specific rendering checks.

### Documentation

- Update [../../../word-library.md](../../../word-library.md) if new read or navigation APIs are added.
- Keep `README.md` high-level unless the module positioning materially changes.

## Complexity

- Estimate: `M`

## Risks

- Read-path APIs can become too broad if they try to model full Word search semantics.
- Test growth can become noisy unless it stays focused on the supported public workflow.

## Dependencies

- recommended dependency on `WORD-001` through `WORD-003`
- depends on the evolving shape of the public Word authoring model

## Subtasks

- [x] Expand read and update regression coverage.
- [x] Finalize the first navigation helpers.
- [x] Add focused implementation coverage for earlier Word tasks.
- [x] Update detailed documentation if the public API changes.

## Acceptance criteria

- The Word module has materially stronger read and update confidence than before.
- New navigation helpers stay small and understandable.
- Tests cover the main supported write and read workflows.

## Progress log

### 2026-07-27 — delivered

Package `3.0.0` → `4.0.0`. Word tests 199 → 241, all green on `net8.0`, `net9.0`, and `net10.0`, with
the one permanently skipped test replaced rather than left in place. Whole solution: 598 tests per
framework, zero failures, zero skipped.

**Three defects found, all in the read and update paths the earlier tasks never exercised.** Each was
confirmed by a test that failed first.

1. **Projected collections went stale on removal.** `ElementWrapperList` built its list from the tree
   once and then hand-synchronized additions. That held while content could only be appended and broke
   as soon as anything removed an element: `ITableCell.SetText` removes the cell's paragraph and adds
   another, so a cell whose `Paragraphs` had been read once reported two paragraphs for one. This is the
   WORD-001 bug in a second form, and the lesson is that caching *part* of a derived value is still
   duplication — it only makes the drift rarer. The cache is gone; `Items` reads the tree on each access
   and only the facade per element is cached, so object identity still holds.

2. **`HeadersAndFooters` reported only session-created containers.** A document opened from disk claimed
   to have none, so every read and template workflow silently skipped its headers. Now derived from the
   `w:headerReference` and `w:footerReference` children of `w:sectPr`, with the wrapper cache keyed by
   kind so `AddHeader` and `HeadersAndFooters` hand back the same instance.

3. **`Close(saveDocument: false)` could not discard anything.** The SDK's `AutoSave` defaults to on and
   writes the package on disposal, so the parameter skipped a redundant explicit `Save()` and then saved
   regardless. The decision belongs at open time: documents are now created and opened with auto-save
   off, and `Close` is the only thing that writes. Notable for *why* no existing test caught it — they
   all wrote and then read, which is the case where saving is wanted. Negative assertions are what found
   it.

**Cross-run text replacement** is the substance of the task. A run boundary in a `.docx` carries no
meaning: Word starts a new run wherever spell-check state or revision tracking changes, so `{{customer}}`
is routinely stored as `Dear {{` + `customer` + `}}, thank you.` and a per-run search finds nothing. The
implementation flattens the paragraph once through `RunContent.Enumerate` — the single definition of what
the text is and which element produced which characters, shared with `RunContent.Read` so the two cannot
disagree — then maps each match back to elements.

The first implementation edited as it walked the matches right to left, which is correct for offsets and
wrong for element identity: rewriting an element detaches it, so a second match inside the same element
found a parent-less node and did nothing. Four tests caught it. Split into plan-then-apply: every offset
is computed against the pristine text, then each element is rewritten exactly once.

Behaviour worth recording, all covered by tests: the replacement takes the formatting of the run the
match starts in; a run the replacement empties is removed so repeated template fills do not accumulate
litter; a match may run through a `w:br` or `w:tab` and a `\n` in the replacement becomes a real break;
`xml:space="preserve"` is reapplied; a match never crosses a paragraph boundary; and the non-ordinal
`StringComparison` values go through `CompareInfo.IndexOf(..., out int matchLength)` because a
culture-sensitive comparison can match a span of a different length than the pattern.

**Public API added.** Deliberately small, per the task's own warning about a query framework:

| Type | Members |
| --- | --- |
| `IParagraph` | `SetText`, `ReplaceText` |
| `IBlockContainer` | `GetAllParagraphs`, `FindParagraphs`, `ReplaceText`, `Remove(IParagraph)`, `Remove(ITable)` |
| `ITable` | `Remove(ITableRow)` |
| `IWordprocessing` | `ReplaceText` |

`GetAllParagraphs` descends into table cells at any depth while `Paragraphs` stays this container's own
children — the two answer different questions, and having both is what keeps a document-wide sweep from
needing to know the table structure. `Remove` verifies ownership and returns `false` rather than taking
an element out of a container the caller did not name.

**Test hardening.** `TestKit/ForeignDocuments.cs` builds documents the way another producer writes them —
split runs, `w:proofErr` markers, and the `w:sectPr` every real document carries — through the SDK rather
than as a checked-in binary, so the input is reviewable in a diff and deterministic across platforms.
This replaces `ReadExternalDocument_ReturnsParagraphText`, which depended on a file outside the
repository and had been skipped since it was written. New files: `TextReplacementTest` (18),
`NavigationTest` (14), `ReadAndUpdateTest` (9).

**Not done, and not attempted:** regular-expression search, a document-query abstraction, replacement
that crosses paragraphs, and replacing text inside a field result. The remaining Word gaps are unchanged
and listed in [../../../word-library.md](../../../word-library.md).
