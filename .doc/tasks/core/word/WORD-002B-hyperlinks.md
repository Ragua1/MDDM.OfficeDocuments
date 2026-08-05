# WORD-002B Hyperlinks

- Module: `OfficeDocuments.Word`
- Priority: `P0`
- Status: `Delivered` (2026-07-27 — see [Progress log](#progress-log))

## Business goal

Hyperlinks are required for actionable reports, letters, documentation, and generated documents that must reference websites, tickets, portals, or knowledge resources.

## Why this belongs in the core backlog

Hyperlinks are a normal part of business-document authoring and do not represent an advanced or optional niche feature.

## Functional description

The library should support:

- adding a hyperlink to document content
- keeping display text separate from the target URL
- supporting typical external-link scenarios in the first iteration

## Technical guidance

### Public API direction

- Integrate hyperlinks naturally into the paragraph or run model.
- Avoid a detached hyperlink API that does not fit the current fluent authoring flow.
- Support the common external-link case first.

### Implementation steps

- Introduce the required relationship handling on the document part.
- Reuse the result of `WORD-001` so hyperlink insertion aligns with the chosen paragraph or run seam.
- Keep the first iteration small and predictable.

### Tests

- Add a test that writes a hyperlink and verifies the relationship plus the relevant document elements.
- Add at least one test for hyperlink display text.

### Documentation

- Add a hyperlink example to [../../../word-library.md](../../../word-library.md).

## Complexity

- Estimate: `M`

## Risks

- Hyperlink behavior can become awkward if the paragraph or run seam is not settled first.
- Relationship handling can leak document internals if the API is shaped too low-level.

## Dependencies

- strong dependency on `WORD-001`
- depends on document-part access and stable paragraph or run composition

## Subtasks

- [x] Finalize hyperlink placement in the public fluent API.
- [x] Implement relationship creation and document insertion.
- [x] Add focused tests.
- [x] Update detailed Word documentation.

## Acceptance criteria

- Consumers can add a hyperlink with display text and a target URL through the public API.
- The API remains fluent and consumer-oriented.
- Tests verify both relationship creation and document structure.

## Progress log

### 2026-07-27 — delivered

API: `IParagraph.AddHyperlink(text, url, format?)`, returning the paragraph so it composes into the
existing fluent chain.

The blocker this task's readiness audit predicted was real and is fixed: **`IParagraph.Runs` now reads
descendants rather than direct children**, because a hyperlink wraps its run in a `w:hyperlink`
container. Without that change a link's text was invisible to `Runs` — `GetTexts()` already walked
descendants, so the two disagreed. `ElementWrapperList` now takes a reader delegate, which is what made
this a one-line change per collection rather than a redesign.

Also delivered:

- `TextFormat.StyleId` for character styles, added because a hyperlink needs one. It makes hyperlink
  runs readable through the normal `Format` property instead of being a special case.
- The built-in `Hyperlink` character style, so a link is blue and underlined rather than plain text.
  `BuiltInParagraphStyles` was renamed `BuiltInStyles` and now handles character styles;
  `DocumentContext.EnsureParagraphStyle` became `EnsureStyle` to match.
- A caller's format layers **over** the hyperlink style rather than replacing it, so a recoloured link
  is still recognisably a link.

Only absolute external targets are accepted, per this task's scope; a relative or malformed URL throws.
Internal links to bookmarks are not implemented, since there are no bookmarks yet.

14 tests added, covering the relationship, the reference that points at it, the style definition, and
the `Runs` enumeration through the container.
