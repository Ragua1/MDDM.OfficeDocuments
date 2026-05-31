# WORD-002B Hyperlinks

- Module: `OfficeDocuments.Word`
- Priority: `P0`
- Status: `Open`

## Business goal

Hyperlinks are required for actionable reports, letters, documentation, and generated documents that must reference websites, tickets, portals, or knowledge resources.

## Why this belongs in the core backlog

Hyperlinks are a normal part of business-document authoring and do not represent an advanced or optional niche feature.

## Functional description

The library should support:

- adding a hyperlink to document content
- keeping display text separate from the target URL
- supporting typical external-link scenarios in the first iteration

## Technical guidance for GHC

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

- [ ] Finalize hyperlink placement in the public fluent API.
- [ ] Implement relationship creation and document insertion.
- [ ] Add focused tests.
- [ ] Update detailed Word documentation.

## Acceptance criteria

- Consumers can add a hyperlink with display text and a target URL through the public API.
- The API remains fluent and consumer-oriented.
- Tests verify both relationship creation and document structure.
