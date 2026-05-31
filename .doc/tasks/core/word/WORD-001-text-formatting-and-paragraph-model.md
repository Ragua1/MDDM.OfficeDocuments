# WORD-001 Text Formatting and Paragraph Model

- Module: `OfficeDocuments.Word`
- Priority: `P0`
- Status: `Open`

## Business goal

The current Word module supports text and line breaks, but without formatting it only covers the simplest generated documents. The goal is to enable professional-looking documents such as letters, reports, templates, and records without forcing consumers to work with OpenXml directly.

## Why this belongs in the core backlog

Basic text and paragraph formatting is part of the minimum viable `.docx` authoring surface. Without it, the Word module does not meet even common business-document scenarios. This is a direct usability requirement, not an advanced add-on.

## Functional description

The library should support:

- run-level formatting such as bold, italic, underline, font, size, and color
- paragraph-level formatting such as alignment, spacing, and indentation
- heading-like and paragraph-style scenarios
- the existing fluent usage pattern

Target ergonomics:

- `GetBody().AddParagraph().AddText("Title")...`
- or a similarly readable API if a small options-based model produces a more stable surface

## Technical guidance for GHC

### Public API direction

- Extend `IParagraph`, and introduce `IRun` only if the current model cannot express formatting cleanly.
- Re-evaluate whether `AddText(...)` should continue returning `IParagraph` or whether a run-oriented return value is needed for chained formatting.
- Preserve backward compatibility as far as practical if the fluent chain changes.
- For paragraph formatting, prefer either:
  - focused fluent setters on `IParagraph`
  - or a small `ParagraphStyleOptions`-style object if that produces a cleaner design

### Implementation steps

- Review `src/OfficeDocuments.Word/Interfaces/IParagraph.cs`, `DataClasses/Paragraph.cs`, `Run.cs`, and `Text.cs`.
- Keep the internal model aligned with the WordprocessingML hierarchy: paragraph -> run -> text.
- Do not introduce a broad style engine in the first slice.
- Start with the most common text and paragraph properties.
- If a new public run object is introduced, it must stay easy to test and consistent with the existing fluent model.

### Tests

- Add focused tests for run formatting.
- Add focused tests for paragraph alignment, spacing, and indentation.
- Validate persisted document structure and properties where realistic.

### Documentation

- Update `README.md` only if the top-level overview needs to mention the expanded Word surface.
- Add or update detailed Word examples in [../../../word-library.md](../../../word-library.md).

## Complexity

- Estimate: `M`
- Recommended delivery shape: one iteration for run formatting, one iteration for paragraph formatting

## Risks

- A weak fluent design here will make later Word API work harder.
- An unclear boundary between paragraph and run APIs can cause avoidable future churn.
- An overly ambitious style system would inflate the core without immediate value.

## Dependencies

- no hard dependency on another Word task
- depends on a clear decision between fluent setters and small options-based formatting
- builds directly on the current `Paragraph`, `Run`, and `Text` model

## Subtasks

- [ ] Finalize the text-formatting fluent model.
- [ ] Implement run formatting.
- [ ] Implement paragraph formatting.
- [ ] Add first-iteration heading or paragraph-style scenarios.
- [ ] Add focused tests.
- [ ] Update detailed Word documentation.

## Acceptance criteria

- Consumers can create visually distinct paragraphs without using OpenXml directly.
- The API remains readable and coherent in fluent usage.
- Tests verify the persisted core text and paragraph properties.
