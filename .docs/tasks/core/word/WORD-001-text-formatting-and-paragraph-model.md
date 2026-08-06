# WORD-001 Text Formatting and Paragraph Model

- Module: `OfficeDocuments.Word`
- Priority: `P0`
- Status: `Delivered` (2026-07-27 — see [Progress log](#progress-log))

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

## Technical guidance

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

- [x] Finalize the text-formatting fluent model.
- [x] Implement run formatting.
- [x] Implement paragraph formatting.
- [x] Add first-iteration heading or paragraph-style scenarios.
- [x] Add focused tests.
- [x] Update detailed Word documentation.

## Acceptance criteria

- Consumers can create visually distinct paragraphs without using OpenXml directly.
- The API remains readable and coherent in fluent usage.
- Tests verify the persisted core text and paragraph properties.

## Progress log

### 2026-07-27 — delivered, after a foundation pass the task did not anticipate

The open design decision in this task — fluent setters versus a small options object — was resolved
in favour of **immutable format records**, `TextFormat` and `ParagraphFormat`, matching the
`Options/*` records already used in the Excel module. Fluent setters on `IParagraph` were rejected
because they cannot express a reusable format: a record can be defined once and varied per call with
`with` or `Merge`, which is what a document with a consistent look actually needs. `AddText(string)`
keeps its existing signature and return type; formatting arrives through overloads.

A `null` property means "inherited", not "off", so `Bold = false` is a real override of a bold style.
That distinction is the reason the properties are nullable rather than defaulted.

**Foundation work done first, because the existing model could not carry the feature.** Two defects
were found by running the code rather than reading it, and both were load-bearing:

- `Body.AddParagraph()` never added the paragraph to `Body.Paragraphs`, which was a snapshot taken
  when the body was wrapped. `GetAllTexts()` therefore returned an empty string for any document this
  library authored. Fixed at the root by projecting the collections from the document tree, so the
  read model cannot drift from the write model (`DataClasses/ElementWrapperList.cs`).
- `Close()` inside a `using` block — the pattern this project's own documentation showed — threw
  `ObjectDisposedException`, because `Dispose()` called `Close()` again and saved an already-disposed
  package. `Close()` is now idempotent.

Also fixed or added in the same pass:

- text fidelity: `xml:space="preserve"` is written when a run needs it, newlines become `w:br`,
  reading no longer trims, and `GetAllTexts()` keeps empty paragraphs
- `w:sectPr` stays the last child of `w:body`, so appending to an opened real document no longer
  produces a file Word has to repair
- `DocumentContext`, the document-scoped seam the [readiness audit](../../../architecture/word-002-readiness-audit.md)
  asked for, now exists and owns style materialization
- built-in style definitions, so `StyleId` and `AddHeading` render instead of only being referenced
- `isEditable` on the constructors, so reading a document does not rewrite it
- dead code removed: `DataClasses/Break.cs`, the commented-out constructor bodies, the unused
  `_isEditable` field
- `test/OfficeDocuments.Word.TestKit` with an `OpenXmlValidator` gate, mirroring the Excel test kit
- the package `README.md` documented an API that never existed (`Wordprocessing.Create`, `doc.Save()`,
  `GetParagraphs()`, `GetText()`); it now matches the real surface

Test count: 3 (1 skipped, 2 smoke tests writing into the working directory) → 67 (1 skipped), green
on net8.0, net9.0, and net10.0.

**Breaking changes**, justified by removing OpenXml leakage from the public surface:

- `IBody.Paragraphs` is `IReadOnlyList<IParagraph>` rather than `List<IParagraph>` — the mutable list
  invited callers to add to a collection that had no effect on the document
- the `Body`, `Paragraph`, and `Run` constructors are `internal`; they previously took raw
  `DocumentFormat.OpenXml.Wordprocessing` types as public parameters
- `Paragraph.RunList` is replaced by `IParagraph.Runs`, typed as `IRun` instead of the concrete
  `Run` wrapper, which had no usable members
- `GetTexts()` and `GetAllTexts()` no longer trim or drop content

Package version raised to `2.0.0`.

**Not in this slice**, and the obvious next additions to the format records: highlight colour,
superscript and subscript, `PageBreakBefore`, `KeepWithNext`, and list or numbering support.
