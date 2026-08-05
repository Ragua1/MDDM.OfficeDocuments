# WORD-003 Headers, Footers, Sections, and Metadata

- Module: `OfficeDocuments.Word`
- Priority: `P0`
- Status: `Delivered` (2026-07-27 — see [Progress log](#progress-log))

## Business goal

Formal documents often need repeated branding, page framing, section behavior, and document metadata. This task supports more complete business-document output such as letters, statements, contracts, and reports.

## Why this belongs in the core backlog

Headers, footers, sections, and metadata are still common document-authoring features, especially for branded or formal output. They are part of a credible minimal Word story once the main paragraph and content model is usable.

## Functional description

The library should support:

- adding headers and footers
- setting basic document metadata
- section-level scenarios such as page setup boundaries or repeated structural content
- keeping the public API understandable and not overly layout-centric

## Technical guidance

### Public API direction

- Keep the API task-oriented and consumer-friendly.
- Start with a smaller model that covers the most common business scenarios.
- Avoid trying to expose the full Word section model in the first iteration.

### Implementation steps

- Add the minimum part and relationship handling required for headers and footers.
- Introduce metadata helpers only for fields that clearly matter to consumers.
- Keep section behavior scoped to practical first-iteration scenarios.

### Tests

- Add tests for header and footer creation.
- Add tests for core metadata persistence.
- Validate the underlying document-part relationships and the persisted document structure.

### Documentation

- Add practical header, footer, and metadata examples to [../../../word-library.md](../../../word-library.md).

## Complexity

- Estimate: `M-L`

## Risks

- The Word section model can grow quickly in scope.
- Metadata, sections, and repeated content can introduce awkward API seams if they are all designed at once.

## Dependencies

- recommended dependency on `WORD-001`
- partial dependency on `WORD-002`, depending on whether shared document-context infrastructure is introduced

## Subtasks

- [x] Finalize the first header and footer API.
- [x] Implement practical metadata support.
- [x] Add a small first section model if needed.
- [x] Add focused tests.
- [x] Update detailed Word documentation.

## Acceptance criteria

- Consumers can create branded or formal documents with basic repeated structure and metadata.
- The API remains readable and focused on common business scenarios.
- Tests verify the expected parts, relationships, and persisted settings.

## Progress log

### 2026-07-27 — delivered

API on `IWordprocessing`: `AddHeader(kind)`, `AddFooter(kind)`, `HeadersAndFooters`, `PageSetup` /
`ApplyPageSetup(...)`, and `Metadata` / `SetMetadata(...)`. Headers and footers are `IHeaderFooter`,
which is an `IBlockContainer` — so they took paragraphs, headings, lists, tables, and images for free
from the extraction done in `WORD-002A`.

**Two silent-failure traps are handled**, both of which produce a perfectly valid document in which the
header simply never appears:

- a first-page header needs `w:titlePg` on the section
- an even-page header needs `w:evenAndOddHeaders` in the document settings part

`AddHeader` sets whichever applies, so choosing a `HeaderFooterKind` is enough.

`w:sectPr` has no typed SDK properties for its children, so `DataClasses/SectionPropertiesOrderer.cs`
applies the schema sequence explicitly: references first, then `w:pgSz`, then `w:pgMar`, with
`w:titlePg` much later. A test deliberately applies the page setup *before* adding the header to prove
the order holds regardless of call order.

`ApplyPageSetup` swaps the stored width and height for landscape rather than only setting `w:orient`;
setting the attribute alone leaves Word laying text out on a portrait page. Creating the margin element
seeds Word's defaults, so setting one margin does not zero the others.

Metadata maps to the package core properties, which is what a document management system or a search
index reads. `Author` maps to `dc:creator` and `Description` to `dc:description`, matching the labels
Word shows.

`AddHeader`/`AddFooter` are idempotent per kind and reuse an existing part when a document is reopened,
so appending to a document does not orphan its original header behind a second reference.

**Sections are deliberately singular.** `PageSetup` describes the one body-level section, per this
task's guidance to avoid exposing the full Word section model. Multiple sections with different page
setups remain out of scope; if that changes, `PageSetup` becomes a per-section object.

26 tests added.
