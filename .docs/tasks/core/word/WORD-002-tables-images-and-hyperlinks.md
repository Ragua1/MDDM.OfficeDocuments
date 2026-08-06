# WORD-002 Tables, Images, and Hyperlinks

- Module: `OfficeDocuments.Word`
- Priority: `P0`
- Status: `Open`

## Business goal

Without tables, images, and hyperlinks, the Word module is difficult to use for invoices, offers, reports, records, formal letters, and branded documents. This task turns the Word module into a practical tool for common business and operational `.docx` scenarios.

This document is the umbrella task for the following delivery slices:

- [WORD-002A Basic tables](WORD-002A-basic-tables.md)
- [WORD-002B Hyperlinks](WORD-002B-hyperlinks.md)
- [WORD-002C Images](WORD-002C-images.md)

## Why this belongs in the core backlog

Tables, images, and hyperlinks are still part of a realistic minimal `.docx` authoring feature set. Without them, the library misses basic business scenarios such as logo-bearing documents, tabular reports, and documents with actionable links.

## Functional description

The library should support:

- creating a table and adding rows and cells
- basic table formatting such as borders, width, and cell alignment
- inserting an image from a stream and from a file
- inserting a hyperlink into paragraph content

## Current readiness audit

The detailed readiness assessment is in [../../../architecture/word-002-readiness-audit.md](../../../architecture/word-002-readiness-audit.md).

The most important conclusions are:

- the current public model is still effectively `Body -> Paragraph -> Text`
- tables are the most realistic first slice because they do not require media parts or hyperlink relationships
- hyperlinks will need document context and likely a more stable run seam from `WORD-001`
- images are the most demanding slice and should come last within `WORD-002`

## Technical guidance

### Public API direction

- Extend `IBody` with `AddTable(...)`.
- Introduce `ITable`, `ITableRow`, and `ITableCell` only if they are genuinely needed; otherwise start with a smaller expandable model.
- For images, decide whether the first API belongs on `IBody`, `IParagraph`, or a future run-oriented seam based on alignment and text-flow needs.
- Hyperlinks should fit naturally into the paragraph or text model.

### Recommended delivery order

1. basic tables
2. hyperlinks
3. images

### Implementation steps

- Work through the WordprocessingML block model rather than assembling raw XML strings.
- Deliver create-and-fill table behavior before more advanced table styling.
- Design hyperlink behavior only after the paragraph or run seam from `WORD-001` is sufficiently clear.
- Add image sizing only in the third slice, once document-part and media infrastructure exists.
- Keep hyperlink display text distinct from the target URL.

### Tests

- Add a test for table creation with multiple rows.
- Add an image-insertion test from a stream.
- Add a hyperlink-insertion test.
- When visual rendering is not practical to validate, validate the document structure and relevant XML elements.

### Documentation

- Update [../../../word-library.md](../../../word-library.md) with example scenarios as slices land.

## Complexity

- Estimate: `M-L`
- Recommended delivery shape: three focused PR slices under one umbrella task

## Risks

- The table object model can become unnecessarily wide if it tries to cover the full Word table system too early.
- Image support can trigger sizing, wrapping, and stream-vs-file API questions.
- Hyperlink behavior must remain consistent with the paragraph or run model.
- Missing document context can lead to poor architecture if images or hyperlinks are added before the run model stabilizes.

## Dependencies

- strong dependency on `WORD-001`, especially for hyperlinks and images
- dependency on a stable body and paragraph block structure
- optional coordination with `WORD-003` if document layout work happens in parallel

## Subtasks

- [ ] Deliver `WORD-002A`.
- [ ] Deliver `WORD-002B`.
- [ ] Deliver `WORD-002C`.
- [ ] Add focused tests and detailed documentation.

## Acceptance criteria

- Consumers can build a document with a table, an image, and a hyperlink.
- The API remains simple enough for everyday business-document authoring.
- Tests verify the resulting document structure and relevant persisted XML.
