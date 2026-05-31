# WORD-002A Basic Tables

- Module: `OfficeDocuments.Word`
- Priority: `P0`
- Status: `Open`

## Business goal

Tables are required for invoices, summaries, protocols, offer documents, and other business documents where structured data must be readable in Word output.

## Why this belongs in the core backlog

Basic table creation is part of normal `.docx` authoring. It delivers clear business value without requiring the broader advanced-feature positioning that heavier layout or media scenarios might need.

## Functional description

The library should support:

- creating a table from the body
- adding rows and cells
- writing text into cells
- applying basic formatting such as width, borders, and simple alignment

## Technical guidance for GHC

### Public API direction

- Prefer a small body-first creation API such as `GetBody().AddTable(...)`.
- Introduce table interfaces only if they improve clarity enough to justify the extra surface.
- Keep the first version focused on creation, fill, and simple formatting.

### Implementation steps

- Model the feature around WordprocessingML table, row, and cell elements.
- Avoid building a full table-style engine in the first iteration.
- Keep the first cell-content workflow aligned with the current paragraph and text model.

### Tests

- Add a test for a multi-row table.
- Add a test for basic table formatting persistence.
- Validate structure first; do not wait for full rendering validation.

### Documentation

- Add a table example to [../../../word-library.md](../../../word-library.md).

## Complexity

- Estimate: `M`

## Risks

- The table surface can expand too quickly if the first version tries to handle every Word table scenario.
- Cell-content rules can become inconsistent if they diverge from the existing paragraph model.

## Dependencies

- benefits from `WORD-001`, but can start with a smaller direct body-level model if needed
- depends on clean paragraph insertion rules for cell content

## Subtasks

- [ ] Finalize the first table API.
- [ ] Implement create and fill behavior.
- [ ] Add basic formatting options.
- [ ] Add focused tests.
- [ ] Update detailed Word documentation.

## Acceptance criteria

- Consumers can create and populate a basic table without using OpenXml directly.
- The first API stays small, readable, and extensible.
- Tests verify the expected table structure and persisted core settings.
