# WORD-002C Images

- Module: `OfficeDocuments.Word`
- Priority: `P0`
- Status: `Open`

## Business goal

Images are required for logos, signatures, screenshots, and visual content that commonly appears in generated business documents.

## Why this belongs in the core backlog

Basic inline image support is part of a realistic baseline `.docx` authoring story. It becomes advanced only when layout, wrapping, or richer media-management scenarios substantially widen the feature scope.

## Functional description

The library should support:

- inserting an image from a stream
- inserting an image from a file path
- basic inline placement in the first iteration
- explicit sizing when needed

## Technical guidance for GHC

### Public API direction

- Keep the first iteration focused on inline images only.
- Place the public API where it best matches the settled authoring model from `WORD-001` and `WORD-002B`.
- Avoid starting with a broad image-layout API.

### Implementation steps

- Add the required media-part creation and relationship handling.
- Build the minimum document-context infrastructure needed for image insertion.
- Defer richer placement and wrapping behavior until there is a strong use case.

### Tests

- Add a stream-based image test.
- Add a file-based image test if it can be kept deterministic.
- Validate the image part, relationship, and inline drawing structure.

### Documentation

- Add an image example to [../../../word-library.md](../../../word-library.md).

## Complexity

- Estimate: `M-L`

## Risks

- Image APIs can expand quickly into sizing, wrapping, anchoring, and media-lifecycle questions.
- The feature can force premature document-context abstraction if delivered before the supporting seams are stable.

## Dependencies

- strong dependency on `WORD-001`
- recommended dependency on `WORD-002B` if hyperlink-related document-context work produces reusable infrastructure

## Subtasks

- [ ] Finalize the first inline-image API.
- [ ] Implement media-part creation and relationship handling.
- [ ] Add deterministic tests.
- [ ] Update detailed Word documentation.

## Acceptance criteria

- Consumers can add an inline image from a stream or file without using OpenXml directly.
- The first API remains small and predictable.
- Tests verify the expected document parts and drawing structure.
