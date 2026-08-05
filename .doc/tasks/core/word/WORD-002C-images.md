# WORD-002C Images

- Module: `OfficeDocuments.Word`
- Priority: `P0`
- Status: `Delivered` (2026-07-27 — see [Progress log](#progress-log))

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

## Technical guidance

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

- [x] Finalize the first inline-image API.
- [x] Implement media-part creation and relationship handling.
- [x] Add deterministic tests.
- [x] Update detailed Word documentation.

## Acceptance criteria

- Consumers can add an inline image from a stream or file without using OpenXml directly.
- The first API remains small and predictable.
- Tests verify the expected document parts and drawing structure.

## Progress log

### 2026-07-27 — delivered

API: three `IParagraph.AddImage(...)` overloads — from a stream with the format inferred, from a stream
with an explicit format, and from a file path with the format inferred from the extension — plus the
`ImageSize` factory (`Intrinsic`, `Exact`, `FromWidth`, `FromHeight`).

`Formatting/InlineImageBuilder.cs` builds the whole `w:drawing → wp:inline → a:graphic → pic:pic`
structure across four namespaces. This is the clearest single case for the library existing: about a
dozen elements are required before an image appears, and Word rejects the document if any is missing.

**Intrinsic sizing reads the image's own header** (`Formatting/ImageMetadata.cs`), covering PNG, JPEG,
GIF, and BMP including resolution where the format records it. `System.Drawing` was rejected for this:
it is Windows-only on .NET Core and the CI builds run on Linux, so an image library would have been a
new dependency for a document library. Content whose header cannot be read still works, but the caller
must then pass both the type and an `ImageSize.Exact(...)`, and the exception says so.

**A relationship bug was found by schema validation, not by a round-trip.** An image added to a header
registered its `ImagePart` on the main document part, so the `r:embed` id did not resolve from
`header1.xml` and Word would report the file as corrupt. A round-trip through this library could never
detect it — reading back resolves nothing and the document is self-consistent. `DocumentContext` now
derives the owning part by walking the element to its tree root and asking that root which part it
belongs to, rather than assuming the main document part owns everything. That mirrors how the
collections avoid drift: derive from the tree instead of carrying a duplicate.

Drawing identifiers are seeded above whatever an opened document already uses, across the main document
and all header and footer parts, so appending to a document cannot collide.

Inline placement only, per this task's scope. Floating and wrapped images, cropping, and effects are not
implemented.

19 document tests plus 12 direct tests of the header reader — the parser is pure logic over four binary
layouts and the document tests only exercise the PNG path.
