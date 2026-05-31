# WORD-002 Readiness Audit

Date: 2026-05-31

This document summarizes the current readiness assessment for `WORD-002` based on the present implementation of the Word module.

## Current code state

The current public Word surface is still very small:

- `IWordprocessing -> IBody -> IParagraph -> IText`
- `IBody` currently supports only `AddParagraph()` and paragraph reading
- `IParagraph` currently supports only `AddText(...)`, `AddBreak(...)`, `GetTextElements()`, and `GetTexts()`
- `Run` is an internal wrapper rather than a public run-oriented API

Current architectural gaps relevant for `WORD-002`:

- `Body` currently holds only a paragraph model, not a general block model for paragraph and table mixes
- neither `Body` nor `Paragraph` carries document-context access such as `MainDocumentPart`
- hyperlink relationship helpers are missing
- media and image infrastructure for `ImagePart`, sizing, and inline placement is missing
- the current test surface is still minimal and does not cover document relationships or richer block structures

## Readiness conclusions

### Tables

Tables are the most realistic next slice because they can be added as body-level blocks without media parts. Even here, the first iteration should stay small:

- `IBody.AddTable(...)`
- a minimal `ITable` model only if it is truly necessary
- create-and-fill workflow before advanced styling

### Hyperlinks

Hyperlinks are moderately ready. They are less complex than images, but they still need relationship creation on the document part and consistent integration with the paragraph or run model. That strongly suggests stabilizing the run-oriented seam first.

### Images

Images are the least ready slice. They require:

- access to `MainDocumentPart`
- media-part creation
- relationship IDs
- a sizing API
- a clear decision on whether the first scope is inline only or also covers layout options

For that reason, images should remain the last delivery slice inside `WORD-002`.

## Recommended delivery order for `WORD-002`

1. basic tables
2. hyperlinks
3. images

## Backlog impact

`WORD-002` should stay a core backlog area, but it should not be delivered as one large PR. It should be handled as an umbrella item with three smaller delivery slices. The main technical dependency is still the paragraph and run model: hyperlinks and images should not force a second, parallel fluent model.

## Implementation recommendations

- Do not introduce a full Word block-tree framework immediately.
- Start tables with a simple body-level append model.
- Introduce document context for hyperlinks and images only to the degree that is required.
- Keep `WORD-002` materially smaller than a full Word authoring engine.
