---
name: Word guidance
description: Object model, formatting records, and WordprocessingML child-order rules for the Word library.
applyTo:
  - "src/OfficeDocuments.Word/**/*.cs"
  - "test/OfficeDocuments.Word.*/**/*.cs"
---

# Word guidance

`src/OfficeDocuments.Word` is the smaller, actively growing module. Keep changes additive and
focused; do not import Excel's architecture to make the two modules symmetrical.

## Object model

- Hierarchy: `IWordprocessing → IBlockContainer → IParagraph → IRun` / `IText`, with
  `IBody`, `IHeaderFooter`, and `ITableCell` all implementing `IBlockContainer`, and
  `ITable → ITableRow → ITableCell` for tables.
- Fluent authoring stays the primary usage pattern: `GetBody() → AddParagraph() → AddText(...)`.
- **New block-level content goes on `IBlockContainer`, not on `IBody`.** The body, a header, a footer,
  and a table cell hold block content on identical terms, and `DataClasses/BlockContainer.cs`
  implements it once for all four. Adding to `IBody` alone means headers and cells silently lack the
  feature.
- **Never keep a copy of the child order.** Every collection — `IBlockContainer.Paragraphs`, `.Tables`,
  `IParagraph.Runs`, `ITable.Rows`, `ITableRow.Cells` — is read from the document on each access through
  `DataClasses/ElementWrapperList.cs`, which caches only the facade per element. This bug has now been
  fixed twice, the second time because caching *part* of a derived value is still duplication: WORD-001
  built the list once and never updated it, so `GetAllTexts()` returned `""` for every authored
  document; WORD-004 found the same list going stale on removal, so `ITableCell.SetText` made a
  one-paragraph cell report two. If you add an operation that removes or reorders elements, nothing
  needs updating — keep it that way.
- Document-scoped work — relationships, media parts, style and numbering definitions, drawing ids —
  goes through `DataClasses/DocumentContext.cs`. Do not open a second path to `MainDocumentPart`.
- Text content goes through `DataClasses/RunContent.cs`, which handles `xml:space="preserve"` and
  translates newlines into `w:br`. Do not construct `w:t` directly.
- **The library, not the SDK, decides when the package is written.** Documents are opened and created
  with auto-save off, so `Close(saveDocument:)` means something. Reinstating auto-save silently breaks
  the discard path: skipping the explicit `Save()` does nothing when disposal saves anyway.

## Text is a property of the paragraph, not of the run

A run boundary in WordprocessingML carries no meaning — Word starts a new run wherever spell-check
state or revision tracking changes, so a placeholder typed as one word arrives as three runs. Any
feature that searches, measures, or rewrites text has to work on the flattened paragraph:

- `RunContent.Enumerate` is the one definition of what a document's text is and which element produced
  which characters. `RunContent.Read` and `DataClasses/TextReplacer.cs` both derive from it. Do not add
  a third walk — an offset computed against one notion of the text is meaningless against another.
- `TextReplacer` plans every edit against the text as read, then rewrites each element once. Editing
  while walking the matches does not work, because rewriting an element detaches it and a second match
  inside it then points at markup the document no longer has.
- A match must not cross a paragraph boundary. Replacing across one would have to merge or delete a
  paragraph, which is a structural edit disguised as a text edit.

## Relationships belong to the part that references them

`r:id="rId4"` means something different in `document.xml` than in `header1.xml`. An image or hyperlink
inside a header must be registered on that `HeaderPart`; registering it on the main document part
produces an id the header cannot resolve and Word reports the file as corrupt.

`DocumentContext` resolves the owning part by walking the element up to its tree root and asking that
root which part it belongs to. Pass the content element to `AddImagePart(...)` and
`CreateExternalRelationship(...)` and the right part is found automatically — do not reintroduce an
assumption that the main document part owns everything.

This defect reached the test suite once and was caught only by schema validation: a round-trip through
this library resolves nothing and reads as perfectly consistent.

The same principle applies to reading. `Wordprocessing.HeadersAndFooters` derives the list from the
`w:headerReference` and `w:footerReference` children of `w:sectPr`, not from what this instance created.
Reporting only session-created containers meant every document opened from disk claimed to have no
headers, and nothing about the result looked wrong — the headers were simply never visited.

## Child order is not optional

Rule 3 in [AGENTS.md](../../AGENTS.md) explains why this matters. WordprocessingML sequences are
especially strict and especially easy to break by accident:

- `w:sectPr` must be the **last** child of `w:body`. Append block content through
  `BlockContainer.AppendBlock`, which `Body` overrides to preserve that. A plain `AppendChild`
  produces a document that is invalid on disk.
- `w:pPr` must be the first child of `w:p`, `w:tblPr` and `w:tcPr` the first child of `w:tbl` and
  `w:tc`, and the children of `w:rPr`, `w:pPr`, `w:style`, and `w:lvl` each follow a fixed sequence.
- **Build properties by assigning the SDK's typed properties, never by `AppendChild`:**

  ```csharp
  // Correct — the typed setter places the child at its schema position.
  runProperties.Bold = new Bold();
  paragraphProperties.Justification = new Justification { Val = JustificationValues.Center };

  // Wrong — compiles, round-trips, and fails to open in Word.
  runProperties.AppendChild(new Bold());
  ```

  This is the single rule that prevents the child-order bug class that hit Excel three times.
- Two sequences have **no** typed SDK properties and so need explicit orderers:
  `w:sectPr` (`DataClasses/SectionPropertiesOrderer.cs` — references, then `w:pgSz`, then `w:pgMar`,
  with `w:titlePg` much later) and `w:numbering` (all `w:abstractNum` before all `w:num`, handled in
  `DocumentContext`).

## Formatting

The consumer-facing contract — what `null` versus `false` means, the points-in/OOXML-units-out rule,
the colour rules, and what `Merge` does — is documented once in
[../word-library.md](../word-library.md). Do not restate it here; do preserve it. What that means
when you write code:

- Formatting is expressed by immutable records in `Formatting/`: `TextFormat`, `ParagraphFormat`,
  `TableFormat`, `TableCellFormat`, `PageSetup`, `DocumentMetadata`. Extend those rather than adding
  fluent setters — only a record stays reusable across many runs and paragraphs.
- Every record carries `IsEmpty` and `Merge(overrides)`. A new record must have both, and a new
  property must be nullable so the inherited/off distinction survives.
- `Format` properties stay get-only; changes go through `ApplyFormat(...)`, which merges. This
  deliberately avoids Excel's set-only `Styles.*` weakness and the ambiguity of assignment semantics.
- Convert units in `Formatting/Measure.cs` (half-points, twips, line units) and `InlineImageBuilder`
  (EMU). Validate colours through `Formatting/HexColor.cs`, and validate measurements in the
  record's `init` accessor so an invalid value throws where the caller wrote it.

## Definitions have to exist before a reference means anything

Three features write only a *pointer*, with the appearance defined elsewhere in the package; writing
the pointer alone produces a document that looks like nothing happened. The consequences for a
consumer are in [../word-library.md](../word-library.md); the seams you have to go through are:

| Reference | Definition | Handled by |
| --- | --- | --- |
| `w:pStyle`, `w:rStyle` | a `w:style` in the styles part | `DocumentContext.EnsureStyle` from `Formatting/BuiltInStyles.cs` |
| `w:numPr` | a `w:abstractNum` + `w:num` in the numbering part | `DocumentContext.EnsureListNumbering` from `Formatting/ListNumbering.cs` |
| `w:headerReference` for `First` / `Even` | `w:titlePg` on the section / `w:evenAndOddHeaders` in settings | `Wordprocessing.EnableHeaderFooterKind` |

Never invent a definition for an identifier the library does not know — writing it through untouched
is what lets a template keep its own styles.

## Images

- Intrinsic sizing reads the file's own header (`Formatting/ImageMetadata.cs`, covering PNG, JPEG, GIF,
  and BMP). `System.Drawing` is deliberately not used — it is Windows-only on .NET Core and this
  library runs on Linux in CI.
- When the header cannot be read, fail with a message telling the caller to pass an `ImageType` and an
  `ImageSize.Exact(...)`. Never guess a size.
- Resolution round-trips are lossy by nature: PNG stores it as a whole number of pixels per metre, so
  300 DPI is not exactly recoverable. Assert sizes to a tolerance, not exactly.

If a future feature needs to decode or transform an image rather than read its header, the replacement
for `System.Drawing` is **SkiaSharp** (MIT, Microsoft-maintained). Not `SixLabors.ImageSharp`: it moved
to the Six Labors Split License in v3, which would push a commercial-licensing obligation onto every
consumer of this package.

## Tests over foreign input

A document this library wrote and then read back only proves it agrees with itself. Use
`TestKit/ForeignDocuments.cs` for the read and update paths: it builds the run splitting that is the
defining property of real Word markup, through the SDK rather than as a checked-in binary, so the input
stays reviewable in a diff and deterministic on every platform. Its documents carry `w:sectPr`, which is
what makes an append bug fail rather than pass.

Assertions that something did *not* happen are worth more here than assertions that it did. The
`Close(saveDocument: false)` defect survived every existing test because they all wrote and then read —
the scenario where saving is what you want.

## Out of scope for now

The list is in [../word-library.md](../word-library.md). One consequence matters for design decisions
here: `PageSetup` describes a single section, so adding multi-section support turns it into a
per-section object rather than a document-level one. Do not build against it as if it were already
per-section.
