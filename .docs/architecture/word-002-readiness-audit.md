# WORD-002 Readiness Audit

Date: 2026-05-31. Revised 2026-07-28, after the whole Word core backlog — `WORD-002A` through
`WORD-004` — was delivered.

**This document is now a record, not a plan.** It assessed how ready the Word module was for tables,
hyperlinks, and images; all three shipped on 2026-07-27, along with headers, footers, page setup,
metadata, and then search and update. It is kept because its reasoning about delivery order proved
correct and because the conclusions it drew are worth checking against what actually happened.

For the state of the module now, and for what is left as advanced-layer work, see
[../tasks/core/word/README.md](../tasks/core/word/README.md).

For the current state of the Word API, see [../word-library.md](../word-library.md). For what was built
and why, see the progress logs in [../tasks/core/word/](../tasks/core/word/).

## What the audit got right

- **Tables first.** They needed no media parts and no relationships, so they were the cheapest way to
  reach useful block-level content — and extracting a block container for them is what made headers,
  footers, and table cells nearly free afterwards.
- **Stabilise the run seam before hyperlinks.** The audit said hyperlinks "should not force a second,
  parallel fluent model", and the specific defect it named — `IParagraph.Runs` enumerating only direct
  children — was exactly what had to be fixed, because a hyperlink wraps its run in a container.
- **Images last.** They did need the most infrastructure: a media part, a relationship, a four-namespace
  drawing structure, and a sizing decision.

## What the audit did not anticipate

- **Relationships are per-part.** The audit treated "document-context access such as
  `MainDocumentPart`" as the requirement. That was necessary but not sufficient: an image or hyperlink
  inside a header must register on the `HeaderPart`, not on the main document part. Registering on the
  wrong part produces a document that round-trips through this library perfectly and is rejected by
  Word. Only schema validation caught it. `DocumentContext` now derives the owning part from the
  element's position in the tree.
- **The block model mattered more than the media infrastructure.** The audit's main structural concern
  was `Body` holding "only a paragraph model, not a general block model". That turned out to be the
  load-bearing item, but for a reason it did not state: the same block model is needed by table cells
  and by headers and footers, so extracting it once removed most of the work from the two later tasks.
- **A reference is not a feature.** Three of the delivered features are only a pointer in the paragraph,
  with the appearance defined in another part: styles, list numbering, and first- or even-page headers.
  Writing the pointer alone produces a document where nothing appears to have happened. Each needed a
  definition to be materialised on first use.
- **The audit assessed writing, and every remaining defect was on the read side.** It judged readiness
  by what the module could produce, which is why it declared the collection model fixed once
  `ElementWrapperList` existed. `WORD-004` found that the list was still cached and hand-synchronized,
  so it went stale the moment anything *removed* an element — invisible to a suite that only appends.
  It also found an opened document reporting none of its own headers, and a `Close(saveDocument: false)`
  that saved anyway. A readiness audit for an authoring surface does not tell you whether the reading
  surface is ready; those are different questions and want different evidence.

## Original readiness conclusions, for the record

The audit recommended delivering `WORD-002` as three slices in the order tables → hyperlinks → images,
rather than as one large PR, and keeping the scope materially smaller than a full Word authoring engine.
Both held: each slice landed separately with its own tests, and the delivered surface stops well short
of footnotes, bookmarks, tracked changes, and multiple sections.

## Guidance that still applies

- Do not introduce a full Word block-tree framework. `IBlockContainer` covers the four containers that
  genuinely share a content model; anything beyond that needs a concrete requirement.
- Reuse `DocumentContext` for relationship, media, style, and numbering work instead of adding a second
  document-access path.
- Reuse the format-record pattern (`IsEmpty`, `Merge`, nullable properties meaning "inherited") for any
  new formatting surface, so the library keeps one formatting idiom.
- End every test that produces a complete document with `OpenXmlValidation.AssertValid(...)`. It has now
  caught defects in both modules that no round-trip test could see.
