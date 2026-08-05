# Word Core Tasks

Date: 2026-07-28

**The Word core backlog is complete.** Every task below was delivered on 2026-07-27, taking the module
from a ~200-line prototype with 3 tests to 241 tests green on `net8.0`, `net9.0`, and `net10.0`, and the
package from `1.0.0` to `4.0.0`.

Each task document carries a progress log recording what shipped, which decisions were made without
asking, and which bugs were found on the way. Read the log before changing the area it covers — several
of them explain why an obvious-looking approach was rejected.

| Task | Delivered | What it added |
| --- | --- | --- |
| [WORD-001](WORD-001-text-formatting-and-paragraph-model.md) | 2026-07-27 | The formatting model: `TextFormat`, `ParagraphFormat`, built-in styles, and the projected-collection fix the rest of the backlog depended on |
| [WORD-002](WORD-002-tables-images-and-hyperlinks.md) | umbrella | Split into `WORD-002A`, `WORD-002B`, and `WORD-002C` rather than delivered as one PR |
| [WORD-002A](WORD-002A-basic-tables.md) | 2026-07-27 | Tables, and the `IBlockContainer` extraction that made headers, footers, and cells share one block model |
| [WORD-002B](WORD-002B-hyperlinks.md) | 2026-07-27 | External hyperlinks, with the run-nesting fix that `IParagraph.Runs` needed |
| [WORD-002C](WORD-002C-images.md) | 2026-07-27 | Inline images, with intrinsic sizing read from PNG, JPEG, GIF, and BMP headers |
| [WORD-003](WORD-003-headers-footers-sections-and-metadata.md) | 2026-07-27 | Headers and footers, page setup, and document metadata |
| [WORD-004](WORD-004-search-navigation-and-test-hardening.md) | 2026-07-27 | Navigation, search, text replacement across run boundaries, structural removal, and the test hardening that found three read-path defects |

## What the delivery order taught

The backlog was not delivered in its own order, and the deviations were the useful part:

- **`WORD-001` was the right feature but not the right first step.** The underlying object model was
  broken — `Body.AddParagraph()` never reached the collection it was supposed to update — so a
  foundation slice went first.
- **`IBlockContainer` was extracted before tables**, out of order, because a header holds block content
  on exactly the same terms as the body. Doing it once made `WORD-003` nearly free.
- **Three defects were only ever visible from the read side**, and `WORD-004` is where they surfaced:
  a projected collection going stale on removal, an opened document reporting none of its own headers,
  and `Close(saveDocument: false)` saving anyway. All three passed every authoring test.

## If you are picking up Word work now

Nothing here blocks anything. The candidates are all advanced-layer, and none has a concrete
requirement behind it yet:

multiple sections with differing page setups (this turns `PageSetup` from a document-level object into a
per-section one), floating and wrapped images, bookmarks and internal links, footnotes and endnotes,
comments, tracked changes, a generated table of contents, row height, cell-level borders, regular-
expression search, and replacement inside a field result.

Splitting `OfficeDocuments.Word.Tests` into tiers the way Excel is split is the other open option; see
[../../../../test/README.md](../../../../test/README.md).

Before writing code, read [../../../ai-instructions/word.md](../../../ai-instructions/word.md). It holds
the invariants that are not visible from the source: which sequences have no typed SDK setters, why
relationships belong to the referencing part, why a reference is not a feature, and why text is a
property of the paragraph rather than of the run.
