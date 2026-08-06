---
name: Excel guidance
description: Object model, style pipeline, element-order invariants, and module layout for the Excel library.
applyTo:
  - "src/OfficeDocuments.Excel/**/*.cs"
  - "test/OfficeDocuments.Excel.*/**/*.cs"
---

# Excel guidance

`src/OfficeDocuments.Excel` is the primary, mature module. It carries more invariants than Word, and
most of them exist because a bug got through.

## Object model

- Hierarchy: `ISpreadsheet → IWorksheet → IRow → ICell`, with `IRange` as the range-centric seam.
- The public boundary is `Interfaces/*`. OpenXml manipulation stays inside implementation classes.
- `Spreadsheet` owns `SpreadsheetDocument`, `WorkbookPart`, and the stylesheet lifecycle. Do not
  open a parallel path to the document.
- Row and column indexes are **1-based**. Invalid indexes throw `ArgumentException`.
- `Row.CreateCell` backfills missing earlier cells on purpose — OpenXml requires ascending cell
  order within a row. Do not replace it with sparse append logic.
- `Worksheet` creates `Columns` and `MergeCells` lazily. Preserve that; eager creation writes empty
  elements into every sheet — and an empty `<mergeCells/>` is not merely untidy, `CT_MergeCells`
  requires at least one child, so it fails the schema validator. Read the existing merges with
  `GetFirstChild<MergeCells>()`; touching the `MergeCells` property creates the element.
- Merged ranges must not overlap. Excel reports a workbook whose merges share a cell as damaged, so
  `AppendMergeReference` rejects an overlap with `ArgumentException` while treating an exact repeat
  as a no-op. A range covering one cell is not a merge and writes no element.

## Module layout after EXCEL-010

`Spreadsheet` and `Worksheet` are **coordinators**, split across partial files, delegating to
internal collaborators in `DataClasses/`:

| Partial files | Collaborators |
| --- | --- |
| `Spreadsheet.cs`, `.Styles.cs`, `.Tables.cs`, `.NamedRanges.cs`, `.Protection.cs` | `TableManager`, `NamedRangeManager`, `WorkbookProtector`, `WorkbookElementOrderer` |
| `Worksheet.cs`, `.Annotations.cs`, `.Comments.cs`, `.Drawing.cs`, `.ElementOrder.cs` | `CommentWriter`, `WorksheetImageWriter`, `DataValidationWriter`, `ConditionalFormattingWriter`, `HyperlinkStore`, `WorksheetElementOrderer` |

New feature work goes **into a collaborator**, not back into the coordinator. If a feature does not
fit an existing collaborator, add one rather than growing `Worksheet` again — it was 1183 lines
before this split.

`Factory/*` was deleted on 2026-07-24 as dead code. Do not reintroduce a factory layer; construction
goes through `Spreadsheet` and the `IWorksheet`/`IRow` members.

## Element order: the recurring bug class

Rule 3 in [AGENTS.md](../../AGENTS.md) explains why this matters. The Excel-side mechanics:

- Child order is owned by `WorkbookElementOrderer` and `WorksheetElementOrderer`. Insert new elements
  through them, never with a bare `AppendChild`.
- `Utils.OpenXmlElementsEqual` is **order-insensitive**. Style deduplication therefore silently
  accepts a mis-ordered element and only leaks it to disk when it happens to be the first of its
  combination. Focused unit tests cannot catch this — the schema validator can. See
  [testing.md](testing.md).

## Style pipeline

- Create through `Spreadsheet.CreateStyle(...)`, compose through `IStyle.CreateMergedStyle(...)`,
  and merge parts with `Utils.MergeFonts` / `MergeFills` / `MergeBorders`. Do not write a fourth
  merge implementation.
- Inferred number formats — keep these defaults unless the task explicitly changes them:

  | Value type | Built-in format id |
  | --- | --- |
  | integer | `1` |
  | floating point / decimal | `4` |
  | date | `14` |
  | string | `49` |

- **A parent style is applied once, when the child is created — never on a later access.**
  `Worksheet.GetOrCreateRow` seeds a new row with the sheet style, `Row.GetOrCreateCell` seeds a new
  cell with the row style, and `Row.CreateCell` does the same for the cells it backfills. Re-applying
  it to something that already exists overlays the *wider* level on top of the narrower one, which
  inverts the precedence above; it stays invisible until something touches the same cell twice, and
  `Range.Merge()` does exactly that after `ApplyStyle`. This cost one defect already.
- Alignment and number-format merging is intentionally shallower than font/fill/border merging.
  Do not "fix" the asymmetry as a drive-by; it is a behaviour change.
- Colours are normalized to 8-digit ARGB hex. `Font.ArgbHexColor` and `Fill(string)` previously
  wrote un-normalized input straight into the file — keep the normalization.

## Public API evolution

- Prefer `AddCell(...)`. `AddCellWithValue(...)` is obsolete, retained for compatibility.
- The remaining OpenXml leakage — the `Styles.*` `Element` getters, `Utils.Merge*`, raw `Stylesheet`
  and `CellFormat` access, and the compatibility members on `IWorksheet` / `IStyle` / `ICell` /
  `ISpreadsheet` — is known debt tracked in [../excel-state-verdict.md](../excel-state-verdict.md)
  and the advanced roadmap. Do not widen it.
- `InternalsVisibleTo` exposes internals to `OfficeDocuments.Excel.UnitTests` only, so the
  collaborators can be unit-tested directly.

## Not yet validated

These have been verified as *absent* from the source and are fair game to add, with tests:
sheet-name length and character rules, `double.NaN` / `Infinity` handling, XML control characters in
strings, and Excel row/column upper bounds.
