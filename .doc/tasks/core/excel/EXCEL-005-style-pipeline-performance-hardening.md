# EXCEL-005 Style pipeline performance hardening

Date: 2026-05-31

## Business goal

Keep workbook styling fast and predictable for larger sheets, repeated range styling, and style composition scenarios.

## Why core or advanced

Core. Styling is already part of the default Excel workflow and directly affects normal document generation throughput.

## Functional description

The Excel library should keep current style behavior while reducing repeated OpenXml DOM serialization, repeated XML parsing, and repeated full stylesheet scans during style creation and style merging.

## Technical guidance for GHC

Current hotspots are concentrated in the style merge and equality path:

- `src/OfficeDocuments.Excel/Utils.cs`
- `src/OfficeDocuments.Excel/Extensions/XElementExtensions.cs`
- `src/OfficeDocuments.Excel/DataClasses/Style.cs`
- `src/OfficeDocuments.Excel/Styles/Font.cs`
- `src/OfficeDocuments.Excel/Styles/Fill.cs`
- `src/OfficeDocuments.Excel/Styles/Border.cs`
- `src/OfficeDocuments.Excel/Styles/Alignment.cs`

Observed issues:

- style merge currently converts OpenXml nodes to `OuterXml`, reparses them into `XDocument`, merges them, then constructs new OpenXml elements
- equality checks repeatedly compare `OuterXml` through XML normalization helpers
- style lookup repeatedly materializes `Fonts`, `Fills`, `Borders`, `CellFormats`, and `NumberingFormats` collections with `ToList()`

Implementation direction:

- replace `OuterXml` plus `XDocument.Parse(...)` equality checks with direct OpenXml-aware value comparison helpers where possible
- avoid reserializing OpenXml elements during merge when a shallow element clone plus child/attribute merge is sufficient
- introduce local caches or single-pass lookup helpers inside `Style` for fonts, fills, borders, numbering formats, and cell formats
- keep behavior identical for built-in number format ids and current alignment merge semantics
- validate with existing `StyleTest` and `UtilsTest`, then add a regression test if equality semantics need adjustment

## Complexity

Medium to high.

## Risks

- style equality changes can create duplicate style records or incorrect reuse
- number format handling must keep current built-in and custom id behavior
- alignment merge semantics are intentionally shallow and must not change accidentally

## Dependencies

- none

## Subtasks

- map each style comparison path and classify whether it needs equality, merge, or dedup only
- introduce internal helper methods for stylesheet lookups without repeated `ToList()` allocations
- replace XML-string-based comparisons with direct structural comparisons
- replace XML-string-based merge for font, fill, and border with direct OpenXml cloning and merge logic
- run focused Excel style tests and roundtrip tests

## Acceptance criteria

- style creation and style merge no longer depend on repeated `OuterXml` plus `XDocument.Parse(...)` in hot paths
- current public style behavior stays unchanged in existing tests
- custom number formats still start at the current custom id baseline and remain stable per workbook
- style-related tests pass across supported target frameworks
