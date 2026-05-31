# EXCEL-007 OpenXml compatibility surface isolation

Date: 2026-05-31

## Business goal

Reduce coupling between the public Excel API and raw `DocumentFormat.OpenXml` types so future library evolution stays safer and more predictable.

## Why core or advanced

Core. The default value of the library is a smaller and more consumer-friendly API than raw OpenXml.

## Functional description

The Excel library should keep existing compatibility members available where needed, but all raw OpenXml exposure should be clearly marked as compatibility-oriented and discouraged on the main consumer path.

## Technical guidance for GHC

Relevant files include:

- `src/OfficeDocuments.Excel/Interfaces/IRow.cs`
- `src/OfficeDocuments.Excel/Interfaces/IWorksheet.cs`
- `src/OfficeDocuments.Excel/Interfaces/IStyle.cs`
- `src/OfficeDocuments.Excel/Interfaces/IOpenXmlWrapper.cs`
- `src/OfficeDocuments.Excel/Interfaces/ISpreadsheet.cs`
- `src/OfficeDocuments.Excel/Spreadsheet.cs`

Observed issue:

- most raw OpenXml compatibility members are already hidden behind `EditorBrowsable(Never)` and `Obsolete(...)`
- `IRow.Element` still exposes `DocumentFormat.OpenXml.Spreadsheet.Row` without the same compatibility warning layer

Implementation direction:

- mark remaining raw OpenXml members consistently with `EditorBrowsable(EditorBrowsableState.Never)` and targeted `Obsolete(...)` guidance
- preserve binary compatibility and avoid removing members in the current major version
- keep tests focused on public behavior first; only use raw element access where serialization-critical assertions require it
- as follow-up work, review whether more internal test assertions can move from raw element checks to public API checks

## Complexity

Low.

## Risks

- new obsolete warnings can surface in tests or downstream consumers that still use compatibility members
- inconsistent messages across interfaces can confuse migration guidance

## Dependencies

- none

## Subtasks

- audit public interfaces for remaining raw OpenXml members without compatibility annotations
- annotate the remaining members consistently
- keep existing members intact for compatibility
- update or suppress internal test usage only where warnings become too noisy

## Acceptance criteria

- remaining raw OpenXml compatibility members are consistently marked as non-primary API surface
- no public member required for compatibility is removed
- Excel projects still build and existing behavior remains unchanged
