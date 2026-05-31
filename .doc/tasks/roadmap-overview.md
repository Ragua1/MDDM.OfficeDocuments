# OfficeDocuments Roadmap Overview

Date: 2026-05-31

This document turns the current backlog into one roadmap view across Excel and Word.

Goals:

- provide a fast overview of the core and advanced backlog
- show a reasonable delivery order
- simplify iteration planning by priority, complexity, and dependencies

## Legend

- `Scope`: `Core` or `Advanced`
- `Priority`: `P0` is the highest priority, `P2` is later-stage work
- `Complexity`: rough delivery estimate
- `Status`: current task state

## Recommended roadmap

| Order | Task | Scope | Module | Priority | Complexity | Status | Dependencies | Primary goal |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 1 | [WORD-001](core/word/WORD-001-text-formatting-and-paragraph-model.md) | Core | Word | P0 | M | Open | none | Add minimal text and paragraph formatting so `.docx` creation becomes practically usable |
| 2 | [WORD-002A](core/word/WORD-002A-basic-tables.md) | Core | Word | P0 | M | Open | `WORD-001` recommended | Add the first small but practical Word table workflow |
| 3 | [WORD-002B](core/word/WORD-002B-hyperlinks.md) | Core | Word | P0 | M | Open | `WORD-001` strongly recommended | Add hyperlink support without breaking the paragraph or text model |
| 4 | [WORD-002C](core/word/WORD-002C-images.md) | Core | Word | P0 | M-L | Open | `WORD-001` strongly recommended, `WORD-002B` recommended | Add basic inline image support with minimal media infrastructure |
| 5 | [WORD-003](core/word/WORD-003-headers-footers-sections-and-metadata.md) | Core | Word | P0 | M-L | Open | `WORD-001` recommended, `WORD-002` partial | Add branded document structure and document metadata |
| 6 | [EXCEL-001](core/excel-roadmap.md#excel-001-range-centric-api-and-public-surface-cleanup) | Core | Excel | P0 | M-L | Delivered | none | Establish range-centric spreadsheet workflows and reduce public OpenXml leakage |
| 7 | [EXCEL-002](core/excel-roadmap.md#excel-002-bulk-insert-and-tabular-import-workflows) | Core | Excel | P1 | M | Delivered | `EXCEL-001` recommended | Add efficient data import and bulk insert workflows |
| 8 | [EXCEL-003](core/excel-roadmap.md#excel-003-worksheet-operations-and-workbook-usability) | Core | Excel | P1 | L | Delivered | `EXCEL-001` recommended | Add worksheet lifecycle and workbook usability features |
| 9 | [EXCEL-004](core/excel-roadmap.md#excel-004-validation-formatting-and-annotations) | Core | Excel | P1 | L | Delivered | `EXCEL-001` recommended, `EXCEL-003` partial | Add validation, annotations, worksheet images, and protection features |
| 10 | [WORD-004](core/word/WORD-004-search-navigation-and-test-hardening.md) | Core | Word | P1 | M | Open | `WORD-001` to `WORD-003` recommended | Strengthen read and edit scenarios and stabilize test coverage |
| 11 | [EXCEL-006](advanced/excel-roadmap.md#excel-006-openxml-interop-surface-extraction) | Advanced | Excel | P2 | L | In Progress | `EXCEL-001`, architecture decision | Extract raw OpenXml-oriented compatibility surface from the minimal core |
| 12 | [EXCEL-007](advanced/excel-roadmap.md#excel-007-factory-and-raw-style-plumbing-extraction) | Advanced | Excel | P2 | M-L | Open | `EXCEL-001`, `EXCEL-006` partial | Remove public factory and raw style plumbing from the preferred consumer surface |
| 13 | [EXCEL-009](advanced/excel-roadmap.md#excel-009-factory-internalization-and-entry-point-simplification) | Advanced | Excel | P2 | M | In Progress | `EXCEL-007` recommended | Turn the factory cleanup direction into a concrete simplification slice |
| 14 | [EXCEL-008A](advanced/excel-roadmap.md#excel-008a-table-create-and-lookup-hardening) | Advanced | Excel | P2 | M | Delivered | `EXCEL-001`, `EXCEL-002` recommended | Add a stable structured-table create and lookup workflow |
| 15 | [EXCEL-008B](advanced/excel-roadmap.md#excel-008b-table-lifecycle-operations) | Advanced | Excel | P2 | M | Delivered | `EXCEL-008A` | Add rename, resize, and remove operations for existing tables |
| 16 | [EXCEL-008C](advanced/excel-roadmap.md#excel-008c-table-style-and-options) | Advanced | Excel | P2 | M | Partially delivered | `EXCEL-008A`, `EXCEL-008B` recommended | Add richer table options and styling controls |
| 17 | [EXCEL-005](advanced/excel-roadmap.md#excel-005-rich-output-and-template-workflows) | Advanced | Excel | P2 | XL | Open | `EXCEL-001` to `EXCEL-004`, `EXCEL-008A` to `EXCEL-008C` optional | Evaluate richer output features for an advanced layer or separate library |

## Milestone view

### Milestone A: Word minimum viable authoring

| Task | Why first |
| --- | --- |
| [WORD-001](core/word/WORD-001-text-formatting-and-paragraph-model.md) | Without text and paragraph formatting, Word output remains too limited for common use |
| [WORD-002A](core/word/WORD-002A-basic-tables.md) | Tables are the most ready and immediately useful body-level Word extension |
| [WORD-002B](core/word/WORD-002B-hyperlinks.md) | Hyperlinks add a common business-document primitive after the paragraph or run seam is stable |
| [WORD-002C](core/word/WORD-002C-images.md) | Images should follow only after the document and media context is ready |
| [WORD-003](core/word/WORD-003-headers-footers-sections-and-metadata.md) | Completes the baseline for branded and formal `.docx` generation |

### Milestone B: Excel data ergonomics

| Task | Why here |
| --- | --- |
| [EXCEL-001](core/excel-roadmap.md#excel-001-range-centric-api-and-public-surface-cleanup) | Establishes the abstraction needed for several later Excel features |
| [EXCEL-002](core/excel-roadmap.md#excel-002-bulk-insert-and-tabular-import-workflows) | Directly improves the main data-export workflow |

### Milestone C: Excel editing and usability

| Task | Why here |
| --- | --- |
| [EXCEL-003](core/excel-roadmap.md#excel-003-worksheet-operations-and-workbook-usability) | Improves workbook handling and everyday usability |
| [EXCEL-004](core/excel-roadmap.md#excel-004-validation-formatting-and-annotations) | Adds controlled input and richer but still core-safe workbook behavior |

### Milestone D: Read and update hardening

| Task | Why here |
| --- | --- |
| [WORD-004](core/word/WORD-004-search-navigation-and-test-hardening.md) | Stabilizes read and edit flows after the main authoring model is in place |

### Milestone E: Advanced expansion

| Task | Why separate |
| --- | --- |
| [EXCEL-006](advanced/excel-roadmap.md#excel-006-openxml-interop-surface-extraction) | The current Excel API still exposes low-level interop pieces that do not fit the preferred minimal core |
| [EXCEL-007](advanced/excel-roadmap.md#excel-007-factory-and-raw-style-plumbing-extraction) | Factory and raw style plumbing are architecture seams, not primary consumer features |
| [EXCEL-009](advanced/excel-roadmap.md#excel-009-factory-internalization-and-entry-point-simplification) | Public factory contracts still look more like historical plumbing than a real consumer seam |
| [EXCEL-008A](advanced/excel-roadmap.md#excel-008a-table-create-and-lookup-hardening) | Structured tables need a stable create-and-lookup foundation before richer operations |
| [EXCEL-008B](advanced/excel-roadmap.md#excel-008b-table-lifecycle-operations) | Rename, resize, and remove are a distinct risk cluster from initial table creation |
| [EXCEL-008C](advanced/excel-roadmap.md#excel-008c-table-style-and-options) | Table options and styling should remain an explicit follow-up after the table base is stable |
| [EXCEL-005](advanced/excel-roadmap.md#excel-005-rich-output-and-template-workflows) | Rich output features have higher complexity and may belong outside the minimal core |

## Dependency notes

### Hard dependency clusters

- `EXCEL-001` is the main Excel API foundation for later range-based behavior.
- `WORD-001` is the main Word API foundation for later authoring behavior.
- `EXCEL-006` and `EXCEL-007` depend on a stable view of what the Excel core should keep.
- `EXCEL-008A` to `EXCEL-008C` are grouped under the structured-table roadmap in [advanced/excel-roadmap.md](advanced/excel-roadmap.md).
- `WORD-002A` to `WORD-002C` are grouped under the umbrella task [WORD-002](core/word/WORD-002-tables-images-and-hyperlinks.md).

### Soft dependency clusters

- `EXCEL-002`, `EXCEL-003`, and `EXCEL-004` benefit from a stable range abstraction first.
- `WORD-002`, `WORD-003`, and `WORD-004` benefit from a stable paragraph and run model first.

## Planning notes

- If only one short delivery window is available, prefer `WORD-001` or the remaining Excel public-surface cleanup work.
- If the next release should strengthen Excel exports, prefer Excel cleanup and higher-level import/export ergonomics.
- If the next release should make Word broadly usable, prefer `WORD-001` followed by `WORD-002`.
- Heavier Excel output features should happen only after the current core-facing API direction is considered stable enough.

## Current Excel parts that likely do not belong in the preferred minimal core

Based on the current code, the strongest candidates remain:

- raw OpenXml-oriented compatibility members in `IWorksheet`, `IStyle`, `ICell`, `ISpreadsheet`, and `Spreadsheet`
- public factory abstractions in `Factory/*`
- public raw style plumbing around `Stylesheet`, `CellFormat`, style IDs, and related compatibility seams

These are tracked by:

- [advanced/excel-roadmap.md](advanced/excel-roadmap.md)

## Current Word architecture observation

Unlike Excel, the current Word module does not yet expose the same kind of strong non-core architecture seam. Public Word interfaces are still small and do not expose raw OpenXml surface to the same degree. The recommended direction is therefore not to split a `Word.Advanced` package yet, but to strengthen the Word core first.

For a more detailed readiness assessment of `WORD-002`, see [../architecture/word-002-readiness-audit.md](../architecture/word-002-readiness-audit.md).
