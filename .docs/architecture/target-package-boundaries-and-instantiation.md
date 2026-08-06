# Target Package Boundaries and Instantiation

Audit date: 2026-05-31 · Last reconciled: 2026-07-27

This document records the architectural audit around package boundaries, construction patterns, and
the role of factory or DI layers in `MDDM.OfficeDocuments`, and the direction that came out of it.

## Audit summary

The audit found:

- `OfficeDocuments.Excel` exposed public factory interfaces and implementations in `Factory/*`
- those factory classes were thin wrappers over `new Spreadsheet(...)`, `new Worksheet(...)`, `new Row(...)`, `new Cell(...)`, and `new Style(...)`
- no real DI or container integration such as `IServiceCollection`, `AddSingleton`, `AddScoped`, or `AddTransient` was present in the repository
- the Word module exposed no comparable public factory or DI layer

The factory layer therefore did not behave like a real runtime-composition seam. It mostly widened the public surface without a clear business payoff.

**Resolved 2026-07-24:** the entire `Factory/*` layer was removed. Nothing below asks for it back.

## Recommended direction

### 1. No DI layer in the minimal core

The minimal core should not require DI registration or a service-based composition model. A normal consumer should be able to use the library through direct constructors and simple entry points.

Recommended public entry points:

- `new Spreadsheet(...)`
- `Spreadsheet.CreateDocument(...)`
- `Spreadsheet.OpenDocument(...)`
- `new Wordprocessing(...)`

If better internal composition is needed later, it should stay internal rather than become a consumer-facing DI seam.

### 2. No public factory layer — done

The factory layer was internalized, found to be unused, and removed on 2026-07-24. Construction now
goes through the `Spreadsheet` constructors and the `Spreadsheet.CreateDocument` /
`Spreadsheet.OpenDocument` static factories.

The rule this leaves behind: a construction abstraction is added back only when a real integration
scenario needs one, not to make the surface look more extensible.

### 3. Target package boundary

The recommended target structure stays conservative and should evolve only when real advanced needs appear:

- `OfficeDocuments.Excel`
  - minimal core for workbook, worksheet, row, cell, range, style, and common data workflows
- `OfficeDocuments.Word`
  - minimal core for body, paragraph, text, and common `.docx` authoring scenarios
- `OfficeDocuments.Excel.Advanced`
  - optional layer for heavier Excel features such as charts, pivot tables, templates, or broader import/export helpers
- `OfficeDocuments.Excel.Interop`
  - only if the project chooses to preserve or intentionally expose a raw OpenXml interop surface outside the core

There is currently no strong reason to create `OfficeDocuments.Word.Advanced`. The Word module first needs a broader and more stable core authoring model.

## What this means for implementation

### Excel

- Public OpenXml-heavy members should continue to move out of the preferred core surface.
- Structured tables and other heavier Excel features should remain clearly separated from the minimal-core story.
- Safe cleanup usually means deprecating or hiding raw interop entry points first and extracting them later.

### Word

- No separate extraction task is needed yet for architecture cleanup alone.
- The priority remains completing the small-core authoring model.
- If later work introduces raw OpenXml leakage or a heavy optional feature layer, that should be re-evaluated only after the Word core is materially more mature.

## Decision rules

A new abstraction layer should exist only when it solves a real problem:

- testability that cannot be achieved more cheaply another way
- separation of a heavier advanced feature set from a small core
- an intentionally supported interop scenario

A new abstraction layer should not exist only because it is theoretically possible.

## Follow-up backlog areas

- Excel public-surface cleanup
- advanced structured-table and template workflows when justified

## Acceptance status of this proposal

This document is an architecture guide for backlog and PR decisions. It does not by itself imply an immediate physical package split. Its purpose is to guide gradual reduction of the core surface and to clarify when an `Advanced` or `Interop` layer becomes worth creating as a separate project.
