# Target Package Boundaries and Instantiation

Date: 2026-05-31

This document summarizes the current architectural audit around package boundaries, construction patterns, and the role of factory or DI layers in `MDDM.OfficeDocuments`.

## Audit summary

The current repository state shows:

- `OfficeDocuments.Excel` contains public factory interfaces and implementations in `Factory/*`
- those factory classes are thin wrappers over `new Spreadsheet(...)`, `new Worksheet(...)`, `new Row(...)`, `new Cell(...)`, and `new Style(...)`
- no real DI or container integration such as `IServiceCollection`, `AddSingleton`, `AddScoped`, or `AddTransient` is present in the repository
- the Word module does not expose a comparable public factory or DI layer
- the current Word public API is still small and consumer-friendly without an obvious need for an extra construction layer

The current factory layer therefore does not behave like a real runtime-composition seam. It mostly widens the public surface without a clear business payoff.

## Recommended direction

### 1. No DI layer in the minimal core

The minimal core should not require DI registration or a service-based composition model. A normal consumer should be able to use the library through direct constructors and simple entry points.

Recommended public entry points:

- `new Spreadsheet(...)`
- `Spreadsheet.CreateDocument(...)`
- `Spreadsheet.OpenDocument(...)`
- `new Wordprocessing(...)`

If better internal composition is needed later, it should stay internal rather than become a consumer-facing DI seam.

### 2. The factory layer should not be the default public extension surface

If the public factory layer does not have a real usage case, it should move in one of these directions:

- internalization
- obsoletion followed by removal in a later major release
- relocation into an optional advanced or interop layer when a genuine integration scenario appears

The preferred option for the current repository remains internalization.

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

- Public `Factory/*` remains a candidate for removal or internalization.
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
- factory cleanup and internalization
- advanced structured-table and template workflows when justified

## Acceptance status of this proposal

This document is an architecture guide for backlog and PR decisions. It does not by itself imply an immediate physical package split. Its purpose is to guide gradual reduction of the core surface and to clarify when an `Advanced` or `Interop` layer becomes worth creating as a separate project.
