# Excel Advanced Roadmap

Date: 2026-05-31

This document groups Excel backlog slices that are broader, heavier, or more architectural than the preferred minimal-core surface.

## EXCEL-005 Rich output and template workflows

- Status: Open
- Goal: evaluate richer Excel output scenarios such as templates or broader import/export helpers
- Why advanced:
  - this work can widen the surface materially
  - template abstractions can pull the library away from the current small-core positioning

## EXCEL-006 OpenXml interop surface extraction

- Status: In progress
- Goal: move raw OpenXml-oriented compatibility members out of the preferred consumer surface
- Why advanced:
  - this is primarily architecture cleanup rather than a new consumer-facing feature

## EXCEL-007 Factory and raw style plumbing extraction

- Status: Partially delivered (updated 2026-07-24)
- Goal: remove historical factory and raw style seams from the public default story
- Why advanced:
  - these are architecture seams rather than primary consumer workflows
- Current state:
  - the public factory layer has been removed entirely (see EXCEL-009)
  - raw style plumbing still leaks: the `Styles.*` wrappers (`Font`, `Fill`, `Border`, `Alignment`, `NumberingFormat`) still expose a public, non-obsolete `Element` of a raw OpenXml type, and `Utils.MergeFonts/MergeFills/MergeBorders` remain public
- Remaining follow-up:
  - hide or wrap the raw `Element` getters and OpenXml-typed constructors on `Styles.*`, and internalize the `Utils.Merge*` helpers

## EXCEL-008A Table create and lookup hardening

- Status: Delivered
- Goal: provide a stable structured-table creation and lookup workflow
- Current state:
  - `AddTable(...)`, `GetTable(...)`, and `GetTables(...)` exist

## EXCEL-008B Table lifecycle operations

- Status: Delivered
- Goal: manage existing tables safely after creation
- Current state:
  - `RenameTable(...)`, `ResizeTable(...)`, and `RemoveTable(...)` exist

## EXCEL-008C Table style and options

- Status: Partially delivered
- Goal: enrich structured-table options and styling without overloading the core surface
- Current state:
  - `TableCreateOptions` and `TableStyleOptions` exist
- Remaining follow-up:
  - only add more table behavior when there is a strong consumer payoff

## EXCEL-009 Factory internalization and entry-point simplification

- Status: Delivered (updated 2026-07-24)
- Goal: continue simplifying construction entry points so the preferred API relies on direct workbook and worksheet flows
- Why advanced:
  - this is part of the broader public-surface cleanup story rather than a standalone business feature
- Current state:
  - the factory interfaces had already been made `internal`, but were unused dead code
  - the entire `Factory/*` layer (`ICellFactory`, `IRowFactory`, `ISpreadSheetFactory`, `IStyleFactory` and their implementations) has been removed
  - the preferred construction entry points are the `Spreadsheet` constructors and the `Spreadsheet.CreateDocument` / `Spreadsheet.OpenDocument` static factories
