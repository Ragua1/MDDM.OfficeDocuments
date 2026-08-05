# Documentation Index

This folder contains the canonical repository documentation for `OfficeDocuments`.

All links inside `.doc/` use relative paths so the documentation stays portable inside the repository.

## Consumer documentation

- [excel-library.md](excel-library.md) - current Excel API, usage guidance, and examples
- [word-library.md](word-library.md) - current Word API, usage guidance, and examples
- [terminology.md](terminology.md) - shared terminology and abbreviations used across the project

## Product analysis

- [library-benchmark-report.md](library-benchmark-report.md) - current benchmark against comparable libraries and current positioning
- [feature-gap-backlog.md](feature-gap-backlog.md) - current gap backlog after the latest API review
- [excel-state-verdict.md](excel-state-verdict.md) - independent Excel technical-debt audit, state verdict, and direction comparison
- [excel-performance-baseline.md](excel-performance-baseline.md) - measured cost of the Excel write and read paths, the four known hot spots, and the thresholds the CI performance guards are calibrated against

## Working instructions

- [ai-instructions/README.md](ai-instructions/README.md) - detailed rules for AI coding agents, and for anyone who wants the same conventions written down

The root [../AGENTS.md](../AGENTS.md) is the entry point; the files in `ai-instructions/` hold the depth behind it.

## Architecture notes

- [architecture/minimal-core-pr-guidelines.md](architecture/minimal-core-pr-guidelines.md) - contribution and PR decision rules
- [architecture/target-package-boundaries-and-instantiation.md](architecture/target-package-boundaries-and-instantiation.md) - package-boundary and construction-model guidance
- [architecture/word-002-readiness-audit.md](architecture/word-002-readiness-audit.md) - historical record: what the `WORD-002` readiness audit got right, and the three things it did not anticipate

## Task planning

- [tasks/README.md](tasks/README.md) - planning entry point
- [tasks/roadmap-overview.md](tasks/roadmap-overview.md) - cross-module roadmap view
- [tasks/core/README.md](tasks/core/README.md) - core backlog entry point
- [tasks/core/excel/README.md](tasks/core/excel/README.md) - Excel core tasks
- [tasks/core/word/README.md](tasks/core/word/README.md) - Word core tasks, all delivered 2026-07-27
- [tasks/advanced/README.md](tasks/advanced/README.md) - advanced backlog entry point

## Test documentation

- [../test/README.md](../test/README.md) - test tiers, entry criteria, and the shared test kits

## Relationship to the root README

The root [../README.md](../README.md) stays intentionally short and product-oriented. Detailed guidance, API notes, and planning material live here in `.doc/`.
