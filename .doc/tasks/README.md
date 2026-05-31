# OfficeDocuments Tasks

Date: 2026-05-31

This directory contains implementation-planning documents derived from the benchmark report and the current feature-gap backlog.

The main roadmap view is in [roadmap-overview.md](roadmap-overview.md).

## Product guardrails

These task documents should follow the current product scope:

- the library targets XML Office formats only, primarily `.xlsx` and `.docx`
- legacy binary formats such as `.xls` and `.doc` are out of scope for the minimal core
- the library should remain a wrapper over `DocumentFormat.OpenXml`, with a simpler and more consumer-friendly API than raw OpenXml usage
- `DocumentFormat.OpenXml` remains the default implementation foundation, but not a dogma for every internal helper
- the core library should stay small, fast, and easy to adopt
- the main business value is efficient work with business data plus straightforward document generation and reading
- broader or heavier features should stay explicitly separable from the minimal core story

## What a task document should contain

Each task should capture:

- `Business goal`: why the feature matters
- `Why core or advanced`: why the feature belongs in that backlog layer
- `Functional description`: what the library should be able to do
- `Technical guidance for GHC`: concrete implementation guidance for GitHub Copilot or another coding agent workflow
- `Complexity`: a rough delivery estimate
- `Risks`: main technical or architectural risks
- `Dependencies`: prerequisites or related tasks
- `Subtasks`: concrete implementation steps
- `Acceptance criteria`: completion conditions

## Backlog entry points

- Core backlog: [core/README.md](core/README.md)
- Advanced backlog: [advanced/README.md](advanced/README.md)
- Cross-module roadmap: [roadmap-overview.md](roadmap-overview.md)

## Architecture references

- [../architecture/minimal-core-pr-guidelines.md](../architecture/minimal-core-pr-guidelines.md)
- [../architecture/target-package-boundaries-and-instantiation.md](../architecture/target-package-boundaries-and-instantiation.md)
- [../architecture/word-002-readiness-audit.md](../architecture/word-002-readiness-audit.md)

## Note on GHC

`GHC` in these task documents means GitHub Copilot or another coding-agent workflow. The technical sections are written to reduce repository rediscovery during implementation.
