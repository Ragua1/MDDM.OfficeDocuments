# OfficeDocuments Tasks

Date: 2026-05-31

This directory contains implementation-planning documents derived from the benchmark report and the current feature-gap backlog.

The main roadmap view is in [roadmap-overview.md](roadmap-overview.md).

## Product guardrails

Every task document must fit the product scope and the dependency policy defined in
[../architecture/minimal-core-pr-guidelines.md](../architecture/minimal-core-pr-guidelines.md).
That document is the single source for what belongs in the library at all, and for the core-versus-advanced
decision. A task that cannot answer its checklist does not belong in the backlog yet.

## What a task document should contain

Each task should capture:

- `Business goal`: why the feature matters
- `Why core or advanced`: why the feature belongs in that backlog layer
- `Functional description`: what the library should be able to do
- `Technical guidance`: concrete implementation guidance for whoever implements it, human or coding agent
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

## Who implements a task

Task documents are tool-neutral: `Technical guidance` addresses whoever picks the task up, human or
coding agent. Its purpose is to reduce repository rediscovery during implementation, not to prescribe
a tool.

The working rules a coding agent follows are in [../../AGENTS.md](../../AGENTS.md) and
[../ai-instructions/](../ai-instructions/README.md).
