---
name: Workflow
description: How to approach a task in this repository — orientation, definition of done, backlog handling, commits and pull requests.
applyTo: "**"
---

# Workflow

## Orientation: read before exploring

Broad repository search is the slow path. Start with these, in order:

1. [AGENTS.md](../../AGENTS.md) — baseline rules and repository map.
2. The instruction file for the area you are touching (see [README.md](README.md)).
3. [../excel-library.md](../excel-library.md) or [../word-library.md](../word-library.md) — the real public API surface.
4. The interface folder of the module (`Interfaces/*`) — the contract before the implementation.
5. [../tasks/roadmap-overview.md](../tasks/roadmap-overview.md) — whether the work is already planned, and under which task ID.

Only then search the tree.

## Working mode

The always-on rules are in [AGENTS.md](../../AGENTS.md). What they mean in practice:

- **Scope discipline.** Before adding a feature, decide core vs advanced against
  [../architecture/minimal-core-pr-guidelines.md](../architecture/minimal-core-pr-guidelines.md).
  A feature that widens the public surface a lot, needs heavier infrastructure, or serves a narrow
  audience belongs in the advanced backlog, not in the core package.
- **Contract first.** For a public change the order is: extend the interface, implement, test,
  document. Not the other way round, and not all at once at the end.
- **Probe before you write it up.** Two of the WORD-001 bugs were only visible by running the code,
  not by reading it. When you suspect a defect, prove it with a throwaway test first.

## Definition of done

A task is done when all of the following hold:

- The change satisfies the acceptance criteria of its task document in [../tasks/](../tasks/README.md), if it has one.
- The solution for the module you touched builds — `OfficeDocuments.Excel.slnx` or
  `OfficeDocuments.Word.slnx`, and `OfficeDocuments.slnx` when the change spans both.
- The focused test tier passes, and the full suite passes for anything touching shared code, build
  configuration, or packaging. See [testing.md](testing.md).
- New or changed public behaviour has tests at the right tier, including the invalid-argument case.
- A bug fix has a regression test that fails without the fix.
- Documentation is updated per [documentation.md](documentation.md).
- If a validation step could not run, that is stated explicitly, with the reason.

Reporting a task complete without running the tests is not acceptable. If tests fail, say so and
show the output.

## Task documents

Tasks live in [../tasks/core/](../tasks/core/README.md) and [../tasks/advanced/](../tasks/advanced/README.md),
indexed by [../tasks/roadmap-overview.md](../tasks/roadmap-overview.md).

When you complete a slice of a task:

- Append to that task's **Progress log** section: what was delivered, what was verified, what was
  deliberately left out.
- Update the `Status` column in the roadmap table.
- Convert relative dates to absolute (`2026-07-27`, not "today").

## Commits and pull requests

- Do not commit or push unless the user asks.
- Commit messages: imperative mood, one line of subject, body only when the "why" is not obvious.
- Group a commit around one intent. Build/config churn does not belong in a feature commit.
- The pull request checklist in [../../.github/pull_request_template.md](../../.github/pull_request_template.md)
  is the default gate — it is a real checklist, not decoration, and it enforces the core-vs-advanced
  and API-leakage questions.

## When you are unsure

State the assumption and continue, unless proceeding under any assumption would produce useless or
unsafe work. Do the parts that do not depend on the answer first.
