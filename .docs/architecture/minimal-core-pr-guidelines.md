# Minimal Core PR Guidelines

Date: 2026-05-31

This document defines the decision rules for future PRs, implementation tasks, and architecture choices in `MDDM.OfficeDocuments`.

## Product scope

- Supported formats are XML Office formats only: `.xlsx` and `.docx`.
- Legacy binary formats such as `.xls` and `.doc` are out of scope.
- The library should primarily act as a wrapper over `DocumentFormat.OpenXml`, because that SDK provides the required XML access but remains awkward as the default consumer API.

## Dependency policy

- `DocumentFormat.OpenXml` is the default and preferred implementation foundation.
- It is not mandatory to use `DocumentFormat.OpenXml` for every internal helper when another realistic approach is:
  - simpler
  - safer
  - faster
  - and still respects the XML-only scope of the library
- Every new dependency must be justified against the cost it introduces into the minimal core.
- If a feature needs a heavier dependency or a large helper stack, that is a strong signal that the feature belongs in an advanced layer or a follow-up library.

## Core vs. advanced decisions

A feature belongs in the core backlog when it satisfies most of these points:

- it supports a common document-authoring or data-reading scenario
- it improves efficient work with business data
- it keeps the public API small and easy to learn
- it can be delivered without disproportionate complexity

A feature belongs in the advanced backlog when it satisfies one or more of these points:

- it materially widens the public surface area
- it requires heavier infrastructure, test fixtures, or new conceptual layers
- it serves a smaller subset of use cases
- it makes sense as an optional layer over a stable core

## PR checklist

Every PR should answer these questions explicitly:

1. Does the feature belong in the core or in an advanced layer?
2. How does the feature improve straightforward document generation or efficient data access?
3. Why is the chosen approach better than asking consumers to use raw `DocumentFormat.OpenXml` directly?
4. Is every new dependency genuinely necessary?
5. Does the change leak a new OpenXml detail into the public API without a strong reason?
6. Is the feature covered by focused tests and updated documentation?

For day-to-day repository use, the PR template in `.github/pull_request_template.md` should remain the default checklist entry point.

## Implementation rules

- Prefer a small and understandable public API.
- Prefer task-oriented helpers over exposing raw OpenXml structures.
- Do not create parallel feature islands when the change can build on an existing abstraction.
- When unsure, deliver the smaller, cheaper, and better-testable iteration first.
- For advanced features, document the "outside the core" option explicitly.

## Documentation rules

- Every meaningful feature task should explain why it belongs in the core or why it should live in an advanced layer.
- `README.md` should stay product-oriented and concise.
- Detailed documentation belongs in `.docs/`.
- The task backlog should remain split into `core` and `advanced`.
