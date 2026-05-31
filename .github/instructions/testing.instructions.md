---
description: "Use when creating or modifying automated tests in MDDM.OfficeDocuments. Covers xUnit test design, overlap strategy, reliability, and coverage practices aligned with OSS and Microsoft guidance."
name: "Testing Guidance"
applyTo:
  - "test/**/*.cs"
---
# Testing Guidance

## Objectives

- Protect public behavior, not implementation details.
- Keep tests deterministic, readable, and fast.
- Use layered coverage: focused tests plus end-to-end document scenarios.
- Ensure baseline tests and complex tests intentionally overlap for critical features.

## Test Pyramid For This Repository

- Keep a majority of focused API-level tests around `ISpreadsheet`, `IWorksheet`, `IRow`, `ICell`, and `IRange` behavior.
- Add scenario tests for realistic document workflows (create -> write -> close -> reopen -> read -> verify).
- Avoid over-reliance on raw OpenXml assertions unless validating serialization-critical behavior.

## Overlap Strategy (Critical For This Library)

For every critical capability, create at least two complementary tests:

- A focused test validating the smallest expected behavior and argument/edge handling.
- A roundtrip or scenario test validating the same behavior after save/reopen and in combination with nearby features.

Critical capabilities include:

- Cell value types and formula handling.
- Range operations (set/get values, merge, sort, validation/formatting where relevant).
- Worksheet lifecycle operations (add, rename, move, hide, remove).
- Table operations (create, query, resize/rename/remove constraints).
- Metadata persistence (hyperlinks, comments, protections, named ranges).

## Test Design Rules (xUnit / OSS Best Practices)

- Follow AAA structure (Arrange, Act, Assert) with explicit intent.
- Use one primary behavior per test; avoid broad assertions unrelated to the test name.
- Name tests as `MethodName_StateUnderTest_ExpectedOutcome`.
- Use `[Theory]` with inline/member data when testing the same behavior across multiple inputs.
- Prefer clear, domain-oriented assertions over highly condensed assertion chains.
- Keep test setup local unless shared setup significantly reduces duplication and does not hide intent.

## Determinism And Reliability (Microsoft Guidance)

- Do not depend on machine-local files, external services, locale-specific assumptions, or wall-clock timing.
- Use temp files/streams via existing test helpers.
- Clean up files created by tests unless the current test infrastructure already guarantees isolation.
- Do not introduce sleeps, random-order assumptions, or tests sensitive to parallel execution side effects.

## What To Assert

- Assert through the public API first.
- For document libraries, verify both in-memory behavior and persisted behavior after reopen.
- Prefer semantic assertions (cell values, references, order, table metadata, hidden state) over brittle internal-node ordering.
- When checking exceptions, assert exact exception type and verify meaningful argument constraints.

## Coverage Quality (Not Only Percentage)

- Add tests for happy path, edge cases, and invalid arguments.
- Add regression tests for every bug fix reproducing the original failure.
- Use coverage as a signal, not the only target: prioritize branch-heavy, behavior-critical areas.
- Keep overlap between focused tests and scenario tests for critical paths to reduce false confidence from single-layer coverage.

## Performance-Safe Testing

- Keep routine tests lightweight; reserve larger datasets for targeted stress scenarios.
- Avoid unnecessary repeated document traversal in assertions.
- Ensure tests remain fast enough for frequent local runs and CI usage.

## Anti-Patterns To Avoid

- Testing private implementation details instead of contract behavior.
- Duplicating complex scenario tests when a focused test would isolate the behavior better.
- Writing tests that only assert file existence for behavior that should validate content correctness.
- Introducing flaky tests or broad snapshots without precise intent.

## Repository-Specific Notes

- Excel is the primary mature surface; prioritize robust overlap on Excel critical features.
- Keep row/column index assertions explicitly 1-based.
- Preserve and validate OpenXml ordering invariants through public behavior outcomes.
- If a test must touch obsolete raw wrappers for validation, keep usage minimal and justified.
