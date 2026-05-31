---
description: "Use when writing or modifying C# code in this repository. Covers modern .NET and C# features, performance-minded implementation, readability, validation, and repo-specific coding expectations."
name: "CSharp Repo Guidance"
applyTo: "**/*.cs"
---
# C# Guidance

- Use the latest stable C# and .NET features that are compatible with the SDK pinned in `global.json` and the project's target frameworks.
- The repository language-version policy is centralized and should stay on the latest stable C# supported by the installed major SDK.
- Prefer modern, readable constructs when they improve the code without causing churn: file-scoped namespaces, collection expressions, pattern matching, switch expressions, raw string literals when useful, and `ArgumentNullException.ThrowIfNull(...)` for null guards.
- Keep the code easy to reason about. Avoid clever abstractions, speculative generalization, or framework-style patterns that the repository does not already use.
- Optimize for both correctness and efficiency. Be alert to repeated LINQ materialization, repeated XML traversal, unnecessary allocations, and repeated style-merging work inside loops.
- Treat larger worksheets and repeated operations as real workloads. Prefer single-pass logic and cached local results when the same OpenXml tree or collection would otherwise be enumerated many times.
- Preserve file-local style when touching existing code. This repo is not fully uniform yet, so consistency with nearby code is better than broad restyling.
- Use explicit argument validation and keep exception types aligned with the surrounding API contracts.
- Preserve public API compatibility unless the task explicitly asks for an API change.
- When changing public behavior, update or add focused xUnit tests in the matching test project.
- When adding or changing public library features, keep `README.md` high-level and update the relevant detailed documentation in `.doc/`.
- Documentation changes should stay in English and follow the terminology in `.doc/terminology.md`.
