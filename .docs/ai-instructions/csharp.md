---
name: C# guidance
description: Language level, nullability, argument validation, exception contracts, and performance expectations for C# code in this repository.
applyTo: "**/*.cs"
---

# C# guidance

## Language and framework level

- `LangVersion` is `default` in [`Directory.Build.props`](../../Directory.Build.props), so the
  language level follows the installed SDK (`10.0.300` → C# 14). Do not pin it per project.
- **The multi-target trap:** the libraries build for `net8.0;net9.0;net10.0`. C# *language*
  features are fine because the same compiler compiles every target, but a BCL API introduced in
  .NET 9 or .NET 10 will break the `net8.0` build. There is no `#if NET*` in the codebase today —
  keep it that way by staying on APIs available in .NET 8, or raise the question before adding the
  first conditional compilation block.
- Prefer modern constructs where they genuinely improve the code: file-scoped namespaces, collection
  expressions, pattern matching, switch expressions, primary constructors, raw string literals.
  Do not restyle surrounding code to match.

## Argument validation

Validate at the public boundary, in the method that received the input, and use the framework
throw-helpers already used across the codebase:

```csharp
public void AddNamedRange(string name, IRange range)
{
    ArgumentException.ThrowIfNullOrEmpty(name);
    ArgumentNullException.ThrowIfNull(range);
    ...
}
```

## Exception contracts

Match the existing behaviour — consumers and tests depend on the exact type:

| Situation | Exception |
| --- | --- |
| Invalid index, invalid reference, invalid user argument | `ArgumentException` |
| Required reference argument is `null` | `ArgumentNullException` |
| Broken document state, impossible internal condition | `InvalidOperationException` |
| Feature genuinely not implemented for this input | `NotSupportedException` |

Never silently coerce an invalid index, reference, or sheet state into something valid.
`NotImplementedException` and sentinel return values such as `-1` were removed from the formula
engine deliberately; do not reintroduce either pattern.

## Nullability

- Nullable reference types are enabled repo-wide. Fix the cause of a warning; do not suppress it
  with `!` or `#pragma`.
- In this domain `null` frequently carries meaning — an optional style, an inherited format value.
  When you add a nullable member, document what `null` means, because "not set" and "explicitly
  off" are different things (see [word.md](word.md)).

## Performance

Document generation is the hot path, and "works on the happy-path example" is not the bar. Watch for:

- Repeated traversal of the same OpenXml tree inside a loop — hoist it into a local.
- Repeated LINQ materialization (`.ToList()` per iteration) and repeated style merging per cell.
- Nested scans over ranges. Keep range operations linear in the requested range.
- Known open debt: style deduplication is O(N²) and comment VML generation is inefficient. Both are
  now isolated in `Style` and `CommentWriter`. Do not add new callers to the slow path.

## Style

- Match the file you are editing. The repository is not uniform yet, so local consistency beats
  global tidiness, and a reformat buried in a behaviour change is a review hazard.
- Comments explain *why*, not *what*. The non-obvious invariants in this codebase — element order,
  lazy creation, backfilled cells — are worth a comment; a restatement of the method name is not.
