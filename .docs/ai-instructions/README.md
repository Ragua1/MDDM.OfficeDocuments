# AI instructions

Detailed working instructions for AI coding agents. The entry point is the root
[AGENTS.md](../../AGENTS.md); this folder holds the depth that would otherwise bloat it.

## Rule ownership

Each rule has exactly one home. If you need to change a rule, change it here — not in a second
place — and if you find the same rule stated twice, delete the weaker copy rather than editing both.

| File | Owns | Read it when |
| --- | --- | --- |
| [workflow.md](workflow.md) | How to approach a task, definition of done, backlog, commits, PRs | Starting any non-trivial task |
| [csharp.md](csharp.md) | Language level, nullability, validation, exceptions, performance | Writing or changing any `.cs` file |
| [excel.md](excel.md) | Excel object model, style pipeline, element ordering, module layout | Touching `src/OfficeDocuments.Excel` |
| [word.md](word.md) | Word object model, formatting records, OOXML child-order rules | Touching `src/OfficeDocuments.Word` |
| [testing.md](testing.md) | xUnit design rules and the schema-validation gate | Adding or changing tests |
| [build-and-packaging.md](build-and-packaging.md) | MSBuild layout, CPM, target frameworks, versioning, CI | Touching `*.csproj`, `Directory.*.props`, workflows |
| [documentation.md](documentation.md) | Where docs live, language, terminology, snippet accuracy | Writing or updating any `.md` |

Deliberately **not** owned here:

- Test-tier entry criteria — [../../test/README.md](../../test/README.md) is the authority.
- Core-vs-advanced scope decisions — [../architecture/minimal-core-pr-guidelines.md](../architecture/minimal-core-pr-guidelines.md).
- What to build next — [../tasks/roadmap-overview.md](../tasks/roadmap-overview.md).

## Precedence

When two sources disagree, resolve in this order:

1. The user's explicit request in the current task.
2. The most specific instruction file — module (`excel.md` / `word.md`) beats language (`csharp.md`) beats root `AGENTS.md`.
3. Existing patterns in the file you are editing.
4. Microsoft .NET guidance, then broadly accepted OSS convention.

A deviation from these files is allowed when the user asks for it. Keep it localized and say so in
your response.

## Maintaining these files

- Prefer one short code example over three paragraphs of prose.
- Every command must be copy-pasteable and work as written.
- Do not restate what the code, `README.md`, or git history already says. These files are for what
  an agent cannot infer by reading the repository — invariants, past bugs, and deliberate choices.
- When a rule stops being true, delete it. A stale rule is worse than a missing one, because the
  agent will follow it.
