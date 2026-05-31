# Copilot instructions for MDDM.OfficeDocuments

Purpose: Help GitHub Copilot and other AI coding agents make small, correct, idiomatic changes in this .NET library that wraps OpenXml for Excel and Word.

These instructions are intentionally concrete. Prefer existing project patterns first, then Microsoft .NET guidance, then broadly accepted OSS conventions.

## Project map
- Solution: `OfficeDocuments.sln`
- SDK: use the version pinned in `global.json` (`10.0.300` with `latestFeature` roll-forward)
- Main library: `src/OfficeDocuments.Excel`
- Secondary library: `src/OfficeDocuments.Word`
- Tests: `test/OfficeDocuments.Excel.Tests` and `test/OfficeDocuments.Word.Tests`
- Public entry documentation: `README.md`
- Detailed documentation set: `.doc/README.md`

## Default working mode
- Prefer minimal, root-cause fixes over broad refactors.
- Preserve the existing public API unless the task explicitly requires an API change.
- Prefer the latest stable language and framework features that are compatible with the SDK pinned in `global.json` and the current target frameworks.
- Favor modern development principles: explicit contracts, strong tests, low incidental complexity, and small reversible changes.
- Treat performance and efficiency as first-class requirements when code touches document traversal, style creation, large ranges, XML merging, or repeated allocations.
- Keep edits aligned with the current file's style; when a file is inconsistent, follow the dominant repo style rather than reformatting unrelated code.
- Validate with targeted tests first, then widen only if needed.

## Instruction layout
- Keep this file focused on repository-wide rules.
- Put language-specific C# guidance in `.github/instructions/csharp.instructions.md`.
- Put area-specific rules in `.github/instructions/excel.instructions.md` and `.github/instructions/word.instructions.md`.
- Avoid duplicating the same rule in multiple files unless the duplication improves discovery or prevents ambiguity.

## Excel architecture
- Treat Excel as the primary, mature surface of the repository.
- Public API boundary is the interface layer in `src/OfficeDocuments.Excel/Interfaces`.
- Concrete behavior lives in `Spreadsheet.cs`, `DataClasses/*`, `Styles/*`, `Extensions/*`, and `Utils.cs`.
- Avoid leaking OpenXml types across public boundaries. If a feature needs OpenXml manipulation, keep that inside internal implementation code.
- Prefer implementation choices that keep worksheet and stylesheet operations efficient for larger documents, not just small happy-path examples.

## Excel object model rules
- Work through the established hierarchy: `ISpreadsheet -> IWorksheet -> IRow -> ICell`.
- `Spreadsheet` owns `SpreadsheetDocument`, `WorkbookPart`, and the stylesheet lifecycle. Do not bypass it with parallel document ownership.
- `Worksheet` lazily creates `Columns` and `MergeCells`; preserve that lazy behavior.
- `Row.CreateCell` backfills missing earlier cells to keep OpenXml order valid. Do not replace this with sparse append logic that breaks cell ordering.
- Row and column indexes are 1-based. Invalid indexes should continue to throw `ArgumentException` consistently.

## Excel style rules
- Centralize style creation through `Spreadsheet.CreateStyle(...)` and composition through `IStyle.CreateMergedStyle(...)`.
- Reuse existing style merge helpers in `Utils.MergeFonts`, `Utils.MergeFills`, and `Utils.MergeBorders` instead of inventing new merge logic.
- Keep current defaults for inferred number formats unless the feature explicitly changes them:
	- integers -> built-in format id `1`
	- floating point / decimal -> built-in format id `4`
	- dates -> built-in format id `14`
	- strings -> built-in format id `49`
- Be careful with alignment and number-format behavior: the current merge behavior is intentionally shallow compared to fonts/fills/borders.

## Excel API evolution guidance
- Prefer `AddCell(...)`; `AddCellWithValue(...)` is obsolete and retained for compatibility and legacy tests.
- When adding new public Excel features, extend interfaces first, then implement them in the internal classes.
- Do not add public APIs that require callers to know OpenXml internals.
- If a new operation touches large ranges, keep it linear in the requested range and avoid repeated DOM scans inside nested loops.
- Factory interfaces exist in `Factory/*`, but they are not the main extension seam today. Do not expand them unless the task specifically asks for factory-based composition.

## Word architecture
- Treat Word as a smaller, still-evolving surface. Make additive changes conservative and localized.
- Primary entry point is `src/OfficeDocuments.Word/Wordprocessing.cs` with public interfaces in `src/OfficeDocuments.Word/Interfaces`.
- Preserve the current fluent usage pattern: `GetBody() -> AddParagraph() -> AddText(...) / AddBreak(...)`.
- Avoid over-engineering Word features to match Excel architecture unless there is a concrete requirement.

## Validation and exceptions
- Match existing exception behavior whenever possible:
	- `ArgumentException` for invalid indexes or invalid user arguments
	- `ArgumentNullException` for required null inputs where the method contract requires a value
	- `InvalidOperationException` for broken document state or impossible internal conditions
- Keep validation close to the public or boundary method that receives the input.
- Do not silently coerce invalid indexes, references, or sheet state.

## Tests
- Tests are xUnit.
- Favor focused tests in the affected test project over solution-wide runs.
- Add both happy-path and edge-case coverage for new behavior.
- Name tests as `MethodName_StateUnderTest_ExpectedOutcome` for new or renamed tests.
- Follow the existing patterns in `SpreadsheetTestBase`, `CreationTest`, `CellTest`, `RowTest`, `StyleTest`, `WorksheetTest`, and Word `CreationTest`.
- Keep tests deterministic: prefer temp files/streams, avoid machine-specific paths, and skip external-resource tests when necessary.

## Documentation
- Keep `README.md` general, short, and product-oriented. Put detailed usage guidance, architecture notes, backlog material, and glossary content into `.doc/`.
- Documentation in `README.md` and `.doc/` must be written in English and use clear, technical, unambiguous language.
- Use relative links between documentation files inside `.doc/`.
- Use `.doc/terminology.md` for preferred project terms, abbreviations, and naming consistency.
- When you add or change a public API, update `README.md` only if the high-level product overview changes, and update the relevant detailed docs in `.doc/`.
- Ensure documentation snippets match the actual public API. Verify against the interface definitions, implementation, and tests before copying patterns forward.
- Documentation must be sufficient for a new consumer to understand the library surface, the intended workflow, and the major capabilities of both Excel and Word modules.
- Public-facing detailed docs should explain not only how to create a document, but also how to read values, apply styles, work with ranges, configure columns, add tables, and close or dispose resources correctly where relevant.
- Prefer short, copy-pastable examples that create, use, and dispose documents correctly.

## Typical validation commands
- `dotnet test test/OfficeDocuments.Excel.Tests/OfficeDocuments.Excel.Tests.csproj`
- `dotnet test test/OfficeDocuments.Word.Tests/OfficeDocuments.Word.Tests.csproj`
- `dotnet test OfficeDocuments.sln`
- `dotnet pack src/OfficeDocuments.Excel/OfficeDocuments.Excel.csproj -c Release`

## Build configuration
- `Directory.Build.props` — centralized compilation settings (`LangVersion`, `Nullable`, `ImplicitUsings`, `TargetFrameworks`) and shared NuGet metadata.
- `Directory.Packages.props` — Central Package Management (CPM); all `PackageVersion` entries live here. Individual project files must NOT include `Version` on `PackageReference` items.
- Do not add `GeneratePackageOnBuild` to any project; call `dotnet pack` explicitly in CI/CD.
- Do not add `Newtonsoft.Json` or `System.IO.Packaging` to project files; they are unused and absent intentionally.

## High-value files
- `src/OfficeDocuments.Excel/Spreadsheet.cs`
- `src/OfficeDocuments.Excel/DataClasses/Worksheet.cs`
- `src/OfficeDocuments.Excel/DataClasses/Row.cs`
- `src/OfficeDocuments.Excel/DataClasses/Cell.cs`
- `src/OfficeDocuments.Excel/DataClasses/Style.cs`
- `src/OfficeDocuments.Excel/Utils.cs`
- `src/OfficeDocuments.Word/Wordprocessing.cs`
- `test/OfficeDocuments.Excel.Tests/SpreadsheetTestBase.cs`
- `test/OfficeDocuments.Excel.Tests/CreationTest.cs`
- `test/OfficeDocuments.Excel.Tests/StyleTest.cs`
- `test/OfficeDocuments.Word.Tests/CreationTest.cs`

## Decision checklist for agents
- Is the change on the public API? Update the interface, implementation, tests, and README together.
- Is the change style-related? Reuse stylesheet helpers before touching raw OpenXml.
- Is the change range-related? Preserve OpenXml ordering and avoid quadratic scans.
- Is the change Word-related? Keep it small, focused, and compatible with the current fluent model.
- Is there a more modern language or framework feature that improves clarity or efficiency without forcing churn? Prefer it.
- Does the change add repeated traversal, repeated merging, or avoidable allocations in a hot or wide path? Rework it before finishing.
- Is there already a nearby test covering the same surface? Extend it before creating a new broad test class.

If the task conflicts with these rules, prefer the user's explicit request and keep the deviation localized and documented in the response.
