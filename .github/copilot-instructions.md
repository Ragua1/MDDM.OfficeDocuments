# Copilot Instructions for MDDM.OfficeDocuments

**Purpose**: Help AI agents work productively in this .NET library that wraps OpenXML for Excel (mature) and Word (early). Keep edits idiomatic to this repo and follow Microsoft best practices.

---

## ⚡ Critical Rules (Non-Negotiable)

1. **Dependency Injection**: NEVER instantiate services with `new`. Always use Constructor Injection where applicable (note: this library is primarily used as a utility, not an ASP.NET app, so DI applies mainly to test mocking).

2. **Async/Await**: ALL I/O operations (file streams, network) must be async. Always accept and propagate `CancellationToken` in async methods.

3. **Logging**: When adding logging support, use `ILogger<T>`. STRUCTURED LOGGING ONLY. No string interpolation in log messages—use structured parameters.

4. **Documentation**: Public APIs MUST have XML `<summary>`, `<param>`, and `<returns>` documentation. Update docs when modifying public APIs.

5. **Security**: Validate external inputs immediately using Guard Clauses. Throw `ArgumentNullException` for null checks, `ArgumentException` for validation failures.

6. **Licensing**: Prefer libraries with permissive licenses (MIT/Apache) for open-source compatibility. Avoid paid or restrictive licenses. Current dependencies (DocumentFormat.OpenXML) use MIT license.

7. **Build Integrity**: If code is modified, ALWAYS trigger a build before execution/debugging. For .NET projects, let projects build automatically or run `dotnet build` explicitly. NEVER use `--no-build` flag when testing changes.

8. **Documentation Maintenance**: Maintain `README.md` for the repository and update incrementally with significant changes. Document breaking changes and new public APIs with usage examples.

9. **Configuration & Dependencies**: Dependencies are managed via `.csproj` (NuGet). This library has minimal configuration; settings are passed via constructor parameters or method arguments.

10. **Task Management**: When analysis yields tasks, document them clearly. Organize work into manageable units. Update documentation after completing tasks.

11. **Analysis Considerations**: When analyzing issues, verify unit tests for correctness. Assume tests may contain bugs—validate test logic alongside production code.

12. **Orchestration System**: Follow structured task management. All orchestration files and commit messages are in English for consistency and international collaboration.

---

## 📝 Code Style Guidelines

### General C# Style
- **`var` Usage**: Use `var` when the type is obvious from the right-hand side (e.g., `var list = new List<string>()`). Use explicit types when clarity is needed.
- **Formatting**: 
  - File-scoped namespaces (`namespace MDDM.OfficeDocuments.Excel;`)
  - Target-typed `new()` expressions where applicable
  - Collection expressions `[]` for initializers (C# 12+)
- **Nullability**: `#nullable enable` always. Treat warnings as errors. All new code must be nullable-reference-type aware.
- **Models**: Prefer `record` for DTOs and configuration objects. Use `init` accessors for immutable properties where appropriate.

### Testing Standards
- **Framework**: xUnit for all tests
- **Assertions**: Use FluentAssertions or similar for readable test assertions
- **Mocking**: Use Moq when needed (minimal in this library since it focuses on file I/O)
- **Naming Pattern**: `MethodName_StateUnderTest_ExpectedOutcome`
  - Example: `AddCell_WithNullValue_ThrowsArgumentNullException`
  - Example: `CreateStyle_WithValidFont_ReturnsStyleWithCorrectId`
- **Coverage**: Include both positive (happy path) and negative (error/edge) cases
- **Speed**: Keep tests fast and deterministic. Avoid long-running integration tests in main suite.

### Specific to This Library
- **Indentation**: Use 4 spaces (consistent with existing code)
- **Naming Conventions**:
  - Interfaces: `ISpreadsheet`, `IWorksheet`, etc.
  - Enums: PascalCase (e.g., `BorderStyleValues`, `HorizontalAlignmentValues`)
  - Private fields: `_camelCase` with underscore prefix
- **Exceptions**: Consistent exception types
  - `ArgumentNullException` for null arguments
  - `ArgumentException` for invalid values
  - `InvalidOperationException` for state-related errors

---

## 🏗️ Architecture & Project Structure

### Layout
- **`src/OfficeDocuments.Excel`**: Main library for Excel manipulation (mature, production-ready)
- **`src/OfficeDocuments.Word`**: Word document support (early stage, scaffold only)
- **`test/*`**: xUnit test projects
- **`.github/workflows`**: CI/CD pipeline definitions

### Public API Design
- **Public Interfaces** in `Interfaces/*`:
  - `ISpreadsheet`, `IWorksheet`, `IRow`, `ICell`, `IStyle`
- **Internal Implementations** in `DataClasses/*` and `Styles/*`:
  - `Spreadsheet`, `Worksheet`, `Row`, `Cell`, `Style`
- **Key Principle**: NEVER expose OpenXML types in public APIs. Always use abstraction layer.

### Excel Architecture & Patterns

#### Core Abstractions
```
ISpreadsheet → IWorksheet → IRow → ICell
```
- Use these abstractions; don't pass OpenXML types (`WorkbookPart`, `SheetData`, etc.) across public boundaries.

#### Style Management
- Create styles via `Spreadsheet.CreateStyle(font, fill, border, alignment, numberingFormat)`
- Compose styles with `CreateMergedStyle(otherStyle)` for style inheritance
- Default style initialized in `Spreadsheet.InitStylesheet()`; style ID 0 = default
- **Style Merge Behavior**:
  - Fonts/Fills/Borders: merged deeply by `Utils.MergeFont`, `Utils.MergeFill`, `Utils.MergeBorder`
  - NumberFormat and Alignment: applied but not deeply merged
  - Numeric/Date setters add common defaults if none specified

#### Lazy Creation Patterns
- `Worksheet` lazily creates `Columns` and `MergeCells` parts
- Creating a cell may backfill missing earlier cells to maintain OpenXML document order
- Avoid redundant DOM traversals; prefer lazy initialization

#### Tables
- Use `ISpreadsheet.AddTable(sheetName, startCell, endCell, columns)` to create Excel tables
- Creates `TableDefinitionPart` and updates `TableParts` collection
- Tables enable filtering, sorting, and structured references

---

## 🔒 Conventions & Gotchas

### Index Semantics
- **1-based indexing** for rows and columns (Excel convention)
- Invalid indices (`< 1`) throw `ArgumentException`
- Example: Row 1, Column 1 = Cell A1

### Method Preferences
- **Prefer**: `AddCell(value)` - recommended API
- **Avoid**: `AddCellWithValue(value)` - marked `[Obsolete]`, kept for legacy test compatibility

### Resource Management
- `Close()` saves the document ONLY when it's editable
- Always call `Dispose()` or `Close()` to flush changes and release file handles
- Use `using` statements for automatic disposal

### Input Validation
- Keep argument validation, naming, and null-handling consistent with existing methods
- Validate at the entry point (public API boundary)
- Use descriptive error messages

---

## 🧪 Testing Strategy

### Test Organization
- **Location**: `test/OfficeDocuments.Excel.Tests`
- **Base Classes**: Inherit from `SpreadsheetTestBase` for common test utilities
- **Key Test Files**:
  - `CreationTest.cs` - Document and element creation
  - `UtilsTest.cs` - Utility function tests
  - Style and formatting tests

### Test Requirements
1. **Positive Cases**: Verify happy paths work correctly
2. **Negative Cases**: Verify error handling (invalid inputs, edge cases)
3. **Naming**: Follow `MethodName_StateUnderTest_ExpectedOutcome` pattern
4. **Speed**: Tests should complete quickly (< 1 second each ideally)
5. **Determinism**: No flaky tests; results must be reproducible

### Example Test Pattern
```csharp
[Fact]
public void AddCell_WithValidString_CreatesCellWithCorrectValue()
{
    // Arrange
    var spreadsheet = CreateTestSpreadsheet();
    var worksheet = spreadsheet.AddWorksheet("Test");
    var row = worksheet.AddRow();
    
    // Act
    var cell = row.AddCell("TestValue");
    
    // Assert
    cell.Value.Should().Be("TestValue");
}
```

---

## 🛠️ Development Workflow

### Before Making Changes
1. Understand the existing architecture and patterns
2. Check for similar existing functionality
3. Review related tests to understand expected behavior
4. Ensure you have .NET SDK from `global.json` (9.0.0)

### Making Changes
1. **Interface Changes**: Update `Interfaces/*` first, then implementations
2. **Implementation**: Add/modify code in `DataClasses/*` or `Styles/*`
3. **Tests**: Add tests following existing patterns
4. **Documentation**: Update XML docs for public APIs
5. **README**: Update examples if adding new public features

### Validation Checklist
- [ ] Code builds without warnings (`dotnet build`)
- [ ] All tests pass (`dotnet test`)
- [ ] XML documentation added/updated for public APIs
- [ ] README updated with usage examples (if applicable)
- [ ] No OpenXML types exposed in public APIs
- [ ] Input validation consistent with existing patterns
- [ ] Performance considerations addressed (no O(n²) loops in hot paths)

---

## 🎯 Extending the Library Safely

### Adding New Features
1. **Define Interface**: Add to `Interfaces/*` (e.g., `IWorksheet`)
2. **Implement Internally**: Add implementation in `DataClasses/*` (mark as `internal`)
3. **Style Management**: Reuse `Style`/`Utils` helpers for stylesheet management
4. **Validation**: Validate inputs consistently; maintain 1-based semantics
5. **Testing**: Add tests under `test/OfficeDocuments.Excel.Tests`
6. **Documentation**: Add XML docs and README examples

### Large-Range Operations
- Ensure OpenXML node order is preserved
- Backfill logic should remain O(range), not O(range²)
- Test with large datasets to verify performance

### Common Pitfalls
- ❌ Don't expose `SpreadsheetDocument`, `WorkbookPart`, `Cell` (OpenXML types) in public APIs
- ❌ Don't use 0-based indexing (Excel uses 1-based)
- ❌ Don't modify OpenXML DOM directly in feature code (use abstraction layer)
- ✅ Do use interface types (`ISpreadsheet`, `IWorksheet`, etc.)
- ✅ Do validate inputs at public boundaries
- ✅ Do follow existing patterns for consistency

---

## 📚 Key Files Reference

### Core Implementation
- `src/OfficeDocuments.Excel/Spreadsheet.cs` - Main entry point, workbook management, stylesheet initialization
- `src/OfficeDocuments.Excel/DataClasses/Worksheet.cs` - Worksheet operations
- `src/OfficeDocuments.Excel/DataClasses/Row.cs` - Row operations
- `src/OfficeDocuments.Excel/DataClasses/Cell.cs` - Cell operations
- `src/OfficeDocuments.Excel/DataClasses/Style.cs` - Style management

### Styles & Utilities
- `src/OfficeDocuments.Excel/Styles/*` - Style-related classes (Font, Fill, Border, etc.)
- `src/OfficeDocuments.Excel/Enums/*` - Enumeration types
- `src/OfficeDocuments.Excel/Utils.cs` - Utility functions

### Tests
- `test/OfficeDocuments.Excel.Tests/CreationTest.cs` - Creation and basic operations
- `test/OfficeDocuments.Excel.Tests/UtilsTest.cs` - Utility function tests
- `test/OfficeDocuments.Excel.Tests/SpreadsheetTestBase.cs` - Test base class

---

## 💡 Best Practices Summary

1. **Follow Microsoft Guidelines**: Prioritize official .NET and C# conventions
2. **Senior-Level Approach**: Propose minimal, high-impact changes; explain trade-offs
3. **Consistency**: Match existing patterns in naming, structure, and error handling
4. **Performance**: Avoid O(n²) algorithms; use lazy initialization where appropriate
5. **Testing**: Write tests first or alongside implementation; cover edge cases
6. **Documentation**: Keep XML docs and README current with code changes
7. **Build Quality**: Ensure builds pass with SDK from `global.json` (9.0.0)

---

## ❓ When in Doubt

If anything is unclear (e.g., alignment merge nuances, conditional formatting scope, or new feature design), **flag it** so we can refine these rules and ensure consistency across the codebase.

**Remember**: This library prioritizes clean abstractions over OpenXML complexity. Keep the public API simple, intuitive, and well-documented.
