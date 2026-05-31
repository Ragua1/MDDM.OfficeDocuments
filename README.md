# OfficeDocuments

`OfficeDocuments` is a .NET library for creating and reading XML-based Office documents through a smaller, consumer-friendly API over the Open XML SDK.

The repository currently ships two modules:

- `OfficeDocuments.Excel` for `.xlsx` creation, reading, styling, ranges, tables, annotations, validation, worksheet operations, protection, and worksheet images.
- `OfficeDocuments.Word` for lightweight `.docx` creation and reading through a small fluent API.

## Product direction

The project is intentionally scoped as a focused wrapper around XML Office formats.

- Supported formats: `.xlsx` and `.docx`
- Out of scope: legacy binary formats such as `.xls` and `.doc`
- Primary technical foundation: `DocumentFormat.OpenXml`
- Primary product goal: predictable, server-friendly document generation and data access without pushing OpenXml details into the default consumer workflow

## Platform baseline

- SDK baseline: `global.json` pins SDK `10.0.300` with `latestFeature` roll-forward
- Target frameworks: `net8.0`, `net9.0`, `net10.0`
- Language policy: latest stable C# supported by the installed major compiler

## Current highlights

### Excel

- File and stream workflows for create/open scenarios
- Workbook, worksheet, row, cell, and range APIs
- Typed reads plus formula, hyperlink, and comment support
- Style creation and style merging
- Bulk row insertion and object collection import
- Sorting, auto-filter, validation, conditional formatting, freeze panes, and auto-fit
- Worksheet rename, move, copy, hide, and remove operations
- Named ranges, worksheet/workbook protection, structured tables, and worksheet images

### Word

- File and stream workflows for create/open scenarios
- Fluent paragraph authoring with text and breaks
- Paragraph text reading from existing documents

## Documentation

- [.doc/README.md](.doc/README.md) - documentation index
- [.doc/excel-library.md](.doc/excel-library.md) - Excel consumer guide and API overview
- [.doc/word-library.md](.doc/word-library.md) - Word consumer guide and API overview
- [.doc/terminology.md](.doc/terminology.md) - shared terminology and abbreviations
- [.doc/library-benchmark-report.md](.doc/library-benchmark-report.md) - current capability benchmark and positioning
- [.doc/feature-gap-backlog.md](.doc/feature-gap-backlog.md) - remaining backlog after the current API review

## Contributing

When public behavior changes, update tests and keep `README.md` plus the relevant files in `.doc/` aligned with the real API.

## License

See [LICENSE.md](LICENSE.md).
