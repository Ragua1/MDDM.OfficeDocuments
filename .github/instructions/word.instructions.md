---
description: "Use when changing the Word library, fluent document-building API, paragraph/text behavior, or Word tests in MDDM.OfficeDocuments."
name: "Word Library Guidance"
applyTo:
  - "src/OfficeDocuments.Word/**/*.cs"
  - "test/OfficeDocuments.Word.Tests/**/*.cs"
---
# Word Guidance

- Treat `src/OfficeDocuments.Word` as a smaller, evolving surface. Prefer focused, additive changes over architecture rewrites.
- Preserve the existing fluent usage pattern: `GetBody() -> AddParagraph() -> AddText(...) / AddBreak(...)`.
- Keep the public API small and approachable. Avoid pushing Excel-specific abstractions or OpenXml implementation details into the Word surface.
- When extending Word behavior, keep docs and tests close to the feature so consumers can understand the supported workflow.
- Keep the root `README.md` high-level and place detailed Word guidance in `.doc/word-library.md`.
- Keep terminology aligned with `.doc/terminology.md`.
- Prefer modern, readable C# constructs, but do not add abstraction layers unless the change clearly benefits current functionality.
- Maintain correct document lifecycle handling: consumers should still be able to create, open, read, and close documents predictably.
