# Excel Core Tasks

Date: 2026-05-31

- [EXCEL-005 Style pipeline performance hardening](EXCEL-005-style-pipeline-performance-hardening.md)
  — open; the hot spots are now measured, see [../../../excel-performance-baseline.md](../../../excel-performance-baseline.md)
- [EXCEL-006 Worksheet and row lookup indexing](EXCEL-006-worksheet-and-row-lookup-indexing.md) — open
- [EXCEL-007 OpenXml compatibility surface isolation](EXCEL-007-openxml-compatibility-surface-isolation.md)
  — open, the remaining `P0`; test-suite migration off the compatibility surface delivered 2026-08-06,
  so nothing but consumer impact now blocks removing those members
- [EXCEL-010 God-class decomposition (Worksheet and Spreadsheet)](EXCEL-010-god-class-decomposition.md)
  — Tier 1 delivered 2026-07-27; Tier 2 (physical `Excel.Advanced` split, breaking/v4) delivered 2026-08-06
- [EXCEL-011 Test suite restructuring (unit / integration / verification / performance)](EXCEL-011-test-suite-restructuring.md)
  — phases 1–6 delivered 2026-07-27/28; blind spots B-7..B-15 open
