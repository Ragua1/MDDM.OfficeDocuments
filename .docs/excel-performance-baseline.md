# Excel performance baseline

Measured 2026-07-27 · `EXCEL-011` phase 5

This is the reference the guards in `test/OfficeDocuments.Excel.PerformanceTests` are calibrated
against. When a threshold there looks wrong, the answer is here — and when a hot spot is fixed,
the numbers here are what the fix has to beat.

## How these were produced

```powershell
dotnet run -c Release --project test/OfficeDocuments.Excel.Benchmarks -- --filter '*'
```

| | |
| --- | --- |
| CPU | AMD Ryzen 9 5900X, 12 physical / 24 logical cores |
| OS | Windows 11 26200 |
| Runtime | .NET 10.0.10, X64 RyuJIT x86-64-v3, SDK 10.0.302 |
| Harness | BenchmarkDotNet 0.15.8, `InvocationCount=1 UnrollFactor=1 WarmupCount=1 IterationCount=5` |

One invocation per iteration, because each workload builds and discards a whole workbook and must
not be batched. Absolute values are specific to this machine; **the ratios are the durable part.**

## Summary

Four paths cost more than linear in the size of their input. All four were already known from
[`excel-state-verdict.md`](excel-state-verdict.md); this is the first time they have been
quantified.

| Path | Growth for 4x input | Cost at the largest size measured |
| --- | --- | --- |
| `CreateStyle` with distinct styles | ~16x allocation, worse in time | 1 000 styles → 3.1 s, 1.2 GB |
| `SetComment` | 16x | 200 comments → 166 ms, 59 MB |
| `Row.CreateCell` backfill | ~9x | one cell at column 4 000 → 101 ms |
| `Range.SortByColumn` | 2x allocation over building the range | 2 000 rows → +67 MB |

And two that are linear and expected to stay that way:

| Path | Growth | Cost |
| --- | --- | --- |
| `AddRows<T>` | linear; ~19 KB per 4-column row | 10 000 rows → 1.15 s, 185 MB |
| `CreateStyle` reusing 8 styles | linear | 1 000 calls → 35 ms, 14 MB |

## Style allocation

`Style.GetFontId` and its siblings resolve a style by walking every entry already in the
stylesheet and comparing it structurally, so N distinct styles perform O(N²) comparisons.

| Styles | Distinct | Reusing 8 |
| ---: | ---: | ---: |
| 250 | 86 ms · 75 MB | 8.4 ms · 3.6 MB |
| 500 | 355 ms · 300 MB | 16.4 ms · 7.2 MB |
| 1 000 | 3 065 ms · 1 196 MB | 35.4 ms · 14.3 MB |

Allocation quadruples for every doubling — textbook quadratic. Time grows faster still, because a
gigabyte of garbage starts costing Gen1 and Gen2 collections that the smaller runs never trigger.

That second effect is why the complexity guard for this path asserts on allocation rather than on
time: at 4x the input the wall clock moves by 33x, which sits between quadratic (16x) and cubic
(64x) and therefore identifies neither. Allocation says it cleanly.

**The workaround is real and cheap.** Reusing eight styles across the same number of calls is 87x
faster and allocates 84x less. Callers deriving a style per row should hoist it.

## Comments

`CommentWriter.Set` re-serializes the entire comments part and rebuilds the entire legacy VML
drawing on every single call, so per-call cost grows with the number of comments already attached.

| Comments | With comments | Same cells, no comments |
| ---: | ---: | ---: |
| 50 | 10.3 ms · 5.0 MB | 1.9 ms · 0.4 MB |
| 100 | 34.2 ms · 16.1 MB | 3.6 ms · 0.7 MB |
| 200 | 166.3 ms · 58.8 MB | 8.3 ms · 1.3 MB |

A clean 16x for 4x the input. Per comment this is about 160 KB against roughly 7 KB for a plain
cell — the most expensive feature in the library per unit of content, and the steepest of the four
hot spots. Extrapolating, a thousand comments costs seconds and over a gigabyte.

## Row backfill

`Row.CreateCell` fills in every missing cell up to the requested column so children stay in
ascending order, and each one is positioned by a linear scan in `InsertCell`.

| Column | One cell at that column | N cells written in order |
| ---: | ---: | ---: |
| 1 000 | 11.0 ms · 1.1 MB | 31.0 ms · 5.0 MB |
| 2 000 | 39.5 ms · 2.0 MB | 84.3 ms · 9.8 MB |
| 4 000 | 101.1 ms · 3.4 MB | 145.7 ms · 18.5 MB |

Allocation is linear — the defect is pure CPU spent scanning, which is why this one cannot be
guarded by allocation and needs a timing ratio.

Note the shape: writing a single far cell allocates a third of what writing the whole row does,
yet grows more than twice as fast. A sheet with a wide header and sparse rows pays this on every
row.

## Range sort

`Range.SortByColumn` snapshots the range by deep-cloning every cell, then writes all of them back
through the normal cell API.

| Rows (× 5 columns) | Build only | Build + sort |
| ---: | ---: | ---: |
| 500 | 16.1 MB | 31.5 MB |
| 1 000 | 33.5 MB | 67.2 MB |
| 2 000 | 67.0 MB | 134.1 MB |

Exactly 2.00x at every size: the clone costs one extra copy of the range, no more and no less.
The wall-clock measurements here were the noisiest in the suite and are not worth quoting —
building the range dominates the sort, so allocation is the statement that means something.

## Whole-document write and read

| Rows (× 4 columns) | `AddRows<T>` | `AddCell` per field | Write + close + reopen + read |
| ---: | ---: | ---: | ---: |
| 2 000 | 200 ms · 38.7 MB | 131 ms · 53.1 MB | 217 ms · 51.7 MB |
| 5 000 | 402 ms · 96.7 MB | 383 ms · 132.6 MB | 445 ms · 129.5 MB |
| 10 000 | 1 148 ms · 184.6 MB | 1 178 ms · 248.2 MB | 1 466 ms · 260.5 MB |

Allocation is flat per row — about 19 KB — at every size, so the path is linear. The wall clock
grows a little faster than the data because of collection pressure, which is the general reason a
timing ratio cannot on its own prove linearity here.

`AddRows<T>` allocates about 27% less than writing the same data cell by cell, despite the
reflection, and matches it on time from 5 000 rows up.

**A practical figure to quote: 10 000 rows by 4 columns, written and read back, is about 1.5
seconds and 260 MB.**

Measured on its own — opening a finished workbook and reading every row, with the writing done
outside the measured region — the read path costs 4.7x for 4x the rows, which is linear within
measurement error. Both axes of cell lookup go through a dictionary, so that is the expected
result; the guard exists to notice if either one ever becomes a scan.

Worth stating explicitly, because the reverse is a tempting shortcut: **the combined round trip
cannot be used as the read-path guard.** Reading is the smaller share of it, so a read path that
turned quadratic would still leave the combined ratio near 8 — under any ceiling loose enough not
to flake. A growth ratio only means something when it is taken over one path at a time.

## What this does and does not justify

The guards in `test/OfficeDocuments.Excel.PerformanceTests` pin these numbers so they cannot get
worse. That is all they do. A passing `KnownHotSpotGuards` run does **not** mean the library is
fast on these paths — it means it is no slower than it was on 2026-07-27.

Fixing them is `EXCEL-005`. The order the numbers argue for:

1. **Style dedup** — worst absolute cost, and the fix is well understood: hash the candidate
   instead of scanning. A dictionary keyed on the element's structural content turns O(N²) into
   O(N).
2. **Comments** — steepest growth, and the cause is a `Save()` and a full VML rebuild inside a
   per-call method. Defer both to close.
3. **Row backfill** — the insertion scan is only needed because cells are kept in a list;
   the row already maintains `_cellsByColumnIndex`.
4. **Range sort** — the clone is the cost. An in-place reorder of the existing elements avoids it.

When one is fixed, move its guard from `KnownHotSpotGuards` to `LinearScalingGuards` and update
the table at the top of this file. That migration is the definition of done.
