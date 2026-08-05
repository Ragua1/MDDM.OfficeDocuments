# Support

`OfficeDocuments` is maintained by one person, in spare time, alongside a full-time job. It is used
in production — including mine — and it is not a funded project. This file exists so you can tell
those two facts apart before you depend on it.

**There is no SLA.** Nothing here is a commitment to respond, to fix, or to release within any period
of time. If you need guaranteed support, this is the wrong dependency, and saying so up front is
more useful to you than an unanswered issue three weeks from now.

What *is* a commitment: the library will not be silently abandoned. If maintenance stops, the
repository will say so.

## Where to go

| You want to | Go to |
| --- | --- |
| Report a bug | [Issues](https://github.com/Ragua1/MDDM.OfficeDocuments/issues) — use the bug report template |
| Request a feature | [Issues](https://github.com/Ragua1/MDDM.OfficeDocuments/issues) — use the feature request template |
| Ask "how do I…" | [Discussions](https://github.com/Ragua1/MDDM.OfficeDocuments/discussions), after checking the guides below |
| Report a security issue | See [Security issues](#security-issues) — **not** a public issue |
| Contribute a fix | Read [AGENTS.md](AGENTS.md), then open a pull request |

Before opening anything, check [.doc/excel-library.md](.doc/excel-library.md) or
[.doc/word-library.md](.doc/word-library.md). They document the actual public API and its semantics,
including several behaviours that are correct but surprising — merged styles folding in the workbook
default font, or a style reaching a cell through the sheet becoming a new stylesheet entry. A
question answered there will be closed with a link to it.

## What gets attention first

Roughly in this order:

1. **A file that Excel or Word refuses to open, or offers to repair.** This is the highest-severity
   class in this project. It usually means malformed output that passed every self-consistency check
   — see the *Correctness* section of the [README](README.md#correctness) for why that happens.
2. **Data corruption or a wrong value** written or read — a wrong date, a truncated number, a value
   that changes on reopen.
3. **A crash or an exception** on input the documentation says is supported.
4. **A missing feature that is already on the roadmap.**
5. **Everything else.**

Items 4 and 5 may sit for a long time. That is not a judgement on the request.

## Filing a bug that can be acted on

The single most useful thing you can provide is **a runnable snippet that produces the bad file**,
because that is what turns into a regression test. Roughly twenty lines against a `MemoryStream` is
ideal.

Please include:

- Package and version (`OfficeDocuments.Excel 4.0.0`, or the commit if you built from source).
- Target framework — `net8.0`, `net9.0`, or `net10.0`.
- OS, and which application opens the file (Excel version, LibreOffice, Google Sheets…).
- What you expected, and what happened instead.
- The exception with its stack trace, if there is one.
- **Culture, if any number, date, or decimal is involved.** Culture-dependent bugs are invisible
  without it.

If the problem is a file Excel or Word will not open, attach the produced file, or the exact repair
message. If the input is a document you did not create, please check it for confidential or personal
data before attaching it — a redacted file that still reproduces the problem is worth far more than
one that cannot be shared. If it cannot be shared at all, say so and describe the structure; that is
still workable.

**Please do not** attach a screenshot of code, or paste a full application. Reduce it first — that
step alone resolves a good share of reports.

## Feature requests

Say what you are trying to produce, not only which API you want. The design of this library is
deliberately narrow, and the useful answer is often a different shape than the one requested.

Two standing scope decisions, so nobody spends effort on them:

- **Legacy binary `.xls` and `.doc` are permanently out of scope.** Not "not yet".
- **A formula calculation engine is not planned.** The library writes formulas; Excel evaluates them.

Everything else is negotiable. The current thinking lives in
[.doc/tasks/roadmap-overview.md](.doc/tasks/roadmap-overview.md) and
[.doc/feature-gap-backlog.md](.doc/feature-gap-backlog.md); a request that fits a gap already
recorded there has a much better chance of getting built.

## Versions and breaking changes

- `OfficeDocuments.Excel` and `OfficeDocuments.Word` are versioned **independently**. They ship as
  separate packages and share no code, so their version numbers do not track each other.
- Semantic versioning. A change that makes previously accepted input throw is a breaking change and
  takes a major bump, even though it only ever prevented a broken file.
- **Only the latest major of each package is maintained.** There are no backports to older majors.
- Target frameworks follow .NET support: a framework is dropped once Microsoft ends its support.

## Security issues

Please do **not** open a public issue for a security problem.

Report it privately through
[GitHub Security Advisories](https://github.com/Ragua1/MDDM.OfficeDocuments/security/advisories/new),
or by email to the address on the [maintainer's GitHub profile](https://github.com/Ragua1).

Realistic scope for a library like this one: malformed or hostile `.xlsx` / `.docx` input causing
unbounded memory or CPU consumption, path traversal via package parts, XML entity expansion, or a
crash reachable from untrusted document input. If your service accepts documents from users, those
are the paths that matter.

Expect an acknowledgement when I next have time at a machine — days, not hours. A fix will be
released as a patch on the current major.

## Commercial support

None is offered, and none is planned. This project has never had a commercial intent, and the
license reflects that: [MIT](LICENSE.md), so you can use it, fork it, and ship it in a closed-source
product with no obligation beyond keeping the copyright notice.

If you need something specific and urgent, a pull request is the fastest route by a wide margin.
