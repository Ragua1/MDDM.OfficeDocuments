# CLAUDE.md

[AGENTS.md](AGENTS.md) is the primary instruction file for this repository. Everything in it applies.
This file adds only what is specific to Claude Code.

@AGENTS.md

## Loading instructions

The routing table in `AGENTS.md` points at [`.docs/ai-instructions/`](.docs/ai-instructions/README.md).
Read the one file that matches the area you are touching, when you get there. Do not pull the whole
folder into context up front — it exists so `AGENTS.md` can stay short.

## Local environment

- Windows with PowerShell. `dotnet` commands run from the repository root as written in `AGENTS.md`.
- Three `.slnx` solutions, no `.sln`. Work through the module solution you are changing
  (`OfficeDocuments.Excel.slnx` or `OfficeDocuments.Word.slnx`); `OfficeDocuments.slnx` is for changes
  that span both. Parsing any of them requires the SDK pinned in `global.json`.
- `bin/` and `obj/` are gitignored — ignore them when searching, they contain generated `.cs` files
  that will otherwise pollute results.
- Build output is Czech (`Počet chyb: 0`, `0 upozornění`). Grep for those when checking a build, not
  for `error` / `warning`.
