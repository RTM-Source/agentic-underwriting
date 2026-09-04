# UW Automation — Project Guide

> **Public-clone note:** this is the live working rulebook from the private team repo, published
> verbatim as part of the process demonstration — only deal identities and live-deal values were
> scrubbed. Rules referencing team workflow (private remote, push discipline) describe the private
> repo's operation.

Lean by design: rules only, one line each where possible; rationale lives in the memories, skills,
and reference docs each rule points at. Hard rules are enforced by the PreToolUse hook
(`.claude/hooks/guard-hard-rules.ps1`, wired in `.claude/settings.json`) — blanket Excel kills and
COM `SaveAs` to OneDrive paths are blocked there, not just requested here. When a rule gets broken
anyway, treat it as a bug against this file (or a new hook) and amend.

## Response style
Short and plain. Lead with the result; report outcomes, not process ("Done X. Result Y. Errors: 0").
Surface only what changed, broke, or needs a decision. Thorough work, terse text.

## Model switching — Opus reasons, Sonnet grinds, Haiku measures
Main thread stays on Opus (design, financial-correctness calls, orchestration, flag triage, the
GATE-0 read-back — re-open and check numbers yourself before archiving/reporting). **The main thread
builds nothing bulky itself and never ingests a stage transcript or sheet dump** — delegate every
fully-specified grunt run (incl. `.H` assembly via `h-assembler`) to a Sonnet sub-agent with the
script/skill path + precise pass/fail checks; never let it redesign financial logic. Haiku gets ONE
kind of seat: pure deterministic recompute with zero judgment (`gate-runner`), always backstopped by
a stronger different-model auditor that sample-checks its evidence. Spawn a sub-agent instead of
flipping the main loop's model (kills the prompt cache).

## Repo layout & hygiene
- Root is one folder per team; `Investments/` holds the underwrite pipeline
  (`Data/ Decks/ Docs/ arc/ lib/ scripts/`). `.claude/` stays at root, shared across teams.
- Git tracks `.claude/`, `Investments/{Data,Docs,lib,scripts}/` — `Data/` is tracked as of
  2026-07-21 (deals under `Data/Transactions/<Deal>/`, synced with the private team repo, PRIVATE;
  `*.pptx` still ignored). Commit after changing skills, agents, lib, docs, or deal data — small,
  plain messages. **Commit locally only — NEVER `git push`; Ryan pushes from his own terminal**
  (ruling 2026-07-22). Report unpushed-commit count when a work session ends.
- Edit deliverables by writing a new `_v#` copy, never by saving over the input (memory
  `ps-var-name-case-collision`); after the new version verifies, move the old one to sibling `arc/`.
- Any deliverable built on imperfect data gets a `Notes` tab as sheet 1: the defect, the ruling and
  who made it, the impact on specific figures, what to request from the broker. (Worked-example
  script in the private team repo; not included in this public clone.)

## Environment (non-obvious — do not rediscover)
- No Python on this box: all `.xlsx` work is Excel COM in PowerShell (Office 16). Recalc with
  `CalculateFull()` before `Save()`, not after.
- Excel-touching agents need the `PowerShell` tool in frontmatter and use it natively — batch COM
  ops into single calls (open → do many things → save → close), not write-a-`.ps1`-and-`Bash` (10x
  measured overhead). Reserve `Bash` for file/dir work.
- Build engine: dot-source `Investments/lib/HF-Build-Lib.ps1` (`New-Excel`, `SetV`, `HF-*`/`Map-*`
  writers, `New-ExcelTracked`/`Stop-TrackedExcel`) instead of hand-rolling COM styling. Caller
  contract example: `Investments/scripts/Build-AsterRidge-HF.ps1` — a template to read, deal-specific,
  never run against another community. Write pitfalls: memory `excel-com-write-pitfalls`.
- Pre-run cleanup: dot-source `Investments/lib/HF-Build-Lib.ps1` and call `Clear-OrphanExcel` (kills
  only leftover Excel no *live tracked* run owns — safe with sibling runs in other terminals); create
  Excel via `New-ExcelTracked`, release via `Stop-TrackedExcel` in `finally`. The raw `MainWindowHandle
  -eq 0 | Stop-Process` sweep is now **hook-BLOCKED** — a COM Excel is windowless its whole life, so a
  sweep force-kills a concurrent run's Excel (memory `never-blanket-kill-excel`). **Concurrency (Ryan,
  2026-07-22):** independent deals in separate terminals may run fully in parallel (the 6-terminal
  3-pack = 3 deals × `.RR`+`.H`) because tracked cleanup never cross-kills. **Within one `.H` run**
  (2026-07-28): census ∥ HF-format run in parallel (census in its own scratch workbook); PnL → assembly →
  verify stay serial. In multi-agent runs `Clear-OrphanExcel` is orchestrator-only, called once at setup.
- Wrap every build body in `try { … } finally { $wb.Close(); $xl.Quit(); release }`; assert the
  input path resolved (`if(-not $src){throw}`) before opening.
- New workbook on a OneDrive path: `Copy-Item` a seed file, then `Open` + `Save()` in place — never
  `SaveAs` (hook-enforced; hangs silently, memory `word-com-saveas-hang`). If `Open` fails,
  re-resolve the path first (OneDrive moves files mid-session).
- Gridlines off on every sheet created or touched.

## Skills & pipeline
- Skills: `.claude/skills/{census-formatting,hf-formatting,pnl-mapping,h-underwrite,rr-formatting}/`
  (each with `References/`). Pipeline map: `.claude/skills/H-PIPELINE-ORCHESTRATION.md`.
- `.H` pipeline: (census ∥ format HF) → map P&L → assemble → verify. Capstone `h-underwrite`: the MAIN
  THREAD orchestrates only (sub-agents can't reliably spawn sub-agents), driving `census-formatter` ∥
  `hf-formatter` (parallel — census builds in a local scratch workbook) → `pnl-mapper` → `h-assembler`
  (Sonnet builds the .H; the main thread no longer assembles inline) → `gate-runner` (Haiku mechanical
  recompute) → `h-verifier` (judgment audit of that evidence). `Clear-OrphanExcel` runs ONCE, by the
  orchestrator at setup — never inside sub-agents. Read
  `Investments/Docs/MD Files/HF-PNL-PIPELINE-REFERENCE.md` before any HF/PnL/`.xlsx` build.
- Gold example — styling AND structure/formulas: `Investments/Data/Transactions/Aster Ridge (Demo)/
  Aster Ridge.H Model_v1.xlsx` — the single golden reference for both concerns in this clone.
- Mapping vocabulary: the verbatim list in `pnl-mapping/References/master-tags.md` — use those tags,
  never an invented or renamed one.
- `.RR`: design-locked; read `rr-formatting/RR-FORMATTING-DESIGN.md` + `References/` first
  (`locked-rules.md` is binding). `.claude/skills/rr-formatting/scripts/rr_build_onesite_v2.ps1` =
  proven source-side parser predating the current `.RR` contract: extend it, don't run it as-is.
  Chain (2026-07-28, mirrors `.H`): `rr-formatter` → `gate-runner` (scoped evidence) → triage ladder
  (bounce-once/ledger/stop) → `rr-verifier` (judgment audit) — see `H-PIPELINE-ORCHESTRATION.md` §
  .RR agent chain.
- Running plan: private team repo document (not included in this public clone).

## Validation gates (all of them, every build)
1. Reconcile-before-build — leaf lines sum to the operator's own subtotals across all 12 months
   (value-conserving; a non-zero diff = a dropped/double-counted/misbucketed line).
2. Zero formula errors after recalc — scan with `SpecialCells(-4123, 16)` in try/catch (throws when
   clean), not by matching `#` strings (error cells marshal as Int32 codes; the string test misses
   real errors and false-fires on headers like `#Apts`).
3. PnL map — column-A `Check` returns 0 "No" (every tag verbatim on the master list).

**GATE 0 — read back what gates 1–3 cannot see, from the saved file.** They check formulas and
arithmetic only; assert the rest explicitly: styling (if a script's purpose is formatting, read the
formatting back — memory `ps-var-name-case-collision`), census completeness (all 12 months populated
and non-zero; the Checks tie revenue/NOI and never see census), and null writes (an empty cell read
by formula returns `0`, not blank — memory `ps-null-write-silent-gate-pass`). Run any restyler
against a copy of the golden first: a correct one is a near no-op there.

*A check that cannot fail is not a check.*
