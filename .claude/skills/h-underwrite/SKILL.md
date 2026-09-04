---
name: h-underwrite
description: >
  Run the full .H underwrite pipeline end to end and assemble the final .H
  (Historical Financials) tab. This is the ORCHESTRATOR + capstone: it drives the three
  upstream skills via sub-agents (census-formatting → Census tab, hf-formatting →
  Formatted HF, pnl-mapping → PnL Mapping), then builds the .H tab that links back to all
  three — occupancy block linked to Census, a detail passthrough that mirrors the Formatted
  HF via XLOOKUP, a SUMIFS roll-up into the firm's standard categories, and two reconciliation
  Checks. Use when the user wants to "run the .H," "underwrite this deal," "build the
  historical financials end to end," "assemble the .H," or hands over a raw operator HF +
  census to turn into a finished model tab. For a single stage only, use that stage's skill
  directly (census-formatting / hf-formatting / pnl-mapping); this skill is for the whole chain
  or the assembly step.
---

# `.H` Underwrite — orchestrator + assembly

## What this skill is
The capstone of the `.H` pipeline. It owns the **connective tissue** of an underwrite: it sequences
the stage skills, carries the small shared state between them, and owns the reconciliation gate. ALL
heavy per-stage work — including the final `.H` assembly (the `h-assembler` sub-agent, 2026-07-28) —
is delegated to **sub-agents** so the orchestrator's own context stays lean.

Read first (cheap, prevents re-deriving): `../H-PIPELINE-ORCHESTRATION.md` (the container map),
`Investments/Docs/MD Files/HF-PNL-PIPELINE-REFERENCE.md` (environment + pitfalls), and this skill's three
References — `h-skeleton.md`, `rollup-category-map.md`, `join-and-census-contract.md`.

## Pipeline (what produces what)
```
                census-formatting ─► Census tab ─────────────────────────────┐
raw operator ─► hf-formatting ─► <CN> Historical Financials ─► pnl-mapping ─► <CN> PnL Mapping ─┤─► .H assembly ─► <CN>.H (Review)
   files          (Formatted HF: col C = code  -  label, E:P values)         (tag per line)     │   (this skill)
```
Five tabs in the finished workbook (naming convention, left→right): `<CN>.H (Review)`, `Census`,
`<CN> PnL Mapping`, `<CN> Historical Financials`, `<CN> Unformatted HF`. (`<CN>` = community name.)
`Census` = the raw census sheet pasted in (values + number formats) with the block appended beneath —
**no separate `Occupancy Source` tab** as of 2026-07-22; an extra source tab appears only when the
window needs a second file.

---

## Agentic topology — one orchestrator (this thread), cold sub-agents, a separate verifier

**The orchestrator is THIS conversation running this skill — not a nested agent.** In Claude Code a
sub-agent cannot reliably spawn its own sub-agents, so the main thread does the delegating. That also
matches the design in `H-PIPELINE-ORCHESTRATION.md` ("orchestrator = main agent"). You (the orchestrator)
spawn each stage as a sub-agent via the **Agent tool**, collect a compact result, and move on.

Use these agent definitions (in `.claude/agents/`):

| Stage | Sub-agent (`subagent_type`) | Skill it loads | Returns |
|-------|------------------------------|----------------|---------|
| 1 ∥ 2 | `census-formatter` (runs in PARALLEL with stage 2) | `census-formatting` | Census SCRATCH workbook built (local `%TEMP%`); its path; window used; flags; the Census cell map (capacity/occupied/Avg-Occ-% IL-AL-MC cells + first month col) |
| 2 ∥ 1 | `hf-formatter` | `hf-formatting` | Formatted HF built; `hfLast`, `hfRevRow`, `hfNoiRow`; reconcile result; flags |
| 3 | `pnl-mapper` | `pnl-mapping` | PnL Mapping built off HF col C; `mapLast`; col-A Check = 0 "No"; flags |
| 4 | `h-assembler` | this skill's §4 + References | Census merged into the model workbook + the `.H` tab built + saved; `detailLast`; gate readings; flags |
| QA-a | `gate-runner` (Haiku — mechanical instrument) | (none — deterministic recompute) | raw evidence table: tie-out vectors, error scan, join samples, census read-backs — numbers, no verdicts |
| QA-b | `h-verifier` (judgment audit) | (none — audits QA-a evidence + targeted reads) | sample-audit of the evidence, tie-out/join/census/styling JUDGMENT calls, silent-failure hunt |

**Why sub-agents.** Each stage reads a 200–300-row sheet; doing them inline floods this context. Sub-agents
**start cold and do NOT inherit skills or this conversation** — so when you spawn one you MUST hand it: the
absolute workbook path, the `<CN>`, its skill path, the specific References to read, the exact input
tab/output tab, and the compact result you want back. The agent files already encode the standing rules;
your prompt supplies the per-deal specifics.

**Sequencing — census runs IN PARALLEL with HF-format; the rest stays serial** *(2026-07-28 — supersedes
the old all-serial rule; adopted from the agent-swarm economics review)*. The old serial constraint
existed only because census and HF shared one physical workbook; census now builds in its **own scratch
workbook** on a local `%TEMP%` path (see `agents/census-formatter.md`) and the `h-assembler` merges the
Census tab into the model workbook at assembly step 0. So: **spawn `census-formatter` and `hf-formatter`
in the SAME message** (two Agent calls, they run concurrently — different workbooks, no shared file),
collect both results, then run **PnL → assembly** serially (PnL keys off HF col C; assembly is gated on
all three). `gate-runner` then `h-verifier` run after assembly + save. **Because two stages now run
concurrently inside one deal, `Clear-OrphanExcel` is called ONCE by the orchestrator at setup — the
sub-agents must NOT call it mid-run** (a sweep while a sibling is between spawn and lock registration is
a race; their agent files say so too). *(Cross-run concurrency is unchanged: independent deals in
separate terminals — the 6-terminal 3-pack — still run fully in parallel; the raw windowless sweep stays
hook-blocked; see memory `never-blanket-kill-excel`.)*

**Model choice when spawning (see `CLAUDE.md` → *Model switching*).** Builders — **including
`h-assembler`** *(2026-07-28: assembly is the most fully-templated stage; it no longer runs inline on
this thread's frontier model)* — default to **Sonnet** (set in their agent files): bulky, well-specified,
and backstopped by the verification chain. You (Opus, this thread) keep ONLY the connective tissue:
sequencing, judgment overrides, flag triage, the census-completeness gate, and the final GATE-0
read-back. **Keep your own context lean — never ingest a stage transcript or a sheet dump; compact
results only.** Two overrides to make consciously via the Agent tool's `model` param:
- **Messy / novel operator HF** → spawn `hf-formatter` with `model:"opus"`. The reconcile-before-build
  regroup (misbucketed lines, sign conventions, GLTL/vacancy, Management-Fee placement) is the hardest
  judgment in the chain and the first place Sonnet slips.
- **If you bump any builder to Opus, spawn `h-verifier` as Sonnet** for that run. The verifier's value is
  a *different model than the builders*, not "always Opus" — keep the two reasoning paths distinct.

Don't reach for Haiku on any JUDGMENT stage; every builder carries flag-don't-guess calls. The one Haiku
seat is **`gate-runner`** — a pure deterministic recompute instrument with no judgment by construction
(anything ambiguous is reported raw, and `h-verifier` sample-audits its evidence, so a slip cannot
silently pass).

---

## Orchestrator procedure

### 0. Set up
- Resolve the workbook path (OneDrive may have moved it — `find` if `Open` fails) and the community name
  `<CN>`. Confirm the raw HF tab is present (the census stage pastes the raw census into its own scratch
  workbook itself — hand it the deal folder + the scratch path, not a pre-imported tab and not the model
  workbook). **Version up NOW — never overwrite:** YOU (the orchestrator) make the new `_v#` copy
  (`<CN>.H Model_v2.xlsx` from `_v1`) here in setup, before any stage spawns; every stage (HF, PnL,
  assembly) builds on that copy. The h-assembler never versions anything.
- **Resolve the 12-month window YOURSELF, here** *(2026-07-28 — census and HF now spawn together, so the
  window cannot come from hf-formatter's return)*: read the raw HF tab's month-header row directly
  (one cheap `Range.Value2` read; convert serials) and derive the T12 window (oldest → newest month-end).
  Hand this window to BOTH `census-formatter` and `hf-formatter` in their spawn prompts. When
  `hf-formatter` returns, its reported window is the CROSS-CHECK — if it differs from what you derived,
  STOP: the census stage was built over the wrong window.
- Pre-run cleanup: dot-source `Investments/lib/HF-Build-Lib.ps1` and call `Clear-OrphanExcel` **ONCE,
  here, before any stage spawns** — the sub-agents do NOT call it themselves (census ∥ HF run
  concurrently; a mid-run sweep races a sibling between spawn and lock registration). Excel is created/
  released per stage via `New-ExcelTracked`/`Stop-TrackedExcel`. Never the raw `MainWindowHandle -eq 0`
  sweep (hook-blocked — it cross-kills a concurrent deal's Excel; the user also keeps workbooks open).
- Pick the **census scratch path**: a local `%TEMP%` file (e.g. `$env:TEMP\<CN>-census-scratch.xlsx`) —
  local so `Save`/`SaveAs` can't hit the OneDrive hang. Hand it to `census-formatter` in its prompt.

### 1–3. Run the three stages (sub-agents)
Spawn **`census-formatter` and `hf-formatter` together in ONE message** (two Agent calls — they run
concurrently: census in its scratch workbook, HF in the model workbook; see Sequencing above). When BOTH
have returned, spawn `pnl-mapper` (it keys off HF col C). Collect each agent's compact result. **Surface
every flag** each raises up to the user — those are the ~10% to review.
If a stage's own validation fails (HF doesn't reconcile, mapping has a "No"), **stop and report** — do not
proceed to assembly on a broken input.

**Checkout-as-you-go triage ladder (Ryan, 2026-07-28 — Phase 2).** After each stage returns, spawn
`gate-runner` in **scoped stage-check mode** on just that stage's output (census scratch / Formatted HF /
PnL Mapping) — a cheap independent measure taken while the builder's context is still HOT. Triage every
finding into exactly one class:
- **Class A — mechanical, objectively spec'd** (formatting misses, wrong number format, gridlines on, a
  formula anchored one row off): **bounce it back to the SAME builder agent** (SendMessage to the agent
  that built it — its context is hot, the fix is a couple of turns), then have gate-runner re-measure
  ONCE. **Hard cap: one bounce.** Still failing → it graduates to Class B.
- **Class B — judgment** (bucketing calls, tag choices, alignment resolutions, an *explained* residual):
  **do NOT fix.** Append to the **defect ledger** — a compact running list you carry (a few lines per
  stage: finding, stage, builder's stated rationale) and hand to `h-verifier` in its spawn prompt.
- **Class C — a broken gate** (a reconcile that doesn't tie, an unexplained residual, a short census):
  **STOP the pipeline and report.** Never bounced, never auto-fixed, never "resolved" by adjusting
  financial logic until a check reads zero.
The ladder catches defects BEFORE downstream stages consume them. It never overrides the standing rule:
a real mismatch stops for a human.

Carry forward this shared state from the results: `hfLast`, `hfRevRow`, `hfNoiRow`, `mapLast`, the Census
cell map, and the 12-month window (confirm HF window == Census window; if not, STOP and flag).

**Census completeness is YOUR gate — the Checks cannot see it.** Both reconciliation Checks tie revenue
and NOI, neither of which touches census, so a census with missing months reconciles perfectly while
occupancy is wrong. Before assembly, confirm the `census-formatter` returned **all 12 months populated
and non-zero**; a short census is a STOP, not a flag (`census-formatting/SKILL.md` §"A short census is a
HARD STOP"). Tell the census stage the **HF's window** explicitly — it must not pick "the most recent 12
months" of whatever file it happens to open. Also point it at the **deal folder**, not one file: Deal B's
census stage read a single-month detail report while a full multi-year history sat unread beside it, and
shipped a 1-of-12-month block through every gate.

### 4. Assemble the `.H` tab (the `h-assembler` sub-agent — NOT this thread)
*(2026-07-28 — assembly is delegated; it was previously done inline by the orchestrator, which billed the
bulkiest, most-templated stage at frontier rates and flooded this context with the assembly transcript.)*
Spawn **`h-assembler`** (Sonnet) with: the versioned-up workbook path, `<CN>` + tab-prefix token,
`hfLast`/`hfRevRow`/`hfNoiRow`, `mapLast`, the census scratch path + cell map + window + which
month-mapping resolution census used. Its step 0 merges the Census scratch tab(s) into the model
workbook (`Worksheet.Copy` is safe there — the scratch is our clean, self-contained build); it then
executes the procedure below and returns compact gate readings. **You do not open the workbook during
assembly; you never ingest its transcript.**

The build itself (the assembler's script — it reads this section) is `<CN>.H (Review)` per
`h-skeleton.md` + `rollup-category-map.md` + `join-and-census-contract.md`. It is a **template
parameterized on five things** (`<CN>`, `hfLast`, `mapLast`, `hfRevRow`/`hfNoiRow`, Census map + first
month) — don't hand-roll a bespoke layout. Order of operations:

1. **Block B (detail) first.** Read HF col C rows 8→`hfLast` with `Range.Value2` and write them
   **byte-identical** into `.H` D starting at row 28 (`.H` row = HF row + 20). Then write the B tag-XLOOKUP
   and the F:Q value-XLOOKUP formulas per row. `detailLast` = last detail row written.
   **Skip the HF's `Margin` row** — it is a ratio with no tag (see `h-skeleton.md` §Block B); dropping it
   shifts every row below up by 1, which is fine since they are all derived from `detailLast`.
2. **Block A (occupancy).** F5 = oldest month date; G5:Q5 `=EOMONTH`. Wire **rows 9–11 (Capacity)** and
   **rows 19–21 (Occ %)** to the resolved `Census!` cells — **Occ % is the pulled input** (green);
   **rows 14–16 (Occupied) are DERIVED in-sheet** as `=F19*F9` (black). Rows 12/17 SUM; row 22
   `=+IFERROR(F17/F12,"NA")`; 24 PRD. *(Ryan, 2026-07-16 — inverts the old pull-Occupied/derive-% shape;
   pull what the operator measured, derive the rest.)* **Confirm which month-mapping resolution the
   census stage used** — date-match (default) or a ruled calendar-year proxy (`census-formatting/
   SKILL.md` §"Aligning the source columns to the window") — and that it's documented on the Notes tab
   if it's the proxy. An *unruled* positional fill is invisible to every gate: no error fires, both
   Checks stay 0.
3. **Block C (roll-up).** Write the fixed category block (221–296) with `detailLast` in the SUMIFS ranges;
   the two Check rows (234, 300) pointing at `hfRevRow`/`hfNoiRow`. **First roll-up row = `detailLast + 3`**
   (two blank rows after the detail) — derive it; a hardcoded gap shipped an extra blank row on Deal A's first build.
4. **Block D (metrics)** 302–336, fixed template.
5. **Block E (totals columns S:U) + styling — NOT optional; the tab is unfinished without them.**
   Full spec (formulas + every color/border/format, verified against the golden): `h-skeleton.md` §"Totals
   columns S/T/U" and §"Styling contract". Headlines: S/T/U = T12 / T6×2 / T3×4 windowed totals on every
   value row (`SUMIFS` over `$E{r}:$Q{r}` bounded by the row-4/5 window dates ×row-2 multiplier;
   `AVERAGEIFS` for occupancy stock rows); the S-column Checks tie to the HF's annual-total **col Q**;
   R col = metric-tag validation XLOOKUPs with plain-black `Ok` literals (an earlier convention used red —
   the current golden is black). **Font-color semantics (per cell):
   green `32768` = pulled from another tab; blue `16711680` = hardcoded input (F5 date, 1.00x cell);
   black = in-sheet formula** — e.g. row 5 is blue at F5 only, black on the G5:Q5 EOMONTH walk. Detail
   col B tag cells = **green `32768`, Calibri 9, no fill** (cross-tab pull into the Mapping tab —
   confirmed in the golden; the older blue-on-cream convention is obsolete). Detail **total** rows (`^TOTAL `/`NET OPERATING INCOME`)
   = bold + thin top/bottom border on `D:Q` and `S:U`. Check rows = green fill `4697456`, fmt `#,##0.00`.
6. Gridlines **off**. Number formats per the model (accounting for $, `#,##0.00x` for residents, % for
   margins/% rows). Recalc happens on `$wb.Save()` — set `xlManual` during build, `CalculateFull()` once
   before save (see `Investments/lib/HF-Build-Lib.ps1` caller contract; reuse its COM helpers).

Engine: **Excel COM in PowerShell only — no Python.** Honor the §5 COM pitfalls in the pipeline reference
(`${r}` brace formula colons, `InvokeMember` scalar writes, one block `Value2=` per array, COM retry).

### 5. Verify (two-stage: mechanical instrument, then judgment audit)
*(2026-07-28 — the single Opus grind-everything verify is split; the arithmetic runs on a cheap
instrument, the expensive model spends its turns on judgment.)* The build-side gates already ran inside
`h-assembler` (its step 4) — read its gate readings before proceeding; a failing reading is a STOP.
Then, on the **saved** file:
1. **`gate-runner`** (Haiku, read-only) recomputes every deterministic gate from scratch — tie-out
   vectors both ways, `SpecialCells(-4123,16)` error scan (never the `#`-string scan), Check-row
   read-backs, census 12-month coverage + header-date pairs, ~10-row join byte-compare, tag-coverage
   list — and returns a raw **evidence table**. Numbers only; it makes no calls.
2. **`h-verifier`** (different model from the builders) gets the evidence table AND the **defect
   ledger** (the Class-B judgment items accumulated by the triage ladder above) IN ITS SPAWN PROMPT,
   sample-audits the evidence against the workbook (never trusts the instrument blindly), then spends
   its effort on interpretation: ruling on each ledger item, what a diff pattern means, month-alignment
   resolution, styling contract, silent-failure hunt. A compromised evidence sample → it falls back to
   the full legacy recompute.
3. **Cap iterations — a real mismatch STOPS and flags; never loop or silently auto-fix.** Your own
   GATE-0 read-back (CLAUDE.md) still applies before archiving/reporting.

### 6. Report (terse, per CLAUDE.md)
Lead with the result: tabs built, both Checks (0/non-0), formula errors (count), and the consolidated
flag list from all stages. Point the user at the new `_v#` file.

---

## Hard rules (every run)
- **No Python** — Excel COM in PowerShell; reuse `Investments/lib/HF-Build-Lib.ps1`. Recalc on `.Save()`.
- **Version up, never overwrite** the input workbook.
- **Gridlines off** on every sheet touched.
- **Clean up with `Clear-OrphanExcel` + tracked create/release**, never the windowless sweep (hook-blocked).
- **Sub-agents are cold** — preload skill path + References + per-deal inputs in every spawn prompt.
- **Flag, don't force** — surface judgment calls; never invent a tag or fabricate census.
- **Reconcile is the gate** — both Checks at 0, or stop.

## Related
- Container map + rationale: `../H-PIPELINE-ORCHESTRATION.md`. Stage skills: `../{census-formatting,
  hf-formatting,pnl-mapping}/SKILL.md`. Env + pitfalls: `Investments/Docs/MD Files/HF-PNL-PIPELINE-REFERENCE.md`.
  COM engine: `Investments/lib/HF-Build-Lib.ps1`. Gold example (styling AND structure/formulas):
  `Investments/Data/Transactions/Aster Ridge (Demo)/Aster Ridge.H Model_v1.xlsx`.
- Agents: `.claude/agents/{census-formatter,hf-formatter,pnl-mapper,h-assembler,gate-runner,h-verifier}.md`.
