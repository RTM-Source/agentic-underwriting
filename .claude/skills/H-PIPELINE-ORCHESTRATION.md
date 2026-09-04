# .H Pipeline — Orchestration & Agentic Design

> **The `.H` tab is the deliverable; the three skills here are its inputs.** This file is the
> container-level map: how `census-formatting`, `hf-formatting`, and `pnl-mapping` connect, and how a
> main agent drives them with sub-agents to produce the final `.H`. Read this first when running or
> building the end-to-end pipeline. The capstone is now **built**: `h-underwrite/`
> (orchestrator + assembly) with its References (`h-skeleton.md`, `rollup-category-map.md`,
> `join-and-census-contract.md`) and the sub-agents in `.claude/agents/`. This file remains the
> container-level rationale; `h-underwrite/SKILL.md` is the operational playbook.
>
> Companion docs: `Investments/Docs/MD Files/HF-PNL-PIPELINE-REFERENCE.md` (the deep environment/crosswalk reference),
> `Investments/lib/HF-Build-Lib.ps1` (the COM build engine), and each sub-skill's `SKILL.md` + `References/`.

---

## The four stages and what each hands to the next

```
 census-formatting ─► Census tab ───────────────────────────────┐
 hf-formatting ─► Formatted HF tab ─► pnl-mapping ─► PnL Mapping ─┤─► .H assembly ─► .H tab
                     (code  -  label, E:P values)   (tag per line) │
```

| Stage | Skill | Produces | Consumed by |
|-------|-------|----------|-------------|
| 1 | `census-formatting` | **Census** tab: the fixed 12-month Capacity/Occupied block | `.H` **rows 1–24** (linked straight from Census) |
| 2 | `hf-formatting` | **Formatted HF** tab: house skeleton + operator codes; col C = `code  -  label`, E:P = 12 monthly values | `pnl-mapping` (reads col C) **and** `.H` (XLOOKUP pulls HF values) |
| 3 | `pnl-mapping` | **PnL Mapping** tab: one master tag per Formatted-HF line | `.H` (the tag routes each HF line to its model row) |
| 4 | **`.H` assembly** *(the `h-assembler` sub-agent, per `h-underwrite/SKILL.md` §4)* | the **`.H`** tab | the model |

**The `.H` tab layout (from the deal Ryan described):**
- **Rows 1–24** — pulled from the **Census** block (uniform; link back to the Census tab).
- **Row 25 → the first hardcoded historical line** (always begins `AL Rent` / `Rent Revenue - AL`) **down to where the
  historical ends** (line 342 in that example — the stop row is deal-specific, wherever the historical ends).
- **Rows 25–342** — a **mix of XLOOKUP(Formatted HF) + PnL-mapping tags**: the hardcoded skeleton on the left,
  values pulled by `XLOOKUP` keyed on the mapped tag. The hardcoded part needs **no** reformatting — that's
  the fixed model.

---

## Agentic topology: one orchestrator, cold sub-agents, a separate verifier

The pipeline is a fixed sequence with a heavy, isolatable workload per stage — a good fit for a main
**orchestrator** that owns the connective tissue and **sub-agents** that each do one stage. Why this shape:

- **Orchestrator (main agent)** — sequences the stages, carries the small connective state (sheet names,
  the stop row, the tag-key column, flags raised), triages flags, and runs the final GATE-0 read-back.
  **It builds nothing itself** *(2026-07-28 — assembly moved to the `h-assembler` sub-agent; the
  orchestrator previously did it inline, which billed the bulkiest templated stage at frontier rates and
  fattened this context with the transcript)*. Delegating ALL bulky per-stage work keeps its context lean
  (a single HF dump is 300+ rows).
- **Per-stage sub-agents** — one each for census / HF-format / PnL-map / **assemble (`h-assembler`,
  Sonnet)**. Census runs **in parallel** with HF-format (census builds in its own local scratch workbook;
  the assembler merges the Census tab in at step 0 — this removed the shared-workbook serialization).
  **Sub-agents start cold and do NOT inherit skills**, so the orchestrator must hand each one: its skill
  path, the specific References it needs, and the exact inputs/outputs. A sub-agent returns a compact
  result (what it wrote + its reconciliation numbers + flags), not a transcript.
- **Verification (two stages, 2026-07-28)** — first **`gate-runner`** (Haiku): a read-only mechanical
  instrument that recomputes every deterministic gate on the **saved** file and returns a raw evidence
  table, no verdicts. Then **`h-verifier`** (a different model from the builders) — a **verifier, not a
  rebuilder**: it sample-audits the evidence against the workbook, then spends its turns on *judgment*
  (what a diff means, month alignment, styling contract, silent-failure hunt); with no/compromised
  evidence it falls back to the full recompute itself. A different model catches *interpretation* errors
  (bucketing, tagging), not just arithmetic. **Cap iterations; a real mismatch STOPS and flags — never
  loop or silently auto-fix.**

### One-shot reliability: template, don't hand-roll
The `.H` structure is fixed; only **two things change per deal — the HF XLOOKUP array and the census
T12.** Build the assembly as a **template parameterized on those two inputs**, not a bespoke script each
deal. That is what makes a one-shot build dependable.

### Mechanics that apply to every stage
- Read via `Range.Value2` (1-based 2-D array), never screenshots. The verifier reopens the **saved** file
  (same-session manual-calc reads can be stale).
- No Python — all Excel work is **Excel COM in PowerShell**; reuse `Investments/lib/HF-Build-Lib.ps1` (`HF-*` writers,
  `Map-*` writers, `New-Excel`, `SetV`/`SetBlock`). Recalc happens on `$wb.Save()`.
- Clean up via `New-ExcelTracked`/`Stop-TrackedExcel` (lib); `Clear-OrphanExcel` is **orchestrator-only,
  once at setup** (2026-07-28 — census ∥ HF run concurrently, so a mid-run sweep races a sibling). The raw
  `MainWindowHandle -eq 0` sweep is hook-blocked — it cross-kills a concurrent deal's Excel, and the user
  keeps workbooks open.

---

## Validation (run as one pass, then the verifier confirms)
The three standing gates (also in `CLAUDE.md`):
1. **Reconcile-before-build** — HF leaf lines sum to the operator's own subtotals across all 12 months.
2. **Zero formula errors** after recalc.
3. **PnL map** — column-A `Check` returns **0 "No"** (every tag verbatim on the master list).

Plus the assembly's own tie-out: the `.H` totals reproduce the Formatted-HF totals, and rows 1–24 match
the Census tab. Surface every flag (HF judgment calls + PnL `Questions & Comments`) up to the orchestrator
so the underwriter sees the ~10% to review.

---

## Capstone skill — BUILT as `h-underwrite/` (this section is the original design note)

```yaml
# h-underwrite/SKILL.md   (built — was planned as dot-h-assembly)
name: dot-h-assembly
description: >
  Assemble the final .H (Historical Financials) tab from the three upstream outputs — the Census block
  (rows 1–24), the Formatted HF, and the PnL Mapping — wiring rows 1–24 to Census and the historical rows
  to XLOOKUP(Formatted HF) keyed on the mapped tag, against the fixed hardcoded model skeleton. Use when
  the user wants to "build the .H," "assemble the historical tab," "wire up the .H," or run the full
  HF→PnL→Census→.H pipeline end to end. Orchestrates the per-stage skills (via sub-agents) and reconciles.
```

The recognizer/sequencer can live inside this skill (it invokes the others) rather than as its own skill.
Promote a stage to richer tooling only if it gets heavy. Keep the **verifier** separate from the builder.

## .RR agent chain (builder/verifier pair added 2026-07-14; swarm-economics split 2026-07-28)
The rent roll's chain now mirrors the `.H` redesign: **`rr-formatter` → `gate-runner` (scoped .RR
evidence) → triage ladder → `rr-verifier` (judgment audit)**.
- **`agents/rr-formatter.md`** (Sonnet) — builds the full .RR from a raw export via the
  `rr-formatting` skill; builds on a copy; runs all four gates itself as a first pass; returns
  judgment calls as PROPOSED. Accepts one bounce-back message naming Class-A defects (see its
  "Bounce-back protocol" section) before escalating.
- **`agents/gate-runner.md`** (Haiku — the SAME shared instrument the `.H` chain uses) — reopens the
  saved .RR, read-only, and recomputes the deterministic gates enumerated in
  `rr-formatting/References/validation-and-build.md` § "Mechanical gate list": reconcile sums to the
  cent, residual *components* (not the ruling), the `SpecialCells` error scan, Check-block read-backs,
  the five-header/label/row-height/`blockStart`/gridlines/zero-dash/count read-backs. Numbers only, no
  verdicts.
- **Triage ladder (run by the orchestrator — main thread)**: reads gate-runner's evidence and sorts it —
  **Class A** (cleanly mechanical defects: a formula error, a missing header, a wrong row height, a
  non-tying dollar sum) bounces **ONCE** back to `rr-formatter`; **Class B** (judgment-flavored: tier
  reasonableness, a residual needing a decomposition ruling, an anomaly needing a look) accumulates in a
  **defect ledger** handed to `rr-verifier`; **Class C** (broken gates — workbook won't open, a
  structural read fails, an unreconcilable dollar gap) **STOPS** the pipeline.
- **`agents/rr-verifier.md`** (Opus — verifier ≠ builder) — reopens the saved workbook fresh,
  sample-audits the gate-runner evidence (falling back to the full legacy recompute if the evidence is
  absent or a sample disagrees), rules on every Class-B ledger item, and spends its judgment on the
  reconcile-residual decomposition (rule 12), tier-assignment audit, and the silent-failure hunt
  (stacked turnover contracts, vacant carry-over, phantom S-line beds, wrong MC `#Apts`).

The .RR chain is standalone (not a `.H` stage). It still differs from the `.H` design in one way
(intentional): it stays **single-builder serial** (one COM automation, no parallel census-style
stage to race). **`Clear-OrphanExcel` ownership depends on context** — when the .RR runs standalone,
`rr-formatter` calls it at its own startup (safe: no sibling stage runs concurrently within an .RR
build). When the .RR runs **alongside other automation in one session** (e.g. a `.RR` + `.H` pass on
the same deal, or a multi-deal swarm), the **orchestrator owns cleanup** — called once at setup, same
rule as the `.H` chain — and `rr-formatter`/`gate-runner` must NOT call it mid-run.
