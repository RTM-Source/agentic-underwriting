# Rent Roll (.RR) Formatting — Design & Build Context

> **STATUS: golden model established (v1 — 2026-07-13). Skill finalized; engine build pending.**
> `SKILL.md` now exists (the invokable orchestrator) and the target format is **locked to the golden
> workbook** `Investments/Data/Transactions/Aster Ridge (Demo)/Rent Roll/Aster Ridge-RR_v1.xlsx` (tab
> `Aster Ridge.RR`, fully built out — in the private team repo the golden is a real-deal build; this
> public clone substitutes the fabricated demo deal). The output contract is the **17-col AW:BM unit block
> + the analysis table** — see `References/output-columns.md` and `References/analysis-table.md`. The
> firm rulings (incl. everything v1 settled) live in `References/locked-rules.md`. What remains is
> extending the OneSite parser in `Investments/scripts/` to emit the v1 shape (it emits a 16-col AW:BL block and a
> summary grouped on the wrong key) — see "What the next build must change."

---

## What this skill does (the job)

Operators send a **rent roll** (a per-resident occupancy + rent snapshot). The underwriting team needs a uniform
set of model-ready columns built **to the right of** the operator's hardcoded export — linking back
with formulas, hardcoding only where a formula can't express it. The deliverable is the right-side
per-unit block (columns **AW:BM**), the derived **analysis table**, and a reconciliation that ties to
the operator's own totals — the full `.RR` tab (golden model v1).

**The hard part is the source, not the target.** The right-side columns are a fixed, uniform contract.
The left-side export is wildly non-uniform across operators (OneSite, Yardi, PCC, …) and even across
exports from the same operator. So the skill is built around that asymmetry.

---

## Planning Mode (intake gate — run BEFORE the build)

The expensive mistakes on a rent roll are **interpretation**, not arithmetic: which unit codes roll up
to which care level, what SqFt bands define Studio/1BR/2BR, and how Memory Care shared rooms are
structured. The rent roll often does **not** make these unambiguous on its own (flat placeholder SqFt,
a "P/SP" flag that says "P" for everyone, codes like `MC-0JJ` whose meaning only the operator knows).
So the workflow opens with a short **planning pass**, and where the grouping is not clear from the file,
it **asks the underwriter before building** rather than guessing.

**Do the planning pass first:**
1. **Inventory the care levels present** — IL / AL / MC (and any sub-levels). Confirm every unit line
   maps to one; flag stray/junk rows.
2. **Propose the analysis buckets** — the blue summary groups each care level by **SqFt band** into
   Studio / 1BR / 2BR (deluxe/"-D"/"SD" variants fold into their bedroom type). Derive bands from the
   data's natural breaks, but **anchor the labels to the OM's stated SqFt ranges** when available.
3. **Resolve MC room structure** — classify each MC code as:
   - **Private** — 1 resident / room. `#Apts = #Units`; SqFt = the room.
   - **Semi-Private (shared)** — **1 door, 2 beds** (one room, two residents). `#Apts = #Units / 2`.
   - **Jack & Jill** — **2 doors, 2 beds** (two bedrooms sharing a bath). `#Apts = #Units / 2`.

   Both shared types collapse `#Apts = #Units / 2` and restate **SqFt per individual** (a flat door
   SqFt like 500 → ~250/bed). The door/bed config only changes the **name**, not the math — so name it
   right (don't call a 1-door semi-private a "Jack & Jill").

**Ask the underwriter when any of these is unclear from the rent roll** (don't invent a default):
care-level mapping of ambiguous codes; OM SqFt ranges per type; and the MC bed/door structure per code.
A quick "explain the IL / AL / MC breakdown and the shared-MC layout" up front prevents a rebuild.
Feed the answers into the build and **record them in the Notes tab** (see the analysis-table notes).
Per-property MC SqFt and the shared-room layout vary (one property's MC-0S wing may be 1 door/2 beds
Semi-Private, another's MC-0JJ wing 2 doors/2 beds Jack & Jill), so confirm per deal.

---

## Architecture: split "read the source" from "emit the standard output"

A single "overlay one format" skill is fragile because the input varies so much. Decompose into three
stages, with the variance isolated in the middle:

```
  recognize ──► read (per-source) ──► emit + reconcile (central, uniform)
   (geometry)     (sub-agent work)        (identical every deal)
```

1. **Recognize** — identify the export family (OneSite / Yardi / PCC / …) and map the table geometry:
   header row, the left data columns, where the right target columns start, the detail region, and the
   operator's own section/property total rows. Output: a small "source profile."
2. **Read (per-source)** — parse the messy source into **canonical unit records** (see
   `References/canonical-record.md`). This is where format variance lives; one reader per export family.
   It is heavy and self-contained → a good **sub-agent** job (it can chew through a 650-row dump and hand
   back compact records, keeping the orchestrator's context lean). Sub-agents start cold, so the reader
   prompt must point at this folder.
3. **Emit + reconcile (central)** — write the AW:BM unit block as formulas linking back to the hardcoded
   left, derive the analysis table (group on Group Unit Code), set number formats, gridlines off, and run
   the **reconciliation gate** (see `References/analysis-table.md` + `References/validation-and-build.md`).
   This stage is identical for every deal and must never be duplicated per-source.

Two things stay central no matter how the readers multiply:
- **The reconciliation gate** — source-agnostic; it caught every double-count in the first build.
- **The canonical record schema** — the single contract between every reader and the one emitter.

This skill is expected to **grow with use**: each new operator adds a reader; each new error/fix adds a
flag pattern. Best-effort + reconcile + flag beats forcing a rigid format the files won't honor.

### Model assignment (bake in at build time — mirrors the `.H` pipeline)
Follow the same Opus-reasons / Sonnet-grinds split the `.H` sub-agents use (see `CLAUDE.md` → *Model
switching* and `.claude/agents/`):
- **Reader sub-agents** (`rr-read-onesite`, future `rr-read-yardi`/`rr-read-pcc`) → **`model: sonnet`**.
  Each chews a 600–650-row dump into canonical records — bulky, self-contained, well-specified. Ideal
  Sonnet job; the token volume lives here, so this is the main cost/latency win.
- **Orchestrator + emit/reconcile** → **Opus** (main thread). The reconciliation gate is the
  correctness backstop, and the interpretation calls (the **BL join re-stamp** so Market/In-Place/Care
  don't go blank — memory `rr-analysis-join-key-bl`; the phantom/2nd-occupant capacity=0 count fix —
  `rr-om-tieout-unit-count`; de-proration formula logic — `rr-om-tieout-unit-count`) are reasoning, not
  grunt work. Keep them on the reasoning model.
- **If you add a QA verifier** for the RR (recompute the reconcile + count tie-out independently), make
  it a **different model than the readers** (Opus while readers are Sonnet) — same invariant as
  `h-verifier`: a different reasoning path is what catches interpretation errors.
- **Recognizer** (export-family classification) is the one spot cheap enough for **Haiku** if it ever
  becomes its own step — it's pure pattern-matching on banner/geometry telltales.

---

## Planned skills (template frontmatter for the next session to create)

Keep readers as the swappable layer. Suggested split (adjust as it grows):

```yaml
# rr-formatting/SKILL.md   (orchestrator the user invokes)
name: rr-formatting
description: >
  Build the firm's standard right-side rent-roll columns (Order, Unit, Unit Type, SqFt, Care Type,
  Capacity, Occupancy, Move-in, Market/In-Place Rent, In-Place Care, Unit Code, 2nd Resident, Basic
  Unit Code) from an operator rent roll, linking back with formulas and reconciling to the operator's
  own totals. Use whenever the user wants a rent roll "formatted," "the .RR columns built," "the right
  side filled in," or a unit/bed roster + in-place-rent build from a OneSite/Yardi/PCC export — even if
  they don't name the format. Recognizes the source, dispatches the matching reader, then emits + reconciles.
```

```yaml
# rr-formatting/rr-read-onesite/SKILL.md   (one reader per export family; preload for sub-agents)
name: rr-read-onesite
description: >
  Parse a OneSite Senior Living rent roll export into canonical per-unit (per-bed) records. Use when
  the recognizer tags the source as OneSite (telltales: "OneSite Senior Living" banner, "Floor Plan -
  SQ FT" header, 'NNN-NNN TOTAL' subtotal rows, transaction-code table at the bottom).
```

The **emit + reconcile** stage and the **recognizer** can start as references/scripts the orchestrator
uses inline rather than separate skills; promote them to their own skills only if they get heavy.
Future readers (`rr-read-yardi`, `rr-read-pcc`, …) follow the same shape.

---

## Reference map (read these as needed — progressive disclosure)

| File | What it holds |
|------|---------------|
| `References/locked-rules.md` | The firm rulings (granularity, capacity, occupancy, no-couples, 2nd-resident, care-type, finance). **Read first.** |
| `References/canonical-record.md` | The per-unit record schema — the reader→emitter contract. |
| `References/output-columns.md` | The **AW:BM** per-unit column map + exactly how each cell is derived (formula vs hardcode). |
| `References/analysis-table.md` | The summary derived from the unit block: tiering, unit-mix, Second Residents, Rent Adj, Check, move-in windows, reconcile. |
| `References/onesite-reader.md` | OneSite source structure + the parsing recipe (anchors, leaf lines, blocks, detail-end). |
| `References/validation-and-build.md` | The reconciliation gate + COM/`.xls` build mechanics and pitfalls. |
| `Investments/scripts/rr_build_onesite_v2.ps1` | **The proven OneSite parser — the starting point, not the engine.** Source-side (anchors, blocks, leaf-sum finance, reconcile-to-the-cent) is correct and safe (builds on a COPY, corrupted-source abort guard, `$w`-prefixed column constants); don't re-hand-roll it. Output side predates v1 — see below. **Extend; don't run as-is.** |

---

## What the next build must change (bring the engine to the v1 golden shape)

`Investments/scripts/rr_build_onesite_v2.ps1` already emits the **16-col AW:BL** block and a per-unit-type summary.
Its source-side parsing is sound; the gap to v1 is on the output side.

**Already correct in v2 — keep:** the AW:BL column set (Order, Unit, Unit Type, P/SP, Care Type, SqFt,
Resident Type, Capacity, Occupancy, Move-in, Market, In-Place Rent, In-Place Care, 2nd Resident, Unit
Code, Basic Unit Code) · **Capacity** `=IF(BB>0,1,0)` and **Occupancy** `=IF(BF="",0,1)` (rules 9–10) ·
**2nd Resident as the $ fee**, not a Y flag · the composite **Unit Code (BK)** and **Basic Unit Code (BL)** ·
the inactive-down occupancy fix · build-on-copy + the corrupted-source abort guard.

**Unit block — extend AW:BL → the 17-col AW:BM** (`References/output-columns.md`): add the **Group Unit
Code (BM** = `Care|Type|Tier`**)** — the analysis join key, and the one genuinely missing column. Convert
**SqFt (BB)** and **Move-in (BF)** from formulas to **hardcoded inputs** (v2 parses both off the source;
v1 hand-enters them — raw SqFt is a flat placeholder and raw move-in is dirty). They stay **black/
automatic**, *not* color-coded — see rule 13. Zero the shared "S" row's fee and fold it onto the primary.
Add the three special-case formula patterns (P/S pair, semi-private companion, mid-month de-proration).

**Replace the summary — emit the v1 analysis table** (`References/analysis-table.md`): v2's blue
per-unit-type block is the *wrong grouping* (it groups on Unit Code, i.e. per SqFt) — v1 groups on **Group
Unit Code** (care × A/B/C tier). Rip it out and emit: the tiering helper, the main unit-mix summary,
Second Residents, Rent Adjustments, the Check block, and the two trailing move-in windows. Plus the
**days-in-period** input that drives de-proration.

**Reconcile** per the refined gate (rule 12): Market & 2nd-Occ tie to the cent; In-Place Rent delta = the
de-proration gross-ups (traceable); Care within rounding.

**Parameterize.** v2's paths are hardcoded to the test files (`Investments/Data/Rent Roll Tests`, the two
OneSite test communities, a literal `TOTAL <property name>` property-total match). Take the source
workbook as an argument and detect the property-total row generically.

All of these are now settled rulings — see `References/locked-rules.md` rules 8–14. (The ID-column meaning
that used to be open is resolved there: Unit / Unit Code / Basic Unit Code / **Group Unit Code**.)

---

## Open questions → rules (from the O&O pass; v1 settled most)

Mirrored as a checklist in `Investments/Docs/Automation Running Plan_v2.docx` §4. **When an answer lands:
encode it in `References/locked-rules.md` (or `output-columns.md`) and strike it here.** Until answered,
the skill must ASK, not guess — same rule as the Planning Mode gate.

**✅ Resolved by the golden model v1 (→ `locked-rules.md`):** #1 three/four ID columns (rule 8 — Group
Unit Code is the join key; the old `BL=BN` framing was wrong) · #2 licensed vs visible beds (rule 14 —
`.RR` capacity = beds visible in the RR; the licensed overlay is a downstream model-layer concern) · #3
scope (rule 14 — the skill emits the full analysis table too) · #4 de-proration (rule 11 — gross
partial-month move-ins to full monthly in the emitter) · #11 second-resident summary (the Second
Residents block, `analysis-table.md` §2).

**Still open (ASK, don't guess):**

5. **Care vs other charges** — which charge columns sum into In-Place Care (level-of-care
   upcharges, med management, ancillaries)? Where do respite/short-stay charges land?
6. **Credits/concessions** — net against In-Place Rent, ignore, or carry separately so the `.AP`
   concession line can trace to it?
7. **Blank/stale Market Rent** — substitution hierarchy: BL re-stamp from unit type → Comps tab →
   flag-and-leave?
8. **Non-revenue units** (employee/model/office) — capacity 1 / occupancy 0, or excluded with a
   flag?
9. **Respite residents** — occupied beds with rent, or excluded and flagged Other Revenue
   (consistent with the PnL mapping ruling)?
10. **Broken operator subtotals** — when the operator's own subtotal ≠ their own leaf lines, does
    the reconcile gate abort or reconcile-to-leaves and flag the delta?
11. ~~2nd-resident summary~~ ✅ **resolved** — the Second Residents block (`analysis-table.md` §2):
    count, utilization, fee and fee/resident per care level. Surfaces the second-resident spread.
12. **Staleness stamp** — record the RR as-of date and flag it against the census/transition date.
13. **Downstream refresh** — after the versioned `—RR_v#` file is built, does the skill also
    refresh the deal workbook's rent-roll support tab, or is that always a manual carry?
