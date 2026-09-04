---
name: rr-formatting
description: >
  Transform a raw operator rent roll (OneSite / Yardi / PCC export) into the firm's finished .RR tab.
  Given only the operator's hardcoded export, build the standard per-unit block (Order, Unit, Unit Type,
  SqFt, Care Type, Capacity, Occupancy, Move-in, Market / In-Place Rent, In-Place Care, 2nd-Resident fee,
  Unit / Basic / Group Unit Codes) linking back with formulas, then derive the analysis summary (unit-mix
  by care level and A/B/C rent tier, Second Residents, Rent Adjustments, trailing move-in windows, Check)
  and reconcile to the operator's own wing totals. Use whenever the user wants a rent roll "formatted,"
  "the .RR built," "the right-side columns filled in," a unit/bed roster with in-place rent, an occupancy
  + rent-tier analysis, or hands over an operator export to turn into the model-ready rent roll — even if
  they don't name the format. Golden standard: `Aster Ridge-RR_v1` (`Aster Ridge.RR`). NOT for the .H,
  census, or P&L mapping — those are separate skills (census-formatting / hf-formatting / pnl-mapping /
  h-underwrite).
---

# Rent Roll (.RR) Formatting

**Input:** a raw operator rent roll — a per-resident occupancy + rent snapshot, hardcoded, wildly
non-uniform across operators. **Output:** the finished `.RR` tab — a uniform per-unit block + a derived
analysis summary + a reconciliation that ties to the operator's own totals. You are **given the raw data
and you transform it into everything else.** The hard part is the *source*, not the target: the right side
is a fixed contract; the left side varies by operator and even by export. Design around that asymmetry.

**Read first (progressive disclosure):**
`References/locked-rules.md` (the firm rulings — read before anything) → **`References/formatting.md`
(the VISUAL CONTRACT — read before writing a single cell; formatting is a gate, not polish)** →
`References/output-columns.md` (the 17-col unit-block contract) → `References/analysis-table.md` (the
summary you derive from it) → `References/onesite-reader.md` (how to parse a OneSite source) →
`References/validation-and-build.md` (the reconcile gate + COM/`.xls` mechanics). Golden reference
workbook: `Investments/Data/Transactions/Aster Ridge (Demo)/Rent Roll/Aster Ridge-RR_v1.xlsx`, tab
`Aster Ridge.RR` (in the private team repo the golden is a real-deal build; this public clone
substitutes the fabricated demo deal).

⚠ **Column letters in these docs are the golden's geometry, NOT the contract.** The raw export's width varies by
operator. Compute **`blockStart = lastRawCol + 1`** and treat every position as an **offset** from it
(`formatting.md` §1). A build with hardcoded `AW` is wrong on the next deal.

## The transform (raw → finished .RR)

```
  recognize ─► PLAN (ask if unclear) ─► build unit block (AW:BM) ─► derive analysis ─► reconcile
   geometry      care levels / tiers        formulas + inputs        group on BM         dollar gate
```

1. **Recognize** the export family (OneSite / Yardi / PCC) and map its geometry — header row, the left
   data columns (Unit, Floor Plan-SqFt, Resident, Move-in, Market, Actual, Care, 2nd-Occ Fee), where the
   operator's own section/wing subtotal rows are. Output a small source profile. (OneSite telltales +
   column map: `References/onesite-reader.md`.)
2. **Plan (intake gate — the expensive mistakes are interpretation, not arithmetic).** Before building:
   inventory care levels (IL/AL/MC); cluster each level's SqFt into **A/B/C tiers by natural breaks**;
   resolve MC room structure (Private vs Semi-Private vs Jack & Jill). Where the file doesn't make these
   unambiguous, **ASK the underwriter — do not invent a default.** Record answers. (Details:
   `RR-FORMATTING-DESIGN.md` → Planning Mode; tiering: `analysis-table.md` §0.)
3. **Build the per-unit block (AW:BM)** — one row per unit (**unit = bed = contract**), row-aligned to the
   left. Formulas link back (`VALUE()`-wrapped); hardcode only SqFt, cleaned move-in, and the tier/label
   codes. Handle the three special cases (P/S 2nd occupant, semi-private companion, mid-month
   de-proration) per `output-columns.md`.
4. **Derive the analysis table** — group the unit block on **Group Unit Code (BM)**: main unit-mix
   summary, Second Residents, Rent Adjustments, Check (QA), and the two trailing move-in windows.
   (`analysis-table.md`.)
5. **Reconcile** — built sums vs the operator's own wing subtotals: Market & 2nd-Occ Fee tie to the cent;
   In-Place Rent ties up to the traceable de-proration delta; Care within rounding. (`validation-and-build.md`.)

## Model assignment (Opus reasons, Sonnet grinds — mirrors the .H pipeline)
- **Reader** (parse a 600+-row dump into canonical unit records) → **Sonnet sub-agent**. Bulky,
  self-contained, well-specified; the token volume lives here. Point its prompt at this folder (it starts cold).
- **Orchestrate + plan + emit + reconcile** → **Opus (main thread).** The interpretation calls (tiering,
  the BM group key, 2nd-occupant vs companion, de-proration) and the reconcile gate are reasoning, not grunt.
- **QA verifier** (recompute the reconcile + counts independently), if added → a **different model** than
  the reader (same invariant as `h-verifier`).

## Non-negotiable mechanics (this box — see CLAUDE.md + memory)
- **No Python.** All Excel work = **Excel COM in PowerShell** (Office 16). `.xls` opens fine via COM.
- **BUILD ON A COPY — never write the operator's original.** A var-name case-collision once saved over
  two source rent rolls (memory `ps-var-name-case-collision`). The three mitigations are baked into
  `Investments/scripts/rr_build_onesite_v2.ps1` and must carry into any new engine: build-on-copy, `$w`-prefixed
  write-column constants, and the abort guard on an already-corrupted source.
- **Version up, never overwrite** the deliverable (`—RR_v2` → `_v3`); archive the superseded copy to `arc/`.
- **STOP at v1 for human review** (ruling 2026-07-15). Build v1, run the gates, REPORT verifier findings
  alongside it — do not auto-build them into a v2. Ryan decides what goes into the next version; only
  version up after his sign-off.
- Clean up with `Clear-OrphanExcel` + `New-ExcelTracked`/`Stop-TrackedExcel` (lib); the raw windowless sweep
  is hook-blocked (cross-kills a peer deal's Excel). One COM automation at a time **within a run**; a
  different deal's build in another terminal may run concurrently. Gridlines OFF on every tab. Recalc before
  save. (memory `never-blanket-kill-excel`, `excel-com-write-pitfalls`.)

## Validation gates (all four, every build)
1. **Reconcile** — Market & 2nd-Occ Fee tie to the cent against the operator's own wing subtotals;
   In-Place Rent delta = the de-proration gross-ups (traceable, not mysterious). Value-conserving, so a
   stray non-zero diff = a dropped/double-counted leaf or a bad block boundary.
2. **Zero formula errors** after recalc.
3. **Check block returns 0** on every line (independent recompute from the unit block).
4. **Formatting gate** — `formatting.md` §"Formatting gate". Right numbers in a wrong-looking tab is a
   FAILED build: this is an investor-facing deliverable. The eye-catching failures are the two header
   styles (yellow Calibri unit block vs navy Arial analysis), `$` signs leaking into the analysis table,
   `0` displayed where a dash belongs, and total rows bordered top *and* bottom.

## Status of the engine
**There is no current engine — this skill is the spec; the build is the next item.**
`Investments/scripts/rr_build_onesite_v2.ps1` is a *proven OneSite parser*, not the engine: its source-side half
(anchor detection, block assembly, leaf-sum finance, reconcile-to-the-cent) is correct and worth reusing,
but it predates the v1 contract — it emits the **16-col AW:BL** block (no **Group Unit Code (BM)**, no
tiering) and a per-unit-type summary rather than the v1 analysis table, and its paths are hardcoded to
the two OneSite test files it was built against. **Extend it; don't run it as-is.** Gap list:
`RR-FORMATTING-DESIGN.md` → "What the next build must change". Until then the golden v1 tab is the
by-hand reference to match.
