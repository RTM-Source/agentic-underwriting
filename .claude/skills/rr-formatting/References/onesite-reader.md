# OneSite Senior Living — reader recipe

> ⚠ **This reader is for ONE source family (OneSite). It does NOT describe every rent roll.** The
> `.RR` *contract* (the 17-col block, the analysis table, the gates, the formatting) is source-agnostic;
> **this file is not.** If the source lacks the OneSite telltales below, none of this parsing applies —
> recognize the family first (next section), then parse to fit. Do **not** force a non-OneSite file
> through this recipe.

Source family for the first two test deals — Family 1/2 (OneSite — the O&O communities), e.g.
`Investments/Data/Rent Roll Tests/{O&O Community 1,O&O Community 2} Rent Roll.xls`.
Telltales the recognizer keys on: a `OneSite Senior Living` banner (row 1), a `Floor Plan - SQ FT` header,
`NNN-NNN TOTAL` per-unit subtotal rows, and a long transaction-code amount table after the detail.

## Other families seen (parse to the SAME contract, different source shape)
**Yardi / RealPage flat report** (e.g. Deal C, 2026-07-14 — an **IL** community):
- One flat line per unit; **no** OneSite banner, **no** `Floor Plan - SQ FT` composite, **no**
  `NNN TOTAL` subtotal rows, **no** trailing transaction table, **no** care-level section headers.
- **Two-physical-row header** (e.g. `Market` over `Rent`), not one wrapped row — so `headerRow` =
  the header block's **bottom** row − 1 (rule 15's "−1" still holds, just pick the completing row).
- Columns are **true numbers**, not text — the `IFERROR(VALUE(…),"")` wrapping (output-columns.md) is
  unnecessary; direct cell refs tie exactly. (Still harmless to wrap.)
- **No Care Fees, no 2nd-Occupant Fee columns** → those two reconcile lines are **N/A**, and
  de-proration (rule 11) typically does **not** apply (verify: a move-in on the report's own As-Of date
  posting full un-prorated rent proves proration is off). Say "N/A + why", don't fabricate a $0-tie.
- **IL-only** → `Care Type` is a flat `"IL"` (rule 6's `LEFT(UnitType,2)` gives garbage on codes like
  `ns_A1B`; its own caveat already said "revisit if IL appears"). Unit-type codes are opaque — map to
  Studio/1BR/2BR by SqFt family, mark PROPOSED, and the **populated tier rows go in the IL section**
  (the AL/MC rows collapse to blank placeholders — the inverse of the OneSite examples).

**The invariant:** recognize the family, parse the source to fit, but the **output contract does not
bend** — same 17-col block, same analysis table, same gates, same formatting.

## Geometry (example files; verify per file)
- Header row **13**. Left data columns: **A** Unit · **E** `Floor Plan - SQ FT` (`AL-1 - 545`) ·
  **I** Resident name / `Vacant` / `NNN TOTAL` · **N** Move-in · **R** Moved-onto-property ·
  **V** Market · **Z** Actual · **AA** Var-to-market · **AC** Vacancy var · **AE** Total var ·
  **AG** 2nd-Occupant Fee · **AI** Care Fees · **AK** Other Fees · **AN** Credits · **AQ** TOTAL.
- Right target columns start at **AW** (49). The finished block runs **AW:BM** (17 cols) + the analysis
  table to its right — see `output-columns.md` / `analysis-table.md`.
- Values are **text** (`"5,885.00"`, `"(759.00)"`). Negatives in parens, thousands commas.

## Line types
- **Section header** — col A is `Assisted Living` / `Memory Care` (no floor plan). Sets care context; no row.
- **Anchor (unit line)** — col A non-empty **and** col E matches `<type> - <sqft>`. Starts a unit block.
- **Continuation (leaf) line** — blank col A; col I is a resident name or `Vacant` (a within-block detail).
- **Unit subtotal** — blank col A; col I ends in `TOTAL` (`112-112 TOTAL`, `139 TOTAL`). **Ignore for value.**
- **Section / property total** — col I starts with `TOTAL ` (`TOTAL Assisted Living`, `TOTAL Aster Ridge…`).

## Block assembly (per anchor)
1. Block = anchor row through the row before the next anchor.
2. Walk it, collecting **status lines** (anchor + blank-col-A residents/`Vacant`) in order; note any
   `NNN TOTAL` row but **do not** use its value. Stop at any `TOTAL ` (section) row.
3. `leafRows` = the status lines (the unit's own detail rows). Finance = sum of these (see rule 7).
4. **Occupancy** = period-end = kind of the **last** status line: a resident ⇒ occupied; `Vacant` or an
   `Inactive-Down unit` ⇒ unoccupied (occ 0, capacity stays per SqFt). `moveInRow` = the current
   resident's row (the last resident line) when occupied, else null.
5. **Capacity** = `IF(sqft>0,1,0)`. A shared **"S" line** (code ends `S` / SqFt 0) is **zeroed** (cap 0,
   occ 0); its 2nd-occupant **fee** lands on the primary's 2nd-Resident column (BJ) — not a `Y` flag
   (locked-rules rules 3 & 5, revised).

## Two traps that cost real time (already solved in the script)
- **Companion / P-S double-count:** a `NNN TOTAL` (e.g. `139 TOTAL`, `135 TOTAL`) rolls up *sibling*
  anchors (`139-139A`+`139-139B`, or a `P` line + its `S` line). Reading that cell double-counts.
  **Fix:** value each unit from its **own** `leafRows`, never the TOTAL cell.
- **Detail end:** the `TOTAL Aster Ridge…` property total lives in a **shifted column block**, so a
  column-I search misses it and the last anchor's block sweeps the bottom transaction-code amount table.
  **Fix:** bound the detail at the **last in-detail section total** (`TOTAL Memory Care`); ignore below it.

## Per-deal counts (for sanity, not gates)
- O&O Community 1: roughly 115 unit rows; occupied count ties the operator's own "Units Occupied" total.
- O&O Community 2: roughly 108 unit rows; occupied count moved by a couple of units once the
  inactive-down ruling was applied (rule 3) — worth a sanity re-check on any deal, not just this one.
- Reconciliation (the real gate) ties to the cent — see `validation-and-build.md`.
