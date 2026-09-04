# Census Format — Reference

> **Companion to `census-formatting/SKILL.md`.** Load this when running the skill.
> It captures the target block spec verbatim, the source shapes census arrives in,
> and the care-type vocabulary — the things the skill needs that aren't obvious from
> a single example. Derived from three validated examples: **Example HF-1**, **Example HF-2**,
> **Example HF-3**.

---

## 1. The target block (authoritative spec)

A fixed 13-row table appended below the source data. Columns are described relative
to an **anchor**: `b` = bucket-label column (default **B**), `c = b+1` (category),
`s = b+2` (spacer, empty), month columns `b+3 … b+14` (default **E…P**, 12 wide).

```
r     : [b]=Output                          [b+3..b+14]= 12 month-end dates (mm-dd-yy)
r+1   : (blank)
r+2   : [c]=Days                            [b+3..b+14]= =DAY(month cell)
r+3   : (blank)
r+4   : [b]=Capacity [c]=IL                 [..]= capacity IL
r+5   : [b]=Capacity [c]=AL                 [..]= capacity AL
r+6   : [b]=Capacity [c]=MC                 [..]= capacity MC
r+7   : [b]=Capacity [c]=Total              [..]= =SUM(IL:MC) in that column
r+8   : (blank)
r+9   : [b]=Occupied [c]=IL                 [..]= occupied IL
r+10  : [b]=Occupied [c]=AL                 [..]= occupied AL
r+11  : [b]=Occupied [c]=MC                 [..]= occupied MC
r+12  : [b]=Occupied [c]=Total              [..]= =SUM(IL:MC) in that column
```

**Invariant:** all four category rows (`IL`, `AL`, `MC`, `Total`) exist in **both**
buckets even when a care type is empty — empty means blank month cells, not a deleted
row. The skeleton is always the same; only which cells carry data changes.

### Styling tokens (identical across all three examples)
- Font everywhere in the block: **Aptos Narrow 11**.
- **Header row `r`** (band spans `b … b+14`, incl. spacer): bold · fill `FFD3D3D3` ·
  thin **bottom** border · centered. Month cells: number format `mm-dd-yy`.
- **`Days`** label: bold, left, no fill/border. Day values: regular, right, accounting.
- **IL/AL/MC** rows: labels regular; values regular, right, accounting.
- **Total rows**: bucket label regular; `Total` label **bold**; values **bold** with
  thin **top** border, right, accounting.
- Accounting format string (month value cells):
  `_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)`
- Leave the source sheet's existing column widths as-is; don't impose new ones.

### Placement
- Start the block ~2 blank rows below the last non-empty source row (examples ranged
  2–4; not load-bearing).
- Default anchor column **B**. The examples also show an anchor at **E** (Example HF-2,
  Example HF-3) — the horizontal position is cosmetic; pick one and keep the geometry.
- **Rename the tab `Census`** as the final step. Source data stays intact above.
- **Pipeline (`.H`) runs (Ryan, 2026-07-22; scratch-workbook update 2026-07-28):** the "source data
  above" is the raw census sheet pasted (values + number formats) as the first act of the stage into
  the stage's **local scratch workbook** — the `h-assembler` merges the finished tab into the Deal.H
  Model workbook at assembly step 0. The block is appended beneath the raw data in that same tab. One
  tab, no separate `Occupancy Source`, no block-only Census tab. The block's start row therefore
  floats with the raw data's height; the stage reports the resulting cell map (and the scratch path)
  for `.H` Block A wiring.

### Values: formulas vs. hardcodes
The human examples hardcoded everything. This skill improves two cells to formulas
(identical displayed result, recalc-safe): **Days** = `=DAY(month cell)`, **Total** =
`=SUM(<3 category cells above>)`. Extracted Capacity/Occupied values are hardcoded
numbers (they come from the source). Month headers are literal month-end dates.
**No Python on this box** — build via Excel COM in PowerShell; Excel recalcs natively on `$wb.Save()`,
so the `Days`/`Total` formulas resolve on save (no separate `recalc.py` step). See `Investments/lib/HF-Build-Lib.ps1`.

---

## 2. Source patterns (the interpretation problem)

Census arrives in very different shapes. Detect which one you have, then derive
Capacity and Occupied from it. The three confirmed patterns:

### Pattern A — Occupancy percentages only (e.g. **Example HF-1**, **Deal B**)
Rows are care types (IL / AL / MC / Total); cells are **occupancy %** by month. There
are **no absolute unit/bed or census counts** in the sheet.
- **Capacity:** not present in *this* file — but do **not** conclude it is unavailable.
  Look for a unit/bed count per care type in this order: (1) the underwriter/prompt
  (Example HF-1 used IL 40 / AL 30 / MC 14); (2) **another file in the deal folder** — a
  single-month occupancy *detail* report, an OM, or the rent roll almost always states
  unit counts (Deal B's IL/AL/MC counts came from the `Units` column of a same-folder
  single-month occupancy detail report, sitting beside the % history). Only if it is nowhere,
  **leave Capacity blank and flag.**
- **Capacity is held STATIC across all 12 months** — it is a physical/licensed attribute,
  not a monthly measurement, and no source gives it per-month. State the assumption on the
  Notes tab; it is safe for a stabilised community, so confirm against the rent roll if
  construction/licensing changed in the window.
- **Occupied:** `= Capacity × occupancy%` per month — derivable as soon as Capacity is known.
  Write it as a **live formula**, not a pasted number, pointed at the `Avg Occ %` rows (which in
  turn pull the operator's own % from the raw rows above — see the one-tab design in §1 Placement).
  The underwriter can see and retune the occupancy driving the model right on the sheet. (The
  pre-2026-07-22 builds used a separate blue hardcoded occ% input block — Deal B rows 58–61 — which
  the pasted-raw-data design makes redundant; don't add one.) Without Capacity, **leave Occupied
  blank and flag.**
- Never back into counts from percentages **alone** — but Capacity × % is not "alone", it is
  the intended derivation for this pattern.
- **Pull the operator's care-level TOTAL line — do not re-derive it** *(Ryan's ruling, 2026-07-16)*.
  The occupancy report states a total per care type (Deal B: `Total for Assisted Living(AL)` etc.);
  pull that line straight into the `Avg Occ %` block. **We tie to the operator's own reported
  occupancy** — the same principle as reconciling to the operator's own subtotals. Do not substitute
  a figure of our own construction; the broker's stated occupancy is the number the deal is discussed
  against, and a model that quietly disagrees with it invites a reconciliation nobody asked for.
  - *Known, accepted imprecision:* Deal B's `Summary` totals are `=AVERAGE(sub-types)` — an unweighted
    mean of unit-type percentages, so a 14-unit wing counts as much as a 20-unit one. A unit-weighted
    derivation moves the AL figure by roughly half a point (deltas to ~0.5pt). **Ryan reviewed this and
    chose the operator's line anyway.** If a deal ever turns on a delta that size, raise it — do not
    silently switch methods.
  - Deal B's *detail report* total needs no such caveat: it is built from resident-unit-days and is
    already unit-weighted (a weighted derivation reproduces its figures exactly).
- **Test month presence on the SUB-TYPE rows, not the total.** An operator total can be `#DIV/0!` while
  the sub-types hold data, and vice versa. (On Deal B, one month is genuinely blank at both levels — the
  gap is real, not an artifact of reading the total.)
- **Expect to draw from more than one file — but decide date-match vs. calendar-year proxy FIRST**
  (see `SKILL.md` §"Aligning the source columns to the window" for the full decision). Deal B's data
  room had a thin single-month detail report (the "original" census) plus a multi-year `Summary`
  history whose most complete year doesn't overlap the T12 window at all. Ryan's
  final ruling: use the **Summary sheet's most complete calendar year, straight across, as the occupancy
  proxy** (Resolution B) — not a stitched date-match across multiple Summary years plus the detail
  report, which was an earlier attempt that only partially covered the window and still left a gap.
  **Capacity still comes from the thin original detail report** regardless of which occupancy
  resolution is used — it is the one figure that source reliably has. Prefer the operator's own
  stated `Occupied (%)` / total-line figure over recomputing from unit counts, so the methodology
  stays uniform across all 12 columns (see the TOTAL-line ruling above).

### Pattern B — Explicit unit counts (e.g. **Example HF-3**)
The sheet states counts directly, typically a pair per care type:
`IL Units` (capacity) and `IL Occupied Units` (occupied), etc. Sometimes an
`Occupancy` % row is also present (ignore it — derive from the counts).
- **Capacity** = the "Units"/"Beds" count for the care type.
- **Occupied** = the "Occupied Units" count.
- Care types absent from the source stay blank (Example HF-3 is IL-only → AL/MC blank,
  Totals equal IL).

### Pattern C — Analytical cube / PCC or Yardi export (e.g. **Example HF-2**)
A wide cube with a month dimension and, per care type per month, columns like
**ADC** (average daily census), **Occupied Last Day of Mo.**, **Beds**, **Occupancy**.
Header rows carry month labels (e.g. `January 2019`), care types are rows under a
community.
- **Capacity** = **Beds**.
- **Occupied** = **ADC** (average daily census) — preferred over the last-day count;
  if only the last-day count exists, use it and note the choice.
- Cube month labels may be text or span merged header cells — normalize to month-end
  dates and select the most recent 12.
- These exports are the messiest; expect to write per-file parsing logic to locate
  the care-type rows and the per-month measure columns.

> When a source doesn't match A/B/C cleanly, identify the capacity basis and the
> occupied basis explicitly, document the call in the chat summary, and flag it.

---

## 3. Care-type synonyms (map source labels → IL / AL / MC)

Source labels vary; map by meaning, not literal text:

| Output bucket | Recognized source labels (case-insensitive) |
|---------------|----------------------------------------------|
| **IL** | Independent Living, Independent, IL, Ind |
| **AL** | Assisted Living, AL, Personal Care, PC, Residential Care |
| **MC** | Memory Care, MC, Alzheimer's, Alz, Dementia, SCU (special care unit) |

- A label that maps to none of the three (e.g. **Skilled Nursing / SNF**,
  **Respite**, **Independent + Assisted blended**) is **not** forced into a bucket —
  flag it.
- A care type the community simply doesn't operate → its IL/AL/MC row stays **blank**
  (don't delete). Totals `=SUM()` naturally skip it.

---

## 4. The 12-month window

- **Pipeline (`.H`) runs: the window is the HF's T12 window, handed to you in the spawn prompt** —
  never "the latest 12 months of whatever source you opened." Standalone runs with no window given:
  the latest 12 month-ends in the source's monthly series.
- Normalize every source date to a **month-end** date for the header (month-start,
  Excel serial, or text like `"January 2019"` all become the month-end datetime).
- Order oldest → newest, left → right, in the 12 month columns.
- **Fewer than 12 months available → HARD STOP, not a partial fill** *(supersedes the old
  fill-N-and-flag rule here — that rule authorized the Deal B defect: 1-of-12 months shipped through
  every gate; see SKILL.md §"A short census is a HARD STOP")*. Do not write a partial block; report
  which months you can and cannot cover, and stop for a ruling.
- The header dates drive the `Days` row via `=DAY()`, so they must be true dates.

---

## 5. What the three examples taught us (diffs)

| | Example HF-1 | Example HF-2 | Example HF-3 |
|---|---|---|---|
| Source pattern | A (occupancy %) | C (PCC cube: ADC/Beds) | B (unit counts) |
| Care types present | IL, AL, MC | AL, MC (IL blank) | IL only (AL, MC blank) |
| Capacity source | external unit counts (not in %-sheet) | Beds column | "IL Units" row |
| Occupied source | Capacity × occ% | ADC | "IL Occupied Units" row |
| Occupied is fractional? | yes (avg census) | no (integers) | yes (half-units) |
| Anchor column used | B | E | E |

Takeaways folded into the skill:
- The block skeleton and styling are **identical** in all three; only the source
  interpretation and which care-type cells are populated differ.
- Empty care types are **kept as blank rows**, never deleted (Example HF-2's IL, Example HF-3's
  AL/MC).
- Totals are sums of the present care types (Example HF-3 Total = IL; Example HF-2 Total =
  AL + MC).
- The anchor column is not fixed across examples — the **geometry** is what's fixed.
- Percentage-only sources (Example HF-1) depend on an external capacity input; without it,
  the right move is to leave blank and flag, not to invent counts.

---

## 6. Validation checklist (before presenting)
- [ ] Each `Total` column = IL + AL + MC for that column.
- [ ] 12 month headers are consecutive month-ends matching the window you were given (pipeline: the
      HF's T12 window; standalone with no window given: ending at the source's latest month).
- [ ] `Days` equals the calendar days of each header month.
- [ ] Every care type the source lacks is **blank**, with its row still present.
- [ ] Occupied does not silently exceed Capacity (flag any case that does).
- [ ] Tab renamed `Census`; source data untouched above the block.
- [ ] After `$wb.Save()` (COM recalc), **zero formula errors** on the sheet.
- [ ] Every judgment call appears as a cell comment **and** in the chat summary.

---

## 7. Growth note
Each new deal that introduces a new source shape, a new care-type synonym, or a new
occupied/capacity basis should be added to §2 / §3 here so the skill's reading ability
compounds with each file processed.
