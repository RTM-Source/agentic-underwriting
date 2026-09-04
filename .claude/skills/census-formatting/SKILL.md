---
name: census-formatting
description: >
  Appends the firm's standard Census summary block to a senior-housing census /
  occupancy worksheet. Reads whatever monthly census the broker/operator sent
  (occupancy %, unit counts, or a PCC/Yardi cube), takes the most recent 12
  months, and writes a fixed two-bucket table — Capacity and Occupied, each split
  IL / AL / MC / Total — below the source data, in the exact layout and styling of
  the firm's examples, then renames the tab to "Census." Use it when the user wants
  raw census data "formatted," "conformed to the Census layout," or "the Capacity/
  Occupied block added." Source files vary widely in shape; interpreting them into
  the uniform output is the value. Do NOT use it to build the .H or .RR tabs, or to
  map P&L lines — those are separate skills.
---

# Census Formatting

## What this skill does
Takes an operator/broker **census (occupancy) worksheet** in any of the many shapes
they arrive in and appends the firm's **standard Census block** beneath the existing
data: two buckets — **Capacity** and **Occupied** — each broken into **IL / AL / MC**
plus a **Total**, across the **most recent 12 months**. The source data is left
intact above; the new block is always written in the same layout and styling; the
tab is renamed **Census**.

This is a light, mechanical conform — but the *reading* of the source is the hard
part, because census arrives as occupancy percentages, raw unit counts, average
daily census, or an analytical cube, and the care-type labels differ deal to deal.
Target an underwriter's first pass: produce the standard block, and **flag anything
ambiguous rather than guess** (per the firm's non-negotiables).

Pipeline position: standalone census conform. It does not feed `.H` or `.RR`.

## Input contract (what this skill expects)
A single worksheet containing a **monthly census time series** for one community,
in one of the recognized shapes (see `references/census-format.md` → "Source
patterns"). At minimum the file must let you determine, per care type, either:
- **Capacity** — licensed units/beds (a units count, a "Beds" column, etc.), and
- **Occupied** — a census count or a basis to derive one (occupied units, average
  daily census, or occupancy % **with** a capacity to multiply against).

If the file gives occupancy **percentages only** and no unit/bed counts anywhere,
absolute Occupied cannot be derived — that care type is **left blank and flagged**
(see "Flag, don't guess"). Do not fabricate counts.

## Output: the Census block

### Geometry (fixed, relative to an anchor)
Anchor the block at a **bucket-label column** (default **column B**) and a **start
row** a couple of blank rows below the last row of source content. Let `b` = bucket
column, `c` = `b+1` (category column), `s` = `b+2` (spacer, stays empty), and the
**12 month columns** run `b+3 … b+14` (default **E…P**). Rows below are offsets from
the Output header row `r`:

| Row | Bucket col `b` | Category col `c` | Month cols `b+3 … b+14` |
|-----|----------------|------------------|--------------------------|
| `r` | `Output` | — | 12 month-end dates |
| `r+1` | *(blank)* | | |
| `r+2` | — | `Days` | days in each month |
| `r+3` | *(blank)* | | |
| `r+4` | `Capacity` | `IL` | capacity IL |
| `r+5` | `Capacity` | `AL` | capacity AL |
| `r+6` | `Capacity` | `MC` | capacity MC |
| `r+7` | `Capacity` | `Total` | Total (= IL+AL+MC) |
| `r+8` | *(blank)* | | |
| `r+9` | `Occupied` | `IL` | occupied IL |
| `r+10` | `Occupied` | `AL` | occupied AL |
| `r+11` | `Occupied` | `MC` | occupied MC |
| `r+12` | `Occupied` | `Total` | Total (= IL+AL+MC) |
| `r+13` | *(blank)* | | |
| `r+14` | `Avg Occ %` | `IL` | occupancy % IL — **the `.H` links HERE** |
| `r+15` | `Avg Occ %` | `AL` | occupancy % AL |
| `r+16` | `Avg Occ %` | `MC` | occupancy % MC |

**The `Avg Occ %` block (`r+14 … r+16`) is REQUIRED on every deal** *(Ryan, 2026-07-16)*, whatever the
source pattern — the assembled `.H` pulls its Block A occupancy **from these rows**, not from `Occupied`,
and derives Occupied as `Occ % × Capacity` (see `h-underwrite/References/h-skeleton.md` §Block A).
Store as a real fraction with a `0.00%` format (`0.9523`), never the number `95.23`.
- **%-source (Pattern A):** these rows are the *pulled operator datum* → they carry the source formula/
  value; `Occupied` (`r+9..r+11`) is then `= Capacity × Occ %`.
- **count-source (Pattern B/C):** `Occupied` is the operator's datum; these rows are `= Occupied ÷
  Capacity`. Either way the `.H` sees a uniform shape.

The category labels (`IL`, `AL`, `MC`, `Total`) are **always present even when a
care type has no data** — leave that row's month cells blank; never delete the row.
The anchor column is cosmetic and may be shifted right (the examples show both a
column-B and a column-E anchor); the **internal geometry above is fixed**.

### Cell values
- **Month header (`r`, month cols):** the 12 most recent month-ends as real dates
  (datetimes), oldest → newest, left → right. Normalize any source date (month-start,
  serial, or `"January 2019"` text) to the **month-end** date.
- **Days (`r+2`):** `=DAY(<that column's month-header cell>)` — days in the month.
- **Capacity IL/AL/MC:** licensed units/beds for the care type, per month (usually
  flat across the 12 months; carry whatever the source gives per month).
- **Occupied IL/AL/MC:** the census count per month (may be fractional, e.g. an
  average daily census of `51.45` — preserve decimals).
- **Totals (`r+7`, `r+12`):** `=SUM()` over the three category cells **above in the
  same column** (e.g. `=SUM(E15:E17)`). SUM treats a blank care type as 0, which is
  correct.
- **Avg Occ % (`r+14 … r+16`):** a real fraction, `0.00%` format. On a %-source, point it at the
  operator's own monthly figure **by formula at the raw census rows sitting above the block on the
  same sheet** (provenance stays live; **black font** — same-sheet formula per the provenance rule.
  Green only if a cell must pull from an *additional* imported source tab); on a count-source,
  `= Occupied ÷ Capacity`.
  **There is no `Total` row here** — the blended occupancy is derived in the `.H` (row 22) from
  Occupied-Total ÷ Capacity-Total, which weights by capacity. Never average the three care-level
  percentages; that silently equal-weights a 35-unit MC wing against a 137-unit AL wing.

### Styling (match the examples exactly)
Everything in the block is **Aptos Narrow 11**.

- **Output header row** (`r`, spanning bucket col `b` through the last month col,
  **including the spacer**): **bold**, solid fill **`FFD3D3D3`** (light grey), **thin
  bottom border**, **centered**. Month cells use number format **`mm-dd-yy`**; the
  `Output` label and the blank label/spacer cells in this row carry the same fill +
  border (one contiguous grey band).
- **`Days` label** (cat col): bold, left-aligned, no fill/border. **Day values:**
  not bold, right-aligned, accounting format (below).
- **IL / AL / MC rows:** bucket + category labels **not bold**; data cells not bold,
  right-aligned, accounting format.
- **Total rows** (`r+7`, `r+12`): bucket label (`Capacity`/`Occupied`) **not bold**;
  the `Total` category label **bold**; data cells **bold** with a **thin top border**,
  right-aligned, accounting format.
- **Accounting number format** for all month value cells:
  `_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)` — comma-grouped, negatives in
  parentheses, zero shown as `-`, no decimals.
- **Do not change the source sheet's column widths** or touch the source rows.
- **Turn the sheet's gridlines off** (`Window.DisplayGridlines = false`) — every house
  output is gridlines-off.

## Flag, don't guess
Surface judgment calls as a short cell **comment** on the affected row **and** in the
chat summary — never silently resolve:
- **Percentage-only source** → this is **Pattern A, and it is buildable** — do not reflex to
  blank. Capacity is nearly always stated in a *sibling* file (a single-month occupancy detail
  report, the OM, or the rent roll). Find it, hold it **static** across the 12 months, and make
  Occupied a live formula `= Capacity × occ%` off a visible occ% input block
  (`References/census-format.md` → Pattern A). Only if capacity is genuinely nowhere in the
  deal folder → leave that care type's Capacity *and* Occupied blank and flag for unit counts.
- **A source care type that isn't IL / AL / MC** (Skilled Nursing, Respite, a blended
  "AL/MC" line, etc.) → do not force it into a bucket; flag it for the underwriter.
- **Only a community Total with no care-type split** → leave IL/AL/MC blank, put the
  total where you can, and flag that the split is unavailable.
- **Ambiguous occupied basis** (e.g. both an average daily census and a last-day
  count present) → use ADC, note the choice.
- **Fewer than 12 months** in the source → **STOP. Do not emit a partial block.**
  *(Hard rule since Deal B, 2026-07-16 — see below.)*
- **A value that looks inverted or off** (occupied > capacity, negative count) → carry
  as-is and flag; do not "correct" it.

A flag is the skill doing its job — it marks the rows an underwriter must eyeball.

## A short census is a HARD STOP, not a flag (learned the expensive way)

**An empty census cell is not blank — it is `0`.** `=Census!E48` pointed at an empty cell returns
`0` (Double). That zero then flows into the `.H`: `AVERAGEIFS` in the S/T/U window columns *counts*
it, so each missing month drags T12 occupancy down by ~1/12 **with no error anywhere**. Both `.H`
reconciliation Checks stay at 0 the whole time, because revenue and NOI do not depend on census.
A partial census therefore passes **every** gate we have and ships looking clean.

This is not hypothetical. **Deal B (2026-07-16)** shipped a `.H` whose census had **1 of 12 months**
populated: 121 `#DIV/0!` in the metrics block, and T12 occupancy that read roughly 8 points low against
the true figure. It passed both Checks. The old rule right above ("fill what exists … flag the short
history") is what authorised it.

So:
1. **Before building, resolve the required window** — it is the **HF's T12 window**, not "the most
   recent 12 months of whatever this file has". The orchestrator hands you the window; honour it.
2. **Inventory EVERY candidate census file in the deal folder before picking one.** Deal B's real
   defect was source selection: the stage read a single-month occupancy detail file (named for that one
   month) while a full multi-year Summary workbook — a full monthly history — sat in the same
   folder, unread. A file named for one month is a *detail* report; the history is usually a
   separate "Summary"/multi-year workbook. **Never build off the first census-looking file you find.**
3. **If, after that sweep, the sources cannot cover all 12 months of the required window → STOP and
   report.** Name the months you can cover and the ones you cannot. Do not emit a block with holes;
   do not interpolate, carry forward, or infer occupancy from revenue. Missing census is a
   data-room gap for the broker, not a modelling problem to solve.
   **Exception — a COMPLETE but non-overlapping calendar year is not "missing data."** If a fuller
   source has a genuinely complete 12-month year, just not the window's own year, that is not a gap
   to STOP on — it is a choice between §"Aligning the source columns to the window"'s two resolutions
   (date-match what overlaps, or a ruled calendar-year proxy for what doesn't). STOP is for when
   **no** source — window-aligned or not — has 12 complete months to offer.
4. **Cross-check the source before trusting it.** Deal B's `Summary` workbook had **December as a
   byte-identical duplicate of January** across all 14 unit-type rows — real-looking numbers that
   were not December's. Compare each month against its neighbours and against the same month in
   adjacent year sheets; identical columns, or a month equal to another month cell-for-cell, means
   a broken export. Flag it and get a ruling — never assume duplicated data is coincidence.
5. **If a ruling says to proceed on imperfect data anyway,** record it where it cannot be missed:
   a **`Notes` tab as the first sheet** of the workbook, stating the defect, the ruling, the impact
   on specific figures, and what to request from the broker. Keep the Census tab itself clean.
6. **COPY the raw census sheet into your build workbook and build the block IN that copied tab**
   *(Ryan, 2026-07-22 — supersedes the separate `Occupancy Source` tab design of 2026-07-16. On a
   pipeline run the build workbook is now the SCRATCH workbook the orchestrator names, not the
   Deal.H Model directly — 2026-07-28, see rule 9; the one-tab structure is identical either way)*:
   paste **values + number formats** into a fresh sheet (never `Worksheet.Copy` FROM the operator
   workbook — it drags its external links along), skipping error cells (`#DIV/0!` marshals as the Int32
   `-2146826281` if you let it through). Append the Census block below the pasted data, pull the
   `Avg Occ %` rows **by in-sheet formula from the raw rows above** (black font), and rename the tab
   `Census`. One tab: provenance stays live and checkable, with no separate source tab, no separate
   block-only tab, and no cross-tab wiring to build — this is also the latency win. **If the window
   needs a second source file** (Deal B: multi-year Summary + single-month detail for capacity),
   import each *additional* source as its own tab (keep the `Occupancy Source` naming) and pull
   those cells cross-tab (green); the primary source is always the one pasted into `Census` itself.

## Aligning the source columns to the window — two valid resolutions, pick deliberately

**Never fill by column position without deciding first.** A monthly occupancy export starts at
**January**, but the `.H` window almost never does — a deal window starting in **May**, for instance,
makes the window's first month the source's **5th** column, not its first. An *accidental* positional
drag-fill silently shifts every month (this happened once on Deal B: an early month's occupancy landed
under the wrong month's header). **Nothing catches an accidental shift**: the values are real, the
formulas are valid, there are no errors, and both `.H` Checks stay at 0 — they tie revenue and NOI,
which never touch census. So never fill positionally by accident — only as a named, ruled choice
(Resolution B below).

There are **two legitimate resolutions**. Which one applies is a judgment call for the underwriter —
present both if the source doesn't cleanly cover the window, don't silently pick one.

**Resolution A — DATE MATCH (the default).** Use when the fuller source's calendar period actually
overlaps the window well enough to cover most of it.
1. Read the source's own **month-header row** and resolve each column to a real month-end date.
2. For each of the 12 window months, **look up the source column whose date equals it**. A window
   month with no matching source column is a genuine gap (see the HARD STOP above) — leave it blank,
   never shuffle a neighbouring column into it.
3. **GATE:** read the block back and assert `Census month-header[i] == source month-date[i]` for
   every populated column. Log the mapping (`Census E (05-31-25) <- source col G (05-31-25)`).

**Resolution B — CALENDAR-YEAR PROXY (named override, requires an explicit ruling).** Use when the
fuller source's period does **not** overlap the window at all (a full prior calendar year vs. a
trailing-12 window that starts mid-year), so date-matching would leave most of the window empty.
Ryan's final ruling on Deal B (2026-07-16): the original census is a single month — good for unit
counts, useless for a trend — and the window's own real data is only 2 months plus one later
snapshot. Rather than ship a mostly-blank census, apply the fuller source's **most complete calendar
year** straight across all 12 window slots, **in calendar order** (source Jan → window slot 1, source
Feb → window slot 2, … source Dec → window slot 12) — regardless of what calendar month each window
slot is actually labeled. This is a **proxy for a typical annual occupancy cycle**, not month-matched
data, and must be labeled as such:
1. Confirm capacity/unit counts still come from the **original** (thin) census, not the proxy source —
   the proxy source usually has occupancy % only, no unit counts.
2. Walk the proxy year's 12 columns straight into the window's 12 columns in order — no date lookup,
   no gaps, no skipped months (a complete calendar year has no gaps to leave).
3. **The `.H`/Census month-header dates do NOT change** — they stay the window's real dates (matching
   the HF), because the financial data in Blocks B/C/D is genuinely for those months. Only the
   occupancy values underneath are the proxy.
4. **Document prominently on the Notes tab**: name the proxy year, state plainly that occupancy is
   "a representative annual pattern, not month-matched," and give the column mapping. A reader must not
   mistake the header date for the occupancy date.
5. **GATE:** confirm the mapping is a clean 1:1 calendar walk (source col *k* → window slot *k*, all 12,
   no exceptions) and that a Notes-tab entry documenting the proxy actually exists. The gate cannot
   (and should not) assert header-date == source-date here — that equality is deliberately false by
   design; instead it asserts the *ruling is on record*.

## Procedure
0. **Inventory every candidate source in the deal folder first** (`*occupancy*`, `*census*`,
   `*summary*`, and any multi-year workbook) and list what window each one covers. Pick the
   file(s) that cover the **required window**, not the first one you open — and expect to need
   **more than one** (Deal B needed a multi-year Summary workbook for 10 months **plus** a
   single-month detail report for the 12th). If nothing covers the window, STOP (see above).
1. Read the source via **Excel COM in PowerShell**: `$ws.UsedRange.Value2` for exact
   values and date cells (a 1-based 2-D array; dates come back as serials — convert with
   `[datetime]::FromOADate`). No Python on this box — openpyxl/recalc.py do not run here.
   **Read `.Value2`, not `.Text`** — an `#DIV/0!` month marshals as the Int32 `-2146826281`,
   which the display text hides; that is how an empty month passes for data.
2. **Identify the source pattern** and the care-type rows/columns
   (`references/census-format.md` → "Source patterns" and "Care-type synonyms").
3. **Build the window:** it is the **HF's T12 window** (the orchestrator gives it to you), not
   "the most recent 12 months in this file". Normalize each column to a month-end date and
   confirm you have all 12 before writing anything.
4. **Derive Capacity and Occupied per care type per month** for that window, using the
   rule for the detected pattern. Map source care types to IL / AL / MC; leave any
   absent care type blank.
5. **Write the block** below the source per the geometry + styling above: literal
   month-end dates, `=DAY()` for Days, extracted values for IL/AL/MC, `=SUM()` for the
   Totals. Add flag comments on judgment-call rows.
6. **Rename the worksheet to `Census`** (leave the source data intact above the block).
7. **Recalculate** by saving — Excel recalcs natively on `$wb.Save()`, so the `Days` and
   `Total` formulas resolve; confirm **zero formula errors**.
8. **Validate** before presenting: each Total column equals IL+AL+MC; the 12 month
   headers are consecutive month-ends **matching the HF window**; Occupied never silently
   exceeds Capacity (flag if it does). **GATE: read every Capacity and Occupied cell back and
   assert all 12 months are populated** — a formula-error scan cannot see this, and neither can
   the `.H` Checks. `0` is a populated-looking value; test for it explicitly. If any month is
   empty or zero, that is a FAIL, not a flag.
9. **Where you operate depends on the run mode.**
   - **Pipeline (`.H`) run** *(2026-07-28 — scratch-workbook design, so this stage can run in
     parallel with the HF stage)*: operate on a **fresh SCRATCH workbook saved to the local `%TEMP%`
     path the orchestrator's prompt names** (never SaveAs to a OneDrive path — hook-enforced hang).
     Paste the chosen raw census sheet in as a fresh tab (values + number formats — see HARD-STOP
     rule 6), then append the block below the pasted data and rename that tab `Census`. Keep the
     workbook **self-contained** (every formula in-workbook, zero external links) — the
     `h-assembler` later `Worksheet.Copy`s the tab(s) into the Deal.H Model workbook at assembly
     step 0, which is only safe because the scratch is clean. The block's start row floats with the
     raw data's height — report the resulting **cell map** AND the scratch path so the orchestrator
     can hand both to the assembler.
   - **Standalone run:** operate in place on the provided workbook itself — append the block below
     the source data, rename the tab, save. No copy needed; the source's hardcoded data is never
     altered. (A fresh copy is only for first-time/example builds, e.g. seeding `Census Examples/`.)
   Either way, report the window used, any care types left blank, and every flag.

## Out of scope (do not do here)
- Building the `.H` (Historical Financials) or `.RR` (Rent Roll) tabs.
- Mapping P&L line items to tags → `pnl-mapping` skill.
- Inventing capacity/occupancy when the source can't support it → flag instead.

## Related skills & references
- **References (this skill):** `references/census-format.md` — the target block spec (verbatim), the three
  source patterns (occupancy-% / unit-counts / PCC-Yardi cube), care-type synonyms, and the validation
  checklist. Load it when running.
- **Build:** Excel COM in PowerShell (no Python); recalc on `$wb.Save()`. Gridlines off, like every house
  output. Pipeline runs paste the raw census into a local **scratch workbook** and build there (one tab;
  the `h-assembler` merges it into the Deal.H Model at assembly — 2026-07-28); standalone
  runs operate in place on the given workbook. Either way: append the block + rename the tab `Census`;
  never alter the source data.
- **Pipeline position:** standalone conform, but it is **stage 1** of the `.H` build — the Census block
  feeds **`.H` rows 1–24** (linked straight from the Census tab). See `../H-PIPELINE-ORCHESTRATION.md`.
- **Memory:** `census-formatting` examples (Example HF-1/HF-2/HF-3), `env-no-python-use-excel-com`,
  `never-blanket-kill-excel`.
