# Aster Ridge Senior Living (DEMO) — Fabrication Spec

**All data in this folder is 100% fabricated for public demo / pipeline-testing purposes.**
"Aster Ridge Senior Living" is not a real community. No resident, employee, financial, or
operational data from any real Oakmont deal was read, referenced, or reproduced anywhere in
these files or in the generator script. Every number below was invented and is reproducible
from `Investments/scripts/Build-AsterRidge-Inputs.ps1`.

## Community

84 units — IL 40 / AL 30 / MC 14. T12 window: Jul-2025 .. Jun-2026. Rent roll as-of 06/30/2026.

| Care | Unit type | SqFt | Market rent | Units | Vacant |
|---|---|---|---|---|---|
| IL | Studio | 420 | $3,295 | 8 (101–108) | 104 |
| IL | One Bedroom | 610 | $3,895 | 22 (109–130) | 119 |
| IL | Two Bedroom | 850 | $4,695 | 10 (131–140) | 136 |
| AL | Studio | 380 | $4,495 | 18 (201–218) | 205 |
| AL | One Bedroom | 540 | $5,295 | 12 (219–230) | 225 |
| MC | Private | 320 | $7,495 | 12 (301–312) | 308 |
| MC | Semi-Private (2 beds, per-bed billing) | 400 | $5,995/bed | 2 (313–314) | none |

Occupied as of 6/30/2026: **IL 37 / AL 28 / MC 13** (84 total units; MC has 15 residents
since both semi-private units are double-occupied).

## Generation rules

- **In-place rent**: `RoundTo5(market × ratio)`. Ratio by move-in bucket: 2019=92%, 2020=93%,
  2021=94%, 2022=95%, 2023=96%, 2024=97%, 2025=98%, Jan–Mar 2026=99%, Apr–Jun 2026=100%.
  Buckets are assigned cyclically (ascending unit #) within each care type to the pool of
  occupied units, then overridden for the 5 units below.
- **5 units with a move-in inside Apr–Jun 2026** (global, across the community): IL 106
  (4/8/26), IL 140 (5/5/26), AL 210 (4/22/26), AL 228 (5/19/26), MC 302 (6/12/26).
- **Second residents ("couples")**: 5 IL units (101, 103, 110, 122, 132) and 3 AL units (202,
  206, 221) carry a second resident charged only the $995/mo second-person fee (no separate
  rent/care line). MC semi-private units are NOT couples — both beds are billed independently
  at full market rent + full MC care, per the spec.
- **AL care levels** (one per occupied AL unit, assigned to the primary resident only, by
  ascending unit # — first 12 units → Level 1, next 10 → Level 2, last 6 → Level 3):
  Level 1 = $595 × 12, Level 2 = $1,095 × 10, Level 3 = $1,595 × 6.
- **MC care**: flat $1,200 per resident (all 15 MC residents, private and both semi-private
  beds).
- **Resident names**: sequential placeholders `Resident 001` … `Resident 078` — no invented
  human names.
- T12 revenue lines 4010/4020/4030/4110/4120/4210: **June-2026 = exact roster sum**; the
  other 11 months = `Occupied(month) × roster-derived average rate × (1 ± ~1% smooth wave)`,
  rounded to whole dollars, using the census occupancy schedule below as the driver.
- 4310/4410/4510 (Community Fees / Other Resident Income / Guest Meals): independent smooth
  formulas loosely tied to month-over-month net occupancy change, clamped to the ranges given
  in the build brief ($3,000–$14,000 / $2,500–$4,000 / $1,800–$2,900).
- Expenses: 30 leaf lines (12 wage depts + payroll tax [9.5% of wages] + benefits + workers
  comp + 13 non-labor lines + management fee [exactly 5.0% of that month's Total Revenue,
  rounded] + Other G&A). Utilities (Electricity/Gas/Water) carry a monthly seasonal factor;
  everything else carries a small smooth sinusoidal wiggle for realism. `TOTAL OPERATING
  EXPENSES` and `TOTAL REVENUE` are always the literal sum of their own leaf lines for that
  month — never an independently-set target — so the reconcile gate is exact by construction.
- Census ramp (Jul-2024 .. Jun-2025): linear ramp IL 32→34, AL 25→26, MC 10→11, landing
  exactly on the Jul-2025 starting values of the given Jul-2025..Jun-2026 schedule (smooth
  hand-off, no discontinuity).

## Anchor numbers (June 2026, roster-derived — ties exactly across all 3 files)

| Line | Amount |
|---|---|
| 4010 Rent Income – IL | $140,535.00 |
| 4020 Rent Income – AL | $128,925.00 |
| 4030 Rent Income – MC | $101,595.00 |
| 4110 Care Level Income – AL | $27,660.00 |
| 4120 Care Level Income – MC | $18,000.00 |
| 4210 Second Person Fees | $7,960.00 |
| **Roster grand total (rent + care + fees)** | **$424,675.00** |

Rent-roll section totals (leaf-line sums, to the cent): IL $145,510.00 · AL $159,570.00 ·
MC $119,595.00 · **Grand total $424,675.00**. (Section totals include the 2nd-person fees /
extra semi-private bed charges that sit outside the 4010/4020/4030 rent-only lines above —
e.g. IL section total $145,510.00 = $140,535.00 rent + $4,975.00 second-person fees.)

## T12 annual totals (12-month sum, verified from the saved file)

- **Annual Revenue: $5,041,509**
- **Annual Operating Expenses: $3,426,677** (67.96% of revenue)
- **Annual NOI: $1,614,832** (32.03% margin)

## Files

- `Rent Roll/06.30.2026 Aster Ridge SeniorLivingRentRoll.xls` — OneSite-family raw export
  shape (banner, header row 13, per-unit anchor/continuation/subtotal blocks, section +
  grand totals). Dollar columns stored as text (`"3,295.00"`, `"(200.00)"`) matching the
  documented OneSite raw-export quirk.
- `Historical Financials/06.2026 Aster Ridge T12 Income Statement.xlsx` — sheet "Income
  Statement", values-only GL export, 9 revenue + 30 expense leaf lines, 3 subtotal rows.
- `Historical Financials/Aster Ridge Occupancy Report Jul 2024 - Jun 2026.xlsx` — sheet
  "Occupancy", Capacity + Occupied (month-end) blocks, IL/AL/MC/Total, 24 months.

## Generator

`Investments/scripts/Build-AsterRidge-Inputs.ps1` — re-runnable, regenerates all three files
in place. Builds the full roster in memory, derives every revenue/expense figure from it or
from smooth deterministic formulas (no external randomness), writes all three workbooks in
one Excel COM session, then reopens each saved file fresh and re-derives every gate number
independently (reconcile-to-the-cent, cross-tie, occupancy-count, census-coverage) before
printing a verification report.
