# .RR Analysis Table — the summary derived from the unit block

Everything right of the per-unit block. Locked to `Aster Ridge-RR_v1.xlsx` → `Aster Ridge.RR`. All of it
is **derived** from the unit block by grouping on **Group Unit Code** (the join key — NOT Basic Unit Code,
and there is no "BN" column; memory `rr-analysis-join-key` has the history).

⚠ **Column letters here are the golden's geometry, not the contract.** The raw export's width varies;
compute `blockStart = lastRawCol + 1` and use **offsets** (`formatting.md` §1). Letters below assume the
golden's `blockStart = AW`. **Read `formatting.md` before building** — the visual spec is a gate, and it
carries the exact row plan, the two header styles, the number formats, and the border rules.

## 0. Tiering helper (golden: **BP12:BR…**, header row 12, data from row 13) — human reference, not a live chain
Three columns: **`Unit Sqft` · `Count` · `Group Unit Code`**, under a **yellow** header (same style as the
unit block, *not* the navy analysis style — see `formatting.md` §2). Assigns each distinct SqFt to an
A/B/C tier by natural clusters. Purely a worksheet the underwriter reads while hand-typing the tier into
each unit's Group Unit Code.
- `Unit Sqft = UNIQUE(FILTER(<SqFt col>,<SqFt col>>0))` — distinct real SqFt (range **stops before MC**:
  MC groups by room-type, not SqFt). Zero-SqFt "S" rows filter out by construction (Capacity = 0).
- `Count = COUNTIFS(<SqFt col>, eachSqFt)` — unit count per SqFt.
- Tier column (`A`/`B`/`C`) is **hand-assigned by the underwriter** (rule 9). **PROPOSE, then ASK** —
  never silently infer. The two-pass method: **(1)** cluster by SqFt similarity, **(2)** rebalance toward
  **roughly equal unit counts** per tier, **(3)** but a **high-frequency SqFt takes its own tier** and a
  **small group never splits**. Unit count outranks SqFt distance.
  A wrong first pass is expected and acceptable — tiering is weighted far below the dollar ties.

  **⚠ "Natural breaks" alone does NOT reproduce an underwriter's assignment — do not trust that
  heuristic.** Worked illustration on a hypothetical spread roster (the shipped demo golden's unit types
  each carry a single SqFt, so its tiering is trivially all-A; real rosters look like this instead):

  | Group | SqFt → tier |
  |---|---|
  | Studio | `410,428,442` → **A** · `455,468` → **B** · `489` → **C** |
  | 1BR | `545` → **A** *(alone)* · `556,562,568` → **B** · `615,624` → **C** |
  | 2BR | `860,905,918` → **A** *(all one tier)* |

  Note what a gap-based clustering gets **wrong** here: `545 → 556` is only an **11 SF** gap yet the
  underwriter **splits** it, while `556 → 568` (12 SF) stays together and all three 2BR sizes (860→918, a 58 SF
  spread) collapse into one tier. The driver is **unit count, not SqFt distance** — `545` is the modal
  floorplan (the plurality of the AL units) so it earns its own tier, and 2BR is too small a group to
  split. A cold agent applying "natural breaks" produced 1BR A(545–568)/B(615–624) + 2BR A(860)/B(905–918)
  — reconciles perfectly, ties to every headline number, and is still **not what the underwriter did**.
  **This is exactly why it must be asked, not inferred.**

## 1. Main summary (BT12:CJ30) — one row per group
Row groups: **Total IL** (all "–" when none) · **AL tiers** (one per `Care|Type|Tier`) · **Total AL** ·
**MC room-types** (`Semi-Private` ABOVE `Private` — higher-rent lines sit lower, ruling 2026-07-15) ·
**Total MC** · **Total**. A blank row separates every care group (formatting.md §5). Columns, left→right:

| Col grp | Meaning | Formula (tier row; group `g` = BT{row}) |
|---|---|---|
| Unit Type | the group label | `= "AL \| 1BR \| A"` (matches a Group Unit Code) |
| #Apts | apartments (doors) | **General distinct-apartment count** (engine fix 2026-07-15, `bab12a1`): collapse any companion pair sharing a raw Unit code to 1 door — works even when the pair shares a tier with ordinary singles (Deal A AL 2BR). Semantics: MC Semi-Private (1 door, 2 beds) → 1 · Jack & Jill (2 doors, 2 beds) → 2, *not* halved (rule 19) · singles → 1. **Must tie to the operator's own apartment total.** ⚠ COM: the nested-COUNTIFS formula silently writes as 0 via `.Formula` — use `.FormulaArray`. |
| #Units | beds | `=COUNTIFS($BM,g,$BD,1)` |
| Occ Units | occupied beds | `=COUNTIFS($BM,g,$BE,1)` |
| Occ % | | `=OccUnits/#Units` |
| Min/Avg/Max SqFt | | `MINIFS/AVERAGEIFS/MAXIFS($BB,$BM,g,$BD,1)` |
| Market (All) | avg market, all beds | `=AVERAGEIFS($BG,$BM,g,$BD,1)` |
| Market (Occ) | avg market, occ beds | `=AVERAGEIFS($BG,$BM,g,$BD,1,$BE,1)` |
| In-Place (Occ) | avg actual, occ beds | `=AVERAGEIFS($BH,$BM,g,$BD,1,$BE,1)` |
| (Discount) | In-Place vs Market, **as a %** | `=InPlace(Occ)/Market(Occ) − 1` — a **RATIO, percent-formatted**, NOT a dollar subtraction. (golden: `=+IFERROR(CD/CC−1,0)`.) |
| Care | avg care | per-group `AVERAGEIFS($BI,…)`; **totals use whole-col `SUMIFS($BI,$BA,"AL")/#Units`** |
| NMI Rent | underwritten market rent | **hardcoded BLUE** input per tier. **Set it ≈ In-Place (Occ) × 1.04–1.05, then round to end in `95`** (rule 18 — the golden runs 3.0–5.5%, every value ends in 95). **Blue on the tier/leaf rows ONLY** — Total rows compute it, so they stay black. (An earlier draft said "RED" — wrong; the golden verified blue `16711680`.) |
| Care Adj | | hardcoded (usually 0) |
| Net NMI Rent | | `=NMI Rent + Care Adj` |
| NMI vs In-Place | mark-to-market, **as a %** | `=NetNMI/InPlace(Occ) − 1` — a **RATIO, percent-formatted**, NOT a dollar subtraction. (golden: `=IFERROR(CI/CD−1,"NA")`.) |

⚠ **`(Discount)` and `NMI vs In-Place` are PERCENTAGES, not dollar differences.** They are ratios
(`x/y − 1`) shown with the percent format (`formatting.md` §3). An earlier draft of this table wrote them
as `InPlace − Market` / `NetNMI − InPlace` (dollar subtractions) — **wrong**, and it contradicted the
percent format `formatting.md` already assigned them. Verified against the golden `Aster Ridge.RR`.

**Totals rows** blend the tiers weighted by unit count and add the Rent-Adjustment plug:
`CD24 = SUMPRODUCT(tierAvgs, tierUnits)/totalUnits + BU42`. An older private-repo golden joined MC
tiers on `$BL` (room-type text) rather than `$BM` — same result because MC's BM/BL both key on
Private/Semi-Private; **standardize new builds on `$BM` for every group.**

## 2. Second Residents (BT32:BX38) — IL / AL / MC / Total
**A THREE-COLUMN block**: row 32 carries the headers `IL` | `AL` | `MC`, and every row below has **one
formula per care level**. On an AL/MC-only property the **IL column correctly shows dashes** — not a bug;
read all three columns before judging it (see `formatting.md`).
The paying-2nd-occupant rollup (from `BJ>0`). Rows and formulas (care = `"AL"` etc.):
- **# 2nd Residents** `=COUNTIFS($BA,care,$BJ,">0")`
- **Utilization** `=#2ndRes / #Units(care)`
- **Rent** `=SUMIFS($BJ,$BA,care)` · **Rent / 2nd resident** `=Rent/#2ndRes`
- **Care** `=SUMIFS($BI,$BJ,">0",$BA,care)` · **Care / 2nd resident**
(golden demo: 5 IL / 3 AL / 0 MC — the demo's MC companions are two real beds per semi-private door,
not fee-paying 2nd occupants, so MC correctly shows none. This block is what surfaces the
second-resident spread across a portfolio.)

## 3. Rent Adjustments (BT40:BU43) — manual plugs
`IL / AL / MC Rent Adj`, hardcoded **RED** (usually 0). Feed the `+BU4x` term in the §1 Totals so a
portfolio-level rent override traces to one cell.

## 4. Check (BT45:BU50) — QA block, entire block BLUE font
Each line = the summary figure **minus** an independent recompute straight from the unit block; **every
line must be 0.** Covers Units/Apartments, Occupied Units, Occupied %, Rent, Care. E.g.
`BU49 = CD30*BW30 − SUMIFS($BH,$BD,1,$BE,1)`. Plus **`BV52` = Days in period** (30/28/31), the proration
denominator used by In-Place Rent de-proration (`output-columns.md`).

## 5. Move-in boxes (CL9:CT20) — two trailing leasing windows
Anchor NMI/market rent to what recent move-ins actually signed. Two date-bounded windows (hand-set RED
start/end serials, e.g. box 1 `CL10:CL11`, box 2 `CQ10:CQ11`). Per group row:
`MINIFS/AVERAGEIFS/MAXIFS($BH,$BM,g,$BD,1,$BE,1,$BF,">="&start,$BF,"<="&end)` + a `COUNTIFS` count. Zero
shows as a dash here.

## 6. Reconciliation (AW402:BC409) — the dollar gate
Built sums vs the operator's **own** in-detail wing subtotals (`TOTAL Assisted Living` row + `TOTAL
Memory Care` row), read at the raw columns:

| Line | Built | Operator | Must |
|---|---|---|---|
| Market | `=SUM(BG17:BG388)` | `=VALUE(V{TOTAL_AL})+VALUE(V{TOTAL_MC})` | tie to the cent |
| In-Place Rent | `=SUM(BH…)` | `=VALUE(Z…)+VALUE(Z…)` | ties **only pre-de-proration**; built exceeds operator by the sum of mid-month gross-ups (traceable) |
| In-Place Care | `=SUM(BI…)` | `=VALUE(AI…)+VALUE(AI…)` | tie within rounding |
| 2nd-Occ Fee | `=SUM(BJ…)` | `=VALUE(AG…)+VALUE(AG…)` | tie to the cent |

Market and 2nd-Occ Fee are value-conserving → **any non-zero diff = a dropped/double-counted leaf or a
bad block boundary** (see `validation-and-build.md`). In-Place Rent legitimately diverges by the
de-proration adjustment; that residual must be **explained** (equal to the gross-ups), not forced to 0.
The golden shows In-Place +$2,145.60 — the de-proration delta, left visible.

## Color & format convention → **see `formatting.md` (authoritative)**
Summary only; `formatting.md` has the verified spec and is the gate:
- **BLUE** (`16711680`) = judgment input: **NMI Rent (tier rows only)**, move-in-window dates,
  days-in-period. **ORANGE** (`14745600`) = Rent Adjustment plugs. **RED** (`255`) = the whole Check block.
- The **per-unit block is entirely black** — its hardcodes are not colour-coded.
- **Two header styles:** unit block + tiering helper = **yellow `13434879`, Calibri 11 bold black**;
  analysis summary + move-in boxes = **navy `4990985`, Arial 10 bold white**. Row height 30, wrapped.
- Analysis table uses the **dash** accounting/percent formats and carries **no `$` signs**; a displayed
  `0` is a formatting bug (memory `zero-display-dash`). The per-unit `$` block may show `$0` —
  it is working data.
- **Total rows:** bold, **top border only** (no bottom border).
- ⚠ Excel `Font.Color` is a **BGR long** — `16711680` hex-prints `FF0000` but is **blue**; `255` is
  **red**. Write the longs; don't infer colour from the hex.
