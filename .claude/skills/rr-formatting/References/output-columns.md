# .RR Per-Unit Block (17 cols) — exact derivation

the firm's standard right-side block, **locked to the golden model** `Aster Ridge-RR_v1.xlsx` →
`Aster Ridge.RR`. Each right cell is **row-aligned** to its left unit line, so `=A17`, `=V17` read the
same row. Prefer a **formula linking back** to the hardcoded left; **hardcode** only what a formula can't
reliably express (order, actual SqFt, cleaned move-in, the tier/label codes). Operator numbers are stored
as **text** → always wrap in `VALUE()` with `IFERROR(…,"")`.

⚠ **The column letters below are the golden's geometry, not the contract** (raw export `A:AV` → block at
`AW:BM`, headers row **12**, data rows 17–388). The raw width **varies by operator**: compute
**`blockStart = lastRawCol + 1`** and treat `AW`=`+0` … `BM`=`+16` as **offsets** (rule 15;
`formatting.md` §1). Styling, number formats and the header treatment: **`formatting.md`** — it is a
build gate, read it first.

## The 17 columns (row `a` = the unit's anchor row)

| Col | Header | Kind | Derivation |
|----|--------|------|------------|
| AW | Order | hardcode | sequential `1..N`. Label only — nothing downstream references it. |
| AX | Unit | formula | `=A{a}` — the operator's raw unit code (`103P`, `139-139A`). |
| AY | Unit Type | formula | `=TRIM(LEFT(E{a},FIND(" - ",E{a})-1))` → `AL-1D`. |
| AZ | P/SP | hardcode | `P` = primary bed · `S` = shared 2nd occupant (from the code suffix / operator). Drives the 2nd-occupant handling below. |
| BA | Care Type | formula | `=LEFT(AY{a},2)` → `AL` / `MC`. |
| BB | SqFt | **hardcode** | the unit's **actual** SqFt. The raw floor-plan SqFt (col E) is a flat placeholder, so hand-enter from the unit mix / OM. `0` marks a shared 2nd-occupant ("S") line. |
| BC | Resident Type | reserved | header present, blank in the golden. Leave empty unless a deal needs it. |
| BD | Capacity | formula | `=IF(BB{a}>0,1,0)` — beds. SqFt-0 "S" line → **0**. |
| BE | Occupancy | formula | `=IF(BF{a}="",0,1)` — occupied iff a move-in date is present. |
| BF | Move in Date | **hardcode** | date serial of the **period-end occupant**; blank if vacant/inactive/down (→ occ 0) and on every "S" line. Cleaned/pasted, not linked (raw col N is dirty). |
| BG | Market Rent | formula | single leaf `=IFERROR(VALUE(V{a}),"")`; multi-leaf unit sums its **own** leaves `=IFERROR(VALUE(V{a}),0)+IFERROR(VALUE(V{leaf2}),0)+…`. |
| BH | In-Place Rent | formula | same over col **Z** (Actual). Mid-month move-in **de-prorated** to full month (see below). **Vacant → literal `0`.** |
| BI | In-Place Care | formula | same over col **AI** (Care Fees) — **including the de-proration** (rule 11). **Vacant → literal `0`.** On a P/S pair the "S" row's care **folds onto the primary**: `=(AI{P}+AI{S})*<factor>`. |
| BJ | 2nd Resident | formula | `=IFERROR(VALUE(AG{a}),0)` — the 2nd-occupant **FEE** ($). **`>0` = a paying 2nd resident.** NOT a Y flag, never a name. |
| BK | Unit Code | formula | `=IF(AX{a}="","",BA{a}&" \| "&AY{a}&" \| "&BB{a})` → `AL \| AL-1D \| 545` (composite billing key). |
| BL | Basic Unit Code | **hardcode** | human label → `AL - 1BR (545-646 SF)` (care + bedroom type + SqFt band). MC uses room-type text (`MC - Private`, `MC - Semi-Private`). |
| BM | Group Unit Code | **hardcode** | `AL \| 1BR \| A` — care + type + **tier**. **THIS is the analysis join key** (see `analysis-table.md`). MC: `MC \| Private`, `MC \| Semi-Private` (no SqFt tier). |

Left source columns (OneSite; **verify per source via the recognizer**): A=Unit · E=`Floor Plan - SQ FT`
· I=Resident/`Vacant`/`TOTAL` · N=Move-in · V=Market · Z=Actual · AG=2nd-Occ Fee · AI=Care · AK=Other ·
AN=Credits · AQ=TOTAL.

## The four ID columns — RESOLVED by the golden (was the standing open question)

- **Unit (AX)** = the operator's raw unit code, verbatim (`=A{a}`).
- **Unit Code (BK)** = composite `CareType | UnitType | SqFt` — the model-readable billing key.
- **Basic Unit Code (BL)** = human-readable type + size label.
- **Group Unit Code (BM)** = `Care | Type | Tier` — the **join key** the analysis table groups on.

(The earlier guess — Unit = door digits, Basic Unit Code = suffix-stripped type — is superseded. AX carries
the full raw code, and a separate **Group Unit Code** was added to carry the A/B/C tier.)

## The three special cases (formulas are NOT uniformly dragged — match the case)

1. **Shared 2nd occupant** (P/S pair, one bed — e.g. `103P` + `103S`): the **"S" row is zeroed** — `BB=0`,
   `BF` blank → `BD=0`, `BE=0`, and `BG/BH/BI/BJ=0`. The 2nd resident's revenue is captured as the
   **2nd-occupant fee on the PRIMARY row's `BJ`** (the primary's finance cells sum both rows' leaves, e.g.
   `BG=IFERROR(VALUE(V{P}),0)+IFERROR(VALUE(V{S}),0)`). So `#Units` (count of `BD=1`) is not inflated, and
   the 2nd resident surfaces only via `BJ>0` in the Second Residents block.
2. **Semi-private / companion** (two beds sharing a door — e.g. `139-139A` + `139-139B`): **both rows keep
   `BD=BE=1`** and standard formulas — two units, one apartment. The apartment count is recovered at the
   summary via `#Apts = #Units / 2` (never here). Do **not** confuse this with case 1 (name it right).
3. **Mid-period turnover / mid-month move-in.** Bespoke per unit — driven by `leafRows` and the move-in
   date, never a dragged fill. **Verified against the golden (2026-07-14):**

   - **De-proration applies to BOTH In-Place Rent AND In-Place Care**, same factor:
     ```
     BH = Z{leaf} *$Days/($Days-DAY(IFERROR(DATEVALUE(N{leaf}),N{leaf}))+1)
     BI = AI{leaf}*$Days/($Days-DAY(IFERROR(DATEVALUE(N{leaf}),N{leaf}))+1)
     ```
     `$Days` = the days-in-period input. **Wrap the move-in in `IFERROR(DATEVALUE(…),…)`** — the raw
     column is TEXT and a bare `DAY()` throws on it.
   - **Turnover = the CURRENT occupant's leaf ONLY** — do *not* sum both contracts. Stacking them puts
     two residents on one bed (golden unit 118: correct `$1,095` care vs stacked `$2,190`, against a
     `$4,995` market rent) and **it still reconciles to the cent**, so no downstream gate catches it.
   - **Vacant/unoccupied → `BH = 0` and `BI = 0`** (literal zero). Do not carry the departed resident's
     rent or care.

## Formatting
- Number format on `BG:BJ` = `$#,##0_);($#,##0)`. Gridlines OFF (house standard, every tab).
- `BB`/`BF`/`BL`/`BM` are hardcoded, but in the **per-unit block they are NOT color-coded** — black/
  automatic font, same as the formulas around them. The input-color convention (blue/orange = input,
  red = Check) applies only to the **analysis table's** judgment inputs — see `analysis-table.md`.
