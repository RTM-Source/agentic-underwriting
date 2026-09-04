# .RR Locked Rules (firm rulings — 2026-06-29)

These were ruled on by the underwriter. They are settled defaults, not flags. Encode them in the
reader/emitter. The *why* matters more than the letter — read the reasoning so edge cases resolve sensibly.

## 1. Granularity: one row per UNIT, where **unit = bed = contract**
- Think of an **apartment** as a physical door, and a **unit** as a physical contract (a bed).
- An apartment can be shared: **1 apartment with 2 beds = 2 units (2 rows).**
- So the table is **line-level** — one row per billed occupant line — never consolidated to apartments.
- Do **not** use apartment/door counts for anything in the .RR.

## 2. Capacity = BEDS (never apartments) — key on STRUCTURE, not on SqFt *(refined 2026-07-14)*
**The question capacity answers is: "could someone move into this slot?"** Not "does this row show a
number in the SqFt column."

- **A second resident has NO room of their own.** They are part of the **primary resident's** room. Their
  slot exists only as a dependency of that primary — **nobody can ever move into it**, because there is no
  additional room to move into. So a second-resident line ("S" suffix, e.g. `104S`) →
  **capacity 0, occupancy 0**, whole row zeroed, revenue folded onto the primary (rules 3, 5, 17).
  This is *why* the export usually prints SqFt 0 on those lines — the 0 is a **symptom** of the structure,
  not the cause.
- ⚠ **Do NOT use `Capacity = IF(SqFt>0,1,0)` as the test.** Exports lie: a real export's `127S` line carries
  a real floor-plan SqFt of **540**, and the SqFt rule turned it into a **phantom bed** (AL showed 31 units
  vs the operator's 30 — and every dollar gate still passed). **Key on the second-resident structure**: the `S`
  suffix / the operator's own second-occupant marker / the line being tied to a primary. SqFt is at best a
  corroborating hint.
- A real bed (its own slot someone could occupy) → **capacity 1**.
- A semi-private / companion suite with **two separate beds** (`139-139A` / `139-139B`) → **capacity 1
  each** (two units, one door).
- The written formula may still read `=IF(SqFt>0,1,0)` **only when SqFt has been correctly hand-entered as
  0 on every second-resident row** (rule 10). The *determination* is structural; the formula just reflects it.

## 3. Occupancy = period-end snapshot (occupied bed = 1, otherwise 0)
- Currently occupied by a resident → **1**.
- **Vacant** → **0** (capacity stays 1 — the bed still exists).
- **Inactive / "Down unit"** → **occupancy 0, capacity 1.** This is *not* a third state — it is just an
  unoccupied real bed. (Example: a primary resident moved out and the line is now an inactive-down unit →
  capacity 1, occ 0.)
- **Second occupant present** (the "S" line) → **occupancy 0, capacity 0.** *(REVISED per golden model
  v1.)* They occupy no separate bed, so they contribute 0 to both bed capacity and bed occupancy; the
  whole "S" row is zeroed and their revenue folds onto the primary as the 2nd-occupant fee (rule 5). This
  keeps `occupied beds ≤ capacity beds`. In practice occupancy is driven by the move-in date:
  `=IF(MoveIn="",0,1)`, and the "S" row's move-in is left blank.
- Determine the period-end state from the **last status line** in the unit's block; vacant / inactive /
  down all resolve to unoccupied.

## 4. No couples — treat every resident as an individual
- Do **not** assume two resident lines are a couple, and do **not** merge them into one apartment row.
- Each resident is their own line/contract.
- A within-month **turnover on one bed** (resident A out, resident B in, same unit, with a `NNN-NNN TOTAL`)
  is one bed → collapses to that bed's **period-end occupant** (the current resident).

## 5. "2nd Resident" column = the 2nd-occupant FEE, not a flag *(REVISED per golden model v1)*
- The column holds the **2nd-occupant fee ($)**: `=IFERROR(VALUE(AG{a}),0)` (raw 2nd-Occ Fee column).
  **`>0` marks a paying second resident.** Never a `Y`, never a name.
- On a shared "S" pair, the fee lands on the **primary** row (its finance cells sum both leaves); the "S"
  row is zeroed (rule 3). The Second Residents analysis block counts/sums where this column `>0`.
- **Distinguish the two shared structures:** a P/S pair (one bed, 2nd occupant → fee on primary, "S" row
  zeroed) vs. a semi-private/companion pair (two beds, `139-139A`/`-139B` → **both** capacity 1 &
  occupancy 1, `#Apts = #Units/2` only at the summary). The fee column is for the former, not the latter.

## 6. Care Type from the unit-type prefix
- `Care Type = LEFT(UnitType, 2)` → `AL` / `MC`.
- These communities are **AL/MC only — there are no IL units**, so IL being untested is fine. If an IL
  community ever appears, revisit.

## 7. Finance values reconcile via leaf lines, not the operator's subtotal cell
- Market / In-Place Rent / In-Place Care link back with `VALUE()` (operator stores numbers as text).
- A unit's value = the **sum of that unit's own leaf lines** (its detail rows), **never** the operator's
  `NNN TOTAL` subtotal cell — that subtotal can roll up *sibling* beds (companion A/B, P/S) and
  double-count. Summing own leaves ties to the cent against the operator's section totals. (This is the
  bug that produced an over-count in the first build; see `validation-and-build.md` for the numbers.)

---

# Rulings settled by the golden model (v1 — 2026-07-13)

`Aster Ridge-RR_v1.xlsx` → `Aster Ridge.RR` is the golden standard (in the private team repo the golden
is a real-deal build; this public clone substitutes the fabricated demo deal). It resolves the
previously-open questions.

## 8. The four ID columns (was OPEN)
- **Unit** = the operator's raw unit code (`=A{a}`). **Unit Code** = composite `Care | Type | SqFt`.
  **Basic Unit Code** = human label (`AL - 1BR (545-646 SF)`). **Group Unit Code** = `Care | Type | Tier`.
- **Group Unit Code is the analysis join key** — every summary COUNTIFS/AVERAGEIFS groups on it. (Not
  "Basic Unit Code = Unit Type via BN" — that framing was wrong; no BN column exists.) See `output-columns.md`.

## 9. Tiering: AL/IL group by SqFt tier (A/B/C), MC by room-type — **TWO-PASS, then ASK** *(ruling 2026-07-14)*
Tiering is **two competing objectives**, resolved in this order:

1. **Pass 1 — cluster by SqFt similarity.** Group the distinct SqFt values into natural bands.
2. **Pass 2 — rebalance by unit COUNT.** Aim for **roughly equal unit weight** per tier (30 AL units →
   ~10 / 10 / 10). A tier holding 2 units next to one holding 28 is a bad split even if the SqFt says so.
3. **The frequency override beats both.** A single SqFt carrying a **large share of the units gets its own
   tier**, however close its neighbours are. Worked illustration (hypothetical spread roster — the demo
   golden's types are single-SqFt, trivially all-A): `545` holding **12 of 30 AL units** → tier **A
   alone**, even though `556` is only **11 SF** away. Conversely **don't split a small group**: all 2BR
   (860/905/918, a 58 SF spread, few units) sit in one tier.

⚠ **"Natural breaks" alone is WRONG and will not reproduce an underwriter's assignment** — it merges the
modal floorplan into its neighbours and over-splits the tails. Unit count is the stronger signal.

**The first pass is expected to be wrong, and that is fine.** Tiering is *judgment*, and it is weighted
**far less** than the dollar ties in the analysis table — a mis-tiered build still reconciles and still
ties every headline number. So: **PROPOSE a tiering with your reasoning and the unit counts, and ASK the
underwriter to confirm.** Never silently infer it (`analysis-table.md` §0).

- **MC does not use SqFt tiers** — it groups by **room structure** (rule 19).
- **Tier tokens are always `A` / `B` / `C`** — never the SqFt value. (An early build labeled its single 2BR
  group `AL | 2BR | 918` because there was exactly one 2BR; that was a one-off convenience, **not** the
  standard. Standardize on `A`.)
- **Don't over-weight the tiering.** Get the *methodology* right and the groups **defendable**; the
  underwriter may regroup, and that is expected. It is weighted far below the dollar ties.

## 19. MC room structures — name by the ACTUAL door/bed configuration *(ruling 2026-07-14)*
Never classify by analogy to another deal. Count doors and beds, then name it:

| Structure | Doors / Beds | `#Apts` | Label |
|---|---|---|---|
| **Private** | 1 door, 1 bed | `= #Units` | `MC - Private` |
| **Semi-Private** | **1 door**, 2 beds | **`= #Units / 2`** | `MC - Semi-Private` |
| **Jack & Jill** | **2 doors**, 2 beds (separate rooms sharing a **bathroom**) | **`= #Units`** | `MC - Jack & Jill` |

- **Jack & Jill is "shared" only in the bathroom sense — it is TWO apartments.** Each resident has their own
  door and their own room. So `#Apts = #Units`, exactly like Private.
- ⚠ **This is the trap.** A Jack & Jill *looks* like a shared MC unit, and applying the Semi-Private
  `#Units/2` rule by analogy **understates apartments — and every dollar gate still passes.** Example: an
  8-bed Jack & Jill wing = **8 apartments**, not 4. The operator's own apartment total is the tie-breaker;
  use it.
- The operator's floor-plan code usually tells you (`MC-0JJ` = Jack & Jill, `MC-0P` = Private, `MC-0S` =
  Semi-Private) — but **confirm against door/bed reality**, don't trust the code alone.

## 20. SURFACE THE ANOMALIES — the flag list is the value-add *(ruling 2026-07-14)*
Rent rolls vary endlessly; **catching the weird stuff is the point of the skill**, not a side effect. After
the build, **append a Notes block at the bottom of the RR analysis** listing every anomaly, each with its
**unit and row reference** so the underwriter can click straight to it, then **ASK for a ruling before
finalising**. Phrase it plainly, e.g.:

> *"Noticed some odd second-resident activity on line 42 (unit 118) — In-Place $8,240 vs Market $4,995
> (+65%), single operator leaf. Linked here for your ruling before I update."*

What must be surfaced (non-exhaustive): In-Place wildly above Market on a unit · second-resident /
"S"-line oddities (a real SqFt on an S row, an S row with its own market rent) · inactive-down and
companion pairs · turnovers and mid-month move-ins · any unit that forced a judgment call · **any count
that disagrees with the operator's own summary totals**. **Never silently "fix" an anomaly** — flag it,
link it, ask.

## 21. Artifact hygiene — build the deliverable IN PLACE, visibly *(ruling 2026-07-14, revised)*
**The workbook is the product,** and the underwriter should be able to **watch it get built** in the deal
folder — an invisible `%TEMP%` build reads as "nothing is happening" to anyone not fluent in the tooling.
Everything needed to trust the file lives *inside* it (the Notes block, rule 20) or in the agent's return
message — never in sidecar files.

**The build-in-place flow (verified 2026-07-14 — `Save()` in place on OneDrive does NOT hang, only
`SaveAs` to a new OneDrive path does):**
1. **`Copy-Item`** the operator original → the deal folder under a **building marker name**:
   **`~building — <Community>—RR_v#.xlsx`**. This is a filesystem copy (~15 ms, no COM) — it does NOT
   write the operator's original, and the deliverable appears in the deal folder **immediately**.
2. **`Open`** that `~building —` file from the deal folder, build the whole tab, and **`$wb.Save()`** in
   place — **`Save()`, never `SaveAs`** (SaveAs to a OneDrive path hangs; Save() on an already-OneDrive-open
   workbook returns in ~1 s). The underwriter sees the file's size/modified-time tick as it builds.
3. **On success:** `Rename-Item` `~building — …` → the final `<Community>—RR_v#.xlsx` (instant, atomic).
   The final name appears **only when the build is complete and gated**.
4. **On ANY failure (`try/finally`): delete the `~building —` partial** so the deal folder is left clean.
   A hard-kill can't be trapped — but then only an obviously-unfinished `~building —` file remains, which
   the marker name flags as do-not-use.

**The `~building —` prefix is a deliberate DO-NOT-OPEN signal** — a file by that name is mid-build; opening
it in Excel takes a lock and breaks the `Save()`. It never survives a successful run.

- **Scaffolding** (any helper scripts, hashes, `%TEMP%` copies) stays in **`%TEMP%`**, never the deal folder.
- **Deal folder after a successful run:** the operator original, the new `<Community>—RR_v#`, and whatever
  was already there — nothing else. Superseded `_v#` → `arc/`.

⚠ **Housekeeping is a tidy-up, NOT a gate.** A stray file never fails a correct build. Hard rail:

> **Delete ONLY files this run created** (its own `~building —` partial, its `%TEMP%` scaffolding). Never
> the raw operator export, never a finished workbook, never anything that predates the run or that you did
> not write. Unsure? Leave it and say so.

`GAPS.md` is a **stress-test instrument only** (a doc-quality probe, not a production concept) — when
requested at all, it belongs in the scratchpad, not the deal folder.

## 22. The golden model is a REFERENCE to check against, not a secret *(ruling 2026-07-14)*
`Investments/Data/Transactions/Aster Ridge (Demo)/Rent Roll/Aster Ridge-RR_v1.xlsx` is the **golden
standard** — the formatting contract and structural layout made concrete. On a **real/new deal, OPEN the
golden and cross-check your finished tab against it** before declaring done:

- **Every block the golden has, yours has** — unit block, tiering helper, main summary (with **headers**),
  Second Residents, Rent Adjustments, Check, both move-in boxes, reconciliation. (Deal C's v1 build
  shipped the main summary and tiering helper **headerless** — a golden cross-check catches that in one
  look; hiding the golden is what let it ship.)
- **Structure, headers, number formats, fonts/fills, borders, the row-plan shape** all match the golden's
  pattern (re-anchored to this deal's geometry and care-mix — an IL-only deal populates the IL rows the
  golden leaves blank, etc.).
- Use it as a **checklist, not a copy source** — the *numbers* come from this deal; the *shape* matches
  the golden.

**The ONE exception — the cold stress test.** When the task is to *rebuild the golden's own source deal*
to test whether the skill can reproduce it from raw, the golden is that deal's own answer, so it is
**off-limits for that test only** (checking against yourself is circular). For every other property the
golden is an expected, encouraged reference.

## 18. NMI Rent = In-Place + 4–5%, always ending in 95 *(ruling 2026-07-14)*
- **NMI (new-move-in) Rent** = the underwritten market-anchored rent per tier. Set it **~4–5% above that
  tier's In-Place (Occ)**, then **round so the figure ends in `95`** (`5,995` · `6,295` · `7,295`).
- Verified across v1: premiums run **3.0%–5.5%** (clustered 4–5%) and **every NMI ends in 95**.
- It stays a **BLUE hardcoded input on the tier/leaf rows only** — Totals compute it, so they are black
  (rule 13). Propose the 4–5% figure; the underwriter may override per tier.

## 10. SqFt and Move-in are hardcoded inputs, not formulas
- The raw floor-plan SqFt (col E) is a flat placeholder and raw move-in (col N) is dirty, so **SqFt (BB)
  and Move-in Date (BF) are hand-entered** (from the unit mix / OM / cleaned dates), and everything
  downstream (Capacity `=IF(BB>0,1,0)`, Occupancy `=IF(BF="",0,1)`) keys off them.
- **They are NOT color-coded** — black/automatic font, like every other cell in the per-unit block. Only
  the *analysis table's* judgment inputs get blue/orange, and red is the Check block (rule 13). (An earlier
  draft of this rule said "RED-font inputs" — wrong, and it contradicted rule 13; v1 verified black.)

## 11. De-proration: gross partial-month move-ins to a full month — **RENT *AND* CARE**
- A mid-month move-in's In-Place figures are **de-prorated to full monthly** using a days-in-period input.
- ⚠ **The SAME factor applies to In-Place CARE, not just Rent.** (Verified in v1, 2026-07-14 — an earlier
  draft of this rule said Rent only, and a cold build that de-prorated Rent alone left Care short.)
  ```
  In-Place Rent:  =Z{leaf} *$Days/($Days − DAY(moveIn) + 1)
  In-Place Care:  =AI{leaf}*$Days/($Days − DAY(moveIn) + 1)      <-- same factor, do not omit
  ```
- **Parse the move-in defensively** — the raw column is TEXT in these exports, so v1 wraps it:
  `DAY(IFERROR(DATEVALUE(N{leaf}),N{leaf}))`. A bare `DAY(N{leaf})` throws on the text form.
- So built In-Place **exceeds** the operator's (prorated) totals by the sum of these gross-ups — an
  **expected, traceable** residual, not a break.

## 12. Reconciliation is refined, not "everything to zero"
- **Market and 2nd-Occ Fee tie to the cent** (value-conserving — a non-zero diff is a real bug).
- **In-Place Rent** ties pre-de-proration; the post-de-proration delta must equal the rule-11 gross-ups
  and be left **visible + explained** (v1: +$2,145.60).
- **In-Place Care does NOT tie exactly either** — it is grossed up by the same rule-11 factor, and vacant
  units are zeroed (rule 17). Its residual must likewise be *explained*, not forced to zero. (An earlier
  draft said "Care ties within rounding" — wrong, and it would send a builder hunting a phantom bug.)

## 17. Vacancy zeroes the money; the "S" row folds CARE too *(v1-verified 2026-07-14)*
- **Unoccupied unit → `In-Place Rent = 0` and `In-Place Care = 0`** (v1 hardcodes a literal `0`, not a
  blank and not the departed resident's stale figure). Market Rent still populates — the bed exists.
  A build that carries the prior occupant's care on a vacant unit overstates care.
- **P/S pair:** the shared "S" row's **In-Place Care folds onto the primary** exactly like the 2nd-occupant
  fee — the primary sums *both* leaves, then de-prorates:
  `BI = (AI{primary} + AI{S}) * <de-proration factor>`. Rule 5 only said the *fee* folds; **care folds too.**
- **Turnover (two contracts on one bed):** take the **current occupant's leaf only** for Rent *and* Care.
  Summing both contracts stacks them (v1 unit 118: correct $1,095 vs stacked $2,190) — and it still
  reconciles, so nothing catches it downstream.

## 13. Color convention: BLUE/ORANGE = analysis input, RED = Check
- In the **analysis table**: **BLUE font** (`16711680`) = judgment inputs (NMI Rent, move-in-window
  dates, days-in-period); **ORANGE** (`14745600`) = Rent Adjustment inputs; **RED** (`255`) = the
  Check/QA block. Header = white-bold on dark fill.
- In the **per-unit block**: hardcodes (SqFt, move-in, Basic/Group code) are **black/automatic**, not
  color-coded.
- ⚠ Excel `Font.Color` is a **BGR long** — `16711680` prints as hex `FF0000` but is **blue**; `255` is
  **red**. Write the longs; a red↔blue mix-up is easy (it bit the first structural read of v1).

## 14. Scope: the skill emits the analysis table too
- The deliverable is **not** just the unit block + reconcile — it is the **full `.RR` tab**: the 17-col
  unit block **+** the analysis table (tiering helper, main summary, Second Residents, Rent Adjustments,
  Check, move-in boxes) **+** reconciliation. See `analysis-table.md`.
- **Capacity = beds visible in the rent roll** (`SqFt>0`); no licensed-bed overlay lives in `.RR` (that
  reconciliation is a model-layer concern, done downstream — not here).

## 15. Geometry is RELATIVE — the raw export's width is arbitrary *(2026-07-14)*
- The raw operator report occupies `A:<lastRawCol>`, and **`lastRawCol` varies by operator and export**.
  In v1 it lands on `AV`, which is why our block starts at `AW` — **a coincidence of that file, not the
  contract.**
- **Compute `blockStart = lastRawCol + 1`** and express every column as an **offset** from it, where
  **`lastRawCol` = the last column of the raw sheet's `UsedRange`** — *not* the last cell holding a value.
  The report owns its trailing formatting. (April-2026 O&O Community 1: last populated col `AQ`=43, but
  `UsedRange` ends at `AV`=48 → `blockStart = AW`=49, matching the golden exactly. The naive reading
  starts 6 columns early, inside the operator's own report.)
- Likewise rows: our header row = **the raw report's own header row** (the SAME row, titles beside
  titles — ruling 2026-07-15, Deal A v3, supersedes the earlier −1 rule; the v1 golden predates this
  and sits at −1); unit rows are row-aligned to the raw anchors; the analysis blocks run down from the
  header on their own plan.
- **A build with a hardcoded `AW` is wrong on the next deal.** Offsets: `formatting.md` §1.

## 16. Formatting is a GATE, not polish *(2026-07-14)*
- The `.RR` is an investor-facing deliverable. **Right numbers in a wrong-looking tab is a failed build.**
- The visual contract is `formatting.md`, verified cell-by-cell against v1. Its gate runs on every build
  alongside the reconcile. The high-frequency failures: mixing the two header styles (yellow/Calibri unit
  block vs navy/Arial analysis), `$` signs leaking into the analysis table, a displayed `0` where the
  zero-as-dash format belongs, and total rows bordered top *and* bottom instead of top only.
