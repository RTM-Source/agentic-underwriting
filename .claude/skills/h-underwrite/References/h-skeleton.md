# `.H` Skeleton — the assembled-tab structural target

> **Golden (styling AND structure/formulas) = `Investments/Data/Transactions/Aster Ridge (Demo)/
> Aster Ridge.H Model_v1.xlsx`.**
> **Re-verified against the golden by COM read-back, 2026-07-22.** A few legacy quirks from older
> conventions are NOT part of the contract and should not be reproduced: red-font `Ok` literals /
> red Marketing metric row (the current golden is plain black), explicit white fills, and stray bold
> in col E.
>
> This is the `master-format.md` analog for the assembled `.H`. The golden's row numbers are the
> **canonical template**; only a handful of things move per deal (see "What changes per deal").
> Read with `rollup-category-map.md` (the fixed roll-up + metrics block) and
> `join-and-census-contract.md` (the cross-tab keys).

## The five tabs in a finished model (naming convention)
`<CN>` = community name (e.g. `Aster Ridge`). A complete model workbook holds, left → right:

| Tab | Built by | Role in `.H` |
|-----|----------|--------------|
| `<CN>.H (Review)` | **this skill** | the deliverable |
| `Census` | `census-formatting` | occupancy block (rows 5–24) links here |
| `<CN> PnL Mapping` | `pnl-mapping` | tag per line (col B XLOOKUP) |
| `<CN> Historical Financials` | `hf-formatting` (the Formatted HF) | values (cols F:Q XLOOKUP) + reconcile anchors |
| `<CN> Unformatted HF` | raw operator source | not referenced by `.H` |

## Column map (the `.H` tab)
- **A** – section/marker text (mostly blank).
- **B** – in detail rows: the mapped **tag** (XLOOKUP into mapping). In roll-up/metric rows: the **display name**.
- **C** – in roll-up/metric rows: the exact **tag key** (the SUMIFS criterion) or a label. Blank in most detail rows.
- **D** – in detail rows: the **`code  -  label` join key** (hardcoded text, mirrors HF col C). In metric rows: the tag key the metric divides.
- **E** – spacer (blank column; width 12.14 like the month columns — per the golden; not narrow).
- **F : Q** – the **12 monthly columns** (F = oldest … Q = newest). *(Q is the 12th month, not a row-total here.)*

Everything is keyed on those 12 month columns F:Q. Gridlines **off** (every output).

---

## Block A — Occupancy (rows 5–24) → links to `Census`

> **Governing principle (Ryan, 2026-07-16): pull what the operator actually MEASURED; derive the rest.**
> The green cell is the operator's own datum; everything downstream of it is a black in-sheet formula.
> Senior-housing operators almost always report **occupancy %**, not unit counts — so the default below
> makes **Occ % the pulled input** and **Occupied the derived value**. This inverts the old skeleton
> (which pulled Occupied and derived Occ %). Arithmetically identical; the point is provenance and
> control — the underwriter sees and flexes the % that actually drives the model.

| Row | Col C | Col D | F:Q formula (shown for col F) | Font |
|-----|-------|-------|-------------------------------|------|
| 5 | | | F5 = first month-end (hardcoded date); **G5:Q5 = `=EOMONTH(F5,1)`** walking right | F5 blue, rest black |
| 7 | | `Days` | `=DAY(EOMONTH(F5,0))` | black |
| 9 | `Capacity` | `IL` | `=Census!<cap-IL-cell>` | **green** |
| 10 | `Capacity` | `AL` | `=Census!<cap-AL-cell>` | **green** |
| 11 | `Capacity` | `MC` | `=Census!<cap-MC-cell>` | **green** |
| 12 | `Capacity` | `Total` | `=SUM(F9:F11)` | black, bold |
| 14 | `Occupied` | `IL` | **`=F19*F9`** (Occ % × Capacity) | black |
| 15 | `Occupied` | `AL` | **`=F20*F10`** | black |
| 16 | `Occupied` | `MC` | **`=F21*F11`** | black |
| 17 | `Occupied` | `Total` | `=SUM(F14:F16)` | black, bold |
| 19 | `Occ %` | `IL` | **`=Census!<occ%-IL-cell>`** — the pulled input | **green** |
| 20 | `Occ %` | `AL` | **`=Census!<occ%-AL-cell>`** | **green** |
| 21 | `Occ %` | `MC` | **`=Census!<occ%-MC-cell>`** | **green** |
| 22 | `Occ %` | `Total` | `=+IFERROR(F17/F12,"NA")` — the capacity-weighted blend, still derived | black |
| 24 | | `PRD` | `=+F17*F7*F<residentsRow>` (Occupied Total × Days × Residents/Occupied-Unit) | black |

No circularity: 19–21 are pulled, 14–16 derive from them, 17 sums, 22 divides the totals.
**Number format on 19–21 = `0%`, right-aligned.**

**Block A total borders (Ryan, 2026-07-16 — easy to miss, they are NOT part of the roll-up border set):**
each of the three groups gets a rule above its `Total` row, drawn on **both** adjacent rows across
**`F:Q` only** (not `S:U`, not `B:E`), thin (`LineStyle=1`, `Weight=2`):

| Rows | Border |
|------|--------|
| 11 / 16 / 21 (last care line: `MC`) | **bottom** thin, F:Q |
| 12 / 17 / 22 (the `Total` rows) | **top** thin, F:Q |

**When the operator reports COUNTS, not %** (Pattern B/C census — explicit `Occupied Units`, or a PCC
cube with ADC/Beds): the same principle flips which cell is green. The Census tab still exposes an occ%
row (derived there from the counts), and `.H` Block A can keep this identical shape — `=Occ% × Capacity`
round-trips to the operator's own count exactly. Prefer keeping the shape uniform across deals.

The exact `Census!` cells are **resolved per deal** — the Census output block can be anchored at
any column (the golden lands it at column R/U). See `join-and-census-contract.md` for how to locate
Capacity/Occupied IL/AL/MC and the first month column on the Census tab and align the 12-month window.

---

## Block B — Detail passthrough (rows 26–217) → mirrors the Formatted HF
`REVENUE` banner at row 26; first line at row 28. **Each `.H` detail row mirrors one Formatted-HF
row**, in the same order, including the HF's own section headers and totals (`TOTAL RENT REVENUE`,
`NET OPERATING INCOME`, `Margin`, …). The HF lines used here are rows **`hfFirst` → `hfLast`** —
**locate `hfFirst` on the HF, don't assume the golden's 8**: it is the first VALUE row after the
`REVENUE` banner (the golden: 8→197; Deal C: **7**→213 — its first line, Gross Mkt Rent Potential,
sat at row 7 and skipping it made both Checks off by a material amount every month). `.H` row 28 = HF `hfFirst`
(the golden: `.H detailRow = hfRow + 20`). Per row `r`:

- **D{r}** = literal join-key string, copied **whitespace-exact** from HF col C (`=IF(S="",T,S&"  -  "&T)` result).
  This is the only hardcoded content of the block — everything else is a formula.
- **B{r}** (tag from mapping):
  ```
  =LET(r,XLOOKUP(D{r},'<CN> PnL Mapping'!$C$3:$C$<mapLast>,'<CN> PnL Mapping'!$B$3:$B$<mapLast>,""),IF(OR(r="",r=0),"",r))
  ```
- **F{r}:Q{r}** (values from HF), walking the HF value columns E,F,G,…,P for the 12 months:
  ```
  F{r} =XLOOKUP($D{r},'<CN> Historical Financials'!$C$8:$C$<hfLast>,'<CN> Historical Financials'!E$8:E$<hfLast>)
  G{r} = …same, value col F …   …   Q{r} = …value col P …
  ```
  (HF value cols E:P map to `.H` month cols F:Q — a one-column shift.)

Rows that are HF section headers/totals still carry their D string and pull through the same way
(their HF total cell resolves the value). Blank HF rows → blank `.H` rows.

**Do NOT carry the HF's `Margin` row into the detail block** *(current convention; an older golden kept
it as its own detail row at row 217)*. Margin is a **ratio, not a dollar line** — it has no master tag,
contributes nothing to the SUMIFS roll-up, and has to be special-cased out of every S/T/U window sum.
The `.H` states margin once, in Block C's own EBITDAR-margin row. Ryan deleted the Margin passthrough
by hand on an early build; skip it at build time. Skipping it shortens the block by one and shifts `detailLast` (and everything below) up by 1 —
which is fine, because those are all derived.
**When Margin is MID-BLOCK (below-the-line detail follows NOI/Margin, not just the end)** *(Deal D,
2026-07-23 — the first such deal; the golden had Margin as the last line)*: the operator's
Formatted HF can carry below-the-line rows (depreciation/interest/owner-allocated, GL-coded) AFTER the
Margin row. Mirror those rows too (they map to `Ignore` → excluded from the roll-up + both Checks), and
skip-and-**compact**: detail `.H` row = HF row + N above the Margin, HF row + (N−1) below it. Don't just
stop the detail at NOI, and don't leave a gap where Margin was — the below-the-line block must sit flush
under NOI so `detailLast` and the roll-up placement stay correct. (Worked-example script in the
private team repo; not included in this public clone.)

---

## Block C — Roll-up by tag (rows 221–296) — FIXED skeleton

**Placement (hard rule): the roll-up's first row = `detailLast + 3`.** Verified across every built model
(the golden 218→221, Deal B 236→239, Deal F 240→243, Deal A 347→350). `detailLast` is the SUMIFS range end,
which sits one row past the last XLOOKUP line — that trailing blank is intentional and harmless.
**Derive this offset from `detailLast`, never carry a hardcoded row number** from a previous deal — every
row below (Checks, metrics, Block D) shifts with it. See memory `h-skeleton-rows-not-fixed`.

The standardized house category block. Full ordered list, every B-display ↔ C-tag-key pair, and the
subtotal rows are in **`rollup-category-map.md`**. Each category row:
```
F{r} =+SUMIFS(F$25:F$<detailLast>,$B$25:$B$<detailLast>,$C{r})
```
(`detailLast` = last detail row, golden 218 — one past the last line, harmless.) Subtotals use `=SUM(...)`
over their member rows; `EBITDARM=Total Revenue−Total Opex`; `EBITDAR=EBITDARM−Management Fee`.

**Reconciliation Checks (hard gate — must be 0 every month):**
- Row 234 `Check` (rev): `=F232-'<CN> Historical Financials'!E<hfRevRow>` (golden E30 = HF Total Revenue).
- Row 300 `Check` (NOI): `=F296-'<CN> Historical Financials'!E<hfNoiRow>` (golden E196 = HF NOI/NET OPERATING INCOME).
- `hfRevRow`/`hfNoiRow` are located by label on the HF, not hardcoded.

---

## Block D — Metrics (rows 302–336) — FIXED skeleton
Denominators (302–308): `Units=F12`, `Occupied Units=F17`, `Residents/Occupied Units=1.00` (fmt `0.00"x"`,
G=`=+F304`), `Resident Days=F24`, `Total Labor=F254`, `Total Benefits=SUM(F256:F262)`, `Total Revenue=F232`.
Per-line metrics (310–336) are `SUMIFS` over the **roll-up** block (`$221:$291`) divided by a denominator
(`/F$302` $/unit, `/F$305` $/PRD, `/F308` %rev, `/F306` %labor, `/1000` $K/property). Captured verbatim in
`rollup-category-map.md` §Metrics — treat as a fixed template (a few rows add specific detail cells, e.g.
Marketing adds contract+referral, Admin adds bad-debts/COVID/legal).

---

## Block E — Totals columns S/T/U (T12 / T6 / T3) — REQUIRED (missed on an early build)
Three windowed-totals columns to the right of the months, all verbatim from the golden. Rows below are the
canonical golden rows — apply the same per-deal offset as Blocks C/D.

**Header (rows 2–5):** `S2/T2/U2 = 1 / 2 / 4` (annualization multipliers); `S3/T3/U3 = T12 / T6 / T3`;
`S4 ==F5`, `T4 ==+EOMONTH(S4,6)`, `U4 ==+EOMONTH(T4,3)` (window starts); `S5 ==Q5`, `T5 ==+S5`,
`U5 ==+T5` (window end = newest month). Format `m/d/yyyy`.

**Header styling (Ryan, 2026-07-16 — applies to the whole `S2:U5` block):**
- **Every cell in `S2:U5` is aligned RIGHT** (`xlRight`, `-4152`). The golden files are inconsistent here
  (`S4`, `S2:U2`, `T5`, `U5` carry no alignment) — make the whole block right, no exceptions.
- **The multipliers `S2:U2` (`1` / `2` / `4`) are font BLUE `16711680`.** They are hardcoded inputs an
  underwriter may retune, so the standing font-colour rule governs. (Deal A already had this; the older
  golden has them black and is wrong.) Rows 3–5 keep their own semantics: `S4`/`S5` are black in-sheet formulas.

**Per row type** (write S, T, U each; `{c}` = the column letter):
- **Flow rows** — Days 7, PRD 24, every value-bearing detail row (skip blanks; there is no `Margin`
  row to special-case any more — Block B drops it), every roll-up category row incl. Management Fee:
  `=IFERROR(SUMIFS($E{r}:$Q{r},$E$5:$Q$5,">="&{c}$4,$E$5:$Q$5,"<="&{c}$5),0)*{c}$2`
- **Occupancy stock rows 9–12 / 14–17:** `=+IFERROR(AVERAGEIFS($E{r}:$Q{r},$E$5:$Q$5,">="&{c}$4,$E$5:$Q$5,"<="&{c}$5),"NA")`
- **Occ % 19–22:** `=+IFERROR({c}14/{c}9,"NA")` etc. **Subtotals/derived/denominators/metrics:** same
  formula as col F translated to the S/T/U letters.
- **Checks (S only — T/U stay blank, number-format only):** `S<chkRev> = S<totRev> - '<CN> Historical
  Financials'!Q<hfRevRow>` — ties to the HF's **annual-total column Q** (the Formatted HF must have one).

**R column (width 30.43):** metric-validation on the metric rows only:
`=+XLOOKUP(D{r},$C$<laborFirst>:$C$<otherTaxes>,...)` (golden `$C$236:$C$291`); rows whose tag has no
roll-up home in that span (Benefits & PR Taxes, Management Fee, Capex Reserve — golden 312/333/335) get
the literal `Ok` in **plain black** *(current golden, 2026-07-22; the old red `255` was a legacy
convention — do not paint it)*.

## Styling contract (COM colors; re-verified against the golden 2026-07-22)
**Font-color semantics (the governing rule — apply per CELL, not per row):**
- **Green `32768`** = value pulled from ANOTHER TAB (`=Census!…` links).
- **Blue `16711680`** = HARDCODED input the underwriter may edit (F5 date, the 1.00x residents cell).
- **Black `0`** = formula computed WITHIN the sheet (EOMONTH walks, SUMs, Occ %, PRD, SUMIFS…).
Golden-file exception to note: the detail-block F:Q XLOOKUPs into the HF tab stay **black** — they are
a passthrough mirror of the HF, not a lookup the underwriter reads as a cross-tab pull. (The col-B tag
cells ARE green; see below. That rule changed on an early build — the older blue-on-cream B cells are obsolete.)

- **Month header row 5:** all bold, fill gray `13882323`, `m/d/yyyy`, thin bottom border across F:U.
  **Only F5 (the hardcoded date) is blue; G5:Q5 EOMONTH formulas are black** (do not paint the whole
  band blue — that mistake shipped on an early build and Ryan corrected it in-file).
- **Census-linked cells F9:Q11 + F19:Q21: font green `32768`** (cross-tab pull — Capacity and Occ %,
  per the 2026-07-16 pull-%/derive-Occupied inversion). **F14:Q16 (derived Occupied) are BLACK** —
  the old "F14:Q16 green" wording predated the inversion; confirmed black in the golden.
  Occupancy total rows 12/17 bold, black (in-sheet SUM).
- **Detail col B tag cells — font green `32768`, Calibri 9, NO fill, NOT bold.** *(current convention;
  supersedes an older blue `16711680` / Aptos Narrow 9 / cream `13369343` convention.)* Clear bold
  **explicitly** — some builds bold the whole HF-total row and catch the tag cell with it (one early
  build had 18 such cells). The B tag is an XLOOKUP into the **Mapping tab**, so
  the green cross-tab rule governs; it is a formula, **not** an input, so it takes **no cream fill**.
  Apply to **every B cell that CARRIES the tag formula**, whether or not it resolves to a tag — a blank
  result is a mapping gap to see, not a cell to restyle. **Never paint B cells that hold no formula**
  (that is the band-painting mistake — memory `font-color-semantics`).
- **Detail total rows** (col D matches `^TOTAL ` or `NET OPERATING INCOME`): **bold + thin top AND bottom
  border across `D:Q` and `S:U`** — leave A/B/C and R untouched. *(current convention; supersedes the old
  "label bold, values plain, no borders".)* Section **banners** (`REVENUE`, `EXPENSES`) get **D bold only**, no border.
- **Check rows:** fill green `4697456`, fmt `#,##0.00`, not bold — on F:Q **and S** (T/U format only).
- **Residents/Occupied Units input cell (F of that row):** orange fill `49407`, fmt `#,##0.00x`; the
  rest of the row carries `=+<prev>` and no fill.
- **Roll-up bolds:** values bold ONLY on Total Opex / EBITDARM / EBITDAR; Total Revenue, labor
  subtotals, Margin, Checks have bold **labels** (col C/B) but plain values.
- **Borders (F:U):** Total Revenue top+bottom thin; In-House Labor / Total Labor / Total L&B top thin;
  Total Opex top+bottom thin; EBITDARM and EBITDAR top thin + **double bottom** (`-4119`).
- **Number formats:** accounting `_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)` everywhere $-valued
  (zeros render "-"); `#,##0_);(#,##0)` on Total Opex/EBITDARM/Units/denominators;
  `0%` on Occ % rows 19–21; `#,##0.0%_);(#,##0.0%)` on the roll-up Margin row; `0.00%` on Block-D
  metric %-rows (verified in the golden on Benefits-%-Labor + Management Fee — the old `#,##0.0%`/
  `#,##0.00%` variants there were from a legacy convention); Marketing metric row values plain black
  *(current golden; the old red `255` is gone)*.
- **Widths:** A 8.57 · B 25 · C 19 · D 36.57 · E:Q 12.14 · R 30.57 · S:U 11 (as read from the golden;
  ±0.14 vs older docs is font-metric rounding — don't chase it). Rows 1–4 otherwise empty
  (no community-name cell — B2 is blank in the golden).

## What changes per deal (everything else is fixed template)
1. **`<CN>`** in all sheet names.
2. **`hfLast`** (last HF row) → sets detail-block length → cascades into `detailLast` and the roll-up SUMIFS ranges.
3. **`mapLast`** (last mapping row).
4. **`hfRevRow` / `hfNoiRow`** reconcile anchors (locate by label on the HF).
5. **First month F5** + the **Census cell map** (Block A links) — from the Census tab's actual anchor.

Build it as a **template parameterized on those five**, not a bespoke script per deal — that is what makes
the one-shot assembly reliable.

---

## Engine clarifications (2026-07-28 — resolved from the golden during the Invoke-HBuild build; verified by 0-diff reproduction of both the golden and the regression build)

The engine (`Investments/lib/H-Assembly-Lib.ps1`, contract in `../scripts/ENGINE-SPEC.md`) encodes these;
listed here so a manual build or audit does not re-stumble on them:
1. Occupancy TOTAL rows 12/17 use `AVERAGEIFS` in S/T/U **without** the `IFERROR` wrapper the member
   rows (9-11/14-16) carry.
2. The HF's own internal `REVENUE`/`EXPENSES` banner rows mid-detail still get full D/B/F formulas —
   only a truly-blank HF cell yields a blank `.H` row; banners get bold D only, no border.
3. Total Opex / EBITDARM / EBITDAR use plain `#,##0`, not the accounting-dash format.
4. Roll-up total-row borders span **F:U** (including R) — unlike Block A's F:Q-only rule.
5. The Residents / Occupied-Units input cell stays **black** despite the general blue-input rule
   (golden deviation, reproduced as-is).
6. Metric-row R-column number format (1- vs 2-decimal percent) follows the metric KIND, independent of
   whether R holds the `Ok` literal or a live XLOOKUP.
7. The single blank row between Total Opex and EBITDARM stays `General` format while sibling blank rows
   in the block carry accounting format (golden quirk, reproduced as-is).
