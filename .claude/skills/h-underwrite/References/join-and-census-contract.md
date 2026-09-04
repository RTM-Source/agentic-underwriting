# Join-key + Census-cell contract — the single biggest source of silent breakage

> The `.H` is three keyed joins. If a key is off by one character of whitespace, or the Census anchor
> is read positionally instead of by label, the build looks fine and is silently wrong. Get these right.

## 1. The detail join key  (`.H` col D ↔ HF col C ↔ mapping col C)
- The Formatted HF col C is `=IF(S="",T,S&"  -  "&T)` → e.g. `4100-0200  -  Rent Revenue - AL`. The separator is
  **two spaces, hyphen, two spaces** (`  -  `). The label half keeps the operator's leading/trailing
  whitespace exactly as `hf-formatting` wrote it.
- `.H` col D must be the **byte-identical string**. Build the `.H` detail by reading HF col C values
  (`Range.Value2`, rows 8→`hfLast`) and writing them straight into `.H` D — never retype or re-derive
  the join key. Same strings then key both the value XLOOKUP (into HF) and the tag XLOOKUP (into mapping).
- **`pnl-mapping` must run off the same HF col-C strings** so mapping col C == HF col C == `.H` col D.
  This is the locked decision from the runbook: one shared key, trivial join. If mapping was built from
  a different string source, the tag lookups return `""` and col B comes back blank.

## 2. The value pull  (`.H` F:Q ← HF E:P)
`F{r}=XLOOKUP($D{r},'<CN> Historical Financials'!$C$8:$C$<hfLast>,'<CN> Historical Financials'!E$8:E$<hfLast>)`,
then G→value-col F, H→G, … Q→P. The HF holds 12 value cols **E:P**; the `.H` holds 12 month cols **F:Q**
(one-column shift). The key range `$C$8:$C$<hfLast>` is **absolute**; the value column letter walks and is
**relative** (`E$8:E$<hfLast>` → `F$8` → …). Lock the row span absolute, let the column ride.

## 3. The tag pull  (`.H` B ← mapping)
```
=LET(r,XLOOKUP(D{r},'<CN> PnL Mapping'!$C$3:$C$<mapLast>,'<CN> PnL Mapping'!$B$3:$B$<mapLast>,""),IF(OR(r="",r=0),"",r))
```
The `LET`/`IF(OR(r="",r=0),"",r)` wrapper blanks structural/untagged rows (mapping returns "" or a 0)
so headers and subtotals don't show a stray tag. Mapping data starts at **row 3** (row 1 blank, row 2 headers).

## 4. The Census cell map (Block A links) — resolve by label, never hardcode
The Census **output block anchor moves per deal**. In the golden it sits at bucket col **R**, category col **S**,
spacer **T**, and the **12 month columns start at U** (U…AF). The `.H` links are positional cell refs
(`=Census!U21`), but you must **find** those cells, not assume U21:

1. On the Census tab, locate the cells whose text is `Capacity` (col holds the bucket) — the three rows
   immediately at/after it labelled `IL`, `AL`, `MC` in the next column are Capacity IL/AL/MC. (Golden: rows
   21,22,23.) Same for `Occupied` (golden: 26,27,28). The `Days` row is just above Capacity (golden: 19).
2. The **first month column** = the column where the month-header / Days values begin (golden: U). The 12
   months run that column → +11.
3. Align the window: `.H` F5 (oldest month) must equal the **oldest** of the Census block's 12 months, so
   `.H` month col `F+i` links to Census month col `first+i`. Confirm the HF's 12 months and the Census's 12
   months are the **same window** (same oldest→newest) before wiring — if they differ, STOP and flag.
4. Wire: `.H` F9 → Census(capIL row, firstMonthCol); G9 → next month col; … and likewise rows 10/11
   (cap AL/MC) and 14/15/16 (occ IL/AL/MC). Totals/Occ%/PRD are internal `.H` formulas (see h-skeleton §A).

## 5. Reconcile anchors on the HF (locate by label)
- `hfRevRow` = the HF row labelled **`TOTAL REVENUE`** (golden row 30) → drives row-234 Check.
- `hfNoiRow` = the HF row labelled **`NET OPERATING INCOME`** (golden row 196) → drives row-300 Check.
- `hfLast` = last used HF row (golden 197 = `Margin`). `mapLast` = last used mapping row (golden 163).
- Find these by scanning HF col C / the used range — do **not** assume the golden's numbers; they shift with
  the operator's chart of accounts.

## 6. Why the two Checks are the whole ballgame
Row 234 (`Total Revenue − HF Total Revenue`) and row 300 (`EBITDAR − HF NOI`) must be **0 in all 12
months**. Because the roll-up is a value-conserving regroup of the detail (which is itself a keyed image of
the HF), any non-zero means exactly one of: (a) a mapping tag that has no C-key home in the roll-up block,
(b) a mis-spelled tag/key, (c) a join-key whitespace mismatch dropping a line to blank, or (d) the Census
window misaligned (Block A only — won't move the Checks, but breaks PRD/metrics). Diagnose in that order.
