# `.H` Assembly Engine — contract (v1, 2026-07-28)

**Goal:** freeze the proven `.H`-assembly COM code into `Investments/lib/H-Assembly-Lib.ps1` so
`h-assembler` never hand-writes it again. The sub-agent's job shrinks to: resolve the per-deal
anchors (hfLast/hfRevRow/hfNoiRow/mapLast/census map), emit **config**, invoke the engine, report
its gate readings. Modeled on `Investments/lib/RR-Build-Lib.ps1` / `rr-formatting/scripts/ENGINE-SPEC.md`.

## Split of responsibility
- **h-assembler (varies per deal):** locate the workbook + tabs, read HF/mapping/Census anchors by
  label (never hardcode), resolve which Census tabs need merging and their scratch path, emit
  `config.json`, invoke `Invoke-HBuild`, relay the returned gate readings + any thrown reason.
- **ENGINE (identical every deal):** census merge (step 0), Block A occupancy, Block B detail
  passthrough, Block C roll-up, Block D metrics, Block E S/T/U totals + the full styling contract
  (`h-skeleton.md`, `rollup-category-map.md`, `join-and-census-contract.md`), self-run gates,
  save-only-if-clean. The engine IS the golden (`Aster Ridge.H Model_v1.xlsx`), made executable.

## Inputs
`Invoke-HBuild -ConfigPath <config.json>`

```json
{ "workbook": "<abs path - built IN PLACE, Save() only, never SaveAs>",
  "cn": "Aster Ridge", "tabPrefix": "Aster",
  "hTab": "Aster.H (Review)", "hfTab": "Aster Historical Financials", "mapTab": "Aster PnL Mapping",
  "censusScratch": "<abs path or null - when non-null, engine merges its tabs into the workbook first>",
  "censusTabs": ["Notes","Census","Occupancy Source"],
  "hfDetailStart": 8, "hfLast": 215, "hfRevRow": 36, "hfNoiRow": 214, "mapLast": 180,
  "censusMap": { "headerDateRow": 26, "capacityRows": {"IL":30,"AL":31,"MC":32},
                 "occRows": {"IL":40,"AL":41,"MC":42}, "firstMonthCol": "E" },
  "windowStart": "2025-07-31" }
```

| Field | Meaning | Locate by |
|---|---|---|
| `workbook` | abs path, built in place (`Open` + `Save()`, never `SaveAs`) | — |
| `cn` / `tabPrefix` | community name used in sheet names / formulas | — |
| `hTab`/`hfTab`/`mapTab` | the three existing tab names the engine reads/writes | exact tab names |
| `censusScratch` | abs path to a scratch workbook holding Census tabs to merge, or `null` if already merged into `workbook` | orchestrator/h-assembler prompt |
| `censusTabs` | ordered list of scratch sheet names to copy (`Notes` lands FIRST in the model workbook; others land after `hTab`'s position) | scratch workbook's own sheet list |
| `hfDetailStart` | first VALUE row on the HF, right after its `REVENUE` banner | first non-blank HF col-C row after the "REVENUE" banner row |
| `hfLast` | last used HF row (its `Margin` row) | last non-blank HF col-C row |
| `hfRevRow` | HF row labelled `TOTAL REVENUE` | scan HF col C |
| `hfNoiRow` | HF row labelled `NET OPERATING INCOME` | scan HF col C |
| `mapLast` | last used PnL-Mapping row | mapping tab's used range |
| `censusMap.headerDateRow` | Census row holding the 12 month-end serials | scan Census col A/B for "Output"/date-header row |
| `censusMap.capacityRows.{IL,AL,MC}` | Census rows labelled `Capacity` × `IL`/`AL`/`MC` | scan Census bucket column |
| `censusMap.occRows.{IL,AL,MC}` | Census rows labelled `Avg Occ %` (or equivalent) × `IL`/`AL`/`MC` | scan Census bucket column |
| `censusMap.firstMonthCol` | first of the 12 month columns on Census (mirrors `headerDateRow`) | column letter of the date-header row's first date |
| `windowStart` | oldest month-end (ISO date) — must equal the Census block's oldest month | HF's own month-header row / cross-checked against Census |

## What the engine derives itself (never hand these in)
- **`detailFirst`** = 28 (fixed template row — Block A always ends at 24, `REVENUE` banner at 26, first
  detail line at 28).
- **`detailLastContent`** = `detailFirst` + (count of mapped HF rows, hfDetailStart..hfLast, MINUS the
  dropped `Margin` row) − 1.
- **`detailLast`** = `detailLastContent` + 1 (the SUMIFS/XLOOKUP range boundary — one row past the last
  written line; this trailing blank is intentional, matches `h-skeleton.md`).
- **`rollupFirst`** = `detailLast` + 3 (two blank rows after the detail block). **Never hardcode this —
  it shifts with every deal's HF length** (memory `h-skeleton-rows-not-fixed`).
- **`rollupLast`** = `rollupFirst` + 70 (the fixed 71-row roll-up-category span used as the metrics'
  SUMIFS criteria range).
- Every Block C/D/E row number = `rollupFirst` + a **fixed offset** from the `rollup-category-map.md`
  baseline (baseline anchor 221 ⇒ offset 0). E.g. `rowTotalRevenue = rollupFirst+11`,
  `rowCheckRev = rollupFirst+13`, `rowEBITDAR = rollupFirst+75`, `rowCheckNOI = rollupFirst+79`,
  denominators `rollupFirst+81..+87`, metrics `rollupFirst+89..+115`. These offsets are baked into
  the engine as data tables (`$script:RollupCats`/`LaborCats`/`TempCats`/`BenefitCats`/`OpexCats`/
  `MetricRows`) — verified cell-by-cell against the golden and reproduced identically on a
  second, independently-lengthed HF (the regression file — same `hfLast`, different Census/mapping).

## Engine procedure (Block order, per `h-skeleton.md` §4)
1. **Step 0 — census merge** (only if `censusScratch` is non-null). Copies ALL listed scratch
   worksheets in **one grouped** `Worksheets(@(...)).Copy($beforeSheet)` COM call — never one sheet at
   a time (a one-at-a-time copy rewrites intra-workbook cross-sheet formulas into **external links** to
   the scratch file's path; this shipped 2228 errors in a prior run). `Notes` is moved to be the first
   sheet in the model workbook after the copy.
2. **Block B (detail) first.** Reads HF col C in ONE bulk `Range.Value2` read (rows `hfDetailStart..
   hfLast`), classifies each row `blank` (HF cell empty → blank `.H` row, no formulas) or `normal`
   (everything else, INCLUDING the HF's own internal `REVENUE`/`EXPENSES` section banners — those still
   get full D/B/F formulas, verified against the golden; only the truly-blank HF cell skips formulas).
   The `Margin` row (matched by trimmed text, wherever it falls) is dropped and the rows below it
   compact up by one. Writes col B (tag `LET`/`XLOOKUP`/`IF(OR(...))` formula), col D (byte-identical
   HF col-C string), cols F:Q (12 `XLOOKUP`s into HF value cols E:P) in **one bulk `SetBlock` range
   write** covering the whole detail block.
3. **Block A (occupancy, rows 5–24).** F5 = `windowStart` (blue); G5:Q5 = `EOMONTH` walk (black). Rows
   9–11/19–21 pull `Census!<col><row>` (green) using `firstMonthCol`+offset per month; rows 14–16 derive
   `=Occ%×Capacity` (black); rows 12/17/22 are in-sheet totals. Row 24 (`PRD`) is written AFTER
   `rollupFirst` is known (it references the Residents/Occupied-Units metric row).
4. **Block C (roll-up, `rollupFirst`..`rollupFirst+79`).** Fixed category/subtotal/Check rows per the
   offset tables above, `SUMIFS` over `$25:$<detailLast>`.
5. **Block D (denominators + metrics, `rollupFirst+81`..`rollupFirst+115`).** Denominators are simple
   in-sheet references; metrics are `SUMIFS` over `$<rollupFirst>:$<rollupLast>` keyed on col C, divided
   by a denominator per the row's `kind` (`unit`/`k1000`/`prd`/`rentcare`/`benpr`/`healthdental`/
   `adminK`/`mgmtfee`/`revpct`/`marketing` — verbatim formula shapes captured from the golden). The R
   column carries a tag-presence `XLOOKUP` self-lookup over the labor-block span
   (`$<rollupFirst+15>:$<rollupLast>`) for rows whose D-tag exists there, else the literal `"Ok"`
   (Benefits & PR Taxes / Management Fee / Capex Reserve — the three tags with no roll-up home in that
   span).
6. **Block E (S/T/U totals).** Three formula SHAPES depending on row kind:
   - **Flow** (Days/PRD, every detail value row, every roll-up category+Management-Fee row):
     `IFERROR(SUMIFS($E{r}:$Q{r},$E$5:$Q$5,">="&{c}$4,...),0)*{c}$2`.
   - **Stock** (Block A rows 9–11/14–16): `IFERROR(AVERAGEIFS(...),"NA")`. **Rows 12/17 (Totals) use
     the SAME `AVERAGEIFS` windowed shape but WITHOUT the `IFERROR` wrapper** — verified divergence
     from the member rows, golden-confirmed.
   - **Translate** (every subtotal/denominator/metric row): literally the SAME formula template used
     for col F, regenerated with the column letter parameter set to `S`/`T`/`U` instead of `F` — NOT a
     re-windowed month sum (these already derive from the roll-up's own windowed totals).
   - **Check (S only):** `=S<rowTotal>-'<hfTab>'!Q<hfRevRow|hfNoiRow>` — ties to the HF's **annual
     total column Q**; T/U carry number-format only, no formula.

## Styling contract highlights the engine self-enforces (see `h-skeleton.md` for the full list)
- Detail col-B tag cells: green `32768`, Calibri 9, no fill, not bold.
- Detail **total** rows (`^TOTAL ` or `NET OPERATING INCOME`, trimmed): bold + thin top/bottom border
  across **D:Q and S:U**. The HF's own internal `REVENUE`/`EXPENSES` banner lines get **D bold only**
  (no border, formulas intact — they are NOT the same case as a total row).
- Roll-up borders span **F:U** (unlike Block A's `F:Q`-only total-line rule) — includes the R column.
- Total Opex / EBITDARM / EBITDAR use the plain `#,##0` int format (not the accounting dash format)
  across F:U; the blank row directly below Total Opex (offset 72) stays `General` (golden read-back
  quirk — not swept with the surrounding accounting format).
- Metric-row R column carries a legacy percent format for the percent-kind metrics regardless of
  whether R holds a formula or the `"Ok"` literal (1-decimal for `benpr`/`rentcare`/`healthdental`,
  2-decimal for `mgmtfee`/`revpct`); all dollar/K-property/PRD-kind metric rows leave R at `General`.
- Residents/Occupied-Units row (rollupFirst+83): F is the **orange-fill** hardcoded input (font stays
  **black**, not blue — golden deviates from the general blue-input rule here); G:Q chain
  `=+<prevCol><row>` off the PREVIOUS column, not all off F.

## Self-run gates (inside `Invoke-HBuild`, before `Save()` — throws with a precise reason on failure,
workbook left unsaved)
1. Zero formula errors workbook-wide via `SpecialCells(-4123,16)` in try/catch per sheet (never the
   `#`-string scan).
2. Revenue Check row and NOI Check row = 0 across all 12 month columns (tol `1e-6`); S-column Checks
   on both = 0.
3. Every mapping tag (PnL Mapping col B, rows 3..mapLast) with a non-blank value has a roll-up C-key
   home among the category/Management-Fee/Ignore keys.
4. External-link scan = 0 hits — every formula workbook-wide scanned for `[` or `http`.
5. Census: all 12 Block-A capacity + Occ% cells non-zero; `.H` row-5 dates == Census header-row dates
   at the aligned month columns.
6. Styling spot-assertions: B28 green/Calibri; Check-row fill; S2 blue + right-aligned; the last detail
   content row bold.

## Hard environment rules honored
- Excel COM only, en-US culture, `xlManual` during writes → `CalculateFull()` (twice, with a settle
  sleep) → gates → `Save()` in place (never `SaveAs`). Gridlines off. `New-ExcelTracked`/
  `Stop-TrackedExcel` in try/finally. The engine never calls `Clear-OrphanExcel` (orchestrator-owned).
- Multi-dimensional COM SafeArray indexing: parenthesize any arithmetic inside `[...]`
  (`$arr[($i+1),1]`, not `$arr[$i+1,1]`) — PowerShell mis-parses the unparenthesized form as a method
  call and throws `does not contain a method named 'op_Addition'`.
- `if{...}else{...}` as an expression may be the RHS of a plain assignment but **cannot** be wrapped in
  `(...)` with a chained method call (`$x = (if(...){}else{}).Trim()` fails to parse) — assign, then
  call the method on a separate line.

## Acceptance tests performed (2026-07-28)
- **GATE A** — regression workbook (a second, independently-lengthed HF build, `.H` tab
  deleted, Census already merged, `censusScratch=null`, `mapLast=180`, census rows Capacity 30/31/32 /
  Occ% 40/41/42): rebuilt via `Invoke-HBuild`, diffed cell-by-cell A1:U353 (Value2 tol 1e-6, formula
  string, number format, font color, bold, fill, top/bottom borders) against the original regression
  `.H` tab. **Result: 0 differing cells** (after 6 iterative fix passes — see below).
- **GATE B** — golden (`Aster Ridge.H Model_v1.xlsx`, `.H` tab deleted, `mapLast=210`, census
  rows Capacity 48/49/50 / Occ% 58/59/60, `headerDateRow=44`): rebuilt with the golden's own derived
  config, diffed the same way. **Result: 0 differing cells on the first attempt** (the engine, tuned
  against the regression file, generalized to the golden's different geometry with no further changes).
- Both runs: 0 formula errors, both Checks 0 across all 12 months and the S-column T12 tie-out, 0
  external-link hits, census gate + styling spot-assertions passed.
- `%TEMP%\hbuild-testA.xlsx` / `hbuild-testB.xlsx` deleted after validation; the golden and regression
  workbooks were opened read-only throughout and are untouched (confirmed via `git status`).

## Sources to harvest (proven patterns — do not re-derive)
- `Investments/lib/RR-Build-Lib.ps1` — house COM style, config-driven engine pattern, gate style.
- `Investments/lib/HF-Build-Lib.ps1` — `New-ExcelTracked`/`Stop-TrackedExcel`, `SetV`/`SetBlock`.
- Golden `Investments/Data/Transactions/Aster Ridge (Demo)/Aster Ridge.H Model_v1.xlsx` — the ground
  truth for every formula/format/color/border in this contract (read-only, never write it).
