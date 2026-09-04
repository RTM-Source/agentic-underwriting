# RR Build Engine — contract (v1, 2026-07-14)

**Goal:** freeze the proven output-side COM code into `Investments/lib/RR-Build-Lib.ps1` so an agent never
hand-writes it again. Agent's job shrinks to: recognize source → judgment calls → emit **config** →
invoke engine → verify. Target wall-time ≤ ~10 min/deal.

## Split of responsibility
- **AGENT (varies per deal):** parse the raw export (any family), resolve judgment (care mix, A/B/C
  tiers, MC structure, anomalies), emit `config.json` + `units.csv` to `%TEMP%`.
- **ENGINE (identical every deal):** everything right of `blockStart` — unit block, tiering helper,
  main summary, Second Residents, Rent Adjustments, Check, move-in boxes, Notes, reconcile block,
  the full formatting contract (`References/formatting.md`), gates, rule-21 build-in-place flow
  (`~building —` → `Save()` → rename). The engine IS the golden made executable.

## Inputs
`Invoke-RRBuild -ConfigPath <config.json>`

**config.json** (global):
```json
{ "community": "Deal C", "sourcePath": "...", "dealFolder": "...", "version": 2,
  "headerRow": 5, "blockStart": 14, "unitRowFirst": 8, "unitRowLast": 139,
  "colMap": { "unit":"A", "market":"F", "actual":"G", "moveIn":"J", "secOcc":null, "care":null },
  "careLevels": ["IL"], "daysInPeriod": 30, "deprorate": false,
  "operatorTotals": { "market": 530815.00, "actual": 353339.49, "care": null, "secOcc": null,
                      "units": 132, "apartments": 132 },
  "tiers": [ {"group":"IL | Studio | A", "sqfts":[592]}, ... ],
  "moveInWindows": [ {"start":"2025-10-01","end":"2026-03-31"}, {"start":"2026-04-01","end":"2026-04-30"} ],
  "notes": [ "PROPOSED: units 403/413 forced vacant — moved out on As-Of date", ... ] }
```

**units.csv** (one row per unit line; agent resolves ALL interpretation before the engine runs):
`anchorRow, leafRows(;-sep), unit, unitType, psp(P/S), careType, sqft, moveIn(ISO or blank),
capacity(0/1), occupied(0/1), secFold(0/1 fee+care fold onto this primary), basicUnitCode, groupUnitCode`

## Engine obligations (all locked rules apply — cite by number)
1. Rule 21 flow: `Copy-Item` source → `~building — <C>—RR_v#.xlsx` in dealFolder → Open → build →
   `Save()` (never SaveAs) → rename on success; `try/finally` deletes partial on failure.
2. Unit block (17 cols at blockStart): formulas link back via colMap (`IFERROR(VALUE(x),"")` — harmless
   on numeric sources); hardcodes per output-columns.md; capacity/occupancy from units.csv (STRUCTURAL,
   rule 2); de-proration (rent AND care, rule 11) only when `deprorate=true`; vacant → literal 0 (rule 17);
   turnover = current-occupant leaf only; S-fold per `secFold` (rule 17).
3. Analysis: tiering helper; main summary grouped on Group Unit Code (percent ratios `x/y−1` for
   (Discount) & NMI-vs-In-Place); Second Residents 4-col IL/AL/MC/Total; Rent Adj; Check (all must
   compute 0); two move-in boxes; **Notes block** from config.notes; reconcile block vs operatorTotals.
   Row plan is DYNAMIC per care mix (populated levels get real rows; absent levels get the golden's
   placeholder convention). NMI = blank BLUE input cells unless config provides values.
4. Formatting per `References/formatting.md` — **ALL FIVE headers**, yellow/navy styles, heights 30/15,
   dash formats no `$` in analysis, blue/orange/red, totals bold top-border-only with labels, sub-block
   titles plain bold, gridlines off, tab `<Community>.RR`.
5. Gates, self-run, returned as an object AND printed: reconcile diffs (each decomposed), SpecialCells
   error count, Check-block values, five-header presence check. Non-zero unexplained diff → throw
   (finally-block still cleans up).
6. **Performance:** ONE Excel session; `ScreenUpdating=$false`; `xlCalculationManual` → `CalculateFull()`
   once; bulk 2-D `Value2`/`Formula` range writes wherever rows are contiguous (never per-cell loops for
   the unit block); en-US culture; `$w`-prefixed constants (case-collision rule); `Clear-OrphanExcel` +
   `New-ExcelTracked`/`Stop-TrackedExcel` cleanup (the windowless sweep is hook-blocked); retry-wrapped Open/Save.

## Acceptance test (required before "done")
Rebuild **Deal C** from its raw file via a hand-written config, output as `_vTEST` in `%TEMP%`
(NOT the deal folder), and diff against the verified `Investments/Data/Stress Test/Deal C—RR_v2.xlsx`:
identical header labels (all 5 blocks), identical unit-block values/formulas modulo cosmetic spacing,
Market/In-Place sums tie to the cent, Check all 0, 0 formula errors, formatting spot-checks (fills,
fonts, heights, dash formats) match. Report any intentional divergence.

## Sources to harvest (proven code — do not re-derive)
- `Investments/Data/Stress Test/arc test/Community 1 Test 2/build_rr_community1.ps1` — v1-contract build
  incl. analysis table + formatting (best single source).
- `Investments/Data/Stress Test/arc test/Community 1 Test 1/rr_build_community1.ps1`, `Community 2 Test 1`
  scripts — parsing + reconcile patterns.
- `Investments/lib/HF-Build-Lib.ps1` — house COM style (New-Excel, SetV, retry, tracked instances). Follow it.
- Golden `Investments/Data/Transactions/Aster Ridge (Demo)/Rent Roll/Aster Ridge-RR_v1.xlsx` — the ground
  truth to mirror.
