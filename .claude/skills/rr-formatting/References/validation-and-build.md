# Validation gate + build mechanics (.RR)

## The reconciliation gate (central, source-agnostic — the real correctness check)
After emitting the AW:BM block, prove the right-side block sums to the operator's **own** detail totals
(refined by rule 12 in `locked-rules.md`: Market & 2nd-Occ Fee tie to the cent; In-Place Rent ties up to
the traceable de-proration delta; Care within rounding):
- `SUM(Market BE) == SUM(In-detail section totals' Market)` (e.g. `TOTAL Assisted Living` + `TOTAL
  Memory Care`, read at the data columns V/Z/AI).
- Likewise `SUM(In-Place Rent BF)` and `SUM(In-Place Care BG)`.
- Each must tie **to the cent**. Because the regroup is value-conserving (every source leaf belongs to
  exactly one unit), any non-zero diff means a leaf was dropped, double-counted, or a block boundary is
  wrong. This is what caught the **+$612 (O&O Community 1) / +$249 (O&O Community 2)** over-count in the
  first build — both were block-boundary bugs (companion/PS TOTAL misread; detail-end not bounded), not
  arithmetic.
- Also scan the written range for formula errors (must be 0) — **see the correct method below; the
  obvious one is broken** — and report counts (units, occupied) and every per-unit flag. Occupancy/unit
  counts are **sanity checks**, not gates — the dollar tie is the gate.

## ⚠ The formula-error scan — the obvious implementation is BROKEN BOTH WAYS
**Do not** scan a `Value2` array for strings beginning with `#`:
```powershell
if($x -is [string] -and $x.StartsWith("#")){ $errs++ }     # WRONG - silently useless
```
Excel COM marshals an **error cell as an `Int32` error code**, not a string (`=1/0` → `-2146826281`), so
this **never sees a real error** — while it **false-fires on legitimate headers** like `#Apts` / `#Units`,
which *are* strings starting with `#`. A build shipped a live `#NUM!` under this check and the gate
reported "0 errors" (verified empirically 2026-07-14).

**Correct** — ask Excel, don't pattern-match:
```powershell
$err = $null
try { $err = $ws.UsedRange.SpecialCells(-4123, 16) } catch { $err = $null }  # xlFormulas, xlErrors
if($err -ne $null){ throw "FORMULA ERRORS: $($err.Count) at $($err.Address())" }
```
(`SpecialCells` **throws** when nothing matches — that is the "no errors" path, hence the try/catch.)
Equivalent alternative: test each cell's `.Text` for a leading `#`, or check `Value2 -is [int]` on a cell
whose `.HasFormula` is true.

Write a small SUMMARY block in the right columns below the property total: unit count, occupied,
`SUM(BE/BF/BG)`, the operator totals, and the diffs — so the underwriter sees the tie at a glance.

## Build mechanics (this box — non-obvious; see also CLAUDE.md and memory)
- **No Python.** Do all Excel work via **Excel COM in PowerShell** (Office 16). `.xls` opens fine via COM.
- **Operator numbers are text** → wrap links in `VALUE()`; use `IFERROR(VALUE(x),"")` so blanks/parens
  don't throw `#VALUE!`. (`VALUE("5,885.00")`→5885; `VALUE("(759.00)")` errors → keep credits out of the
  positive finance fields.)
- **`.xls` save can silently no-op.** Save in a short retry loop, then **reopen read-only and re-sum the
  written column** to confirm the data is actually on disk before declaring done.
- **`.xls` open can RPC-reject right after** (`Call was rejected by callee`) — wrap Open/Close/Save/Quit in
  a small retry-with-sleep.
- **Don't touch the hardcoded rent roll.** Only write the right-side columns + the summary block. Gridlines
  off per house convention is a view setting (fine); data is untouched.
- **Concurrency-safe Excel cleanup:** dot-source `Investments/lib/HF-Build-Lib.ps1`, call `Clear-OrphanExcel`
  at startup, create via `New-ExcelTracked`, release via `Stop-TrackedExcel`. The raw `MainWindowHandle -eq 0
  | Stop-Process` sweep is **hook-blocked** — a COM Excel is windowless its whole life, so a sweep also kills
  a concurrent deal's Excel in another terminal (the 6-terminal 3-pack). The user also keeps workbooks open.
  (See memory `never-blanket-kill-excel`.)
- **Read with `UsedRange.Value2`** (1-based 2-D array); note `UsedRange.Row` may be >1 — array index ≠
  sheet row then. Don't reuse an array index as a sheet row in `.Cells()` without the offset.
- **BUILD ON A COPY — never write the operator's original.** Copy `<name>.xls` → `<name> .RR.xls`,
  open and write the copy, leave the source untouched. (2026-06-29 a var-name case-collision —
  `$Order` constant vs `$order` counter, names are case-insensitive — scribbled a diagonal of order
  numbers across the sheet and **saved over both source rent rolls**, destroying cells; one community's
  file was recovered from OneDrive version history, the other's wasn't.) Mitigations now baked into the engine:
  (1) prefix all write-column constants distinctly (`$wOrder`, `$wUnit`, …); (2) build-on-copy;
  (3) a guard that aborts if the source already looks corrupted (e.g. anchor `A17` == `1`). See memory
  [[ps-var-name-case-collision]].

## Mechanical gate list (gate-runner scope, 2026-07-28)
Mirrors the `.H` swarm-economics split: `gate-runner` (Haiku, read-only instrument, shared agent —
`.claude/agents/gate-runner.md`) recomputes everything below from the **saved** file and returns a raw
evidence table, zero verdicts. `rr-verifier` then sample-audits that table (per its own doc) and spends
its judgment on the column marked JUDGMENT-ONLY below.

**DETERMINISTIC (safe for gate-runner):**
1. **Market & 2nd-Occ column sums vs the operator's own raw-side section totals** — to the cent. Pure
   arithmetic; a non-zero diff is a raw number, not a call.
2. **In-Place Rent/Care sums + the raw components needed for residual decomposition** — report the
   built sum, the operator's raw total, the diff, and the traceable components (de-proration gross-ups,
   vacant-zeroing, turnover-drops) gate-runner can read directly off the sheet. **Report components only
   — do NOT judge whether they fully explain the residual or which sign applies.** That ruling is
   `rr-verifier`'s (rule 12).
3. **`SpecialCells(-4123,16)` error scan**, every tab, in try/catch (throws when clean) — per-tab error
   count + each error cell's address and Int32 code. Never the `#`-string scan (false-fires on `#Apts`).
4. **Check-block cell read-backs** — the raw values of every Check-block cell. Reading whether a cell
   equals 0 is deterministic; reading whether the whole *pattern* means something (e.g. only one care
   level is off) is not — that's the verifier's.
5. **Presence/read-backs** (report raw, no interpretation):
   - all **FIVE block headers** present and non-blank at `headerRow` (unit block, tiering helper, main
     summary, both move-in boxes);
   - **Total-row labels** populated (`Total IL/AL/MC/Total` — text present or not; not whether it's the
     *right* label for the structure);
   - **row heights**: header rows = 30, every other row = 15 — spot-read across the used range;
   - **`blockStart` derivation** — confirm `blockStart == UsedRange last column + 1` was actually used
     (no hardcoded column letter in the written formulas);
   - **gridlines off**, per sheet;
   - **zero-as-dash number formats** applied in the analysis table (accounting `_(* "-"_)` / percent
     en-dash `–`) — format-string read-back, not a visual judgment call;
   - **unit/bed/apartment count sums vs the operator summary** — raw counts and the diff, both ways
     (units/beds and `#Apts`).

**JUDGMENT-ONLY (stays with `rr-verifier`, never gate-runner):**
- **Tier reasonableness** — whether the two-pass tiering (rule 9) was actually followed and the result
  is defendable, not just internally consistent.
- **Reconcile-residual DECOMPOSITION rulings (rule 12)** — whether the reported components fully and
  correctly explain the In-Place Rent/Care residual, and which sign applies.
- **Silent-failure patterns** — stacked contracts, vacant carry-over, phantom S-line beds, MC `#Apts`
  *structure* correctness (Jack & Jill vs Semi-Private vs Private — the arithmetic is mechanical once
  the structure is known, but naming the structure from door/bed reality is judgment), label-less
  totals as a pattern, missing/thin Notes-block content.
- **Anything anomaly-shaped** — any reading gate-runner can't resolve to a clean pass/fail without a
  ruling goes to the Class-B defect ledger for the verifier, never guessed at by the instrument.

## Reuse the parser
`Investments/scripts/rr_build_onesite_v2.ps1` is the proven OneSite **parser** — its source-side half (anchor
detection, block assembly, leaf-sum finance, the reconcile gate; builds on a copy, aborts on an
already-corrupted source) reconciles to the cent and should not be re-hand-rolled. Its **output** half is
stale: it predates the v1 contract (16-col AW:BL, no Group Unit Code, no tiering, no v1 analysis table)
and its paths are hardcoded to the test files. **Extend it into the new engine; don't run it as-is.** Gap
list: `../RR-FORMATTING-DESIGN.md` → "What the next build must change".
