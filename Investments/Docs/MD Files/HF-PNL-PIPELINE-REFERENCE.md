# HF Formatting → PnL Mapping — Working Reference

> **Point future sessions here BEFORE running the `hf-formatting` or `pnl-mapping` skills.**
> It captures the environment realities, the input shapes, the proven crosswalk, the
> validation gates, and the PowerShell/Excel-COM pitfalls that otherwise burn tokens in
> trial-and-error. Read this + the relevant `SKILL.md` once; do **not** re-derive what is
> already distilled below.

Pipeline: **format HF (`hf-formatting`) → map P&L (`pnl-mapping`) → populate `.H`.**
Skills live at `.claude/skills/{hf-formatting,pnl-mapping}/`.

---

## 0. Environment hard facts (read first — these are non-obvious and cost the most when missed)

- **No Python on this box.** `python`/`python3`/`py` all fail. The skills' default
  openpyxl + `recalc.py` path **does not work here.**
- **Use Excel COM via PowerShell instead.** `New-Object -ComObject Excel.Application`
  works (Office 16.0). Excel **recalculates natively on `.Save()`** — no separate recalc step.
- **Read values** with `$ws.UsedRange.Value2` → 1-based 2D array. Date serials come back as
  doubles; text as strings. `UsedRange` starts at A1 here, so array indices == sheet rows.
- **Concurrency-safe Excel cleanup (dot-source `Investments/lib/HF-Build-Lib.ps1`):** a failed script leaves
  the workbook locked, so clean up leftovers — but the raw `MainWindowHandle -eq 0 | Stop-Process` sweep is
  now **hook-blocked**. A COM Excel is `Visible=$false` → windowless for its ENTIRE life, so that sweep also
  force-kills (a) a hand-opened workbook? no — that has a real window and survives — but (b) **a concurrent
  deal's live COM Excel in another terminal**, corrupting its in-flight `.Save()` with no crash. Instead:
  - `Clear-OrphanExcel` at startup — sweeps only windowless Excel that **no live tracked run owns** (it reads
    the per-PS-PID lockfiles), so it is safe across sibling TERMINALS (the 6-terminal 3-pack). It also deletes
    stale lockfiles whose owning PowerShell has exited. **Who calls it (2026-07-28): in an orchestrated `.H`
    run, the ORCHESTRATOR only, once at setup — sub-agents (census/HF/PnL/assembler/gate-runner) must NOT
    call it**, because census ∥ HF run concurrently and a mid-run sweep races a sibling inside the spawn-to-
    lockfile grace window. Standalone single-agent runs (e.g. the `.RR` chain) still call it at startup.
  - `New-ExcelTracked` to create Excel (records its PID in this run's lockfile) and `Stop-TrackedExcel` in
    `finally` to release only what this run spawned. Never a name-based `Get-Process EXCEL | Stop-Process`.
- **OneDrive may relocate files mid-session** (this session: `Data/*` → `Data/HF Examples/*`).
  If `Workbooks.Open` throws "couldn't find," re-resolve the path with a `find` before retrying.

---

## 1. Know the input shape BEFORE building (the #1 rework saver)

Two HF archetypes — identify which you have *first*:

| Archetype | Example | Shape | Codes? |
|---|---|---|---|
| **GL export** | Example HF-2 | 5-digit code + label + 12 months, operator subtotals end in `99` | Yes — feeds the `code  -  label` tag cleanly |
| **Mgmt report** | **Example HF-1 (management-report format)** | labels in col N, 13 monthly *actuals* in P–AB, KPI/census block on top, pooled benefits | **No** operating GL codes (only below-the-line dep/interest carry codes) |

A mgmt-report input cannot feed the downstream `code  -  label` contract cleanly and forces
many flags. That is expected — **flag, don't force.**

---

## 2. Example HF-1 specifics (so you never re-dump the file)

- A management-report (analytical-cube) export. Single sheet `Unformatted HF`.
- KPI/census block ≈ rows 21–172; **P&L "Financials" starts row 174.**
- **Trailing-12 window = source cols 17–28 (Apr'25 – Mar'26).** (Source holds 13 actual months.)
- **Reconciliation anchor rows (Example HF-1's own subtotals):**
  Total Revenue **216**, Total Opex excl mgmt **349**, Total Mgmt Fee **355**,
  Net Operating Income **357**, Total Below the Line **388**.
- Labor is dept-split for Regular/OT/Bonus, but **benefits/taxes/agency are pooled** and
  **Healthcare labor is not split AL vs MC** — neither can be departmentalized without an
  allocation basis. Keep each as its own flagged block.

---

## 3. Proven Example HF-1 → model crosswalk (reuse as the pattern)

**Revenue:** Rent group = AL Rent, AL Discounts, MC Rent, MC Discounts, Respite →
`TOTAL RENT REVENUE`. Care = Care Fees → `TOTAL CARE REVENUE`. Other = R&B Concessions,
Community Fee Rev/Incentives, Second Occupant (+Disc), Other Rev → `TOTAL OTHER REVENUE`.

**Expenses (function-based):** Food Service→Culinary · Transportation+Community Life→Activities ·
Healthcare→**own flagged block** (candidate AL/Alz split) · Maintenance+R&M→Maintenance ·
Utilities→Utilities · Housekeeping→Housekeeping · Sales+Marketing non-labor→Marketing ·
Administrative+HR+Professional Svcs+Other Costs→Administration · Insurance+Prop Tax+Mgmt Fee→
Non-Departmental · Benefits/Taxes/Agency→**own pooled flagged block** · Depreciation/Interest
(coded `8xxx`)→Corporate (below NET INCOME, col-A marker `t`).

> Model NOI is **after** mgmt fee, so it ties to Example HF-1's `Net Operating Income` (row 357).

---

## 4. Validation gates (non-negotiable — run all three in ONE pass)

1. **Reconcile-before-build:** extract leaf lines, prove they sum to the operator's own
   subtotals **across all 12 months** before writing anything. A non-zero diff = a line was
   dropped / double-counted / misbucketed (regroup is value-conserving). This catches the
   nested-subtotal trap (e.g. exclude "AL Private and Medicaid Rent").
2. **Zero formula errors** after recalc (scan cells for `#`).
3. **PnL map:** column-A `Check` returns **0 "No"** (every tag is verbatim on the master list).

---

## 5. PowerShell + Excel COM pitfalls (the token sink — get these right on the FIRST script)

- **`$r:P` parses the `:` as a namespace variable** → `"=SUM(E$r:P$r)"` silently became
  `=SUM(E7)`. **Brace it:** `"=SUM(E${r}:P${r})"`.
- **Single-element array-of-arrays collapses.** `$Care = @( @(194,"","Care") )` flattens and
  iterates `194/""/"Care"` as separate lines. Use the **unary comma:** `$Care = @( ,@(194,...) )`.
- **Do not push a 2-D array through the pipeline / `,$x` into `Range.Value2`** (InvalidCast).
  Build values inside the writer and set cells with an explicit `[double]` cast.
- **Parenthesize arithmetic in multidim index:** `$m[0,($c-17)]`, not `$m[0,$c-17]`.
- **PNG export (CopyPicture→ChartObject→`Chart.Export`) is flaky:** set `Visible=$true`,
  `Start-Sleep 500ms` between `CopyPicture` and `Paste`, and retry if the file is < ~9 KB
  (blank). Skip the screenshot unless the user asks.
- **COM "Server execution failed" right after `Stop-Process`:** the server is mid-shutdown —
  `Start-Sleep 2–3s` and retry `New-Object` in a small loop.
- **Write the build as a `.ps1` and run it once.** Iterating cell-by-cell from chat multiplies
  round-trips. Put all the fixes above in before the first run.
- **Keep `.ps1` files ASCII-only.** Windows PowerShell 5.1 reads scripts as **ANSI**, not UTF-8, so a
  UTF-8 em-dash (`—`) arrives as `â€"` and shatters the surrounding string literal into a cascade of
  parse errors ("Unexpected token", "hash literal was incomplete"). Editors here write UTF-8 by default.
  Use `-`, `->`, `x` instead of `—`, `→`, `×` **in scripts**; markdown is fine.
- **PowerShell variable names are CASE-INSENSITIVE.** `$GREEN = 32768` and a counter `$green++` are ONE
  variable — the counter overwrites the constant and every write silently uses the wrong value. Never
  distinguish a constant from a local by case alone; prefix counters `$cnt*` and write-columns `$w*`.
  This has now shipped **four** times (`$TAG`/`$tag`, `$Order`/`$order` — which saved over an operator
  rent roll, `$C`/`$c`, `$GREEN`/`$green`). Memory: `ps-var-name-case-collision`.
- **`.Value2` on an error cell returns an Int32 code, not a string** (`#DIV/0!` → `-2146826281`,
  `=1/0` → `-2146826281`). `.Text` shows `#DIV/0!` but hides it from any numeric test. Test for data
  with `.Value2` and an explicit type check — never by looking for a leading `#`.

---

## 6. Formatting specs easy to miss

**HF (`Formatted HF` tab):** slate banner `#8F96AF` (COM Color `11507343`) spans **C–Q incl
spacer D**; every total band = bold + thin top/bottom border **C–Q incl D**; accounting number
format `_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)`; gridlines off; S/T helper cols hidden;
`C = =IF(S="",T,S&"  -  "&T)`; real community name in C2.

**PnL Mapping tab (verbatim from `PnL Mapping Examples.xlsx`, Example 1):**
widths **A 8.43 / B 19.71 / C 44.86 / D 17.43 / E 21.14**; headers (A,C,D,E) Aptos Narrow 9
**bold + centered**; the **blank B2 header = Aptos Narrow 11, not bold**; **thin bottom border
under the four labeled headers A2, C2, D2, E2** (not B2 — the underline breaks over the blank
gap column); tag cells (B & E) blue font `16711680` on cream fill `13369343`,
`Ignore` red `255`; raw-line/comment cols (C,D) Tahoma 8; gridlines off;
col A `=IF(COUNTIF($E:$E,B{row}),"Yes","No")`; 63 master tags listed in E from row 3.
(COM Color = R + G·256 + B·65536.)

---

## 7. Token-efficiency playbook (where it went, how to spend less)

**Biggest spend this session, in order:**
1. **Iterative COM/PowerShell debugging** — repeated build-script re-runs over the §5 gotchas.
   *Fix:* apply §5 up front; one correct script beats five round-trips.
2. **Full reference + source dumps** — `master-format.md` (~500 lines), the whole Example HF-1
   account hierarchy (~150 rows). *Fix:* this doc's §2–§3 distillations replace re-dumping;
   open the references only to confirm a **verbatim spelling** (e.g. `Alz`, `Commercial Lease
   Revenue`, singular `TOTAL HOUSEKEEPING EXPENSE`).
3. **PNG screenshots** — flaky retries. *Fix:* skip unless asked.

**Standing rules for these tasks:**
- Read each `SKILL.md` **once**; rely on this doc for the rest.
- Validate in a **single pass** (reconcile + error scan together), not per-edit.
- **Flag calibration:** column D is for genuine judgment calls (~10%). Do **not** flag routine
  mappings like `Overtime → Salaries`; that noise dilutes the signal and costs review time.

---

## 8. What worked (keep doing)

- **Reconcile-before-build** caught structural mistakes instantly and made the final tie-out trivial.
- **Value-conserving regroup** → grand totals always reconcile even when individual department
  splits are judgment calls.
- **Keeping "no-Master-home" departments** (Healthcare, Pooled Benefits) as their own flagged
  blocks instead of forcing them into a department — honest, reviewable, and reconciles.
- **Writing memory notes** for the two durable facts (no-Python/COM; Example-HF-1-is-a-management-report) so
  later sessions start from them instead of rediscovering.
