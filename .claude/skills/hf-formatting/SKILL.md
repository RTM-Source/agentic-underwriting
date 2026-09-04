---
name: hf-formatting
description: >
  Conforms a raw operator historical financial (T12 P&L) to the firm's HF Master
  Format: regroups the operator's own line items into the firm's revenue groups
  (Rent / Care / Other) and expense departments, builds one total per group,
  preserves the operator's codes/labels, and emits the `code  -  label` tag
  formula that feeds the PnL mapping step. This is the FIRST step of the .H
  pipeline (format HF -> map P&L -> populate .H via XLOOKUP). Use it when the user
  wants a broker/operator HF "formatted," "conformed to the Master Format," or
  "put into the firm's HF layout," or refers to building/cleaning the historicals
  before mapping. Do NOT use it to assign master-list category tags to lines —
  that is the downstream pnl-mapping skill.
---

# HF Formatting

## What this skill does
Takes an operator's historical financial (a T12 P&L with account codes, labels,
and monthly columns) and rewrites it into a new **"Formatted HF"** tab that
follows the firm's **HF Master Format**: the operator's line items are regrouped
into the firm's fixed group skeleton (revenue Rent / Care / Other; the standard
expense departments), each group gets a single live total, and every line carries
the `code  -  label` formula the mapping step reads.

This is a **first-pass, group-level** tool. It conforms *structure* and preserves
the operator's lines verbatim. It does **not** assign master-list category tags to
individual lines — that is the next skill (`pnl-mapping`). Target: produce what an
underwriter would lay out on the first pass, and **flag** anything ambiguous rather
than guess.

Pipeline position: **format HF (this skill) → PnL mapping → populate `.H` (XLOOKUP).**

## Input contract (what this skill expects)
An operator HF, typically one sheet, where each data line has:
- an **account code** (operator's own, e.g. `45105`, `61585` — 5-digit here, but
  any scheme),
- a **label** (often indented with leading whitespace),
- **12 monthly value columns** plus a **Total** column,
- section headers and the operator's own subtotals interspersed.

Tolerate these variations:
- the **month header row may be row 1 or row 2** (or elsewhere); find the row whose
  cells are dates/serials,
- month headers may be **datetime objects or Excel serials** (convert serials with
  the 1899-12-30 epoch / `[datetime]::FromOADate`),
- read values via **Excel COM in PowerShell** — `$ws.UsedRange.Value2` returns a
  1-based 2-D array of numbers (not the operator's formulas); dates come back as
  serials. (No Python on this box — openpyxl/recalc.py do not run here. See CLAUDE.md
  and `Investments/Docs/MD Files/HF-PNL-PIPELINE-REFERENCE.md`; reuse `Investments/lib/HF-Build-Lib.ps1`.)

The keep file is the authoritative structural/style spec:
**`references/master-format.md`** (and the template workbook
`Master_HF_Formatting.xlsx`). Read it before building.

## Identifying line types in the source
Walk the sheet top to bottom and classify each row:
- **Data line** — has a numeric code, a label, and at least one numeric monthly
  value. These are the only rows that get carried as line items.
- **Subtotal** — skip. Operator subtotals end in `99` (e.g. `45999`, `61899`,
  `70599`) and/or their label starts with `Total` or `Net `. (Belt-and-suspenders:
  skip if the code is a known subtotal code **or** the label starts with
  `total`/`net `.)
- **Section header** — skip. These have a label but **no monthly values**
  (e.g. `Net Rental Income`, `Administrative Salaries`).

Do not carry the operator's subtotals or headers into the output — the Master
skeleton supplies its own groups and totals. The information they carried
(care type, department) survives in each line's own label and in the group it
lands in.

## How to group the lines (the analytical core)

### Revenue → three groups
- **Rent group → TOTAL RENT REVENUE.** All base residential rent economics, net:
  Gross/Potential Rent, Gain/Loss to Lease, Vacancy, and Concessions / Veteran
  (or other) Discounts — across every care type (IL / AL / MC). Keep them together
  so the total equals the operator's net rental revenue.
- **Care group → TOTAL CARE REVENUE.** Care/assisted-care fees and their
  concessions.
- **Other Revenue group → TOTAL OTHER REVENUE.** Everything else: community &
  move-in fees, incontinence/personal-care revenue, second-occupant fees, meals,
  salon, transportation, misc operating revenue, etc.
- **TOTAL REVENUE** = sum of the three group totals.

### Expenses → model departments
Map each operator expense block to the Master department by **function**, then emit
one total per department in Master order. Worked crosswalk from the Example HF-1 trial
(use as the pattern, adapt per operator):

| Operator block | → model department |
|----------------|----------------------|
| Administrative | Administration |
| Marketing | Marketing |
| Maintenance (salaries + contracts + supplies) | Maintenance |
| Utilities | Utilities |
| Culinary | Culinary |
| Assisted Living | Assisted Living |
| Memory Care | **Alzheimers** (Alz = Memory Care — flag the naming) |
| Housekeeping | Housekeeping |
| Activities (incl. vehicle/transportation sub-block) | Activities |
| Taxes / Insurance / Licensure | Non-Departmental |

- **Flatten the operator's two-level subtotals.** Operators often split a
  department into a "Salaries" sub-block and an "Expenses" sub-block, each with its
  own subtotal. The Master has **one** total per department, so merge both
  sub-blocks into a single flat line list under the department and emit **one**
  total. Drop the operator's intermediate subtotals and sub-headers.
- **TOTAL OPERATING EXPENSES** = sum of all department totals.
- **NET OPERATING INCOME** = TOTAL REVENUE − TOTAL OPERATING EXPENSES.
- **Margin** = NOI / TOTAL REVENUE (percent format).
- **Below NOI** (only if the operator reports them): Other (Income)/Expense lines
  → TOTAL OTHER (INCOME) EXPENSES → NET INCOME; Corporate lines → TOTAL CORPORATE
  EXPENSES (col A = `t`). If the operator's P&L ends at NOI, end at NOI + Margin.

### Line-level rules (always)
- **Keep the operator's own codes and labels.** Do not renumber to the Master's
  account codes. Strip leading/trailing whitespace from labels.
- **Preserve source signs exactly.** Vacancy, concessions, discounts, GLTL arrive
  negative — do not flip them.
- **Infer care type from context, not a default** — from the surrounding section
  header or the label suffix (`- MC`, `- AL`), not a blanket assumption.
- **Group-level only.** Do not relocate individual lines across departments to
  match Master convention — line-level category assignment is the mapping step's
  job (see "Flag, don't force" re: Management Fee).

## Flag, don't force
Surface judgment calls instead of silently resolving them — as a short cell
**comment** on the relevant row and in the chat summary:
- **A department with no Master home** (e.g. **Wellness / Nursing**): keep it intact
  as its own block with its own `TOTAL <NAME> EXPENSES`, positioned among the care
  departments, and flag it (candidate: a Direct Care bucket, or split across
  AL/Alz). Do not silently dissolve it.
- **GLTL / Vacancy placement.** The Master *template* lists Gain/Loss to Lease and
  Vacancy Loss under OTHER REVENUE, but operators usually tie them to rent by care
  type. Keep them in the Rent group to match the operator's economics and flag the
  divergence.
- **Management Fee location.** Operators often bury Management Fee inside
  Administrative; the Master convention is Non-Departmental. **Leave it where the
  operator put it** — the mapping step retags it correctly — and flag it.
- **Empty Master groups** (Independent, COGS for an operator that lacks them):
  omit and note.
- **Month-header quirks** (e.g. a 12th month dated the 1st instead of month-end):
  carry as-is and note.

## Output layout & styling
Add a new sheet **"Formatted HF"** as the **first** tab (leave the operator's
original sheet intact behind it). Reproduce the Master Format exactly — full column
map, header block, group skeleton, fonts, fills, number formats, gridlines-off, and
the row-1 spacer are all specified in **`references/master-format.md`**. The pieces
the build code must get right:

- **Column C is a formula:** `=S{r}&"  -  "&T{r}`, with `S{r}` = operator code and
  `T{r}` = clean label. This is what the mapping step consumes.
- **Monthly values land in E–P (12 cols); Total Q = `=SUM(E:P)`** per line.
- **Totals are live formulas:** each group/department total sums its member line
  rows; TOTAL REVENUE sums the three revenue-group totals; TOTAL OPERATING EXPENSES
  sums the department totals; NOI = TOTAL REVENUE − TOTAL OPERATING EXPENSES;
  Margin = `=IFERROR(NOI/REV,0)` in **percent** format.
- **The REVENUE banner carries the 12 month dates** (format `mmm-yy`) and `Total`;
  the EXPENSES banner is filled but has no date text.

### Two corrections that are easy to miss (fold these in)
1. **Real community name in the header (C2)** — e.g. `Aster Ridge (######)`, not the
   `Community Name` placeholder. Keep the `(######)` property-ID slot.
2. **The slate banner fill and the bold + thin-top/bottom-border total band must
   span columns C THROUGH Q contiguously — including the spacer column D.** Styling
   C and E–Q while leaving D bare leaves a visible gap. On banner rows, D gets the
   slate fill + white bold font; on total rows, D gets bold + thin top/bottom
   borders (black).

## Build & validation procedure
1. Read `references/master-format.md` and the relevant `xlsx` skill notes first.
2. Read the source via **Excel COM**: `$ws.UsedRange.Value2` for exact monthly values
   and the month-header row (1-based 2-D array; date headers as serials).
3. Parse data lines (per "Identifying line types"), classify each into a group,
   preserving the operator's encounter order within each group.
4. Build the "Formatted HF" tab per the styling spec, writing S/T helpers, the
   C formula, monthly values, and live total/NOI/Margin formulas.
5. Add flag comments on the judgment-call rows.
6. **Recalculate** by saving — Excel recalcs natively on `$wb.Save()`; no separate step.
   (Set `$xl.Calculation = xlManual` during the build, then `$xl.CalculateFull()` once
   before save to avoid O(n^2) recalc-on-every-write.)
7. **Validate** — both must hold before presenting:
   - **zero formula errors** from recalc, and
   - the rebuilt **TOTAL REVENUE, TOTAL OPERATING EXPENSES, and NET OPERATING
     INCOME reconcile to the operator's own corresponding totals** (their
     Total Revenue / Total Expenses / Total NOI lines) to the cent. Because the
     regrouping is value-conserving, any mismatch means a line was dropped,
     double-counted, or misbucketed — find it.
8. Report the group totals back to the user and list the flags. The C-column tag
   strings are now ready to feed the `pnl-mapping` skill.

## Keep / improve
Log each new formatting decision and operator-structure quirk (new department
mappings, unusual sign conventions, header oddities) to the **HF Formatting
Observations** notepad so the rules here grow with each deal processed.

## Related skills & references
- **References (this skill):** `references/master-format.md` — the authoritative target layout/styling
  (read before building). Template workbook: `Master_HF_Formatting.xlsx`.
- **Build engine:** `Investments/lib/HF-Build-Lib.ps1` — the `HF-*` COM writers, `New-Excel`, `SetV`/`SetBlock`.
  Reuse it; don't re-hand-roll the COM styling. Deep environment + crosswalk: `Investments/Docs/MD Files/HF-PNL-PIPELINE-REFERENCE.md`.
- **Pipeline position:** this is **stage 2** of the `.H` build. Its output (the Formatted HF tab — col C
  `code  -  label`, E:P values) feeds **`pnl-mapping`** next, then the `.H` assembly. See the container map
  `../H-PIPELINE-ORCHESTRATION.md`.
- **Memory:** `formatted-hf-target-format` (the real target = heavy skeleton WITH operator codes),
  `hf-natural-account-operator-split`, `excel-com-write-pitfalls`, `env-no-python-use-excel-com`.
