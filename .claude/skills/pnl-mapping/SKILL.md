---
name: pnl-mapping
description: >
  Maps the line items of a normalized senior-housing historical financial (T12)
  to the firm's master-list category tags and produces a reviewable PnL Mapping
  sheet (Check / Tag / raw line / Questions & Comments / List of Tags). Use this
  AFTER a historical financial has been normalized into a clean account-code +
  label + monthly layout. Do NOT use it to clean up a messy raw broker file —
  that is the upstream HF normalization skill. Trigger when the user wants raw
  HF line items turned into model-ready model tags, or refers to "PnL mapping,"
  "tagging the .H lines," or "mapping the historicals."
---

# PnL Mapping

## What this skill does
Takes a normalized historical financial (HF) and assigns each real line item one
tag from the firm's master list, writing the result into a PnL Mapping sheet that
an underwriter reviews and corrects before it flows into the model. This is a
**first-pass** tool: the target is to match an underwriter's first pass (~90%),
not to be perfect. Speed at the outset is the value; the human does the strategic
work afterward.

## Input contract (what this skill expects)
A normalized HF where each line has, at minimum:
- an **account code** (e.g. `40050`, `44100`, `7010-1110`),
- a **label** (e.g. `Potential Rent`, `IL Studio`, `Payroll Wages - Regular`),
- the structure preserved: **section headers, subtotals, and original row order
  are kept in place** (these carry the context the mapping depends on),
- monthly columns and/or a total (values, not formulas).

This matches the "Unformatted HF" stage of a clean file (e.g. Example HF-2) or the
"Formatted" stage of a messy one (e.g. Example HF-3). If the file is still a raw,
messy broker export (property name in a stray column, budget columns interleaved,
stats mixed with dollars), STOP — it needs the upstream normalization skill first.

## The one hard rule (non-negotiable)
**Map only to tags that appear, verbatim, in `references/master-tags.md`.**
Never invent, rename, re-spell, merge, or extend a tag, and never edit the master
list. If a line cannot be confidently placed on a master tag, it is **flagged**,
not forced. (Reminder: the list spells it `Commercial Lease Revenue` — match it
exactly. And there is no `Care Revenue - IL` tag — an IL care line must be flagged.)

## Standing firm rulings (apply silently — no flag)
Ruled by Ryan 2026-07-08 (first applied: O&O v16). Precedent + detail in
`references/off-list-tags.md`:
- **Holiday pay → `Benefits - <Dept>`**, never Salaries — same treatment as PTO, even when the
  label reads `Salaries - Holiday <Dept>` (`6120-0102  Salaries - Holiday Housekeeping` →
  `Benefits - Housekeeping`; `...Holiday G&A` → `Benefits - Admin`). Regular/OT/double-time/
  accrued wages stay `Wages - <Dept>`.
- **Software line items → `Admin - IT`** regardless of the section they sit in (`Marketing
  Software`, `Sales Software`, `Admin Software`, `Software Services`). Media/comms spend is not
  software — Web Services, Digital Advertising, SEO/Content stay `Marketing`.
- **Personal property tax → `Real Estate Taxes`** (pooled with real + business property tax), never
  `Other Taxes` (`Other Taxes` keeps sales/franchise/misc).

## Output: the PnL Mapping sheet
Reproduce the firm's existing mapping-tab layout AND formatting exactly — both extracted
from `PnL_Mapping_Examples.xlsx` — so the sheet drops straight into how they already work.

**Layout**
- Row 1: blank.
- Row 2: headers.
- Row 3 onward: data.

**Columns**

| Col | Header | Contents |
|-----|--------|----------|
| A | `Check` | Validation formula `=IF(COUNTIF($E:$E, B{row}), "Yes", "No")` — confirms the tag in B exists on the master list. **This is NOT a confidence flag.** |
| B | *(no header)* | the assigned master-list tag (the model-ready category) |
| C | `Tag` | the raw line = `<account code> <label>` (i.e. source `=A&" "&B`, label indentation preserved) |
| D | `Questions & Comments` | the flag/justification on every judgment call or low-confidence line |
| E | `List of Tags` | the 63 master tags, listed from row 3 down — reviewer reference and the source for column A's COUNTIF |

Firm quirks to preserve: column C is headed **"Tag"** but holds the **raw line**; column B
holds the **mapped category** and has **no header**. Don't swap them.

**Cell formatting (match exactly)**
- Header row (row 2): Aptos Narrow 9, **bold**.
- Header underline: a **thin bottom border under the four labeled headers — A2, C2, D2, E2**
  (`Check`, `Tag`, `Questions & Comments`, `List of Tags`). Leave **B2** (the blank tag-column
  header) **without** a border, so the underline breaks over the unlabeled gap column.
- Column A (Check): Aptos Narrow 9, the COUNTIF formula above.
- Columns B and E (tags): Aptos Narrow 9, font color `0000FF` (blue), solid fill `FFFFCB` (pale cream).
  - Exception: the `Ignore` tag renders **red** (`FF0000`), not blue — wherever it appears
    (always in column E's list; in column B if a line is tagged `Ignore`). Cream fill unchanged.
- Columns C and D (raw line, comments): Tahoma 8.
- Column widths: A default (~8.4), B ≈ 20.43, C ≈ 45.57, D ≈ 18.14, E ≈ 21.86.
- Sheet view: **gridlines off** (`$wb.Windows.Item(1).DisplayGridlines = $false` via
  Excel COM) — always. (No Python on this box; build via COM / `Investments/lib/HF-Build-Lib.ps1`.)

**How flags surface.** Column A is a COUNTIF validation that returns "Yes" for any valid
master tag, so it does not distinguish flags, and **no cell is highlighted.** The flag is a
terse note in column D — **a non-empty `Questions & Comments` cell is the flag;** confident
rows leave D blank. Keep notes token-efficient: telegraphic, drop articles, use `→`
(e.g. "Respite/short-term; could be Other Revenue", "Company-wide; no dept Benefits tag").

**Intentional retentions beyond the firm's bare (flat) example**
- Structural passthrough rows (section headers, subtotals, NOI) kept in column C, **bold**,
  untagged — the context anchors the firm's flat example omits.

## Procedure
1. **Walk the HF top to bottom in order.** Order is information — don't sort or
   regroup it.
2. **Pass through structural rows untagged.** Section headers (`REVENUE`,
   `Independent Living`, `Assisted Living`, `Memory Care`), subtotal/total rows
   (`TOTAL POTENTIAL RENT`, `NET RENTAL REVENUE`, `Total Independent Living`), and
   blank rows are structure, not line items. Keep them; never assign them a tag.
3. **Maintain a "current context" as you go**, updated by the structural rows:
   - revenue vs. expense block,
   - care type (IL / AL / MC) of the section you're inside.
4. **For each real line item:**
   - Build the raw-line string `C = "<account code> <label>"`.
   - Decide revenue vs. expense from the block and the account-code range.
   - Decide care type: if the **label** names one (e.g. `IL Studio`, `AL Care
     Revenue`, `MC Care Revenue`), use it; otherwise **inherit the care type from
     the current section** (see Care-type inference below).
   - Pick the single best-matching master tag and put it in `B`. Column `A` carries the
     COUNTIF validation formula (it returns "Yes" for any valid master tag).
   - If confident → tag in `B`, `D` empty.
   - If not confident → best guess in `B`, a terse reason in `D` (non-empty D = the flag;
     no highlighting). Never substitute an off-list tag.
5. **Below-NOI / non-operating lines → `Ignore`.** If you can't tell where the NOI
   line falls, flag it rather than guessing the boundary.
6. **Preserve signs exactly** as they appear in the source. Vacancy, concessions,
   discounts, and marketing incentives normally arrive negative; gain/loss to
   lease is signed. Do not flip a sign. If a sign looks inverted vs. expectation,
   flag it in `D` — don't silently correct it.

## Care-type inference (the central judgment)
Care type usually comes from the **neighborhood, not the label**. In a normalized
HF the line `Private` means nothing alone, but under a `Memory Care` section
header it is MC. Inherit the active section's care type for any line that doesn't
name its own. When a line sits on a boundary or is genuinely ambiguous (e.g. a
`Care Concession - Other` line between an AL block and an MC block), make your best
call, flag it (terse note in `D`), and explain it — exactly the way the human note
"Doesn't specify which, mapped to AL" did. Do not apply a blanket default; read
the context.

## When to flag (terse note in column D)
Always flag — a short note in `D`, nothing else (no highlighting) — for:
- blended or split IL/AL lines;
- care type inferred from context rather than stated;
- vague or non-standard labels with no clean master match;
- lumped categories covering more than one tag;
- IL care/health-services lines (no `Care Revenue - IL` tag exists);
- anything sent to `Ignore` that might actually be operating, or an unclear NOI
  boundary;
- a sign that looks inverted.

A flag is a success, not a failure — it is the skill doing its job of surfacing
the ~10% an underwriter needs to look at.

## Worked references (from real mappings)
- `45710 Concession - Community Fee` (under a revenue block, arrives negative) →
  `Rent Concessions`; preserve the negative sign.
- `45705 Community Fee` → `Community Fees`.
- `Potential Rent` inside a `RENT REVENUE` section of a primarily-IL community →
  `Rent Revenue - IL` (care type from the community/section, not the word "rent").
- `Payroll Wages - Regular` appears many times under different cost centers; the
  account-code prefix / section determines the department (e.g. an admin cost
  center → `Wages - Admin`, a culinary cost center → `Wages - Culinary`).
  Same label, different tag, decided by context.
- `4700-0034 Care Concession - Other`, sitting between AL and MC care lines →
  best guess `Care Revenue - AL`, flagged in D: "care type unstated; inferred from AL block."

## Out of scope (do not do here)
- Cleaning/normalizing a messy raw file → upstream HF normalization skill.
- Building the .H tab or any model formulas → .H formatting skill.
- Editing the master tag list → never.
- Rent roll fields → rent-roll skill.

## Related skills & references
- **References (this skill) — read all three:**
  - `references/master-tags.md` — the 63 legal tags, verbatim. The only strings allowed in column B.
  - `references/mapping-examples.md` — the worked corpus (~2,300 mapped lines). **§1 (context-dependent
    labor labels) is the single biggest driver of correct mapping** — read it first; map labor by the
    cost-center section, not the label.
  - `references/off-list-tags.md` — what to do with lines that don't fit a master tag (`Ignore` for
    below-NOI, best-guess + column-D flag for everything ambiguous, and the 5 recurring pending-ruling flags).
- **Build engine:** the `Map-*` writers in `Investments/lib/HF-Build-Lib.ps1` (header borders, the 63-tag list in E,
  the col-A `COUNTIF` formula, blue/cream tag styling). Reuse them. (No Python — Excel COM only.)
- **Pipeline position:** **stage 3** of the `.H` build. Reads the **Formatted HF** col C (`code  -  label`)
  from `hf-formatting`; its tags route each HF line in the `.H` assembly. See `../H-PIPELINE-ORCHESTRATION.md`.
- **Memory:** `ps-var-name-case-collision` (don't clobber `$TAG` with `$tag` in map builds),
  `excel-com-write-pitfalls`.
