# PnL Mapping — Off-List & Flag Rules

> **Companion to `master-tags.md` (the only legal tags) and `mapping-examples.md` (worked precedent).**
> This file answers: *what do you do with a line that doesn't sit cleanly on a master tag?* The answer is
> almost always **best-guess tag + flag**, never invent a tag. Flagging is the skill doing its job.

## The core rule
Every value in column B must be a tag that appears **verbatim** in `master-tags.md` (match the firm's
spelling exactly, including `Commercial Lease Revenue`). A line you can't place confidently is **flagged**,
not forced onto an approximate tag.

## How a flag surfaces (don't expect a highlight)
Column A `Check` is a `COUNTIF` validation — it returns **"Yes" for any valid master tag** (including
`Ignore`), so it does **not** distinguish flags and nothing is highlighted. **The flag is a non-empty note
in column D (`Questions & Comments`).** Confident rows leave D blank. Keep notes terse (drop articles, use
`→`): "IL care line; no Care-IL tag — inferred", "company-wide pooled; no dept Benefits tag".

## Lines that are legitimately off the operating list → `Ignore`
`Ignore` **is** a master tag (the only off-list "answer"), for below-NOI / non-operating lines:
mortgage interest, depreciation, amortization, income tax, owner/extraordinary items, gain/loss on sale,
below-the-line corporate allocations. If you can't tell whether a line falls below NOI, **flag it** rather
than guessing the boundary.

## Lines with no clean master home → best guess + flag
- **IL care / health-services** — there is **no `Care Revenue - IL`** tag. Best guess `Care Revenue - AL`, leave `Check`
  effect aside (the tag is valid so A shows "Yes"), and note "IL care line; no Care-IL tag — inferred."
  (The corpus sometimes folds `Ind Care Fees` → `Care Revenue - AL`; treat that as precedent for the *guess*, not a
  green light to drop the flag — the flag rule wins over a silent fold.)
- **Company-wide / pooled benefits with no department** — no company-wide Benefits tag exists. Best guess
  `Benefits - Admin`, flag "company-wide pooled; no dept Benefits tag."
- **`Health & Dental`** — departmental → `Benefits - <Dept>`; company-wide G&A → `Admin - Health & Dental`;
  if the section doesn't make it clear, **flag it**.
- **Any labor line whose department you can't read from its section** — flag rather than guess the dept
  (labor labels are identical across departments — see `mapping-examples.md` §1).

## Standing rulings (locked — these are NOT flags)
Ruled by Ryan 2026-07-08 during the O&O v16 mapping pass. Apply silently; column D stays blank:
1. **Holiday pay → `Benefits - <Dept>`** — grouped with PTO/Vacation/Sick, never `Wages - <Dept>`,
   even when the operator labels it `Salaries - Holiday <Dept>`. Department resolves from the label
   suffix / section as usual (`Holiday Food & Beverage` → `Benefits - Culinary`, `Holiday G&A` →
   `Benefits - Admin`, `Holiday SNF` under a nursing block → `Benefits - Direct Care`).
2. **Software → `Admin - IT`** wherever it sits: `Marketing Software`, `Sales Software`, `Admin Software`,
   `Software Services`, computer/hardware-software lines. Media spend is NOT software — Web Services,
   Digital Advertising, SEO/Content/Newsletter stay `Marketing`.
3. **Personal property tax → `Real Estate Taxes`** — all property taxes (real, personal, business) pool in
   `Real Estate Taxes`; `Other Taxes` keeps sales tax, franchise tax, misc state/local taxes.

## Recurring judgment-call flags (pending a standing ruling)
These five recur on nearly every deal as the same call. They stay **flags** (current de-facto defaults
below) until firm leadership issues a standing ruling, at which point the default folds into
`SKILL.md` as a rule and drops off the per-deal flag list. (Tracked in `Investments/Docs/Running Notes for
UW.Skill.docx` → "Open Mapping Conventions".)
1. Pooled benefits, no department → `Benefits - Admin` + flag.
2. Respite / short-term-stay revenue → `Other Revenue` + flag.
3. Telephone / internet sitting inside a Utilities section → `Admin - Telephone` + flag.
4. Care type unstated and not inferable from the section → inferred (usually AL) + flag.
5. Single combined `Care` revenue line with no IL/AL/MC split → `Care Revenue - AL` + flag.

## Signs
Preserve source signs exactly — vacancy, concessions, discounts, GLTL normally arrive negative; gain/loss
is signed. If a sign looks inverted vs. expectation, **flag it** — never silently flip it.

> Net: the master list is closed and the spellings are fixed; your only "escape hatches" are `Ignore`
> (for genuinely non-operating lines) and a **column-D flag** (for everything ambiguous). Reach for the
> flag freely — surfacing the ~10% an underwriter must eyeball is the whole point.
