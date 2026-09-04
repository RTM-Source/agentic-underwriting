> **STATUS 2026-07-06:** partly historical. RESOLVED/OBSOLETE: §1 (doc-vs-reality gap — fixed) and §4
> (recalc.py — no Python on this box; builds go through `lib/HF-Build-Lib.ps1` via Excel COM).
> STILL LIVE: §2 (normalization adapter contract — feeds the planned HF-normalization skill) and §5
> (five open mapping conventions — still need one-time rulings).

# Process Friction Handoff — What Was Slow (HF Format + PnL Map Run)

**Scope:** Running the HF-formatting + PnL-mapping skills across 4 deal files
(Ex-1, Ex-2, Ex-3, Ex-4). This doc captures **only the slow parts and
their causes** — the time sinks, not the wins. Its job is to prioritize the future
**HF normalization / formatting skill** and to feed the notes writeup.

**Status of the run when this was written:** Ex-1 ✅, Ex-2 ✅, Ex-3 ✅
(all validated + saved). Ex-4 ⏳ still pending. Notes deliverable ⏳ pending.

---

## TL;DR — the slow things, ranked by cost

1. **Reverse-engineering the *real* Formatted-HF target** because the skill docs
   describe the wrong artifact. (Biggest single time sink; one-time, now resolved.)
2. **Per-deal bespoke normalizers** — every messy file needed its own hand-built
   state machine because no upstream normalization skill exists. (Recurs per deal.)
3. **Ex-3's Yardi cube dump** — the single hardest file to normalize.
4. **recalc.py invocation friction** — small but repeated, trips up every save.
5. **Open mapping conventions** re-litigated per deal instead of being settled once.

---

## 1. Reverse-engineering the true "Formatted HF" target  ⏱️⏱️⏱️

**What slowed it down:** `SKILL.md` (hf-formatting) and `master-format.md` describe
building the firm's **heavy master-code skeleton** — codes `4100/6001…`, slate banners
`#8F96AF`, accounting-parens number format, Aptos Narrow fonts, total bands spanning
C–Q, *flattening* the operator's two-level subtotals into one total per Master group.
But the actual ground-truth "Formatted HF" example tabs (Ex-1, Ex-2) are
**light-touch**: the operator's own codes, labels, sections, and two-level subtotals
are preserved **verbatim**, with only a column-Q tag string (`=A&" "&B`) and light
styling added.

So the docs point at the wrong deliverable. The heavy skeleton in `master-format.md`
is the **downstream `.H` tab** (built *after* mapping, via XLOOKUP) — not the
Formatted-HF that mapping consumes. `SKILL.md` **conflates the two artifacts.**

**Why it cost time:** had to ignore the written spec, pull the real format off the
example tabs by inspection, and confirm it against two files before trusting it.

**The fix (do once, then it's free):**
- Split the skill docs into two clearly-named targets: **`Formatted HF`** (light-touch,
  operator structure preserved + Q tag) vs. **`.H` tab** (the master skeleton).
- Document the verified Formatted-HF format inline so the next run doesn't rediscover
  it: `A=code, B=label, C–N=12 months, O=Total (=SUM), P=spacer, Q==A&" "&B`; 4-line
  header block in **column A** rows 1–4; row 5 = month headers + "Total", Tahoma 8 bold,
  grey fill `FFD3D3D3`; Tahoma 8 throughout; headers/subtotals bold; gridlines off;
  widths A=11.43 / B=37.14 / C+=12.86; number format `#,##0` (Ex-1 used `#,##0.00`).

---

## 2. Per-deal bespoke normalizers (no upstream normalization skill)  ⏱️⏱️⏱️

**What slowed it down:** Each file's raw shape is different enough that there is no
single "read the HF" routine. Every deal needed a **hand-written state machine** to
walk the sheet, identify section headers / subtotals / data rows, infer care type and
department from context, and set sign conventions — before mapping could even start.
This is the work the **HF normalization skill is supposed to absorb**, but it doesn't
exist yet, so Claude *was* that skill, by hand, four times.

What recurs in every normalizer (i.e. what the skill should own):
- Section-header → care-type / department state machine.
- Labor vs. non-labor toggling (department headers flip it on; `TOTAL … LABOR`
  subtotals flip it off).
- Subtotal detection (`label startswith TOTAL or NET`, or operator `…99` codes).
- NOI-boundary detection → everything below becomes `Ignore`.
- Sign preservation (vacancy/concessions/discounts/GLTL arrive signed — never flip).

**Why it cost time:** the *shared* engine (`oak.py`) was reusable, but the **per-deal
adapter** (which section header means what, which columns hold the trailing 12 months,
how labels are assembled) was net-new each time and is where the hours went.

**The fix:** build the normalization skill as an **adapter contract** — a small,
declarative per-operator config (section→dept map, month-column locator, label-builder,
labor toggles) that feeds one shared walker. Turn the four hand-built state machines
into the first four configs.

---

## 3. Ex-3 — the Yardi analytical-cube dump  ⏱️⏱️⏱️

**What slowed it down:** Ex-3 is a database cube export, not a P&L. Specific drains:
- **No GL/account codes** on most lines — had to synthesize a code from a friendly
  label + a trailing account fragment.
- **KPI / census block (rows 1–175)** mixed in ahead of the financials — had to detect
  and skip it.
- **Hidden duplicate recap rows** (`#hiderow` / `#showrow` tokens in col 1) — below-NOI
  roll-ups (Events-Food-Entertainment, Contracted Services, Printing) that *duplicate*
  above-the-line expenses. Had to keep them out of the operating totals (→ `Ignore`)
  without dropping the real lines.
- **Interleaved 13-month + QTD/YTD columns** — had to locate the correct trailing-12
  window (cols 17–28 = Apr 2025–Mar 2026) among the extra period columns.
- **Department inferred from group codes** (30=Admin, 95=Marketing, 80=Direct Care, …)
  and from **non-labor sub-block totals** (`Total Dining`→Culinary, etc.) rather than
  from clean section headers.

**Why it cost time:** almost none of the normal "find the headers, walk the rows"
assumptions held; nearly every classification needed an Ex-3-specific rule.

**The fix:** flag cube/analytical exports as a **distinct input class** in the
normalization skill — they need their own adapter (token-driven row visibility, group-code
dept map, period-window locator). Don't try to force them through the operator-P&L path.

**Residual judgment call left flagged (correctly, not forced):** `Mgmt Fee-Reimbursables`
sits in the management-fee zone above final NOI — left flagged for human review rather
than guessed onto `Management Fee`.

---

## 4. recalc.py invocation friction  ⏱️

**What slowed it down:** `recalc.py` must be run **from its own directory** because it
imports `office.soffice`; the local copy at `<sandbox>/scripts/recalc.py`
**fails** (missing `office` module). Correct invocation:
```
python3 recalc.py <filepath> 60
```
Hit this once per workbook save; each failure → diagnose → re-run.

**The fix:** bake the exact invocation into the skill's build procedure (one line), and
stop keeping a local copy that shadows it and fails.

---

## 5. Open mapping conventions re-litigated per deal  ⏱️

**What slowed it down:** the same handful of unsettled conventions surfaced as flags on
every deal, each needing a fresh judgment call:
- **Company-wide benefits** with no department → currently `Benefits - Admin` + flag.
- **Respite / short-term revenue** → currently `Other Revenue` + flag.
- **Telephone/internet inside a Utilities section** → currently `Admin - Telephone` + flag.
- **Care-type inference when unstated** → inferred (often AL) + flag.
- **Single combined "Care" revenue** (no IL/AL/MC split) → `Care Revenue - AL` + flag.

These aren't bugs — flagging is the designed behavior — but because they're **unresolved
firm conventions**, they re-appear every run and each one is a small decision tax.

**The fix:** get a one-time ruling from Ryan on each, fold the rulings into the mapping
`SKILL.md` as explicit rules, and the flags drop to genuine per-deal ambiguity only.

---

## What was NOT slow (for contrast / so we don't "fix" what works)

- **Clean operator files (Ex-1, Ex-2)** were fast and high-confidence —
  Ex-1 hit a 0-diff structural match to the reference; Ex-2 mapped with 0 flags.
- **The shared engine (`oak.py`)** paid off — written once, reused across all deals; the
  mapping logic itself was not the bottleneck. The bottleneck was *getting each file into
  a shape the engine could read.*

---

## One-line conclusion for the notes writeup

> The mapping step is solid; the time sink is **normalization** (no upstream skill yet)
> plus a **doc-vs-reality gap** on the Formatted-HF target. Build the normalization skill
> as a per-operator adapter contract, split the Formatted-HF vs. `.H` specs, and settle
> the five open conventions once — and most of this run's slowness disappears.
