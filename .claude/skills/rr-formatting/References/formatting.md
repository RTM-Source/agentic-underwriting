# .RR Formatting Contract — the visual spec (HIGH PRIORITY)

**Read this before you write a single cell.** Verified cell-by-cell against the golden tab
`Aster Ridge-RR_v1.xlsx` → `Aster Ridge.RR` (2026-07-29; in the private team repo the golden is a
real-deal build, this public clone substitutes the fabricated demo deal). The numbers can be right and
the tab still be *wrong* — this deliverable is read by underwriters and shown to investors, so
formatting fidelity is a **gate, not a polish step**. A build that reconciles to the cent but looks
nothing like the golden tab has failed.

Every colour below is an Excel **BGR long**, not RGB. Write the long. (`16711680` hex-prints `FF0000`
but is **blue**; `255` is **red**. This has bitten us twice.)

---

## 1. GEOMETRY IS RELATIVE — never hardcode column letters

The raw operator export's width **varies by operator and by export**. In the golden the raw report
happens to occupy `A:AV`, so our block starts at `AW`. **That is a coincidence of this file, not the
contract.**

> **The rule:** **`lastRawCol` = the last column of the raw sheet's `UsedRange`** (i.e.
> `UsedRange.Column + UsedRange.Columns.Count - 1`), then **`blockStart = lastRawCol + 1`.**
> Every column below is an **offset from `blockStart`**.

⚠ **`lastRawCol` is NOT "the last cell holding a value."** The OneSite report carries trailing
formatting/furniture past its last populated column, and that space belongs to the report. Concrete, from
the April-2026 O&O Community 1 export: the last cell **with content** is `AQ` (43), but `UsedRange` runs to
**`AV` (48)** — so `blockStart = AW` (49), which is exactly where the golden `Aster Ridge.RR` block sits. Take
the naive reading and you start at `AR`, six columns adrift of the golden and sitting inside the
operator's own report space. **Use `UsedRange`.**

| Offset | Width | Region |
|---|---|---|
| `+0 … +16` | 17 cols | **Per-unit block** (Order → Group Unit Code) |
| `+17, +18` | 2 | spacer (the `+18` column is **hidden**, width 0) |
| `+19 … +21` | 3 | **Tiering helper** (Unit Sqft · Count · Group Unit Code) |
| `+22` | 1 | spacer |
| `+23 … +39` | 17 cols | **Main summary** (Unit Type → NMI vs In Place) |
| `+40` | 1 | spacer |
| `+41 … +44` | 4 | **Move-in box 1** (Min · Avg · Max Rent · #) |
| `+45` | 1 | spacer |
| `+46 … +49` | 4 | **Move-in box 2** |

Sanity-check against the golden (`blockStart = AW = 49`): `+23 → BT` ✓ · `+41 → CL` ✓ · `+46 → CQ` ✓.

**Rows.** The header row for *every* region is **`headerRow` = the raw report's own header row —
the SAME row, so our titles sit beside theirs** *(ruling 2026-07-15, Deal A v3 — SUPERSEDES the
earlier −1 rule; the golden predates this ruling and sits at −1, so do NOT fail a new build
against the golden on this one point)*. Per-unit data rows are **row-aligned to the raw
anchor rows** (golden: 17–388). The analysis blocks run **down from `headerRow`** on their own row plan
(§5) and are *independent* of the unit rows. Move-in window date inputs sit **above** the header, at
`headerRow − 2` and `headerRow − 1`, in the box's **first column**.

> ⚠ **EVERY block gets its header row written — no exceptions.** There are **five** header-bearing
> blocks: (1) the per-unit block, (2) the tiering helper, (3) the main summary, (4) move-in box 1,
> (5) move-in box 2. A block that has **data but no column headers** is a **FAILED build** — the reader
> sees a floating grid of unlabeled numbers, which destroys credibility instantly. This is not
> hypothetical: Deal C's v1 build headered the unit block and both move-in boxes but **left the
> main summary and the tiering helper headerless** (2026-07-14). The gate below checks all five.

---

## 2. TWO header styles — do not mix them

The tab uses **two distinct header treatments**. Getting this wrong is the single most visible error.

| | **Yellow header** | **Navy header** |
|---|---|---|
| Applies to | Per-unit block (`+0…+16`), Tiering helper (`+19…+21`) | Main summary (`+23…+39`), **both** move-in boxes |
| `Interior.Color` | **`13434879`** (#FFFFCC pale yellow) | **`4990985`** (#09284C navy) |
| `Font.Name` / `Size` | **Calibri 11** | **Arial 10** |
| `Font.Bold` | True | True |
| `Font.Color` | `0` (black) | **`16777215`** (white) |
| `WrapText` | True | True |
| Alignment | `xlGeneral` (1) | centre (`-4108`); **first column left (`-4131`)** |

### Row heights *(ruling 2026-07-14)*
- **Header rows = 30** (both styles — the wrapped two-line headers need it).
- **EVERY other row on the tab = 15.** This includes the whole per-unit block, not just the analysis
  rows. Excel's default 12.75 is not acceptable — set 15 explicitly across the used range, then stamp
  the header row(s) back to 30.

---

## 3. Number formats — the zero-as-dash standard

Two families. **Do not use the unit block's `$` formats in the analysis table** — that is the mistake
that makes a build look homemade.

**Per-unit block** (working data — `$0` may show as `$0`):

| Col | Format |
|---|---|
| SqFt `+5` | `#,##0` |
| Capacity `+7`, Occupancy `+8` | `0` |
| Move in Date `+9` | `m/d/yyyy` |
| Market / In-Place Rent / In-Place Care / 2nd Resident `+10…+13` | `$#,##0_);($#,##0)` |
| everything else | `General` |

**Analysis table + move-in boxes** (presentation — zeros render as a dash, per the house standard):

- **Accounting / count / $ / SqFt** — *every* numeric column incl. plain counts:
  `_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)`
- **Percent** (Occ %, (Discount), NMI vs In Place, Utilization, Occupied %):
  `_(#,##0.0%_);(#,##0.0%);_("–"_)_%;_(@_)_%`
  ⚠ the percent format's dash is an **en-dash `–` (U+2013)**, the accounting one is a plain hyphen `-`.
  Copy them verbatim.

Note there are **no `$` signs in the analysis table** — rent figures are bare accounting numbers.

---

## 4. Font colours = meaning (per-unit block is NOT colour-coded)

| Colour | Long | Where |
|---|---|---|
| **Blue** = judgment input | `16711680` | **NMI Rent** on the **tier/leaf rows only** · move-in window **dates** · **Days in period** |
| **Orange** = plug input | `14745600` | the three **Rent Adjustment** cells |
| **Red** = QA | `255` | the whole **Check** block — label *and* value |
| **Black** | `0` | everything else — including **every cell of the per-unit block**, and **NMI Rent on the Total rows** (those are computed, not typed) |

**The per-unit block has no colour coding at all.** Its hardcodes (SqFt, Move-in, Basic/Group Unit Code)
are black like the formulas beside them. Colour lives only in the analysis table.

---

## 5. The analysis row plan (golden exact, `headerRow` = 12)

Offsets are **from `headerRow`**, so this transposes to any raw geometry.

| Row | +n | Content |
|---|---|---|
| 12 | +0 | navy header |
| 13–14 | +1,+2 | IL tier rows (**blank when no IL** — the group still exists) |
| 15 | +3 | **Total IL** (bold, top border) |
| 16 | +4 | blank |
| 17–23 | +5…+11 | **AL tier rows** — one per `Care \| Type \| Tier` (golden: Studio A/B/C, 1BR A/B/C, 2BR A) |
| 24 | +12 | **Total AL** (bold, top border) |
| 25 | | blank |
| 26–27 | | **MC room-type rows** — `MC - Semi-Private (Shared)`, `MC - Private` |
| 28 | | **Total MC** (bold, top border) |
| 29 | | blank |
| 30 | | **Total** (bold, top border) |
| 31 | | blank |
| 32–38 | | **Second Residents** — header + `# 2nd Res` · `Utilization` · `Rent` · `Rent/2nd resident` · `Care` · `Care/2nd resident` |
| 39 | | blank |
| 40–43 | | **Rent Adjustments** — header + `IL / AL / MC Rent Adj` (orange) |
| 44 | | blank |
| 45–50 | | **Check** (all red) — header + `Units/Apartments` · `Occupied Units` · `Occupied %` · `Rent` · `Care` |
| 51 | | blank |
| 52 | | **`Days in period (proration)`** — the blue de-proration denominator |
| 54+ | | **`Notes`** — the anomaly/flag block (rule 20). Header styled like the other block headers; one line per flag, each naming the **unit + row** so the underwriter can jump to it. Omit the block only if there is genuinely nothing to flag. |

### Blank rows and within-group ordering *(rulings 2026-07-15, Deal A v3)*
- **At least one blank row at EVERY care-group transition** (Total IL → first AL row, Total AL → first
  MC row, Total MC → grand Total) — the row plan above shows them, and they are load-bearing for
  readability, not optional. The engine once inserted only the last one; that shipped and got flagged.
- **Within a care group, higher-rent lines sit LOWER**: `MC - Semi-Private` ABOVE `MC - Private`
  (Private's rent is much higher). Ryan confirmed this ordering explicitly. IL/AL `A/B/C` tier order
  stays as-is (A first) unless he rules otherwise.

### Every Total row MUST carry its label
`Total IL` · `Total AL` · `Total MC` · `Total` in the first column, **bold**. A build that writes the
total *values* but leaves the label cell empty is a FAIL — it has shipped (Deal C v1, rows 24/28/30:
correct numbers, no labels). Check the label text, not just the numbers.

### Second Residents is a FOUR-COLUMN block — not one column
Row 32 carries the headers **`IL` | `AL` | `MC` | `Total`** across four value columns, and every row
below (`# 2nd Residents`, `Utilization`, `Rent`, `Rent / 2nd resident`, `Care`, `Care / 2nd resident`)
has **one formula per care level plus a Total**, each filtering on its own care type. On an AL/MC-only
property the **IL column correctly shows dashes** — that is not a bug. (A verifier that reads only the
first value column will see `"IL"` everywhere and wrongly report the block as hardcoded. Read all four.)

### Sub-block title rows are PLAIN BOLD, not banners
Only the true **`headerRow`** (and the tiering-helper / move-in-box headers) get the yellow/navy fill
treatment. The **sub-block title rows inside the analysis** — `Second Residents`, `Rent Adjustments`,
`Check`, and the `Total IL/AL/MC/Total` rows — are **bold text on NO fill** (`Interior.Color` =
`16777215` white/none). `Check` is bold **red** (`255`); the rest are bold black. Do **not** navy-banner
them — a build that does looks wrong. (Verified golden `Aster Ridge.RR` 2026-07-29: `Second Residents`
and `Check` both no-fill bold.)

The tier-count is deal-specific — the **blank-row separators and the block order are the contract**, the
absolute row numbers are not.

⚠ **MC group labels use the *Basic Unit Code* text** (`MC - Private`), not Group-Unit-Code pipe syntax.
AL/IL rows use pipe syntax (`AL | 1BR | A`). Match the golden exactly; they are not stylistically uniform.

### Total-row borders
Bold **+ a TOP border only** (`Borders(8).LineStyle = 1`, xlContinuous). **No bottom border**
(`Borders(9) = xlNone`). The old v2 script sets top *and* bottom — that is wrong.

---

## 6. Column widths (golden, as offsets from `blockStart`)

```
+0   7.9   +1  11.0   +2  11.1   +3  12.7   +4  14.0   +5   9.6   +6  17.6   +7  12.9
+8  14.7   +9  17.1   +10 16.3   +11 16.7   +12 16.6   +13 16.6   +14 31.6   +15 19.7
+16 19.7   +17 19.7   +18  0(hidden)   +19 19.7   +20 19.7   +21 19.7   +22  8.1
+23 23.1   +24  8.9   +25  7.0   +26  8.6   +27  7.9   +28  6.3   +29  6.3   +30  6.3
+31 11.0   +32 11.7   +33 12.6   +34  9.4   +35  6.3   +36  8.7   +37  7.7   +38 12.6
+39 13.9   +40  8.1   +41 10.0   +42  8.1   +43  8.1   +44  8.1   +45  8.1   +46  9.4
+47  8.1   +48  8.1   +49  8.1
```

## 7. Move-in boxes
Header navy (§2). The two window bound dates sit in the box's **first column** at `headerRow−2` /
`headerRow−1`: **blue font**, `m/d/yyyy`, **no fill, not bold, no From/To labels, no borders**. Body cells
carry the analysis number formats (§3) so an empty window shows a dash, not `0`.

## 8. Sheet-level
- **Gridlines OFF** (verified false in the golden) — house standard on every tab we create or touch.
- Tab named `<Community>.RR`.

---

## Formatting gate (run before you declare done)
1. `blockStart == lastRawCol + 1`; no column letter hardcoded anywhere in the build.
2. **All FIVE block headers present and populated** at `headerRow` — unit block, tiering helper, main
   summary, move-in box 1, move-in box 2. A data block with a blank header row is a FAIL. Then: yellow
   header on unit block + tiering helper; navy header on summary + both move-in boxes; correct **font
   family per style** (Calibri 11 / Arial 10); header row height **30**, wrapped.
3. **Row heights: every non-header row = 15** (per-unit block included, not just the analysis).
4. Analysis table uses the **dash** accounting/percent formats — **no `$` signs**, no bare `0` displayed
   (the Check block's `0.00` included — it should render as a dash).
5. Blue = NMI (tier rows only) + window dates + days-in-period; orange = Rent Adj; red = Check block;
   **per-unit block entirely black**.
6. Total rows: bold, **top border only**, and **each carries its label** (`Total IL/AL/MC/Total`).
7. MC group labels name the **real structure**: `MC - Private` / `MC - Semi-Private` / `MC - Jack & Jill`.
8. **Notes block present** at the bottom of the analysis with every flagged anomaly + its unit/row
   (rule 20) — unless there is genuinely nothing to flag.
9. Gridlines off.
10. **Header row = the raw report's own header row** (same row, not −1 — ruling 2026-07-15).
11. **Blank row at every care-group transition** in the summary; **within-group ordering puts
    higher-rent lines lower** (`MC - Semi-Private` above `MC - Private`).
