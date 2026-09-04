# HF Master Format — `.H` Tab Reference

> **Authoritative target layout for the HF (`.H`) Formatting Skill.**
> Reproduced verbatim from `HF Master Format` (the canonical empty
> template). This file defines the *fixed skeleton* the formatting skill conforms
> a deal's mapped historicals into: the section order, every account code + label,
> the subtotal/total hierarchy, the column layout, and the visual formatting.
>
> **Do not invent, rename, re-spell, merge, or reorder** any code, label, or
> section. Firm spellings and quirks (`Alz`, `Outing's`, singular
> `TOTAL HOUSEKEEPING EXPENSE`, the duplicated `6960`) are intentional — match
> them exactly. When a deal has no data for a group, **omit that group** rather
> than printing a zero block (per project convention); never add a group that
> isn't below.

---

## 1. Column layout & conventions

The sheet is laid out left-to-right as follows. Structural columns (A, S, T) are
scaffolding the display column (C) is built from; they are not part of the
printed statement.

| Col | Role | Notes |
|-----|------|-------|
| **A** | Structural marker | `x` on every banner and total row; `t` on `TOTAL CORPORATE EXPENSES` (row 353); `x` again on the footer row (356). Drives which rows get banded. |
| **B** | (unused) | left blank |
| **C** | **Display line** | The printed line. For account lines it is a formula `=S{row}&"  -  "&T{row}` — code and label joined by **two-space / hyphen / two-space**. Banner and total rows hold literal text. Column width ≈ **58**. |
| **D** | Spacer | Blank, but **carries banner fill and total-band border** (see §3). Part of the printed band, so format it contiguously with C. |
| **E … P** | Monthly columns | 12 trailing months. **E = period start, P = period end** (header row labels them `Date Start` / `Date End `). Accounting number format (§3). |
| **Q** | **Total** | Row total across E:P. Same accounting format. |
| **R** | (unused) | left blank |
| **S** | Account **code** | Numeric (e.g. `4100`). Source for the C formula. |
| **T** | Account **label** | Text (e.g. `IL Revenue - Rent`). Source for the C formula. |

**Line-display rule:** `C = S & "  -  " & T`. Note the joiner is *two* spaces
either side of the hyphen, distinct from single-hyphen separators that appear
*inside* labels (e.g. `IL Revenue - Rent`).

---

## 2. Header block (rows 2-5)

| Row | Cell | Content (template placeholder) | Format |
|-----|------|-------------------------------|--------|
| 2 | C2 | `Community Name (######)` | Tahoma 8, grey `#505050`, centered — **fill in the real community name + ID per deal** |
| 3 | C3 | `Statement (12 months)` | Tahoma 12, **bold**, centered |
| 4 | C4 | `Period = Month 20XX - Month 20XX` | set to the actual T12 window |
| 5 | C5 | `Book = Accrual ; Tree = wlf_cf` | provenance line |

---

## 3. Formatting spec

Match the template exactly; existing template conventions override any house style.

**Fonts**
- Account line items (column C formula result): **Arial 10**.
- Structural rows — banners, totals, NOI, and the S/T scaffold: **Aptos Narrow 11**.
- Header block: **Tahoma** (8 for the community line, 12 bold for the statement line).

**Banners** (`REVENUE` row 6, `EXPENSES` row 64)
- Bold, **slate fill `#8F96AF`**.
- Fill spans **C through Q, including spacer column D** — one contiguous band.

**Total band** (every `TOTAL …` row, plus `NET OPERATING INCOME` and `NET INCOME`)
- Bold, **thin top + bottom border**, no fill.
- The band spans **C through Q, including spacer column D** — extend it contiguously; don't leave D un-bordered.

**Data cells (E:Q)**
- Accounting number format: `_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)`
  — comma-grouped, **negatives in parentheses**, **zero shown as `-`**, no decimals.
- **Preserve source signs exactly.** Vacancy, concessions, discounts, marketing
  incentives, and GLTL arrive signed; do not flip them.

**`Margin` row (319)** sits directly under NOI, left-aligned, same accounting format
(it holds NOI ÷ Total Revenue).

---

## 4. Section & subtotal hierarchy

The statement nests as below. Each leaf block sums into its `TOTAL`, and the
section totals roll up to the bold summary lines.

```
REVENUE  (banner)
  Rent revenue              -> TOTAL RENT REVENUE
  Care revenue              -> TOTAL CARE REVENUE
  Other revenue             -> TOTAL OTHER REVENUE
                            => TOTAL REVENUE
EXPENSES (banner)
  Culinary                  -> TOTAL CULINARY EXPENSES
  Activities                -> TOTAL ACTIVITIES EXPENSES
  Independent               -> TOTAL INDEPENDENT EXPENSES
  Assisted Living           -> TOTAL ASSISTED LIVING EXPENSES
  Alzheimers (MC)           -> TOTAL ALZHEIMERS EXPENSES
  COGS                      -> TOTAL COGS EXPENSES
  Maintenance               -> TOTAL MAINTENANCE EXPENSES
  Utilities                 -> TOTAL UTILITIES EXPENSES
  Housekeeping              -> TOTAL HOUSEKEEPING EXPENSE   (note: singular)
  Marketing                 -> TOTAL MARKETING EXPENSES
  Administration            -> TOTAL ADMINISTRATION EXPENSES
  Non-Departmental          -> TOTAL NON-DEPARTMENTAL EXPENSES
                            => TOTAL OPERATING EXPENSES
=> NET OPERATING INCOME  (+ Margin)
  Other (Income) Expenses   -> TOTAL OTHER (INCOME) EXPENSES
=> NET INCOME
  Corporate Expenses        -> TOTAL CORPORATE EXPENSES   (below the line; A-marker = "t")
```

---

## 5. Canonical line structure (verbatim)

Each line is `code — label`, in sheet order. Section headings here correspond to
the `TOTAL …` that closes each block.



## REVENUE

### TOTAL RENT REVENUE

- `4100` — IL Revenue - Rent
- `4101` — AL Revenue - Rent
- `4102` — Alz Revenue - Rent
- `4103` — Independent Rent - VLI
- `4104` — Assisted Living Rent - VLI
- `4105` — Alzheimer's Rent - VLI
- `4108` — Revenue - Rent Concessions
- `4111` — Concessions due to water damage

### TOTAL CARE REVENUE

- `4200` — AL Care Fees
- `4201` — Alz Care Fees
- `4203` — Ind Care Fees
- `4206` — Ind Care Concessions

### TOTAL OTHER REVENUE

- `4106` — Non-Refundable Entrance Fee Recognition
- `4107` — 2nd Person Fee
- `4109` — Revenue - Community Fee
- `4110` — Cares Act Revenue
- `4112` — IL Meals
- `4130` — Application Fee
- `4320` — Vacancy Loss
- `4113` — Other Income - Occ
- `4121` — Below Market Lease Intangible Revenue
- `4125` — Non-Refundable Entrance Fee Revenue Post Acquisition
- `4300` — Revenue - Salon, Spa and Massage
- `4301` — Revenue - Guest Meals
- `4302` — Transportation and Escort Fees
- `4303` — Room Service Trays
- `4304` — Late Fees
- `4305` — Business Tenant Rent
- `4306` — Beverages
- `4307` — Outing's Revenue
- `4308` — Apartment Upgrades
- `4309` — Catering / Private Event
- `4310` — Misc Services Sales
- `4311` — Misc Tangible Sales
- `4312` — Revenue - Respite Revenue
- `4313` — Incontinence Care Program
- `4315` — Miscellaneous Revenue - Other
- `4316` — Rebates
- `4317` — Bad Debt
- `4318` — Pet Fees
- `4319` — Gain/Loss to Lease
- `4350` — Clubhouse Management Fee Revenue
- `4360` — Fire related respite
- `4001` — Management Fee Income
- `4520` — Gain/Loss on Sale

**» TOTAL REVENUE** — roll-up of the subtotals above


## EXPENSES

### TOTAL CULINARY EXPENSES

- `6001` — Salaries - Regular
- `6002` — PR Tax / Benefits
- `6003` — Vacation/Sick/Holiday
- `6004` — Overtime
- `6005` — Double-time
- `6006` — Bonuses
- `6007` — 401K
- `6008` — Training Wages
- `6010` — Temporary Labor
- `6011` — Linen Service
- `6012` — Raw Food
- `6013` — Beverages
- `6014` — Supplies / Replacements
- `6015` — Replacements
- `6016` — Equipment Rental
- `6050` — Acquisition expense
- `6060` — Fire related culinary expenses

### TOTAL ACTIVITIES EXPENSES

- `6101` — Salaries - Regular
- `6102` — PR Tax / Benefits
- `6103` — Vacation/Sick/Holiday
- `6104` — Overtime
- `6105` — Double-time
- `6106` — Bonuses
- `6107` — 401K
- `6108` — Training Wages
- `6110` — Temporary Labor
- `6111` — IL Resident Programs / Special Events
- `6112` — IL Entertainment
- `6113` — IL Supplies
- `6114` — Subscriptions
- `6115` — AL Resident Programs / Special Events
- `6116` — AL Entertainment
- `6117` — AL Supplies
- `6118` — Alz Resident Programs / Special Events
- `6119` — Alz Entertainment
- `6120` — Alz Supplies
- `6150` — Acquisition expense
- `6160` — Fire related activities expenses

### TOTAL INDEPENDENT EXPENSES

- `6201` — Salaries - Regular
- `6202` — PR Tax / Benefits
- `6203` — Vacation/Sick/Holiday
- `6204` — Overtime
- `6205` — Double-time
- `6206` — Bonuses
- `6207` — 401K
- `6208` — Training Wages
- `6210` — Temporary Labor
- `6211` — Supplies / Replacements
- `6212` — Medical Waste
- `6260` — Fire related Independent expenses

### TOTAL ASSISTED LIVING EXPENSES

- `6301` — Salaries - Regular
- `6302` — PR Tax / Benefits
- `6303` — Vacation/Sick/Holiday
- `6304` — Overtime
- `6305` — Double-time
- `6306` — Bonuses
- `6307` — 401K
- `6308` — Training Wages
- `6310` — Temporary Labor
- `6311` — Supplies / Replacements
- `6312` — Medical Waste
- `6360` — Fire related assisted living expenses

### TOTAL ALZHEIMERS EXPENSES

- `6401` — Salaries - Regular
- `6402` — PR Tax / Benefits
- `6403` — Vacation/Sick/Holiday
- `6404` — Overtime
- `6405` — Double-time
- `6406` — Bonuses
- `6407` — 401K
- `6408` — Training Wages
- `6410` — Temporary Labor
- `6411` — Supplies / Replacements
- `6412` — Medical Waste
- `6460` — Fire related memory care expenses

### TOTAL COGS EXPENSES

- `6511` — Salon Expenses
- `6512` — Non-Resident Meal Cost
- `6513` — Other COGS
- `6514` — Beverages - Sold
- `6515` — Personal Care Program Expenses
- `6516` — Outing's Expense
- `6517` — Apartment Upgrades
- `6518` — Catering / Private Event

### TOTAL MAINTENANCE EXPENSES

- `6601` — Salaries - Regular
- `6602` — PR Tax / Benefits
- `6603` — Vacation/Sick/Holiday
- `6604` — Overtime
- `6605` — Double-time
- `6606` — Bonuses
- `6607` — 401K
- `6608` — Training Wages
- `6609` — Workers Comp Insurance
- `6610` — Temporary Labor
- `6611` — Maintenance Uniforms
- `6612` — Scheduled Maintenance - Other
- `6613` — Repairs and Maintenance
- `6614` — Maintenance Supplies
- `6615` — Landscaping / Tree Maint
- `6616` — Painting and Decorating
- `6617` — Room Turns
- `6623` — Turnover Carpet Clean/Repair
- `6626` — Turnover Painting
- `6660` — Natural Disaster Damage/Restoration

### TOTAL UTILITIES EXPENSES

- `6618` — Cable
- `6619` — Electric
- `6620` — Gas
- `6621` — Trash
- `6622` — Water / Sewer

### TOTAL HOUSEKEEPING EXPENSE

- `6701` — Salaries - Regular
- `6702` — PR Tax / Benefits
- `6703` — Vacation/Sick/Holiday
- `6704` — Overtime
- `6705` — Double-time
- `6706` — Bonuses
- `6707` — 401K
- `6708` — Training Wages
- `6709` — Workers Comp Insurance
- `6710` — Temporary Labor
- `6711` — Supplies
- `6712` — Small Equipment
- `6713` — Linen Replacement
- `6760` — Fire related housekeeping expenses

### TOTAL MARKETING EXPENSES

- `6801` — Salaries - Regular
- `6802` — PR Tax / Benefits
- `6803` — Vacation/Sick/Holiday
- `6804` — Overtime
- `6805` — Double-time
- `6806` — Bonuses
- `6807` — 401K
- `6808` — Training Wages
- `6810` — Temporary Labor
- `6811` — Advertising
- `6812` — Marketing Software
- `6813` — Website
- `6814` — Printing and Copying
- `6815` — Brochures / Collaterals
- `6816` — Contract Services / Fees
- `6817` — Referral Agency Fees
- `6818` — Direct Mail
- `6819` — Postage
- `6820` — Marketing Supplies
- `6821` — Respite Expense
- `6822` — Networking
- `6823` — Oak Club
- `6824` — Signs / Banners
- `6825` — Special Events
- `6826` — Specialty Advertising
- `6827` — Move-in/Tour/Referral
- `6828` — Signage
- `6829` — Marketing Assessment
- `6830` — Yellow Pages
- `6860` — Fire related marketing expenses

### TOTAL ADMINISTRATION EXPENSES

- `6901` — Salaries - Regular
- `6902` — PR Tax / Benefits
- `6903` — Vacation/Sick/Holiday
- `6904` — Overtime
- `6905` — Double-time
- `6906` — Bonuses
- `6907` — 401K
- `6908` — Training Wages
- `6909` — Workers Comp Claims < $5k
- `6910` — Temporary Labor
- `6911` — Family Fund
- `6912` — Dues and Subscriptions
- `6913` — Education and Training
- `6914` — Equipment Rental
- `6915` — Telephone
- `6916` — Travel and lodging
- `6917` — Meals
- `6918` — Auto - Gas, Reg, Rprs, Z permits
- `6919` — Auto/ Bus Lease
- `6920` — Bank / Payroll Charges
- `6921` — Classified Ads / Recruiting
- `6922` — Computer Services
- `6923` — Employee Testing
- `6924` — Contributions
- `6925` — Employee Appreciation
- `6926` — Resident Retention
- `6927` — Flowers
- `6928` — Health and Dental Benefits
- `6929` — Office Supplies
- `6930` — Postage
- `6931` — Professional Fees
- `6932` — Consulting
- `6933` — Uniforms (All)
- `6934` — Conferences
- `6935` — Corporate Meetings
- `6936` — Associations/Memberships/Subscriptions
- `6937` — COVID
- `6938` — Real Page
- `6939` — Bad Debt
- `6940` — Security Service
- `6941` — Sustainability
- `6943` — Bank Fees
- `6950` — Acquisition expense
- `6955` — Transitions Expense
- `6960` — Fire related admin expenses
- `6960` — Emergency Preparedness
- `6987` — Workers Comp Premiums
- `6990` — Clubhouse Management Fee

### TOTAL NON-DEPARTMENTAL EXPENSES

- `6971` — GLPL Insurance Expense - Owner Captive
- `6974` — Insurance aggregate
- `6973` — Property Insurance Expense - Owner Captive
- `6975` — Real Estate Tax
- `6991` — PY Real Estate Tax
- `6976` — Business Property Tax
- `6977` — Business Insurance - Property
- `6978` — Licensing Fees
- `6979` — Management Fee
- `6980` — Asset Management Fee
- `6983` — Business Insurance - GLPL Premium
- `6984` — Business Insurance - GL/EPLI Reserves
- `6985` — Business Insurance - Other
- `6988` — Operating Expense Rent - Insurance
- `6989` — Construction Management Fee

**» TOTAL OPERATING EXPENSES** — roll-up of the subtotals above

**» NET OPERATING INCOME** — roll-up of the subtotals above

> `NET OPERATING INCOME` is immediately followed by a **Margin** row (NOI ÷ Total Revenue).

### TOTAL OTHER (INCOME) EXPENSES

- `7001` — Franchise Tax
- `7002` — Other Income and expense
- `7003` — Non-Recurring items
- `7004` — Gain/Loss on disposal of assets
- `4015` — Insurance Recoveries - fire damage
- `4016` — Insurance Recoveries - water damage
- `5060` — Business Interruption Expense
- `5061` — Temporary Relocation Costs
- `5063` — Fire related travel costs
- `5064` — Fire related remediation costs
- `5065` — Fire related evacuation costs
- `5066` — Fire related resident credits
- `5068` — Management fee not paid due to fires
- `5070` — Water Damage remediation costs

**» NET INCOME** — roll-up of the subtotals above

### TOTAL CORPORATE EXPENSES

- `5024` — Accounting
- `5027` — Consulting
- `5041` — Travel and Lodging
- `5044` — Office Supplies
- `5050` — Lease Expense
- `5051` — Interest Mortgage
- `5054` — Taxes
- `5058` — Tax Prep Fees
- `5093` — IC Preferred Int Exp
- `5903` — Depreciation
- `6981` — Incentive management fee
- `—` — Placeholder (spare line; keep in place)

---

## 6. Quirks to preserve (do not "fix")

- **Verbatim spellings:** `Alz` (not "Alzheimer's") in the activities/MC line
  labels, `Outing's Revenue` / `Outing's Expense`, `Workers Comp Claims < $5k`,
  `Auto - Gas, Reg, Rprs, Z permits`. Keep apostrophes, abbreviations, and symbols as-is.
- **`TOTAL HOUSEKEEPING EXPENSE` is singular** while every other total is plural.
- **Duplicate code `6960`** appears on two consecutive Admin rows —
  `Fire related admin expenses` and `Emergency Preparedness`. Both are retained;
  do not dedupe.
- **Utilities codes `6618`-`6622`** are numerically inside the Maintenance `66xx`
  range but are **split into their own Utilities section**. Section membership
  follows the template, not the numeric code order.
- **Per-department payroll stub.** Each labor department repeats the same opening
  lines (`Salaries - Regular`, `PR Tax / Benefits`, `Vacation/Sick/Holiday`,
  `Overtime`, `Double-time`, `Bonuses`, `401K`, `Training Wages`, `Temporary
  Labor`) under its own `6X0x` prefix. Identical labels, different codes/sections
  — the department is fixed by the code prefix and the section it sits in.
- **Corporate Expenses sits *below* `NET INCOME`** and is marked with `t` (not `x`)
  in column A. It is below-the-line; treat accordingly.
- **`Placeholder` row** (code-less) is a deliberate spare line — keep it in place.
- **Footer row 356** carries `x` markers (A/C/E/P/S) bounding the printed range.

---

## 7. How the formatting skill should use this

1. Start from this fixed skeleton (sections, codes, labels, order).
2. Drop the deal's mapped monthly values into E:P against the matching account
   line; let Q total across the row and the `TOTAL …` rows sum their blocks.
3. **Omit** any group with no data rather than printing a zero block.
4. Apply the §3 formatting: banners, total bands across C-Q (incl. D), accounting
   number format, header block filled with the real community/period.
5. Surface judgment calls (a deal line with no clean home here, an ambiguous
   section, a sign that looks inverted) as **cell comments + chat flags** — never
   silently resolve.
