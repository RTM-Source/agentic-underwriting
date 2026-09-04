# PnL Mapping — Worked Examples Corpus (Reference)

> **Companion to `master-tags.md` and `off-list-tags.md`.** Representative
> `raw line -> master tag` examples distilled from the 10 validated example tabs
> (~2,300 mapped lines). These encode the firm's mapping conventions better than any
> written rule — when a new line resembles one below, map it the same way.
>
> **Read this first, before the per-tag lists:** the single biggest driver of
> correct mapping is **§1 (context-dependent labor labels)**. Most mapping mistakes
> come from tagging a payroll line by its label alone instead of by the cost center
> it sits under.
>
> Examples are shown as the raw `code label` string exactly as it arrives. Account
> codes are operator-specific and vary deal to deal — **match on the label and the
> section, never on the code number.**

---

## 1. Context-dependent labor labels (the central rule)

A large share of every HF is labor, and labor labels are **identical across
departments.** `Payroll Wages - Regular`, `Overtime`, `401K`, `PR Tax / Benefits`,
etc. appear once per cost center with the same text and a different code. The label
alone is never enough — **the department comes from the section/cost-center the line
sits under**, then you pick that department's tag.

**Resolution rule:**

| Label family (any of these) | Tag family | Resolve department by section |
|---|---|---|
| `Salaries - Regular`, `Payroll Wages - Regular/Other/Overtime`, `Overtime`, `Double-time`, `Accrued Salaries` | `Wages - <Dept>` | Culinary, Activities, Direct Care, Housekeeping, Maintenance, Marketing, Admin |
| `PR Tax / Benefits`, `Payroll Taxes`, `401K`, `Benefits - Medical/Vacation/Sick`, `Vacation/Sick/Holiday`, `Paid Time Off`, `Holiday` pay (even labeled `Salaries - Holiday <Dept>`), `Health & Dental` (when departmental) | `Benefits - <Dept>` | same 7 departments |
| `Temporary Labor` | `Contract Labor - <Dept>` | same 7 departments |
| `Bonuses` | `Bonuses - Admin` / `Bonuses - Marketing` **only** | see special case below |

**Special cases — read carefully:**

- **Bonuses have only two tags:** `Bonuses - Admin` and `Bonuses - Marketing`. A
  bonus line under any *other* department (Culinary, Activities, Direct Care,
  Housekeeping, Maintenance) folds into **that department's `Wages - <Dept>`**,
  not a Bonuses tag. (Confirmed in the corpus: `Bonuses` maps to `Wages - Culinary`,
  `Wages - Direct Care`, etc., as well as to `Bonuses - Admin/Marketing`.)
- **Holiday/Vacation/Sick/PTO pay is a BENEFIT, not salary** (standing ruling, Ryan
  2026-07-08): map to `Benefits - <Dept>` even when the label starts with "Salaries"
  (`6120-0101 Salaries - Holiday Food & Beverage` → `Benefits - Culinary`). Regular,
  overtime, double-time, and accrued-salary wages stay `Wages - <Dept>`.
- **Care staff = `Direct Care`.** Nursing/wellness/care cost centers map their
  salaries to `Wages - Direct Care` and benefits to `Benefits - Direct Care`.
  There is no Nursing or Wellness tag.
- **`Health & Dental`** is genuinely ambiguous: as a *departmental* benefit it is
  `Benefits - <Dept>`; as a *company-wide G&A* line it is the dedicated
  `Admin - Health & Dental`. If the section doesn't make it clear, **flag it.**
- If a labor line is in a section whose department you can't determine, **flag it**
  rather than guessing the department.

---

## 2. Canonical examples by master tag

Each master tag with representative raw lines that mapped to it in the validated
corpus. A tag with no example was not exercised in these 10 deals.


### Revenue

**`Rent Revenue - AL`**
- `4101 - AL Revenue - Rent`
- `4104 - Assisted Living Rent - VLI`
- `4001-002: Assisted Living Base Rent`
- `41400-010 Assisted Living Rent`
- `41400-980 Assisted Living Rent - Concession One Time`
- `41400-990 Assisted Living Rent - Concession Recurring`

**`Rent Revenue - IL`**
- `510050: Gross Mkt Rent Potential`
- `510100: Gain/Loss to Lease`
- `510120: Prior Period Rent Adjustment`
- `510200: Vacancy`
- `510300: Models`
- `510350: Employee Units`

**`Rent Revenue - TC`**
- _(no example in the validated corpus — map by definition; flag if unsure)_

**`Rent Revenue - MC`**
- `4102 - Alz Revenue - Rent`
- `4105 - Alzheimer's Rent - VLI`
- `4001-003: Memory Care Base Rent`
- `41600-010 Memory Care Rent`
- `41600-990 Memory Care Rent - Concession Recurring`
- `4700-1120 Gross Potential Rent - GH`

**`Rent Concessions`**
- `510410: Recurring Concessions`
- `521400: Other Concessions`
- `510415: Membership/Flex Plans`
- `510420: Non-recurring Concessions/Concession Adjustments`
- `4108 - Revenue - Rent Concessions`
- `4700-1500 Incentives`

**`Other Revenue`**
- `513050: Garage`
- `513100: Carport`
- `520100: NSF Fees`
- `520150: Late Fees`
- `520250: Initial Pet Fees`
- `520600: Lease Cancellation Fees`

**`Commercial Lease Revenue`**
- `530050: Corporate/Guest Suite Income`
- `530052: Corporate/Guest Suite Vacancy`
- `4305 - Business Tenant Rent`
- `48100-050 Lease Revenue`
- `49070-00 Commercial Rent`

**`Rent - Second Person`**
- `513325: 2nd Occupant Fee`
- `4107 - 2nd Person Fee`
- `41200-900 Independent Rent - Second Occupant`
- `41400-900 Assisted Living Rent - Second Occupant`
- `41500-900 Mezzanine Rent - Second Occupant`
- `41600-900 Memory Care Rent - Second Occupant`

**`Care Revenue - AL`**
- `4200 - AL Care Fees`
- `4203 - Ind Care Fees`
- `4204 - AL Care Concessions`
- `4206 - Ind Care Concessions`
- `4003-001: Assisted Living Care Fees`
- `42400-010 Assisted Living Services Lvl1`

**`Care Revenue - MC`**
- `4201 - Alz Care Fees`
- `4205 - Alz Care Concessions`
- `4003-002: Memory Care Care Fees`
- `42600-010 Memory Care Services Lvl1`
- `42600-020 Memory Care Services Lvl2`
- `42600-030 Memory Care Services Lvl3`

**`Community Fees`**
- `580950: Community Fee`
- `580990: Community Fee Concessions`
- `4109 - Revenue - Community Fee`
- `4109 - Revenue - Community Fee (lease up)`
- `4020-000: Community Fees`
- `43100-020 Community Fee - Independent`


### Labor — Salaries

**`Wages - Admin`**
- `711310: Executive Director`
- `711330: Business Office Director`
- `711340: Concierge- SH`
- `711860: OT- Admin`
- `6901 - Salaries - Regular`
- `6904 - Overtime`

**`Wages - Culinary`**
- `711970: OT- Dining Services`
- `711971: Culinary Services Director`
- `711972: Dining Room Supervisor`
- `711973: Executive Chef`
- `711974: Wait Staff`
- `711975: Bartender`

**`Wages - Direct Care`**
- `6201 - Salaries - Regular`
- `6204 - Overtime`
- `6205 - Double-time`
- `6206 - Bonuses`
- `5101-000: Salaries - Assisted Living`
- `5104-000: Salaries - Overtime`

**`Wages - Activities`**
- `711981: Lifestyle Director`
- `711983: Lifestyle Coordinator`
- `711985: Driver`
- `711988: OT- Programming`
- `6101 - Salaries - Regular`
- `6104 - Overtime`

**`Wages - Housekeeping`**
- `713800: Housekeeper`
- `6701 - Salaries - Regular`
- `6704 - Overtime`
- `6705 - Double-time`
- `6706 - Bonuses`
- `5701-000: Wages - Housekeeping`

**`Wages - Maintenance`**
- `713200: Maintenance Manager`
- `713400: Maintenance Technician`
- `713855: OT Maintenance`
- `6601 - Salaries - Regular`
- `6604 - Overtime`
- `6605 - Double-time`

**`Wages - Marketing`**
- `712050: Leasing Director`
- `6801 - Salaries - Regular`
- `6804 - Overtime`
- `6805 - Double-time`
- `5301-000: Wages - Marketing`
- `61100-010 Sales & Marketing - Payroll Wages`


### Labor — Bonuses

**`Bonuses - Marketing`**
- `6150-0100 - Sales/Commission` *(commission = bonus, not salary)*
- `714100: Bonus - Leasing`
- `6806 - Bonuses`
- `5320-000: Commissions/Bonuses`
- `61100-030 Sales & Marketing - Bonus_Commission`
- `7020-1140 Marketing Commissions`
- `5620 Wages - Marketing Bonus & Commissions`

**`Bonuses - Admin`**
- `714050: Bonus - Administrative`
- `714070: Bonus- Programming`
- `714150: Bonus - Maintenance`
- `714060: Bonus - Dining Services`
- `6906 - Bonuses`
- `60100-030 Administration - Bonus`


### Labor — Temp Labor

**`Contract Labor - Maintenance`**
- `6610 - Temporary Labor`
- `68100-290 Plant Operations - Temp Labor`
- `6080-1150 Outside Labor`

**`Contract Labor - Marketing`**
- `6810 - Temporary Labor`

**`Contract Labor - Culinary`**
- `6010 - Temporary Labor`
- `62100-290 Dining Services - Temp Labor`
- `51303 Temporary Labor - Culinary`
- `6030-1150 Outside Labor`

**`Contract Labor - Activities`**
- `6110 - Temporary Labor`

**`Contract Labor - Admin`**
- `6910 - Temporary Labor`
- `60100-290 Administration - Temp Labor`

**`Contract Labor - Direct Care`**
- `6210 - Temporary Labor`
- `5108-000: Temp Labor - Assisted Living`
- `5208-000: MC Temp Labor`
- `66600-290 Memory Care - Temp Labor`
- `55303 Temporary Labor - A/L`
- `6050-1150 Outside Labor`

**`Contract Labor - Housekeeping`**
- `6710 - Temporary Labor`


### Labor — Benefits

**`Benefits - Activities`**
- `6102 - PR Tax / Benefits`
- `6103 - Vacation/Sick/Holiday`
- `6107 - 401K`
- `6108 - Training Wages`
- `5802-000: Salaries - PR Tax/EE Benefits`
- `5803-000: Salaries - Vac/Sick/Holiday`

**`Benefits - Culinary`**
- `6120-0101 - Salaries - Holiday Food & Beverage` *(holiday pay ruling — "Salaries" label notwithstanding)*
- `714335: Payroll Taxes - Dining Services`
- `714485: Insurance Benefits - Dining Services`
- `714635: 401k Contribution - Dining Services`
- `716400: Vacation`
- `6002 - PR Tax / Benefits`
- `6003 - Vacation/Sick/Holiday`

**`Benefits - Direct Care`**
- `6202 - PR Tax / Benefits`
- `6203 - Vacation/Sick/Holiday`
- `6207 - 401K`
- `6208 - Training Wages`
- `5102-000: Salaries - PR Tax/EE Benefits`
- `5103-000: Salaries - Vac/Sick/Holiday`

**`Benefits - Marketing`**
- `6802 - PR Tax / Benefits`
- `6803 - Vacation/Sick/Holiday`
- `6807 - 401K`
- `6808 - Training Wages`
- `5302-000: Salaries - PR Tax/EE Benefits`
- `5303-000: Salaries - Vac/Sick/Holiday`

**`Benefits - Housekeeping`**
- `6702 - PR Tax / Benefits`
- `6703 - Vacation/Sick/Holiday`
- `6707 - 401K`
- `6708 - Training Wages`
- `5702-000: Salaries - PR Tax/EE Benefits`
- `5703-000: Salaries - Vac/Sick/Holiday`

**`Benefits - Maintenance`**
- `714300: Payroll Taxes - Maintenance`
- `6602 - PR Tax / Benefits`
- `6603 - Vacation/Sick/Holiday`
- `6607 - 401K`
- `6608 - Training Wages`
- `6002-000: Salaries - PR Tax/EE Benefits`

**`Benefits - Admin`**
- `6120-0100 - Salaries - Holiday G&A` *(holiday pay ruling)*
- `714330: Payroll Taxes`
- `714480: Insurance Benefits`
- `714630: 401K Contribution`
- `6902 - PR Tax / Benefits`
- `6903 - Vacation/Sick/Holiday`
- `6907 - 401K`


### Operating expense — Departmental

**`Raw Food`**
- `746220: Resident Meals`
- `746230: Liquor`
- `6012 - Raw Food`
- `6013 - Beverages`
- `62100-510 Dining Services - Food`
- `62100-511 Food Rebates`

**`Culinary`**
- `746221: Dining Service Employee`
- `746223: Internal Functions/ Parties`
- `746240: Kitchen Equipment`
- `746260: Kitchen Repair/Main`
- `746270: F&B Uniforms`
- `746271: F&B Cleaning`

**`Activities`**
- `746420: Resident Activity Supplies`
- `746440: Entertainment`
- `746450: Fitness Instructor`
- `746470: Other Instructor`
- `746520: Fuel`
- `746560: Vehicle Repairs/Maint`

**`Direct Care`**
- `6209 - Workers Comp Insurance`
- `6211 - Supplies / Replacements`
- `6212 - Medical Waste`
- `6360 - Fire related assisted living expenses`
- `6460 - Fire related memory care expenses`
- `5110-000: Supplies`

**`COGS`**
- `6511 - Salon Expenses`
- `6512 - Non-Resident Meal Cost`
- `6513 - Other COGS`
- `6514 - Beverages - Sold`
- `6515 - Personal Care Program Expenses`
- `6516 - Outing's Expense`

**`Maintenance`**
- `725050: Carpet Cleaning`
- `725100: Vinyl Repair`
- `725200: Painting`
- `725400: Cleaning Supplies`
- `725500: Other`
- `730050: Landscape Contract`

**`Utilities`**
- `720100: Electricity-Common Area`
- `720200: Gas-Common Area`
- `720250: Water`
- `720300: Sewer`
- `720350: Trash Removal- Contract`
- `720380: Cable`

**`Housekeeping`**
- `730620: Contract Common Area Cleaning`
- `746320: Housekeeping/Linen Supplies`
- `746340: Housekeeping/Linen Equipment`
- `6709 - Workers Comp Insurance`
- `6711 - Supplies`
- `6712 - Small Equipment`

**`Marketing`**
- `740050: Adv-Printed Media`
- `740160: Website/Portals`
- `740170: Internet Listing Service (ILS)`
- `740180: Marketing Automation Tools`
- `740195: Digital Search Advertising`
- `740200: Radio/TV`

**`Marketing - Contract Services`**
- `6816 - Contract Services / Fees`
- `5325-000: Contract Services/Fees`
- `7020-3500 Design Consultant`

**`Marketing - Referral Fees`**
- `740700: Locator/Broker Fees`
- `6817 - Referral Agency Fees`
- `61100-580 Sales & Marketing - Referral Fees`
- `7020-2500 Referral Fees`
- `6823 - Oak Club`
- `6660 Marketing - Third Party Referral Fees`


### Operating expense — Admin / G&A

**`Admin`**
- `714800: Uniform`
- `745160: Employee Recruiting`
- `745200: Training & Education`
- `745280: Employee Recognition`
- `745320: Office Supplies`
- `745360: Postal/Express Mail`

**`Admin - Health & Dental`**
- `6928 - Health and Dental Benefits`
- `6590-000: Health & Dental`
- `60100-450 Administration - Health Insurance`
- `60100-451 Employee Wellness Program`
- `61100-450 Sales & Marketing - Health Insurance`
- `62100-450 Dining Services - Health Insurance`

**`Admin - Vehicle Lease`**
- `6919 - Auto/ Bus Lease`
- `6475-000: Equipment Rental - ALL`
- `6545-000: Auto/Bus Gas/Reg/Ins/Repair`
- `6550-000: Auto / Bus Lease`
- `6951 Lease - Auto & Bus`

**`Admin - Travel`**
- `745240: Employee Travel/Mileage`
- `745250: Lodging- Admin`
- `745260: Meals - Admin`
- `745261: Entertainment- Admin`
- `6916 - Travel and lodging`
- `6525-000: Travel/Lodging/Air - ALL`

**`Admin - Telephone`**
- `745040: Telephone`
- `745050: Cell Phone`
- `745060: Internet Service`
- `6500-000: Telephone/Cell/DSL Service`
- `69100-070 Utilities - Telephone`
- `69100-080 Utilities - Cell Phone`

**`Admin - Legal`**
- `6931 - Professional Fees`
- `60100-900 Administration - Legal Fees`
- `7010-5701 Professional Fees - Legal`
- `6759 Professional Fee - Legal`
- `79910-94 Accounting Fee Expense`
- `73450-10 Legal and Professional`
- **NOT `Consulting` — see the pinned ruling below.**

**Pinned rulings (Ryan, 2026-07-28 — Deal B regression; these are DETERMINISTIC, do not re-judge):**
- `6626-000: Consulting` → **`Admin`** (general G&A, not legal spend; the pre-2026-07-28 golden's
  `Admin - Legal` was a legacy grouping and is superseded).
- `6597-000: 401K Employee Expense` (employer 401K contribution) → **`Benefits - Admin`** (a true
  employee benefit).
- `6595-000: 401K Admin Expense` (plan administration fee) → **`Admin`** (a G&A cost, NOT a benefit —
  the one 401K-labeled line that does not follow the Benefits rule above).

**`Admin - Bad Debt`**
- `6939 - Bad Debt`
- `4910-000: Bad Debt W/O`
- `60100-980 Administration - Bad Debt Expense`
- `7010-3800 Bad Debt Expense`

**`Admin - IT`**
- `7260-0000 - Marketing Software` *(software ruling — dept section doesn't matter)*
- `7140-0200 - Sales Software`
- `745680: Computer Services & Fees`
- `6922 - Computer Services`
- `6938 - Real Page`
- `6425-000: Dues & Subscriptions - ALL`
- `6565-000: Computer Maint & S/W`
- `60100-600 Administration - Hardware_Software`

**`Admin - Bank & Payroll Fees`**
- `714870: Payroll Fees`
- `745640: Banking Fees/Charges`
- `6920 - Bank / Payroll Charges`
- `6555-000: Bank/Payroll Charges`
- `60100-490 Administration - Payroll Processing Fee`
- `60100-660 Administration - Bank Charges`

**`Admin - Other Expense`**
- `6937 - COVID`


### Insurance & Taxes

**`Insurance - Workers Comp`**
- `714780: Workers Compensation`
- `714685: Workers Compensation - Dining Services`
- `6909 - Workers Comp Insurance`
- `6909 - Workers Comp Claims < $5k`
- `6986 - Workers Comp Reserves`
- `6987 - Workers Comp Premiums`

**`Real Estate Taxes`**
- `7510-1200 - Personal Property Tax` *(property-tax ruling)*
- `760050: Real Estate Property Taxes`
- `6991 - PY Real Estate Tax`
- `6975 - Real Estate Tax`
- `6976 - Business Property Tax`
- `69400-010 Taxes - Real Property Tax`
- `69400-020 Taxes - Personal Property Tax`

**`Business Insurance`**
- `755050: Property Insurance`
- `755055: Automobile Insurance`
- `755200: Other`
- `6972 - IBNR Expense`
- `6974 - Insurance aggregate`
- `6977 - Business Insurance - Property`

**`Insurance - Other`**
- `6971 - GLPL Insurance Expense - Well Captive`
- `6973 - Property Insurance Expense - Well Captive`
- `6971 - GLPL Insurance Expense - Owner Captive`
- `6973 - Property Insurance Expense - Owner Captive`
- `6988 - Operating Expense Rent - Insurance`

**`Other Taxes`**
- `746235: Sales Tax-Liquor and Food Sales`
- `6650-000: Franchise Tax`
- `7300-1000 State/Local Taxes`
- `73200-10 State Franchise Tax`
- `73300-10 Tax - Other`

> One old deal mapped `6300-9100 Personal property Taxes` here — **superseded**: personal/business
> property tax now goes to `Real Estate Taxes` (standing ruling, Ryan 2026-07-08).


### Other

**`Non-Departmental`**
- `6978 - Licensing Fees`
- `6980 - Asset Management Fee`
- `60100-631 Licenses & Fees`
- `69300-010 License Fees - Business License`
- `69300-020 License Fees - Assisted Living License`
- `69300-040 License Fees - Food & Beverage`

**`Management Fee`**
- `750050: Management Fee`
- `750055: Marketing Management Fee`
- `750225: Management Fee- Dining Services`
- `750250: Management Fee -Wellness Services`
- `6665-000: Management Fees`
- `7020 Management Fee-Base`


### Exclude

**`Ignore`**
- `8002-000: Mortgage Interest`
- `8003-000: Amortization Expense`
- `8004-000: Prior Yr/Ownership/Extraordinary Expenses`
- `8007-000: Ask Accountant`
- `8014-000: Legal Fees`
- `8016-000: Rent`


---

## 3. Coverage note

62 of the 63 master tags appear in the validated corpus. The exception is
**`Rent Revenue - TC`** (transitional care) — no example deal exercised it, so there is no
worked precedent. Map a transitional-care rent line to `Rent Revenue - TC` by definition,
and flag it for review since it is unprecedented in the example set.

**Two caveats where the corpus shows human discretion, not a settled rule:**

- **IL care lines.** Some examples fold `Ind Care Fees` / `Ind Care Concessions`
  into `Care Revenue - AL` (shown above). This conflicts with the rule in
  `off-list-tags.md` that IL care lines must be **flagged** because no `Care Revenue - IL`
  tag exists. Treat these corpus rows as *precedent for the best guess*, not a
  green light: map to `Care Revenue - AL`, leave `Check` blank, and note "IL care line; no
  `Care Revenue - IL` tag — inferred." The flag rule wins over the silent fold.
- **Vacancy / GLTL inside a rent block.** The `Rent Revenue - IL` examples include
  `Vacancy` and `Gain/Loss to Lease` lines mapped to the rent tag. That is a real
  convention (these net into the care-type's rent), not a mis-map — preserve their
  signs as received and keep them with the rent tag for that care type.

This corpus should grow: as each new deal is validated, fold a few of its confirmed
`raw line -> tag` rows back into the relevant tag above, especially any that
resolved a previously ambiguous label.
