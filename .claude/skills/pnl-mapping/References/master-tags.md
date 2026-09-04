# Master Tag List (.H / PnL Mapping)

**This list is the canonical mapping vocabulary — a generic senior-housing chart of
accounts of the kind every underwriting shop maintains. The pipeline treats whichever
list is installed in this file as law; a firm adopting the pipeline swaps in its own.**

Rules for using this file:
- The mapping skill may ONLY assign a tag whose text appears exactly below.
- Do not invent, rename, re-spell, merge, or extend any tag.
- If a line cannot be confidently placed on one of these tags, it is FLAGGED in
  the Questions & Comments column, not forced onto an approximate tag.
- The groupings/headers below are a reading aid only. The canonical item is the
  exact tag string, not the group it sits under.

## Revenue
- Rent Revenue - AL
- Rent Revenue - IL
- Rent Revenue - TC
- Rent Revenue - MC
- Rent Concessions
- Other Revenue
- Commercial Lease Revenue
- Rent - Second Person
- Care Revenue - AL
- Care Revenue - MC
- Community Fees

> Note: there is NO `Care Revenue - IL` tag. An IL care/health-services line cannot be
> mapped to a master tag — flag it.

## Labor — Wages
- Wages - Admin
- Wages - Culinary
- Wages - Direct Care
- Wages - Activities
- Wages - Housekeeping
- Wages - Maintenance
- Wages - Marketing

## Labor — Bonuses
- Bonuses - Marketing
- Bonuses - Admin

## Labor — Contract Labor
- Contract Labor - Maintenance
- Contract Labor - Marketing
- Contract Labor - Culinary
- Contract Labor - Activities
- Contract Labor - Admin
- Contract Labor - Direct Care
- Contract Labor - Housekeeping

## Labor — Benefits
- Benefits - Activities
- Benefits - Culinary
- Benefits - Direct Care
- Benefits - Marketing
- Benefits - Housekeeping
- Benefits - Maintenance
- Benefits - Admin

## Operating expense — Departmental
- Raw Food
- Culinary
- Activities
- Direct Care
- COGS
- Maintenance
- Utilities
- Housekeeping
- Marketing
- Marketing - Contract Services
- Marketing - Referral Fees

## Operating expense — Admin / G&A
- Admin
- Admin - Health & Dental
- Admin - Vehicle Lease
- Admin - Travel
- Admin - Telephone
- Admin - Legal
- Admin - Bad Debt
- Admin - IT
- Admin - Bank & Payroll Fees
- Admin - Other Expense

## Insurance & Taxes
- Insurance - Workers Comp
- Real Estate Taxes
- Business Insurance
- Insurance - Other
- Other Taxes

## Other
- Non-Departmental
- Management Fee

## Exclude
- Ignore
