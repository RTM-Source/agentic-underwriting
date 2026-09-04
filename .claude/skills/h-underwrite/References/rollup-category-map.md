# Roll-up Category Map + Metrics — the FIXED `.H` skeleton

> Verbatim from the Aster Ridge Demo golden (`Aster Ridge.H Model_v1.xlsx`). Blocks C and D of `.H` are the **same every deal** — only the
> SUMIFS detail range (`$25:$detailLast`) and a couple of cross-tab anchors shift. Write this block as
> a template. The **dual-column trick is load-bearing**: col **B = display name** (can repeat), col
> **C = the exact master tag key** (the SUMIFS criterion — must match a tag verbatim, incl.
> `Commercial Lease Revenue`). Don't swap B and C.

Notation: `D` = `detailLast` (last detail row; golden 218). Category-row formula (cols F:Q):
`=+SUMIFS(F$25:F$D,$B$25:$B$D,$C{row})`. Subtotals as noted.

## Roll-up block (rows 221–296)

| Row | B (display) | C (tag key) | Formula |
|----|-------------|-------------|---------|
| 221 | AL Rent | `Rent Revenue - AL` | SUMIFS |
| 222 | IL Rent | `Rent Revenue - IL` | SUMIFS |
| 223 | TC Rent | `Rent Revenue - TC` | SUMIFS |
| 224 | MC Rent | `Rent Revenue - MC` | SUMIFS |
| 225 | Rent Concessions | `Rent Concessions` | SUMIFS |
| 226 | Other Revenue | `Other Revenue` | SUMIFS |
| 227 | Other Revenue | `Commercial Lease Revenue` | SUMIFS |
| 228 | Other Revenue | `Rent - Second Person` | SUMIFS |
| 229 | AL Care | `Care Revenue - AL` | SUMIFS |
| 230 | MC Care | `Care Revenue - MC` | SUMIFS |
| 231 | Community Fees | `Community Fees` | SUMIFS |
| **232** | | **Total Revenue** | `=+SUM(F221:F231)` |
| 233 | | | *(blank)* |
| **234** | Check | | `=F232-'<CN> Historical Financials'!E<hfRevRow>` |
| 235 | | | *(blank)* |
| 236 | Culinary | `Wages - Culinary` | SUMIFS |
| 237 | Activities | `Wages - Activities` | SUMIFS |
| 238 | Direct Care | `Wages - Direct Care` | SUMIFS |
| 239 | Maintenance | `Wages - Maintenance` | SUMIFS |
| 240 | Housekeeping | `Wages - Housekeeping` | SUMIFS |
| 241 | Marketing | `Wages - Marketing` | SUMIFS |
| 242 | Admin | `Wages - Admin` | SUMIFS |
| 243 | Marketing | `Bonuses - Marketing` | SUMIFS |
| 244 | Admin | `Bonuses - Admin` | SUMIFS |
| **245** | | **In-House Labor** | `=+SUM(F236:F244)` |
| 246 | | | *(blank)* |
| 247 | Culinary | `Contract Labor - Culinary` | SUMIFS |
| 248 | Activities | `Contract Labor - Activities` | SUMIFS |
| 249 | Direct Care | `Contract Labor - Direct Care` | SUMIFS |
| 250 | Maintenance | `Contract Labor - Maintenance` | SUMIFS |
| 251 | Housekeeping | `Contract Labor - Housekeeping` | SUMIFS |
| 252 | Marketing | `Contract Labor - Marketing` | SUMIFS |
| 253 | Admin | `Contract Labor - Admin` | SUMIFS |
| **254** | | **Total Labor** | `=+SUM(F245:F253)` |
| 255 | | | *(blank)* |
| 256 | Culinary | `Benefits - Culinary` | SUMIFS |
| 257 | Activities | `Benefits - Activities` | SUMIFS |
| 258 | Direct Care | `Benefits - Direct Care` | SUMIFS |
| 259 | Maintenance | `Benefits - Maintenance` | SUMIFS |
| 260 | Housekeeping | `Benefits - Housekeeping` | SUMIFS |
| 261 | Marketing | `Benefits - Marketing` | SUMIFS |
| 262 | Admin | `Benefits - Admin` | SUMIFS |
| **263** | | **Total Labor & Benefits** | `=+SUM(F254:F262)` |
| 264 | | | *(blank)* |
| 265 | Culinary | `Culinary` | SUMIFS |
| 266 | Culinary | `Raw Food` | SUMIFS |
| 267 | Activities | `Activities` | SUMIFS |
| 268 | Direct Care | `Direct Care` | SUMIFS |
| 269 | Non-Departmental | `COGS` | SUMIFS |
| 270 | Maintenance | `Maintenance` | SUMIFS |
| 271 | Utilities | `Utilities` | SUMIFS |
| 272 | Housekeeping | `Housekeeping` | SUMIFS |
| 273 | Marketing | `Marketing` | SUMIFS |
| 274 | Marketing | `Marketing - Contract Services` | SUMIFS |
| 275 | Marketing | `Marketing - Referral Fees` | SUMIFS |
| 276 | Admin | `Admin` | SUMIFS |
| 277 | Admin | `Admin - Legal` | SUMIFS |
| 278 | Admin | `Admin - Bad Debt` | SUMIFS |
| 279 | Admin | `Admin - Other Expense` | SUMIFS |
| 280 | Admin | `Insurance - Workers Comp` | SUMIFS |
| 281 | Admin | `Admin - Health & Dental` | SUMIFS |
| 282 | Admin | `Admin - IT` | SUMIFS |
| 283 | Admin | `Admin - Telephone` | SUMIFS |
| 284 | Admin | `Admin - Travel` | SUMIFS |
| 285 | Admin | `Admin - Bank & Payroll Fees` | SUMIFS |
| 286 | Admin | `Admin - Vehicle Lease` | SUMIFS |
| 287 | Real Estate Taxes | `Real Estate Taxes` | SUMIFS |
| 288 | Business Insurance | `Business Insurance` | SUMIFS |
| 289 | Business Insurance | `Insurance - Other` | SUMIFS |
| 290 | Non-Departmental | `Non-Departmental` | SUMIFS |
| 291 | Non-Departmental | `Other Taxes` | SUMIFS |
| **292** | | **Total Opex** | `=+SUM(F263:F291)` |
| 293 | | | *(blank)* |
| **294** | | **EBITDARM** | `=+F232-F292` |
| **295** | Management Fee | `Management Fee` | SUMIFS |
| **296** | | **EBITDAR** | `=+F294-F295` |
| **297** | | Margin | `=+F296/F232` |
| 298 | Ignore | `Ignore` | *(label row; no SUMIFS — marks the Ignore bucket)* |
| 299 | | | *(blank)* |
| **300** | Check | | `=F296-'<CN> Historical Financials'!E<hfNoiRow>` |

> **Tag coverage check.** Every master tag that can carry value should appear once as a C key in
> 221–295 (plus `Ignore` at 298). If `pnl-mapping` emits a tag with no home here, the dollars vanish
> from the roll-up silently — the row-234/300 Checks will go non-zero. Treat a non-zero Check as a
> missing/mis-spelled category key first.

## Denominators (rows 302–308)
| Row | B | F formula | G… |
|----|---|-----------|----|
| 302 | Units | `=+F12` | per col |
| 303 | Occupied Units | `=+F17` | |
| 304 | Residents/Occupied Units | `1.00` (number, format `0.00"x"`) | `=+F304` (carry right) |
| 305 | Resident Days | `=+F24` | |
| 306 | Total Labor | `=+F254` | |
| 307 | Total Benefits | `=+SUM(F256:F262)` | |
| 308 | Total Revenue | `=+F232` | |

## Per-line metrics (rows 310–336) — fixed analytics template
Pattern: `=+SUMIFS(F$221:F$291,$C$221:$C$291,$D{row})/<denom>`, where **D{row} = the tag key** being
measured and `<denom>` is one of `F$302` ($/unit), `F$305` ($/PRD, wrapped in IFERROR), `F308` (% rev),
`F306`/`1000`. C{row} holds the unit label shown (`$ Unit `, `$ PRD`, `$K/Property`, `% Revenue`,
`% Labor`, `% Rent+Care`). Verbatim rows:

| Row | B / D (tag) | C (unit) | Denominator / note |
|----|-------------|----------|--------------------|
| 310 | Bonuses - Marketing | `$ Unit ` | `/F$302` |
| 311 | Bonuses - Admin | `$K/Property` | `/1000` |
| 313 | Culinary | `$ Unit ` | `/F$302` |
| 314 | Raw Food | `$ PRD` | `IFERROR(.../F$305,0)` |
| 315 | Activities | `$ Unit ` | `/F$302` |
| 316 | Direct Care | `$ PRD` | `IFERROR(.../F$305,0)` |
| 317 | COGS | `% Rent+Care` | `/SUM(F221:F225,F229:F230)` |
| 318 | Maintenance | `$ Unit ` | `/F$302` |
| 319 | Utilities | `$ Unit ` | `/F$302` |
| 320 | Housekeeping | `$ Unit ` | `/F$302` |
| 321 | Admin | `$K/Property` | `/1000 + (F277+F278+F279)/1000` |
| 322 | Insurance - Workers Comp | `$ Unit ` | `/F$302` |
| 323 | Admin - Health & Dental | `% Labor` | `=+F281/F306` |
| 324 | Admin - IT | `$K/Property` | `/1000` |
| 325 | Admin - Telephone | `$K/Property` | `/1000` |
| 326 | Admin - Travel | `$K/Property` | `/1000` |
| 327 | Admin - Bank & Payroll Fees | `$K/Property` | `/1000` |
| 328 | Admin - Vehicle Lease | `$K/Property` | `/1000` |
| 329 | Real Estate Taxes | `$K/Property` | `/1000` |
| 330 | Insurance - Other | `$ Unit` | `/F$302` |
| 331 | Business Insurance | `$ Unit ` | `/F$302` |
| 332 | Non-Departmental | `$K/Property` | `/1000` |
| 333 | Management Fee | `% Revenue` | `=+F295/F308` |
| 334 | Other Taxes | `% Revenue` | `/F308` |
| 335 | Capex Reserve | `$ Unit` | `/F$302` |
| 336 | Marketing | `$ Unit` | `/F$302 + ((F275+F274))/F302` |
| 312 | Benefits & PR Taxes | `% Labor` | `=+F307/F306` |

(`$ Unit ` / `$ Unit` trailing-space inconsistencies are reproduced from the source — copy as-is.)
