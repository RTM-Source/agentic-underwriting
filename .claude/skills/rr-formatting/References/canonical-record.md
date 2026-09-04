# Canonical Unit Record (the reader → emitter contract)

Every reader (OneSite, Yardi, PCC, …) parses its messy source into a list of these records, one per
**unit (bed/contract)** per `locked-rules.md` rule 1. The emitter consumes only this — it never re-reads
the source's quirks. Keep this schema stable; it's what lets one emitter serve all readers.

| Field | Type | Meaning / how the reader fills it |
|-------|------|-----------------------------------|
| `order` | int | 1..N in source order. |
| `unit` | id | The operator's **raw unit code**, verbatim (`103P`, `139-139A`) → col AX = `=A{a}` (rule 8). |
| `unitCode` | string | Composite billing key `Care \| Type \| SqFt` (`AL \| AL-1D \| 545`) → col BK. |
| `unitType` | string | Floor-plan code parsed from col E (`AL-1D`, `MC-0P`). |
| `pSp` | `P`/`S` | Primary bed vs shared 2nd occupant (from the code suffix) → col AZ. |
| `sqft` | number | The unit's **actual** SqFt (hardcoded input — raw floor-plan SqFt is a placeholder). **0 marks a shared "S" line**. |
| `careType` | `AL`/`MC`/`IL` | From unit-type prefix. |
| `tier` | `A`/`B`/`C` \| "" | AL/IL SqFt tier by natural breaks (Planning-Mode call); "" for MC (grouped by room-type). |
| `basicUnitCode` | string | Human label `AL - 1BR (545-646 SF)` / `MC - Private` → col BL (hardcoded). |
| `groupUnitCode` | string | `Care \| Type \| Tier` (`AL \| 1BR \| A`) → col BM. **The analysis join key** (rule 8). |
| `capacity` | 0/1 | Bed count: 1 if `sqft>0`, else 0 (rule 2). |
| `occupancy` | 0/1 | Period-end: occupied=1; vacant/inactive/down=0; **shared 2nd-occupant "S" line = 0** (rule 3, revised). |
| `moveInRow` | date \| null | The current occupant's cleaned move-in (hardcoded input, col BF); null if unoccupied / "S" line. |
| `leafRows` | source rows[] | The unit's own detail rows (anchor + continuation lines), **excluding** any `NNN TOTAL`. Drives the finance link formulas (rule 7). |
| `anchorRow` | source row | The unit's primary row (where the right-side cells get written, row-aligned). |
| `secondResidentFee` | number | The 2nd-occupant fee (col BJ = `IFERROR(VALUE(AG),0)`); **`>0` = a paying 2nd resident** (rule 5, revised — a fee, not a `Y` flag). On a P/S pair it lands on the primary; the "S" row is zeroed. |
| `flags` | string[] | Per-unit notes for the underwriter: `inactive-down`, `2nd-occupant`, `companion`, `turnover`, `de-prorated`, `vacant`, `sign-anomaly`, `unmapped-status`, etc. A flag is success, not failure. |

Notes:
- The emitter writes the right-side block **row-aligned** to `anchorRow` (so formulas read the same row's
  left cells); continuation/total rows on the left get blank right-side cells.
- The reader carries **rows, not values**, for anything that stays a live link (market, rent, care →
  `VALUE()` formulas). It carries computed values only for what can't be a formula (order, SqFt, cleaned
  move-in, occupancy, capacity, tier, the group/basic codes) — see `output-columns.md` for which is which.
- `leafRows` is the crux of correctness: get the block boundaries right and the reconciliation gate passes
  automatically, because every leaf belongs to exactly one unit.
