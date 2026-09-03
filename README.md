# Master Property Database — Expense Comp Finder

`Master_Property_Database.xlsx` is organized for one job: find comparable expenses. Search by
**OCCC Code** (property use type) and **Building Size (SF)**, then read the matching expenses on
an **annual** or **per-square-foot** basis. All 29 subject properties are in **Adams County /
metro Denver, Colorado**.

## Workbook Tabs (3)

1. **Property Registry** — one row per property: identity (ID, name, address), **OCCC Code**,
   **Building Size (SF)** + its source, year built, tenancy, owner, latest assessor value, and a
   live **Has Expense Data** flag. Sorted by OCCC Code, then Building Size (largest first).
2. **Expenses - Annual** — every operating expense, interest, and other-expense line item, in
   annual dollars. `OCCC Code` and `Building Size (SF)` are pulled live from Property Registry
   (`INDEX`/`MATCH` on Property ID), so they always stay in sync with the registry.
3. **Expenses - Per SF** — the identical line items, normalized to `$/SF` (`Annual Amount ÷
   Building Size`). Every column here is a live formula pointing back at Expenses - Annual, so
   the two tabs never drift apart — edit one fact in Annual and both update.

All three tabs have a frozen header row and AutoFilter enabled — filter Property Registry or
either expense tab down to one OCCC Code, a Building Size range, a Category, or a specific
property in a few clicks.

## How to use it for comps

1. On **Property Registry**, filter `OCCC Code` to the use type you're comping (see key below)
   and/or filter `Building Size (SF)` to a size band. Note which of those properties have
   `Has Expense Data = Yes`.
2. Switch to **Expenses - Annual** or **Expenses - Per SF**, filter `OCCC Code` (and `Building
   Size (SF)` if you want a tighter band), then filter `Category` to `Expense` for pure operating
   costs — or include `Interest` / `Other Income/Expense` if you also want financing and
   depreciation/amortization for that property.
3. Read `Line Item / Account` for the specific cost type (Insurance, Utilities, R&M, Management,
   Real Estate Taxes, etc.) and compare `Annual Amount` or `Amount per SF` across properties.

**Only 14 of the 29 properties have captured expense detail** (the rest are lease/appraisal-only
proforma entries with no income/expense statement in the source documents) — that's what the
`Has Expense Data` column on Property Registry flags before you go looking.

## OCCC Code key

OCCC (Occupancy Classification Code) is **not an official code from any source document** — no
such field existed in the original materials. It's a short code this workbook assigns to every
property from its `Property Type / Use`, purely so properties can be filtered/grouped by use
type. If your organization or the county assessor already issues real OCCC codes, send them over
and this column gets remapped.

| OCCC Code | Description | # Properties |
|---|---|---|
| IND-WH | Industrial – Warehouse | 18 |
| IND-AMZN | Industrial – Amazon fulfillment / logistics | 3 |
| IND-DIST | Industrial – Distribution | 3 |
| IND-MFG | Industrial – Manufacturing / Food Processing | 2 |
| COM-MT | Commercial (multi-tenant industrial) | 1 |
| IND-2U | Industrial – 2-unit (multi-tenant) | 1 |
| IND-FLEX | Industrial / Flex (multi-tenant) | 1 |

## Building Size (SF)

`Building Size (SF)` on Property Registry uses **NRA/NLA (net rentable/leasable area)** where the
source documents give it — the field expense comps are conventionally normalized against — and
falls back to **GBA (gross building area)** only where NRA wasn't captured. The `Size Source`
column says which. One property, **Broadview (125 Bridge St, R0002819)**, has no building size in
any source document; its `$/SF` reads **"n/a"** everywhere rather than a fabricated number.

**2780 N. Tower Road** is captured as one combined income statement for both of its parcels
(R0212546 + R0212547); the expense rows use parcel R0212546's building size (386,000 SF) — see
the note on those rows.

## Data notes

- **Expense figures are extracted facts, not live-recalculated formulas.** The original source
  workbook's `=SUM(...)` cross-check formulas were resolved to their computed values (verified in
  Python against each source document's own printed total, e.g. Broadview's Form 8825 total of
  $316,181) before being placed here, since reorganizing the rows would have broken their
  original cell references. Only the OCCC Code, Building Size, Has Expense Data, and $/SF columns
  are live formulas.
- **A formula-recalculation check (LibreOffice) could not be run in this session** — LibreOffice
  was unable to open any file at all in this environment (confirmed on a blank test workbook),
  which is an environment issue, not a defect in this file. In its place, every formula actually
  written to the file was independently re-parsed and evaluated in Python against known-correct
  values (formula ranges, OCCC/size matches for all 159 expense rows, and a spot-check against a
  known printed total all passed with zero issues). All formulas use standard `INDEX`/`MATCH`/
  `IFERROR`/`COUNTIF`/`IF`/`OR` functions that Excel and Google Sheets compute natively on open —
  they will populate correctly.
- 584-batch expense figures: `Expense` (Category) is operating expense only — it excludes
  mortgage interest and depreciation/amortization, which are captured as separate `Interest` and
  `Other Income/Expense` rows for the same property/period — **except** 2780 N. Tower Road,
  Majestic Bldg 12/11, Broadway Logistics Center, and 22100 E 26th Ave (ASB), whose single
  `Expense` totals already include interest/depreciation (see each property's Notes).
- "In house" on the TempTee survey = owner self-performs the service, entered as `$0`.
- Broadview and 5970 Marion figures include financing/depreciation (cash-flow/tax basis, not
  pure NOI).
- Prior tabs (Tenants & Leases, Assessments & Valuation, Related Parcels, Source Documents) have
  been retired from this version to keep the workbook focused on the comp-search workflow above.
  Ask if you need that lease/valuation detail restored as additional tabs.
