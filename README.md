# Master Property Database — Operating Expense Comps

`Master_Property_Database.xlsx` is an operating-expense comparables tool for Adams County,
Colorado industrial/commercial properties: a registry of subject properties classified by an
"OCCC Code" (property-type code) and building size, plus every operating expense line item
extracted from each property's tax-appeal or financial documents, in both annual-dollar and
per-square-foot form so properties of different sizes can be benchmarked directly.

A reusable Claude Code skill, **`.claude/skills/property-doc-extract`**, captures the
extraction workflow so future document batches can be parsed and merged in quickly.

## Workbook Tabs

1. **Property Registry** — one row per subject property: ID, name, address, city, county, OCCC
   Code, property type/use, building size (SF) & its source, year built, tenancy, owner/client,
   assessment year, latest assessor value, a live `Has Expense Data` flag, and source file. Sorted
   by OCCC Code, then building size (largest first) — filter/sort by either to find comparables.
2. **Expenses - Annual** — operating expense, interest & other expense line items by property, in
   annual dollars. `OCCC Code` and `Building Size (SF)` are pulled live via `INDEX/MATCH` from the
   Property Registry, so re-sorting or editing that tab keeps these columns correct automatically.
   `Category` is one of `Expense`, `Interest`, or `Other Income/Expense` (the latter used for a
   non-operating rollup — debt service, depreciation/amortization, capex, etc. — where the source
   doesn't break it out further). This tab intentionally holds **expenses only**, not rental
   income.
3. **Expenses - Per SF** — the same line items as Expenses - Annual, mirrored 1:1 by row and
   normalized to $/SF of Building Size (`"n/a"` where building size isn't available).

### OCCC Codes in use

| Code | Meaning |
|---|---|
| COM-MT | Commercial (multi-tenant) |
| IND-2U | Industrial - 2-unit |
| IND-AMZN | Industrial - Amazon fulfillment/distribution |
| IND-DIST | Industrial - Distribution |
| IND-FLEX | Industrial / Flex (multi-tenant) |
| IND-MFG | Industrial - Manufacturing |
| IND-SS | Self-Storage |
| IND-SY | Industrial - Multi-Tenant Storage Yard/Warehouse Bays (rented bays, not individual self-storage units) |
| IND-WH | Industrial - Warehouse |

## Notes

- `Has Expense Data` (Property Registry) and the `OCCC Code`/`Building Size (SF)` lookups
  (Expenses tabs) are live formulas referencing fixed ranges (`Property Registry!$A$5:$A$40` etc.,
  sized to the current row count). **Adding or removing property rows requires updating these
  range endpoints** in all three tabs, or the formulas for rows beyond the old range will break.
- Dollar totals were verified in Python against each source document's printed total before being
  entered; `TOTAL EXPENSE`/`TOTAL EXPENSES` rows are the printed figures, not live `SUM` formulas,
  since this schema stores each line item as a plain value (no in-sheet subtotal formulas).
- **R0099616** (All American Mini Storage): the taxpayer's appeal treats 6 Adams County schedules
  (R0099616, R0099615, R0099617, R0180917, R0099618, R0180918) as one economic unit; only the
  primary schedule (R0099616) is registered, representing the combined unit. 2024 expense detail is
  fully itemized (21 categories, verified to $212,162 excl. RE taxes); 2021-2023 are subtotal-only
  (excl./incl. RE-tax figures), as the source Income Analysis exhibit didn't itemize those years.
- **R0169115** (Stor-n-Lock #20): Building Size is calculated (Annual Rent ÷ $16.00/SF NNN asking
  rate), not a stated NRA figure. City was not given in the source (Adams County, unincorporated).
- **R0085523** (IN Self Storage-Fitzsimons): `TOTAL EXPENSE` for 2023/2024 is the printed monthly
  P&L total; the full line-item budget-vs-actual detail (Exhibit B in the source) was not
  individually itemized here.
- **R0071007** (Warehouse and Storage Building(s), N. Pecos Ave — T & G Pecos LLC): a multi-tenant
  storage yard of ~90 individual bays/units (Buildings A-G), classified `IND-SY` rather than
  `IND-SS` since tenants rent bays (mixed individuals & small trade businesses), not classic
  self-storage units. No assessor value or building size is available in its source file (an SB
  11-119 income/expense disclosure, not an appeal package).
- **R0055143** (8700 Devonshire Storage, CubeSmart-managed): its 26-month rolling income statement
  is a dense, rotated scanned table; only clearly legible subtotals were transcribed (`TOTAL
  EXPENSE` and a derived non-operating `Other Income/Expense` rollup), not full line-item
  categories. No assessor value or building size is available in the source.
- **R0043569** (Public Storage #08214, Thornton): only an internal FY2024 monthly Actuals P&L was
  provided — no cover letter, appeal package, assessor/taxpayer value, or building size.
