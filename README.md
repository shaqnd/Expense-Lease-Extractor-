# Master Property Database — Lease & Income/Expense Extract

`Master_Property_Database.xlsx` consolidates lease, income/expense, and valuation data
extracted and cleaned from the source documents (scanned PDFs, text PDFs, a DOCX appraisal,
and XLSX financials). All subject properties are in **Adams County, Colorado**.

A reusable Claude Code skill, **`.claude/skills/property-doc-extract`**, captures the
extraction workflow so future document batches can be parsed and merged in quickly.

## Subject Properties (11)

| Property | Account / Schedule | Parcel # | Address | Type |
|---|---|---|---|---|
| 7205 Gilpin Way | R0178308 | R0178308 | 7205 Gilpin Way, Denver, CO 80229 | Industrial – Warehouse |
| (7194) CO-Aurora | R0169133 | R0169133 | 17608 E. 24th Drive, Aurora, CO 80011 | Industrial – Distribution |
| TempTee Brand Steaks | R0103792 | 0182511406001 | 2011 E 58th Ave, Denver, CO 80216-1517 | Industrial – Mfg / Food Processing |
| 5970 Marion Drive | — | 0182511303018 | 5970 Marion Drive, Denver, CO 80216 | Industrial (2-unit multi-tenant) |
| Washington Business Park | R0103767 | R0103767 | 5650 Washington St, Denver, CO | Industrial / Flex (multi-tenant) |
| SGS – Pennsylvania Industrial | R0024442 | R0024442 | 12260 Pennsylvania St, Thornton, CO | Industrial – Warehouse |
| Broadview | R0002819 | 01569-06-3-13-002 | 125 Bridge St, Brighton, CO 80602 | Commercial (multi-tenant) |
| 6770 E. 56th Ave | R0092302 | R0092302 | 6770 E. 56th Ave, Aurora, CO 80022 | Industrial – Warehouse |
| Thornton (TH) - Dahn - 21548 | R0040829 | 0171914201013 | 10350 Washington Street, Adams County, CO | Specialty – Self Storage |
| Aurora - 18th Avenue - CS - 6730 | R0086010 | 0182132318001 | 15413 E 18th Ave, Aurora, CO 80011 | Specialty – Self Storage |
| Aurora Altura Boulevard CS 6731 | R0085571 | 0182131405022 | 1540 Altura Blvd, Aurora, CO 80011 | Specialty – Self Storage |

## Workbook Tabs

1. **Source Documents** – index of the source files and compilation notes.
2. **Property Master** – one row per property: IDs, parcel, address, owner, assessee, type, tenancy, site/GBA/NRA, year built, land:bldg, occupancy, valuation summary.
3. **Tenants & Leases** – actual rent rolls (125 Bridge St: 11 tenants w/ rent, term, CAM; 5970 Marion units) plus proforma lease assumptions used in the tax appeals.
4. **Assessments & Valuation** – assessed/county values, prior-year values, and income-approach proformas (PGI→NOI→value).
5. **Income-Expense Summary** – total income / expense / net by property and year.
6. **Income-Expense Detail** – full line-item detail; totals use live `=SUM` formulas that match the printed totals.
7. **Related Parcels** – North Side Gardens LLC 4-parcel portfolio co-listed with 7205 Gilpin Way.
8. **Self-Storage Comparison** – side-by-side income/expense/valuation comparison of the three
   Adams County self-storage tax appeals (10350 Washington St, 15413 E 18th Ave, 1540 Altura
   Blvd), pulled from the "R0040829 - 10350 Washington Street" folder under December Hearings.

> Scope: this database is for **lease and income/expense** extraction and storage. Sales
> comparables and other market/appraisal support data in the source files are intentionally
> not stored here.

## Notes

- Total / net-income / average cells are live spreadsheet formulas (`SUM`, `AVERAGE`,
  subtraction); each was verified in Python to equal the source-document printed total to the
  cent. They populate on open in Excel/Google Sheets. (LibreOffice cannot cold-start in the
  build sandbox, so the automated recalc pass is skipped in favor of Python verification.)
- **5650 Washington St** (appraisal DOCX): the rent-roll grid, income proforma, and final
  concluded value are embedded as **EMF vector images** and could not be extracted as text; the
  narrative figures (rent range, vacancy 7%, OpEx $7.17/sf ≈ 47.1% EGI, adjusted comp range
  $115.46–$134.48/sf) are captured. The concluded value is in the appraisal images.
- "in house" on the TempTee survey = owner self-performs the service, entered as `$0`.
- 5970 Marion and Broadview expense totals include financing/depreciation, so they are cash-flow /
  tax figures, not NOI. See the Assessments & Valuation tab for stabilized-NOI proformas.
- R0169133's cover value ($3,408,546) differs from its schedule total ($3,114,815); both are
  reproduced as printed.
- **Self-storage batch (10350 Washington St, 15413 E 18th Ave, 1540 Altura Blvd)**: all three are
  Ryan LLC income-capitalization appeals valued as of 6/30/2024. 10350 Washington's appeal package
  contains only the income-cap proforma — no actual on-site P&L — so its actual-statement cells in
  the Self-Storage Comparison tab are marked "—"; the two Winner Storage/CubeSmart comps (stores
  6730 and 6731) include real Dec 2024 YTD consolidated income statements, reconciled to the cent
  against the printed totals. 1540 Altura Blvd's document header prints "Assessment Year: 2026"
  but its own value table and income statement are labeled 2025 — reproduced as printed.
  `Appeal docs R0040829.pdf` and `R0040829 self storage.pdf` in the source folder are duplicates.
