# Master Property Database — Lease & Income/Expense Extract

`Master_Property_Database.xlsx` consolidates lease, income/expense, and valuation data
extracted and cleaned from the source documents (scanned PDFs, text PDFs, a DOCX appraisal,
and XLSX financials). All subject properties are in **Adams County, Colorado**.

A reusable Claude Code skill, **`.claude/skills/property-doc-extract`**, captures the
extraction workflow so future document batches can be parsed and merged in quickly.

## Subject Properties (13)

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
| 18300 E 28th Ave | R0110351 | R0110351 | 18300 E 28th Ave, Aurora, CO | Industrial – Warehouse/Distribution |
| Majestic Commercenter – Bldg 6 | R0111559 | R0111559 | 3590 N Himalaya Rd, Aurora, CO 80011 | Industrial – Warehouse |
| Majestic Commercenter – Bldg 7 | R0111560 | R0111560 | 20320 E. 36th Dr, Aurora, CO 80011 | Industrial – Warehouse |
| Friesen–Washington | R0114026 | 01825-10-1-02-023 | 6051 Washington St, Adams County, CO | Commercial / Industrial (multi-tenant) |
| Anchor Business Park | R0121833 | 25-31459-0001-CO | 5360 Washington St, Denver, CO 80216 | Industrial – Warehouse |

## Workbook Tabs

1. **Source Documents** – index of the source files and compilation notes.
2. **Property Master** – one row per property: IDs, parcel, address, owner, assessee, type, tenancy, site/GBA/NRA, year built, land:bldg, occupancy, valuation summary.
3. **Tenants & Leases** – actual rent rolls (125 Bridge St: 11 tenants w/ rent, term, CAM; 5970 Marion units) plus proforma lease assumptions used in the tax appeals.
4. **Assessments & Valuation** – assessed/county values, prior-year values, and income-approach proformas (PGI→NOI→value).
5. **Income-Expense Summary** – total income / expense / net by property and year.
6. **Income-Expense Detail** – full line-item detail; totals use live `=SUM` formulas that match the printed totals.
7. **Related Parcels** – North Side Gardens LLC 4-parcel portfolio co-listed with 7205 Gilpin Way.

> Scope: this database is for **lease and income/expense** extraction and storage. Sales
> comparables and other market/appraisal support data in the source files are intentionally
> not stored here.

## Standalone Analyses

`Lease_Comparables_5_Adams_County_Packages.xlsx` — the 20 lease comparables carried by three of the
five Adams County appeal packages (R0110351 Ryan/CoStar, R0121833 Sansone/CoStar, and the shared
Sterling Exhibit F survey used by both R0111559 and R0111560). Tabs: Lease Comparables · Set
Statistics · Notes & Data Gaps. Built by `scripts/build_5pkg_lease_comps.py` (+ `..._tabs.py`).

Notable: the Sansone comp set includes the subject property's own Unit B space at $10.75/SF asking,
while the executed lease in the same package is $9.50/SF NNN. CoStar's published "Average" is
SF-weighted; Sterling's Exhibit F average is a simple mean.

`1521-1527_Peoria_St_Lease_Comparables_Analysis.xlsx` — lease-comparable extract and
adjustment analysis for **Grease Monkey International, LLC**, 1521-1527 Peoria St, Aurora CO
80010 (Adams County, parcel 182335429010, assessment year 2025/2026). All 8 CoStar retail
lease comps from the Ryan, LLC appeal package are captured with full attributes, plus an
adjustment grid concluding a market rent on a triple-net basis. Built by
`scripts/build_peoria_lease_comps.py`.

Tabs: Subject & Conclusion · Comparables · Adjustment Grid · Market Conditions · Source Notes.

> This file is kept **separate from** `Master_Property_Database.xlsx` by design — comparables
> are market-support data, which the master database scope excludes.

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
- **R0114026** (6051 Washington St) is a **partial submission**: the 2023 P&L detail is an excerpt
  (pages 2–3 of 14), so itemised 2023 rents total $214,254.27 against a printed income total of
  $216,726.00 — $2,471.73 sits on a page that was not included. The 06/24 rent roll gives only
  tenant, rate PSF and term (no SF or monthly rent), and its third tenant row is obliterated by a
  black scan artifact.
- **Not true NOIs.** Anchor Business Park's P&L includes loan interest ($44,300.59 in 2024;
  $46,920.74 in 2023). 6051 Washington St's 2024 statement includes depreciation ($78,017.21),
  amortization ($5,811.58) and lease commissions ($34,701.00) while reporting property tax as $0.00.
  The Commercenter figures are operating income before interest and depreciation. Each is annotated
  in the Income-Expense Summary tab.
- **Majestic Commercenter Bldg 7** (R0111560) was 100% vacant at the 6/30/2024 date of value because
  the FedEx Ground lease (200,000 SF at $6.35/SF) expired 4/30/2024. Sterling's requested value
  applies a $3,662,000 excess-vacancy and lease-up discount to a $14,007,200 as-if-stabilized value.
- Sterling's letters cite a Tower Metro District of **22.5 mills** while their Exhibit H labels the
  same 0.63% effective tax rate adder as **41.259 mills**; both are reproduced as printed. Exhibit I
  in R0111560 also carries a typo'd schedule number (`R01115690`).
- Sansone's capitalized values for Anchor Business Park do not recompute exactly from their own
  printed NOI and cap rate (e.g. $347,380 ÷ 9.90% = $3,508,889 vs. the printed $3,510,421); figures
  are reproduced as printed rather than recalculated.
- Rent rolls for both Commercenter buildings are dated **06/30/22**, not the 6/30/2024 date of value.
