# Master Property Database — Lease & Income/Expense Extract

`Master_Property_Database.xlsx` consolidates lease, income/expense, and valuation data
extracted and cleaned from the source documents (scanned PDFs, text PDFs, a DOCX appraisal,
and XLSX financials). All subject properties are in **Adams County, Colorado**.

A reusable Claude Code skill, **`.claude/skills/property-doc-extract`**, captures the
extraction workflow so future document batches can be parsed and merged in quickly.

## Subject Properties (12)

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
| Simply Storage | R0102989 | R0102989 | 4911 West 58th Ave., Arvada, CO 80002 | Self-Storage |
| Stor-n-Lock Self Storage #20 | R0169115 | 0172114202001 | 11210 E 104th Avenue, Adams County, CO | Self-Storage (NNN) |
| All American Mini Storage | R0099616 (+5 related schedules) | 0182504223005 (+related) | 1777 W. 68th Avenue, Denver, CO 80221 | Self-Storage + Office/Residential |
| IN Self Storage – Fitzsimons | R0085523 | R0085523 | 1520 N Fitzsimons Parkway, Aurora, CO 80011 | Self-Storage + Manager Residence |

## Workbook Tabs

1. **Source Documents** – index of the source files and compilation notes.
2. **Property Master** – one row per property: IDs, parcel, address, owner, assessee, type, tenancy, site/GBA/NRA, year built, land:bldg, occupancy, valuation summary.
3. **Tenants & Leases** – actual rent rolls (125 Bridge St: 11 tenants w/ rent, term, CAM; 5970 Marion units) plus proforma lease assumptions used in the tax appeals.
4. **Assessments & Valuation** – assessed/county values, prior-year values, and income-approach proformas (PGI→NOI→value).
5. **Income-Expense Summary** – total income / expense / net by property and year.
6. **Income-Expense Detail** – full line-item detail; totals use live `=SUM` formulas that match the printed totals.
7. **Related Parcels** – North Side Gardens LLC 4-parcel portfolio co-listed with 7205 Gilpin Way; the 6-schedule Kekake Hale LLC / All American Mini Storage economic unit; and the ~29-parcel National Storage Affiliates / Securcare Colorado portfolio attached to the R0169115 authorization letter (reference only, not an economic unit).

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
- **R0102989** (Simply Storage): pages 6-22 of the source PDF are a self-storage cap-rate/market
  survey (out of scope, skipped). The 11-page unit-level rent roll (524 units) is summarized as a
  single aggregate row on the Tenants & Leases tab rather than transcribed unit-by-unit.
- **R0169115** (Stor-n-Lock #20): city was not stated in the source document (Adams County,
  11210 E 104th Ave). Its Letter of Authorization (Exhibit C) lists ~29 other Colorado
  self-storage parcels under the National Storage Affiliates/Securcare family — reproduced on the
  Related Parcels tab for reference; they are separate ownership entities, not part of this
  appeal's economic unit.
- **R0099616** (All American Mini Storage): the taxpayer's appeal treats 6 Adams County schedules
  (R0099616, R0099615, R0099617, R0180917, R0099618, R0180918) as one economic unit; combined
  current value $7,659,751, requested $6,250,000. Two of the six (R0099618, R0180918) are
  detention-pond "drainage" parcels the taxpayer contends the County over-values as buildable land.
- **R0085523** (IN Self Storage – Fitzsimons): 2023 value was reduced from $6,998,300 to
  $5,200,000 by a Board of Assessment Appeals order (Docket 2023BAA2761), with $200,000 allocated
  to the on-site manager's residential apartment; the 2026 appeal in this file requests
  $3,395,400. Pages 27-28 (CoStar apartment sale comps supporting the residential allocation) are
  out of scope and were skipped.
