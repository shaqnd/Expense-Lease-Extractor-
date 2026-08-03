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
| 9345/9335 Elm Ct (TARA LLC) | R0048725 | 0171920102027 | 9345 (9335) Elm Ct, CO | Commercial rental |
| I-25 Corporate Center | R0070622–27 | — | 460–550 E. 76th Ave, Denver, CO | Industrial – Warehouse/Mfg (multi-tenant) |
| Majestic Commercenter | R0083953 | — | 20901 E. 32nd Pkwy, Aurora, CO | Industrial – Warehouse (4 suites) |
| 3254 & 3650 Fraser St | R0084233, R0084262 | — | 3254–3650 Fraser St, Aurora, CO | Industrial – Warehouse (2 bldgs, multi-tenant) |
| 14501 E. 35th Pl | R0084237 | 01821-30-0-02-005 | 14501 E. 35th Pl, Aurora, CO | Industrial – Warehouse/Distribution |

## Workbook Tabs

1. **Source Documents** – index of the source files and compilation notes.
2. **Property Master** – one row per property: IDs, parcel, address, owner, assessee, type, tenancy, site/GBA/NRA, year built, land:bldg, occupancy, valuation summary.
3. **Tenants & Leases** – actual rent rolls (125 Bridge St; 5970 Marion; I-25 Corporate Center 14 units; Majestic Commercenter 4 suites; 3254 & 3650 Fraser St; 14501 E. 35th Pl) plus proforma lease assumptions used in the tax appeals.
4. **Assessments & Valuation** – assessed/county values, prior-year values, and income-approach proformas (PGI→NOI→value), incl. 1st Net's actual-income and stabilized/lease-up approaches for the two WPC properties.
5. **Income-Expense Summary** – total income / expense / net by property and year.
6. **Income-Expense Detail** – full line-item detail; totals use live `=SUM` formulas that match the printed totals.
7. **Related Parcels** – co-listed parcel schedules (North Side Gardens 4-parcel portfolio; WPC E. 76th Ave six schedules; St. Paul Fraser accounts).

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
- **E. 76th Ave (I-25 Corporate Center)**: the protest letter cites a county value of
  $44,788,554 and owner opinion of $32,000,000, while the income-approach sheet and 1st Net
  protest list show a 2025 county value of $47,464,870 and a 1st Net value of $21,725,000 —
  all reproduced as printed.
- **St. Paul (R0084237)** rent rolls cover the full 17-building "Denver Industrial Portfolio";
  only subject-building (14501 E. 35th Pl) leases are stored. 14501's FY2024 revenue collapse
  and net loss reflect Mountain States Logistics vacating 156,838 SF at lease expiration
  3/31/2024; Colorado Industrial Packaging took 32,556 SF from 6/1/2024 (rent abated to 9/1/24).
- **3254/3650 Fraser** income-statement "Total Expense" figures include capital expenditures
  (TI, roof, building, lease commissions) as printed, so their nets are cash-flow figures, not NOI.
