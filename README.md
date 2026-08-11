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
| Denver Distribution Center | R0180834 | — | not stated in source | Industrial – Warehouse / Distribution (553,757 SF) |
| 2780 North Tower Road | R0212546 | — | 2780 N. Tower Road, Aurora, CO | Industrial – Warehouse, single tenant (377,729 SF, 1983) |
| Majestic Commercenter (WPC ABC LLC) | R0083953 | — | 20901 E. 32nd Pkwy, Aurora, CO | Industrial – 4 × 50,000 SF units (200,000 SF, 1985) |
| Commercenter #22 LLC | R0132030 | — | Aurora, CO 80011 | Industrial – multi-tenant (200,090 SF, 24.99% vacant) |

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
- **R0180834 / Denver Distribution Center** — single tenant United Natural Foods Inc, 553,757 SF,
  100% occupied, lease 6/15/2012–10/31/2028 at $0.54/SF/mo ($6.52/SF/yr), stepping to $0.56
  ($6.69/SF/yr) on 11/1/2025. FY2024 accrual: revenue $5,292,302, OpEx $2,376,573, **NOI
  $2,915,729**; financing cost $902,885 is carried separately. Address and year built are not
  stated anywhere in the source document.
- **R0212546 / 2780 N. Tower Road** — absolute net single-tenant lease: property taxes, R&M and
  (in FY2024) utilities are all **$0** to the owner. Rent $3,047,376 = **$8.07/SF** on 377,729 SF
  in FY2024 and $2,979,661 = $7.89/SF in FY2023. Its "Total Expense" is a QuickBooks figure that
  *includes* $456,096 depreciation and $12,124 amortization; cash operating expense is only
  $46,980 ($0.12/SF). The two uploaded PDFs for this account contain the identical P&L, so the
  data is entered once.
- **R0083953 / Majestic Commercenter** — 4 units of 50,000 SF (Zayo Group, Tritz Pallet ×2, Peco
  Pallet), 100% occupied at all four rent-roll dates 2021–2024. Base rent grew $52,083 → $58,458
  per month ($3.12 → $3.51/SF). Three full years of income statements (FY2022–FY2024) are loaded.
  Its 6/30/2024 recap page shows Peco Pallet expiring 3/31/2030 while the Colliers rent roll for
  the same date shows 3/31/2025; both are reproduced as printed.
- None of the five documents in that batch contain an assessor value or a petitioner opinion of
  value, so no rows were added to **Assessments & Valuation** for those three properties.
- **R0132030 / Commercenter #22 LLC** — 200,090 SF in Aurora, only two suites leased (Expeditors
  International 64,835 SF to 7/31/2026; Steelcase 85,253 SF to 12/31/2024) with **50,002 SF
  (24.99%) vacant at all three rent-roll dates** (6/30/23, 12/31/23, 6/1/24). FY2024 accrual:
  revenue $1,648,443, operating expense $872,475, **operating income $775,968**; after $457,398
  mortgage interest and $743,031 of other (income)/expense the printed bottom line is a **net loss
  of $(424,461)**. The petitioner's Exhibit H proforma indicates $14,013,500 as-if-stabilized
  ($70.04/SF) and **$13,402,900** after a $(610,600) excess-vacancy adjustment.
  - ⚠️ *Source-document defect, reproduced not corrected:* the statement's "Total Other
    (Income)/Expense" subtotal of $743,031.13 does not equal its own printed line items, which add
    to $25,001.47. The $718,029.66 gap is exactly twice the $(359,014.83) Depr-Buildings Step Up
    credit — the report adds that credit rather than subtracting it. The printed subtotal is the
    one consistent with the printed net loss, so it is entered as a hard value (not a `SUM`) and
    the line items are entered as printed, with the discrepancy annotated in the cell.
  - The proforma's "Adjustment for 67% Vacancy" label is as printed and does not match the 24.99%
    vacancy on the rent roll.
- The **Dollar General / HighPoint Elevated** article (Mile High CRE, 8/9/2022) covers a 919,000 SF
  build-to-suit at a third-party Aurora development. It carries no rent roll, lease, or
  income/expense data for any subject property, so it is indexed in **Source Documents** for
  context only and has no Property Master row — consistent with the database scope, which excludes
  market/comparable support material.
