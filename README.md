# Master Property Database — Lease & Income/Expense Extract

`Master_Property_Database.xlsx` consolidates lease, income/expense, and valuation
data extracted and cleaned from five source documents (three scanned PDFs and two
management reports). All subject properties are in **Adams County, Colorado**.

## Subject Properties

| Property | Account # | Parcel # | Address | Type |
|---|---|---|---|---|
| 7205 Gilpin Way | R0178308 | R0178308 | 7205 Gilpin Way, Denver, CO 80229 | Industrial – Warehouse |
| (7194) CO-Aurora | R0169133 | R0169133 | 17608 E. 24th Drive, Aurora, CO 80011 | Industrial – Distribution |
| 5970 Marion Drive | — | 0182511303018 | 5970 Marion Drive, Denver, CO 80216 | Industrial (2-unit multi-tenant) |
| TempTee Brand Steaks | R0103792 | 0182511406001 | 2011 E 58th Ave, Denver, CO 80216-1517 | Industrial – Mfg / Food Processing |

## Workbook Tabs

1. **Source Documents** – index of the five source files and compilation notes.
2. **Property Master** – one column per property: ID, parcel, address, owner, type, size, year built, valuation summary.
3. **Tenants & Leases** – actual tenants (5970 Marion) plus proforma lease assumptions used in the two tax appeals.
4. **Assessments & Valuation** – historical assessed values (2024–2025) and the Ryan LLC income-approach proforma.
5. **Income-Expense Summary** – total income / expense / net income by property and year.
6. **Income-Expense Detail** – full line-item detail; totals use live `=SUM` formulas that match the printed totals.
7. **Sales Comparables** – warehouse sale comps supporting the R0169133 appraisal (CoStar).
8. **Related Parcels** – North Side Gardens LLC 4-parcel portfolio co-listed with 7205 Gilpin Way.

## Notes

- Figures are transcribed directly from the source PDFs. Total / net-income / average
  cells are live spreadsheet formulas (`SUM`, `AVERAGE`, subtraction) that recompute
  when inputs change; each was verified to equal the source-document printed total.
- "in house" on the TempTee income & expense survey means the owner performs the
  service internally with no dollar figure reported (entered as `$0`).
- 5970 Marion Drive "expense" totals are cash-basis and include mortgage principal,
  interest, and property taxes — they are cash-flow figures, not NOI.
- R0169133's cover page lists a 2025 Assessor's Actual Value of $3,408,546 while its
  historical-assessment schedule shows a 2025 total of $3,114,815; both are reproduced
  as printed.

## Source Documents

| File | Type | Property | Period |
|---|---|---|---|
| R0169133.pdf | Property tax appeal (income approach + sale comps) | 17608 E. 24th Drive | Assessment Yr 2025 |
| R0178308.pdf | Property tax appeal (income approach + auth letter) | 7205 Gilpin Way | Assessment Yr 2025 |
| R0103792.pdf | Adams County Commercial-Industrial Income & Expense Survey | 2011 E 58th Ave (TempTee) | 2023–2024 |
| 5970_Marion_Drive_2024_PL_Detail.pdf | Income statement (detailed, cash basis) | 5970 Marion Drive | FY2024 |
| 5970_Marion_Drive_2023_Cash_Flow_Statement.pdf | Cash flow statement (cash basis) | 5970 Marion Drive | FY2023 |
