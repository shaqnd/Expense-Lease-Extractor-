# Master Property Database — Lease & Income/Expense Extract

`Master_Property_Database.xlsx` consolidates lease, income/expense, and valuation data
extracted and cleaned from the source documents (scanned PDFs, text PDFs, a DOCX appraisal,
and XLSX financials). All subject properties are in **Adams County, Colorado**.

A reusable Claude Code skill, **`.claude/skills/property-doc-extract`**, captures the
extraction workflow so future document batches can be parsed and merged in quickly.

## Subject Properties (29: 8 original + 21 in the "584" batch)

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

### "584" batch (21 properties — flagged `Batch = 584` in every tab)

| Property | Account / Schedule | Address | Type |
|---|---|---|---|
| Majestic Commercenter Bldg 12 | R0218284 | 20600 E. 35th Drive, Denver | Industrial – Warehouse |
| Lovett 76 Logistics Center | R0213058 | 6196 E Bridge St, Brighton | Industrial – Warehouse |
| Majestic Commercenter Bldg 11 | R0218283 | 20310 E. 35th Drive, Denver | Industrial – Warehouse |
| Denali Building 1 | R0218172 | Denali Buildings 1–3 (Bldg 1), Adams County | Industrial – Warehouse |
| Denali Building 2 | R0218173 | Denali Buildings 1–3 (Bldg 2), Adams County | Industrial – Warehouse |
| Denali Building 3 | R0218174 | Denali Buildings 1–3 (Bldg 3), Adams County | Industrial – Warehouse |
| 2780 N. Tower Road (parcel 1 of 2) | R0212546 | 2780 N. Tower Road, Aurora | Industrial – Distribution |
| 2780 N. Tower Road (parcel 2 of 2) | R0212547 | 2780 N. Tower Road, Aurora | Industrial – Distribution |
| TrusTile Doors | R0198564 | 1111 E 71st Ave, Denver | Industrial – Manufacturing |
| Broadway Logistics Center | R0217136 | 6795 Broadway St, Denver | Industrial – Warehouse |
| DEN3 – Bull Crossing (Amazon) | R0198789-Lease | 14601 Grant St, Thornton | Industrial – AR Sortable |
| DEN7 (Amazon) | R0198530 | 22300 E 26th Ave, Aurora | Warehouse |
| Majestic Commercenter Bldg 15 | R0193588 | 2889 Himalaya Road, Adams County | Industrial – Warehouse |
| DEN2 – PR Park 70 (Amazon) | R0191292 | 22205 E 19th Ave, Aurora | Traditional Non-Sortable |
| Majestic Commercenter Bldg 29 | R0187866 | 19799 E. 36th Drive, Adams County | Industrial – Warehouse |
| Majestic Commercenter Bldg 26 | R0172868 | 19755 E. 35th Drive, Adams County | Industrial – Warehouse |
| Home Depot Denver | R0180551 | 1953 Gun Club Rd, Aurora | Industrial – Warehouse |
| 22100 E 26th Ave (ASB) | R0172848 | 22100 E 26th Ave, Aurora | Industrial |
| Majestic Commercenter Bldg 24 | R0172865 | 3500 N. Windsor Drive, Adams County | Industrial – Warehouse |
| Majestic Commercenter Bldg 22 | R0132030 | 3543 N. Windsor Drive, Adams County | Industrial – Warehouse |
| Shamrock Foods | R0092655 | 5199 Ivy St, Commerce City | Warehouse |

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

## "584" batch (2026-08-25)

Source: Google Drive **Expenses & Leases / 584** folder (18 PDFs, all Adams County, CO). Every
row added by this batch is flagged **`Batch = 584`** in a dedicated column on each tab, so it can
be filtered/identified separately from the original 8 properties.

- **Amazon leasehold appeals** (DEN2/DEN3/DEN7 — R0191292, R0198789, R0198530): tenant-filed
  possessory-interest appeals against the fee-simple assessed value; full lease abstracts
  (term, renewal options, base-rent escalation schedules) captured in Tenants & Leases.
- **Sterling Property Tax Specialists "Majestic Commercenter" filings** (Bldgs 11, 12, 15, 22,
  24, 26, 29 — R0218283/84, R0193588, R0132030, R0172865, R0172868, R0187866): each includes a
  Ryan/Sterling-style income proforma plus an actual 12-month detail income statement (FY2024,
  several also FY2023) pulled from the buildings' own accrual GL. Two buildings (Bldg 11 & 12)
  were 100% vacant at the 6/30/2024 date of value; Bldg 22 was 75% occupied.
- **Denali Buildings 1–3** (R0218172/73/74): one appeal letter covering three schedule numbers;
  Property Master carries 3 rows, Assessments & Valuation splits assessor/taxpayer values per
  schedule, and the income proforma (combined 760,400 SF) is carried as a single row.
- **2780 N. Tower Road** (R0212546/R0212547 — two schedule numbers, one tax-district-split
  parcel): full actual accrual P&L for FY2024 and FY2023, single tenant.
- **TrusTile Doors** (R0198564) and **Shamrock Foods** (R0092655): both 100% owner-occupied,
  no actual tenant; TrusTile is a county income/expense survey (expenses only), Shamrock is a
  DMA appeal using a hypothetical fee-simple market-lease income approach.
- **22100 E 26th Ave / ASB Real Estate** (R0172848): building only 28% occupied at the DOV; has
  an actual rent roll (Pelsue Company lease with 10-year escalation schedule, in Tenants &
  Leases) plus a full accrual-with-CapEx income statement for FY2024 and FY2023 showing large
  net losses driven by leasing capex ahead of a new 313,730 SF lease (Vederra, starts 5/1/2025).
- **Home Depot Denver** (R0180551): related-parcel schedule for the Cherry/Thornton/Huron/Park 76/
  Quintero Owner LLC portfolio (KKR-affiliated) added to Related Parcels; only Adams County
  parcels included, other-county parcels (Douglas, Arapahoe) excluded as out of scope.
- **Data-quality gaps flagged in Income-Expense Detail**: several buildings' actual income
  statements had legible printed **totals** but not full GL-account-level detail (heavy OCR
  noise on monthly columns, or scope/time limits) — those are entered as summary TOTAL rows
  (literal printed figures, not `SUM` formulas) with a note explaining the gap, per property.
  Where full line-item detail was captured (Bldg 12, Bldg 11 income, Broadway, TrusTile, and
  both years of 2780 N. Tower Road), the `TOTAL` rows use live `=SUM` formulas verified in
  Python to match the printed total to the cent.
- The Denali doc's "Property Operating Report" page (Dec 2024 activity) was encoded in a
  symbol/dingbat font and not decodable — flagged as a gap, not entered.
- Lease/rent comps (CoStar reports, Lowery/RealtyRates surveys, market lease-comp tables) in
  every 584-batch document were skipped per the established scope (subject-property data only).
