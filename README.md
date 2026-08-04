# Master Property Database — Lease & Income/Expense Extract

`Master_Property_Database.xlsx` consolidates lease, income/expense, and valuation data
extracted and cleaned from the source documents (scanned PDFs, text PDFs, a DOCX appraisal,
and XLSX financials). All subject properties are in **Adams County, Colorado**.

A reusable Claude Code skill, **`.claude/skills/property-doc-extract`**, captures the
extraction workflow so future document batches can be parsed and merged in quickly.

## Subject Properties (20)

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
| Aurora Altura Boulevard CS 6731 | R0085571 | 0182131405022 | 1540 Altura Blvd, Aurora, CO 80011 | Specialty – Self-Storage (CubeSmart) |
| Aurora – 18th Avenue CS 6730 | R0086010 | 0182132318001 | 15413 E 18th Ave, Aurora, CO 80011 | Self Storage (CubeSmart) |
| Thornton (TH) – Dahn 21548 | R0040829 | 0171914201013 | 10350 Washington St, Thornton, CO 80229 | Self Storage (Mini U Storage) |
| 5720 Holly St | R0091907 | 0182308302003 | 5720 Holly St, Commerce City, CO 80022 | Industrial – Warehouse (multi-tenant) |
| Park Industrial 6125 | R0091909 | R0091909 | 6125 E. 56th Ave, Commerce City, CO 80022 | Industrial / Flex (multi-tenant) |
| Metal Mart (Hillenmark) | R0091912 | — | 6475 E. 56th Ave (56th @ Monaco), Commerce City, CO 80022 | Industrial – single-tenant distribution |
| Park Industrial 6300/6340/6360 | R0091931 | R0091931 | 6300 E. 58th Ave, Commerce City, CO 80022 | Industrial / Flex (multi-tenant) |
| Park Industrial 5800 | R0091933 | R0091933 | 5800 E. 58th Ave, Commerce City, CO 80022 | Industrial / Flex (multi-tenant) |
| Park Industrial 5750 | R0091934 | R0091934 | 5750 E. 58th Ave, Commerce City, CO 80022 | Industrial / Flex (multi-tenant) |
| WPC 50th LLC | R0092711 (+ R0181935) | R0092711 | 6701 E. 50th Ave, Commerce City, CO 80022 | Industrial – Warehouse (multi-tenant) |
| 4950 Olive Street | R0092720 | R0092720 | 4950 Olive St, Commerce City, CO 80022 | Commercial |
| 6475 Franklin Street | R0098149 | 0182502308019 | 6475 Franklin St, Adams County, CO | Commercial office & warehouse (owner-occupied) |

## Workbook Tabs

1. **Source Documents** – index of the source files and compilation notes.
2. **Property Master** – one row per property: IDs, parcel, address, owner, assessee, type, tenancy, site/GBA/NRA, year built, land:bldg, occupancy, valuation summary.
3. **Tenants & Leases** – actual rent rolls (125 Bridge St: 11 tenants w/ rent, term, CAM; 5970 Marion units), proforma lease assumptions used in the tax appeals, and self-storage occupancy / unit-mix snapshots (storEDGE & SiteLink reports as of 12/31/2024).
4. **Assessments & Valuation** – assessed/county values, prior-year values, and income-approach proformas (PGI→NOI→value).
5. **Income-Expense Summary** – total income / expense / net by property and year.
6. **Income-Expense Detail** – full line-item detail; totals use live `=SUM` formulas that match the printed totals.
7. **Related Parcels** – portfolio parcel schedules attached to the source packages: North Side Gardens LLC (4 parcels), Winner Storage / CubeSmart PTA (29 parcels), Dahn Corporation / Mini U Storage (10 parcels), Watumull Properties / WPC (12 Adams County parcels of ~40 listed), and Peoria Way Associates LLC (4 parcels).

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
- Self-storage income statements (CubeSmart stores #6730/#6731, Mini U Storage Thornton) are
  store-level accrual **NOI-basis** statements: property tax and third-party management fees are
  included; interest, depreciation, and extraordinary items are excluded (captured as memo lines
  for 10350 Washington St). The Cushman & Wakefield H2-2024 self-storage investor survey bundled
  in R0040829 is market-support material and was intentionally not stored.
- R0040829's Ryan package prints its value year with a dropped digit ("202"); the LOA covers
  2025/2026 with a 6/30/2024 level of value, so it is recorded as "2025/2026". R0086010 is a
  Value-Year-2026 package whose abatement petition targets Tax Year 2025; both years are noted.
- **Commerce City batch (R0091907–R0098149).** The four Sterling Property Tax appeals (6125 E. 56th
  and 6300 / 5800 / 5750 E. 58th) share one owner and one Yardi chart of accounts. Their Exhibit A
  statements are accrual and property-level: real-estate tax is **inside** operating expense, while
  mortgage interest and depreciation sit below NOI and are carried as memo lines. Sterling's letter
  recasts expense as RE Taxes + CAM Reimbursable + CAM Non-Reimbursable; CAM Reimbursable equals
  Total Operating Expenses less Tax Real Estate to the dollar, and CAM Non-Reimbursable is a
  Sterling-selected subset of the non-CAM block, so the letter's "Total Expenses" runs above the
  statement's Total Operating Expenses. Sterling capitalizes at 7.25% plus a 2.45% tax load (9.70%)
  and therefore excludes RE tax from the capitalized expense; the two 1st Net appeals (5720 Holly,
  6701 E. 50th) capitalize actual rent-roll income at an unloaded rate with **formulaic** expenses
  (management 3% + reserve/owner 5% of EGI), not actual operating statements.
- The file uploaded as **R0091936 is a byte-for-byte duplicate** of the R0091934 package (all 27
  pages render identically); no data for schedule R0091936 was received.
- 5720 Holly St: the owner summary prints annual rent of $484,407 alongside an average rent of
  $8.72/SF, which are mutually inconsistent ($484,407 ÷ 54,300 SF = $8.92/SF); the appeal's income
  analysis uses the $8.72 figure ($473,496). The Ginzel lease end (6/30/2026 on the Yardi roll vs
  11/30/2026 on the owner summary) also differs; both are reproduced as printed.
- 6701 E. 50th Ave: the Colliers rent roll shows 0 vacant SF while the appeal's sales grid states
  14% vacancy and its income analysis applies a 5% vacancy deduction. True World Foods' lease
  expired 4/30/2024 (holdover), and its prorated annual rent reflects a 5/1/2023 rent step.
- 4950 Olive St carries no income/expense or rent-roll data — the package is the protest letter plus
  the prior-year BAA stipulation and order. The BAA docket names the petitioner "First California
  Investments" while the letter of authority names the owner "Peoria Way Associates LLC".
- 6475 Franklin St is an assessor income/expense survey only: 100% owner-occupied, so no rental
  income is reported, and the form prints no expense total (the `=SUM` is computed). Property taxes
  ($90,477) and bank fees ($151) were handwritten on the Comments line rather than in the grid.
