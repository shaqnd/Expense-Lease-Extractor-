# Master Property Database — Lease & Income/Expense Extract

`Master_Property_Database.xlsx` consolidates lease, income/expense, and valuation data
extracted and cleaned from the source documents (scanned PDFs, text PDFs, a DOCX appraisal,
and XLSX financials). All subject properties are in **Adams County, Colorado**.

A reusable Claude Code skill, **`.claude/skills/property-doc-extract`**, captures the
extraction workflow so future document batches can be parsed and merged in quickly.

## Subject Properties (30)

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
| Denver Distribution Center | R0180834 | — | not stated in source | Industrial – Distribution (553,757 SF) |
| 2780 North Tower Road | R0212546 | — | 2780 N. Tower Road, Aurora, CO | Industrial – single tenant (377,729 SF, 1983) |
| Commercenter #22 LLC | R0132030 | — | Aurora, CO 80011 | Industrial – multi-tenant (200,090 SF, 24.99% vacant) |
| 650 W. 104th Ave | R0037193 | 0171910307020 | 650 W. 104th Ave, Northglenn, CO 80234 | Commercial – collision repair shop |
| RJR Investments | R0075352 | 0172116004003 | not stated in source | Industrial – office/shop/storage (~136,776 SF) |
| 7908 Highway 85 | R0077789 | 0172131110004 | 7908 Highway 85, Commerce City, CO 80022 | Industrial – asphalt/concrete recycling |
| 4850 E. 74th Ave | R0077963 | 0172131402020 | 4850 E. 74th Ave, Commerce City, CO 80022 | Industrial – tank trailer/truck (18,720 SF) |
| 14891 E. Colfax Ave | R0085560 | R0085560 | 14891 E. Colfax Ave, Aurora, CO | Commercial – Meineke Car Care (4,264 SF) |
| 11400 Huron | R0030085 | 0171903005007 | 11400 Huron, CO | Commercial – tire store (~3,840 SF) |
| 11450 N. Huron St | R0030089 | 1719-03-0-05-016.019 | 11450 N. Huron St, CO | Commercial – multi-tenant (30,782 SF) |

*(Plus 11 further properties — R0048725, I-25 Corporate Center, the Fraser St / E. 33rd Pl / Moncrieff / Uravan
group — carried over from the consolidated `Master_Property_Database_2` workbook.)*

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
- **R0180834 / Denver Distribution Center** — United Natural Foods, 553,757 SF, 100% occupied, $0.54→$0.56/SF/mo
  ($6.52→$6.69/SF/yr). FY2024: revenue $5,292,302, OpEx $2,376,573, **NOI $2,915,729**.
- **R0212546 / 2780 N. Tower Road** — absolute net: taxes, R&M and (FY2024) utilities all **$0** to the owner.
  Rent **$8.07/SF** (FY2024) and $7.89/SF (FY2023) on 377,729 SF. Its "Total Expense" includes $456,096
  depreciation; cash operating expense is only $46,980 ($0.12/SF).
- **R0132030 / Commercenter #22** — Expeditors International (64,835 SF) and Steelcase (85,253 SF) leased,
  **50,002 SF (24.99%) vacant** at all three rent-roll dates. FY2024 operating income $775,968; net loss
  $(424,461). Proforma value $13,402,900.
  - ⚠️ *Source defect, reproduced not corrected:* its "Total Other (Income)/Expense" subtotal ($743,031.13) does
    not equal its own line items ($25,001.47). The gap is exactly twice the $(359,014.83) depreciation step-up
    credit, which the report adds rather than subtracts. The printed subtotal is the one consistent with the
    printed net loss, so it is stored as a hard value, not a `SUM`.
- **R0085560 / 14891 E. Colfax** — Meineke Car Care, 4,264 SF, 100% occupied, lease 12/1/2021–7/31/2026 with
  printed steps $17.32 → $18.93/SF. Four years of cash-basis P&L (2021-2024). Assessor $1,172,600 vs petitioner
  **$870,000**; agent proforma ($17.85 rent, 7% vacancy, 8% expenses, 7.5% cap) = $868,290.
- **R0030089 / 11450 N. Huron** — 30,782 SF, FY2021 and FY2022 income & expense. County $5,987,885 vs requested
  **$3,500,000**. Expenses $7.13/SF (2021) and $8.04/SF (2022).
  - The schedule prints depreciation, amortization, the owner management fee, quarterly estimated tax and
    general-building capital in a **side column outside the totalled year columns** — they are excluded from
    Total Expenses and from the capitalised net income. Those items are stored separately, tagged
    *Excluded from total*. In-column line items foot to the printed totals within $1–$3 (the schedule rounds
    every line to whole dollars).
- **County income & expense surveys** (R0037193, R0075352, R0077789, R0077963, R0030085) are single-page owner
  responses with no rent roll or full P&L. Specific limits:
  - **R0075352** answers the operating-expense grid with A/B letters showing *who pays* rather than dollar
    amounts, so no expense figures exist for it. It also reports 100% owner-occupancy alongside a $27,825/mo
    NNN rent — a related-party lease.
  - **R0030085** reports building insurance as "included in total CO" (corporate level), so its $64,534 expense
    total is **partial**.
  - **R0077789** and **R0030085** are 100% owner-occupied and report no rent at all.
- The **Dollar General / HighPoint Elevated** article is indexed in Source Documents as market context only and
  has no Property Master row; **R0030089 pages 8-9** are CoStar sale comparables and were skipped — the database
  excludes market/comparable support material.
