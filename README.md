# Master Property Database — Lease & Income/Expense Extract

`Master_Property_Database.xlsx` consolidates lease, income/expense, and valuation data
extracted and cleaned from the source documents (scanned PDFs, text PDFs, a DOCX appraisal,
and XLSX financials). All subject properties are in **Adams County, Colorado**.

A reusable Claude Code skill, **`.claude/skills/property-doc-extract`**, captures the
extraction workflow so future document batches can be parsed and merged in quickly.

## Subject Properties (28)

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
| 11459 5977-5995 N Broadway | R0103518 | R0103518 | 5977-5995 N. Broadway, Denver, CO 80216 | Industrial – Distribution |
| 5945 & 5957 Broadway | R0103519 | R0103519 | 5945 Broadway, Denver, CO | Commercial / Industrial (multi-tenant) |
| 5720 Washington St | R0103704 | 0182511300068 | 5720 Washington St, Denver, CO 80216-1322 | Office / Warehouse (owner-occupied) |
| 1890 E 58th Ave | R0103779 | 0182511400063 | 1890 E 58th Ave, Denver, CO 80216 | Industrial – Warehouse |
| 550 W 53rd Pl | R0104122 | 0182515202002 | 550 W 53rd Pl, Adams County, CO | Industrial / Retail (multi-tenant) |
| 5710 E. 56th Avenue | R0187789 | R0187789 | 5710 E. 56th Ave, Commerce City, CO | Industrial / Flex (multi-tenant) |
| Confluent Center 70 | R0188699 | R0188699 | 15100 E 40th Ave, Aurora, CO | Industrial (30' clear) |
| 17851 E 40th Avenue | R0198179 | — | 17851 E 40th Ave, Aurora, CO 80011 | Industrial – leased to the State of Colorado |
| Majestic Commercenter – Bldg 5 | R0191099 | R0191099 | 3559 N Himalaya Rd, Aurora, CO 80011 | Industrial – Warehouse |
| 3420 Lisbon Street | R0200721 | R0200721 | 3420 N. Lisbon St, Aurora, CO 80011 | Industrial – Warehouse |
| 5690 E. 56th Avenue | R0187787 | R0187787 | 5690 E. 56th Ave, Commerce City, CO | Industrial / Flex (multi-tenant) |
| Alpine Park II | R0157664 | 0172132216014 | 6045 E 76th Ave, Commerce City, CO | Industrial – Warehouse (multi-bay) |
| 9410 Heinz Way | R0164588 | R0164588 | 9410 Heinz Way, Commerce City, CO | Industrial – Distribution (vacant) |
| 2850 Walden Street | R0164287 | 0182128401002 | 2850 Walden St, Aurora, CO 80011 | Industrial – Warehouse / Distribution |
| Park 70 – 1910 N Gun Club Rd | R0180894 | R0180894 | 1910 N Gun Club Rd, Aurora, CO 80019 | Industrial – Warehouse (Class A) |

## Workbook Tabs

1. **Source Documents** – index of the source files and compilation notes.
2. **Property Master** – one row per property: IDs, parcel, address, owner, assessee, type, tenancy, site/GBA/NRA, year built, land:bldg, occupancy, valuation summary.
3. **Tenants & Leases** – actual rent rolls (125 Bridge St: 11 tenants w/ rent, term, CAM; 5970 Marion units) plus proforma lease assumptions used in the tax appeals.
4. **Assessments & Valuation** – assessed/county values, prior-year values, and income-approach proformas (PGI→NOI→value).
5. **Income-Expense Summary** – total income / expense / net by property and year.
6. **Income-Expense Detail** – full line-item detail; totals use live `=SUM` formulas that match the printed totals.
7. **Related Parcels** – North Side Gardens LLC 4-parcel portfolio co-listed with 7205 Gilpin Way; four Majestic Commercenter / Majestic Lisbon buildings (R0111559, R0111560, R0191099, R0200721); the two KEW Realty buildings on E. 56th Ave (R0187787, R0187789); and the 16 Adams County parcels from the First Industrial 40-parcel letter-of-authorization schedule.

> Scope: this database is for **lease and income/expense** extraction and storage. Sales
> comparables and other market/appraisal support data in the source files are intentionally
> not stored here.

## Standalone Analyses

`Lease_Comparables_Adams_County_Packages.xlsx` — the 43 lease comparables carried by eight of the
twenty Adams County appeal packages (R0110351, R0103518 and R0188699 Ryan/CoStar;
R0121833 Sansone/CoStar; R0164588 Ryan/CoStar; R0187789 Sterling (a set reused verbatim for its
sibling R0187787) and R0200721 Sterling; plus the shared Sterling Exhibit F survey used by R0111559,
R0111560 and R0191099). The other twelve packages contain no comparables. Tabs: Lease Comparables · Set Statistics · Notes &
Data Gaps. Built by `scripts/build_5pkg_lease_comps.py` (+ `..._tabs.py`,
`add_R0103518_lease_comps.py`).

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
- **R0103518** prints two internal $1 rounding differences (Total Rent Revenue $330,136 vs. $330,135 of
  components; Total Common Area Repair/Maint. $32,391 vs. $32,392). Both are reproduced as printed and
  annotated on the Income-Expense Detail rows. Its cover sheet gives the ZIP as 80012 while the rent
  roll and parcel schedule say 80216, and the stated NLA (54,431 SF) exceeds the rent roll's project
  area (50,280 SF). Its assessment year prints as 2024 with a date of value of June 30, **2025**.
- **R0103704 and R0103779** are handwritten county income/expense surveys with no printed totals; the
  totals in this database are sums of the handwritten entries. R0103704's expenses were written into
  the *tenant* column despite 100% owner-occupancy; R0103779 marks janitorial and snow/trash "T"
  rather than a dollar amount, and reports both ~50% owner-occupancy and 100% of the building leased.
- **R0104122** is a rent roll only — no income statement, no lease dates for one of its two tenants.
- **R0103519**'s rent roll lists five spaces by square footage with **no tenant names, lease dates or
  terms**, across five snapshot dates (1/1/23, 6/30/23, 1/1/24, 6/30/24, 1/1/25).
- **R0198179** is an executed **Second Amendment to Lease** rather than an appeal: Prologis, L.P. to the
  State of Colorado (Division of Homeland Security and Emergency Management), 82,131 RSF at
  17851 E 40th Ave, renewal term 7/1/2026–6/30/2033. Its 7-step rent schedule, $2.50/SF TI allowance and
  $128,648.31 JLL commission are all captured. Note the schedule deducts a **$2.32/RSF** property-tax
  credit while the accompanying footnote states Adams County taxes are **$4.53/RSF**; both are printed.
- **R0200721** reports its own in-place lease three different ways — $7.30/SF scheduled on the Exhibit C
  survey, $7.7446/SF on the 06/01/24 rent roll, and $7.60/SF in the letter text. All three are recorded.
  Its comp survey averages $9.30/SF scheduled and $8.41/SF effective, yet the income analysis applies
  $11.00/SF.
- The **3559 N. Himalaya Road, Suite 100** lease (Erickson Metals) is a tenant *at* subject property
  R0191099 and is simultaneously used as a rent comparable in the Sterling surveys for R0111559,
  R0111560 and R0200721, where it is shown at $6.97/SF against the $6.75/SF on R0191099's own rent roll.
- **R0188699** carries no rent roll; occupancy (100%) and actual 2024 NNN income are stated at the
  property level only. Its $1,536,047 printed EGI is $1 below the sum of its own components.
- **R0164588** (9410 Heinz Way) was **100% vacant** at the 6/30/2024 date of value — Home Depot's lease
  expired November 2022 and the space had been listed roughly four years. The assessor still raised it
  37.8% to $17,644,243; the appeal asks $8,539,000.
- **R0187787**'s rent roll shows the Techneaux Technology Services unit at **$0.00 monthly rent** at
  6/30/2024 — the lease had commenced two months earlier and was in a rent-abatement period. The letter
  nonetheless cites it as the base-period lease that is "in-line with market".
- **R0164287** (2850 Walden St) is an **owner-filed** CBOE appeal carrying a 2010–2025 rent history
  ($4.86/SF rising to $7.57/SF). Its work sheet derives a cap rate *from* the assessor's value rather
  than from the market, to show the implied 1.14% return. The county survey ticks "Gross" while its own
  comments state the tenant pays every expense except property tax. Building size is given as 27,800 SF
  on the survey and 27,819 SF on the rent schedule and work sheet.
- **R0157664**'s agent analysis and its underlying statements disagree for 2022: the analysis shows
  income of $258,492 and real estate taxes of $59,147, while the income statement shows $258,992.67
  (it omits $500 of damages income) and $66,104.78 of taxes. Both are recorded as printed.
- **R0180894** reports NLA as 163,790 SF on its salient-facts page and 163,386 SF on its income-trends
  page; its cover sheet prints $150.00 PSF against $150.10 in the assessed-value summary.
