# Master Property Database — Lease & Income/Expense Extract

`Master_Property_Database.xlsx` consolidates lease, income/expense, and valuation data
extracted and cleaned from the source documents (scanned PDFs, text PDFs, a DOCX appraisal,
and XLSX financials). All subject properties are in **Adams County, Colorado**.

A reusable Claude Code skill, **`.claude/skills/property-doc-extract`**, captures the
extraction workflow so future document batches can be parsed and merged in quickly.

## Subject Properties (34)

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
| 9345/9335 Elm Ct (TARA) | R0048725 | 0171920102027 | 9345 (9335) Elm Ct, CO | Commercial rental |
| I-25 Corporate Center | R0070622-27 | — | 460-550 E. 76th Ave, Denver, CO | Industrial – Warehouse/Mfg |
| Majestic Commercenter | R0083953 | — | 20901 E. 32nd Pkwy, Aurora, CO | Industrial – Warehouse (4 suites) |
| 3254 & 3650 Fraser St | R0084233 | — | 3254-3650 Fraser St, Aurora, CO | Industrial – Warehouse (2 bldgs) |
| 14501 E. 35th Pl | R0084237 | 01821-30-0-02-005 | 14501 E. 35th Pl, Aurora, CO | Industrial – Warehouse/Distribution |
| 1803 E 58th Ave (Arakouzo) | R0103791 | R0103791 | 1803 E 58th Ave, Denver, CO | Industrial – Warehouse/Bakery |
| 14200 E. 33rd Pl | R0084249 | 01821-30-0-04-010 | 14200 E. 33rd Pl, Aurora, CO | Industrial – Warehouse |
| 3250 Abilene St | R0084250 | bt42569 | 3250 Abilene St, Aurora, CO | Industrial – Warehouse |
| 14509 E. 33rd Pl | R0084253 | bt46265 | 14509 E. 33rd Pl, Aurora, CO | Industrial – Warehouse |
| 14705 E. 33rd Pl | — | bt46267 | 14705 E. 33rd Pl, Aurora, CO | Industrial – Warehouse |
| 14501 E. Moncrieff Pl | R0084268 | — | 14501 E. Moncrieff Pl, Aurora, CO | Industrial – Warehouse/Distribution |
| Aurora - Uravan (Wagner) | R0086203 | R0086203 | 2420 Uravan St, Aurora, CO | Industrial – Warehouse |
| Cast Transportation Sub. (BTS) | R0213618 | R0213618 | Cast Transportation Subdivision, Commerce City, CO | Industrial – Distribution (BTS) |
| Grand Lake | R0179014 | 0182505219006 | 1010 Lousell Blvd, Westminster, CO | Warehouse / Light Industrial |
| 8500 Brighton Partners | R0179050 | R0179050 | 8510 Brighton Rd, Adams County, CO | Industrial – Vehicle Auction/Storage |
| Bohan Family — Auto Body | R0100117 | 0182505224001 | 3580 W. 72nd Ave, Westminster, CO 80030 | Auto Body Shop |
| Northlawn Auto Center | R0100737 | 1973-35-2-41-001 | 6500 N. Federal Blvd, Denver, CO 80221 | Retail – Auto Repair/Service |
| Digby Family — 1225 W 64th | R0099649 | 182504402007 | 1225 W 64th Ave, Denver, CO | Industrial |
| DSP LLC — Newport St | R0092740 | 0182317405014 | 1975 Newport St, Commerce City, CO 80022 | Industrial / Warehouse |
| Meineke Car Care — Colfax Ave | R0085560 | R0085560 | 14891 East Colfax Avenue, Aurora, CO | Retail – Auto Service |
| Diamond Beall Dev — 74th Ave | R0077963 | 0172131402020 | 1750 74th Ave, Commerce City, CO 80022 | Industrial (Tank Trailer/Truck) |
| B and M Equipment — Hwy 85 | R0077789 | 0172131110004 | 7908 Highway 85, Commerce City, CO | Industrial (owner-occupied) |
| RJR Investments — Adams Co. | R0075352 | 0172116004003 | (address illegible in OCR), Adams County, CO | Warehouse / Office / Shop |
| Baizer Properties — Collision Repair | R0037193 | 0171910307020 | (address illegible in OCR), Northglenn area, CO | Collision Repair Shop |
| 11450 Huron St (Wallace Associates) | R0030089 | 1719-03-0-05-016, 019 | 11450 Huron St, Denver, CO | Office / Medical (multi-tenant) |
| Peerless Tyre Co — Huron St | R0030085 | 0171903005007 | 490 Huron St, Denver, CO | Retail – Tire Store / Light Mechanical |

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

### 325 Folder Batch (14 properties, 15 source files)

- **R0179050** (8510 Brighton Rd): the original 8.4 MB PDF could not be downloaded in this
  session after repeated Google Drive MCP timeouts. Its data is transcribed entirely from
  Google Drive's server-side OCR text extract (`read_file_content`), not from a locally
  rendered/parsed PDF. Its cost-approach summary sheet, cover letter, and an embedded lease
  (8500 Brighton LLC / Insurance Auto Auctions, Inc.) all yielded usable text via this route.
- Several County Income & Expense Survey scans in this batch have partly illegible handwritten
  `$` figures per OCR (**R0075352, R0037193, R0179014, R0092740**); illegible fields are left
  blank rather than guessed, and flagged in the relevant Property Master / Tenants & Leases /
  Income-Expense notes.
- **R0030089** (11450 Huron St): the printed 2021/2022 "Total Expenses" exceed the sum of the
  legible line items in the dense OCR'd financial grid. A balancing "Other / not individually
  legible in OCR" line was added in each year so the `=SUM` formula reproduces the printed
  total exactly, rather than silently understating it.
- **R0030085 / R0030088 / R0030089** share adjacent Adams Co. schedule numbers
  (`1719-03-0-05-xxx`), suggesting units within the same Huron St. complex — noted in Related
  Parcels but not confirmed from a plat or ownership record.
- **R0213618**: two source files — the original ground lease (Scannell Properties #534, LLC /
  Performance Food Group, Inc., 20-yr term + 2×5-yr options) and a later Assignment and
  Assumption of Lease (landlord interest → Bel Commerce LLC, 5/24/2024). Neither document
  states the actual rent amount within the extracted text.
- Per the skill's scope rule, CoStar sale-comparable data bundled into several of these PDFs
  (R0100737, R0030089) was **not** stored — only the subject property's own lease/income/expense
  and assessment data was extracted.

### Batches 2-3 (12 properties: R0048725 through R0086203)

These 12 properties were added to the canonical copy of `Master_Property_Database.xlsx` on
Google Drive by other work sessions between this repo's last sync and the 325-folder update
above — they were never committed to this git repo. When updating the Drive file with the
325-folder batch, the current Drive copy (not this repo's stale local copy) was used as the
merge base, so this repo now also reflects those 12 properties. Their extraction caveats
(Sterling appeals' tax treatment, 1st Net's formulaic 3%+5% expense ratios, accrual vs. cash vs.
tax-basis accounting, the R0091936/R0091934 duplicate upload, source-document discrepancies on
2021-22 income data, rent-roll SF counts, and county values) are documented inline in the
workbook's own notes rows on each tab — see **Assessments & Valuation** row 33 and **Source
Documents** rows 30-32 for the full detail, since they were authored by those sessions and are
not duplicated here.

## Syncing with Google Drive

The canonical, most up-to-date copy of this workbook lives on Google Drive in the
**Expenses & Leases** folder as `Master_Property_Database.xlsx`. Multiple sessions/agents may
update that Drive copy directly (e.g. from other document-batch folders like `407` or
`Mega Warehouse`) without necessarily pushing back to this git repo. When asked to "update the
master database," always treat the **current Drive copy** as the source of truth to merge into
— not just this repo's last commit — to avoid silently discarding other sessions' work.
