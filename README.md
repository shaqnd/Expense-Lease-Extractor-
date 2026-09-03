# Master Property Database — Lease & Income/Expense Extract

`Master_Property_Database.xlsx` consolidates lease, income/expense, and valuation data
extracted and cleaned from the source documents (scanned PDFs, text PDFs, a DOCX appraisal,
and XLSX financials). All subject properties are in **Adams County, Colorado**.

A reusable Claude Code skill, **`.claude/skills/property-doc-extract`**, captures the
extraction workflow so future document batches can be parsed and merged in quickly.

## Subject Properties (8)

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

## Workbook Tabs

1. **Source Documents** – index of the source files and compilation notes.
2. **OCCC Legend** – the OCCC (Occupancy Classification Code) key: one short code per distinct
   `Property Type / Use`, with its description and which properties carry it.
3. **Property Master** – one row per property: IDs, parcel, address, owner, assessee, type, tenancy, site/GBA/NRA, year built, land:bldg, occupancy, valuation summary, **OCCC Code**.
4. **Tenants & Leases** – actual rent rolls (125 Bridge St: 11 tenants w/ rent, term, CAM; 5970 Marion units) plus proforma lease assumptions used in the tax appeals.
5. **Assessments & Valuation** – assessed/county values, prior-year values, and income-approach proformas (PGI→NOI→value).
6. **Income-Expense Summary** – total income / expense / net by property and year, with **OCCC Code**.
7. **Income-Expense Detail** – full line-item detail; totals use live `=SUM` formulas that match the printed totals. Includes **OCCC Code**.
8. **Related Parcels** – North Side Gardens LLC 4-parcel portfolio co-listed with 7205 Gilpin Way.

> Scope: this database is for **lease and income/expense** extraction and storage. Sales
> comparables and other market/appraisal support data in the source files are intentionally
> not stored here.

## Filtering by OCCC Code

No OCCC (Occupancy Classification Code) field existed in any source document or in the prior
version of this workbook. To make the properties and their expense rows filterable by use type,
this version assigns a short **OCCC Code** to every property, derived from its existing
`Property Type / Use` value (see the **OCCC Legend** tab for the key):

| OCCC Code | Description | Properties |
|---|---|---|
| IND-WH | Industrial – Warehouse | 7205 Gilpin Way, SGS – Pennsylvania Industrial, 6770 E. 56th Ave |
| IND-DIST | Industrial – Distribution | (7194) CO-Aurora |
| IND-MFG | Industrial – Mfg / Food Processing | TempTee Brand Steaks |
| IND-2U | Industrial – 2-unit (multi-tenant) | 5970 Marion Drive |
| IND-FLEX | Industrial / Flex (multi-tenant) | Washington Business Park |
| COM-MT | Commercial (multi-tenant industrial) | Broadview |

The **OCCC Code** column is on **Property Master**, **Income-Expense Summary**, and
**Income-Expense Detail**, and AutoFilter is enabled on all three tabs (plus the legend), so you
can filter expense line items, summary totals, or the property roster down to one use type in a
click. If "OCCC code" refers to a specific code your organization or the county assessor already
issues (rather than a use-type classification), let me know the actual codes/scheme and I'll
remap this column instead of the assigned ones above.

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
