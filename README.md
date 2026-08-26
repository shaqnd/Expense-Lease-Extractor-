# Master Property Database

`Master_Property_Database.xlsx` — 29 commercial properties, all Adams County, Colorado.
Simplified to **2 tabs** for fast filtering: a property registry and a flat expense ledger.

A reusable Claude Code skill, **`.claude/skills/property-doc-extract`**, captures the
extraction workflow so future document batches can be parsed and merged in quickly.

## Tab 1 — Property Registry

One row per property (29 rows), sorted by OCC Code then Size (SF):

`Property | Property ID | Account/Schedule # | Parcel # | OCC Code | Property Type / Use | Size (SF) | Address | City | ST | County | Tenancy | Occupancy | Year Built | Owner / Client | Assessee | Assessor Value | Assess Yr | Taxpayer Opinion | Source File`

**OCC Code** is a simple use-based classification inferred from each property's `Property Type / Use`
— **IND-WHSE** (industrial warehouse), **IND-DIST** (industrial distribution), **IND-MFG**
(industrial manufacturing), **IND-FLEX** (industrial/flex), **COM-MT** (commercial multi-tenant).
It is **not** an official Adams County CAMA occupancy code — swap in the real code if you have
the county's official list.

## Tab 2 — Expenses

One row per expense line item (344 rows), filterable by **OCC Code**, **Size (SF)**,
**Property Type**, **Category**, and **Year**:

`Property | OCC Code | Property Type | Size (SF) | Year | Category | Line Item / Account | Annual Amount | Amount $/SF | Basis | Notes | Source File | Batch`

- **Annual Amount** is the printed dollar figure from the source document. **Amount $/SF** is a
  live formula (`Annual Amount / Size (SF)`) — blank ("n/a") where the property's SF isn't known
  or the line item has no dollar value.
- **Category** buckets every line item into: `Insurance`, `Utilities`, `CAM & R&M`, `Management`,
  `Real Estate Taxes`, `G&A / Professional Fees`, `Interest Expense`,
  `Depreciation & Amortization` (or `(net)` for a few properties where only a combined
  depreciation/amortization/insurance-credit total was available), `Debt Service (Principal)`,
  `Other / Pass-Through` (tenant billback/DTR accounts, reserves), or `Other / Uncategorized`.
  Filter the Category column down to `Insurance`, `Utilities`, `CAM & R&M`, `Management` for a
  clean operating-expense comp set; add `Real Estate Taxes` if you want that too.
- **Basis** flags `Actual` (from a real income/expense statement or tax return) vs.
  `Proforma Assumption` (the one exception — Shamrock Foods, owner-occupied, no actual tenant,
  so the appeal's hypothetical fee-simple assumptions are used) vs. `N/A` (the 4 NNN leaseholds
  below, where the tenant pays these costs directly and no landlord-side actual exists).
- Every numeric total was verified in Python against the source document's printed total to the
  cent before being entered. Where only a blended total was available (no per-account
  breakdown — a handful of FY2023 statements), the row is tagged Category = `Multiple (Total,
  no line-item detail)` so it's clearly distinguished from real per-category detail.
- **NNN leaseholds** (DEN2, DEN3, DEN7 — the 3 Amazon fulfillment centers — and Home Depot
  Denver) each get a single informational row (`Category = N/A (NNN — tenant-paid)`, no dollar
  amount) instead of expense line items: the tenant pays Insurance/Utilities/CAM/Management
  directly under the lease, so no landlord-side actual $ exists in these appeal documents.
- 10 of the 29 registry properties (7205 Gilpin Way, 17608 E. 24th Dr, Washington Business Park,
  12260 Pennsylvania St, 6770 E. 56th Ave, Lovett 76 Logistics Center, Denali Buildings 1-3, and
  2780 N. Tower Road's second parcel) have **no rows in Expenses** — their source documents were
  tax-appeal income proformas or construction/financing paperwork with no actual expense
  statement, so there's genuinely nothing to report rather than a gap.

## How to filter

Both tabs have AutoFilter enabled on the header row. In Expenses, click the OCC Code, Size (SF),
Property Type, or Category filter arrows to narrow to a comp set (e.g. `OCC Code = IND-WHSE` and
`Category = Insurance` shows every warehouse's actual insurance cost and $/SF side by side).
Select a filtered range of the Annual Amount or Amount $/SF column to see live sum/average in
Excel's status bar — no pivot table needed for a quick comparison.

## Prior structure (history)

Earlier versions of this workbook had 8 tabs (Source Documents, Property Master, Tenants &
Leases, Assessments & Valuation, Income-Expense Summary, Income-Expense Detail, Expense Comps,
Related Parcels) carrying lease terms, assessed values, income-approach proformas, and a source
document index in addition to expense data. That detail wasn't dropped — it's recoverable from
git history (`git log`) on this file — but the working copy was simplified to the 2 tabs above
per request, since the day-to-day need is fast expense-comp filtering, not the full archive.

## Data-quality notes

- Figures are transcribed from the source documents (Ryan LLC / Sterling Property Tax
  Specialists tax appeals, Amazon leasehold appeals, Adams County income & expense surveys, IRS
  Form 8825, P&L / cash-flow statements). Line-item detail is as granular as each source
  document supports — some (2780 N. Tower Road, the 6 re-derived Majestic Commercenter
  buildings, TrusTile, Broadway) have full GL-account detail; others (22100 E 26th Ave/ASB) use
  the statement's own category-level totals because the monthly-column OCR was unreliable.
- "in house" on the TempTee survey = owner self-performs the service (no $ reported), entered as $0.
- 5970 Marion and Broadview totals include financing costs (mortgage interest/principal,
  depreciation, amortization) in their respective Categories — filter those out if you want a
  pure cash operating-expense view.
- Lease/rent comps (CoStar reports, Lowery/RealtyRates surveys, market lease-comp tables) in the
  source documents were skipped throughout — this workbook is scoped to subject-property
  expense data only.
