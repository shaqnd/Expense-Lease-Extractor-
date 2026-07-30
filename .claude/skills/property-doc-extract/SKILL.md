---
name: property-doc-extract
description: Extract lease, income/expense, and valuation data from commercial real-estate property documents (tax-appeal packages, appraisals, P&Ls, cash-flow statements, rent rolls, assessor income/expense surveys, IRS Form 8825) and consolidate them into the master property database spreadsheet. Handles scanned PDFs, text PDFs, DOCX appraisals, and XLSX financials. Trigger whenever the user uploads property/lease/expense documents and wants them parsed, cleaned, organized, or added to the master database.
---

# Property Document Extraction → Master Database

Turn a batch of messy CRE documents (mostly Adams County, CO tax/appraisal material) into
clean rows in `Master_Property_Database.xlsx`. This skill exists because the sandbox has
**no poppler and a broken LibreOffice**, so the naive Read-the-PDF path fails — the steps
below are the workflow that actually works here.

## 0. Environment prep (once per session)

```bash
pip install pypdfium2 openpyxl python-docx pillow   # markitdown/pdfplumber are NOT reliable here
```

- `pdftoppm`/poppler is **absent** — the built-in Read tool cannot render PDF pages. Use `pypdfium2`.
- `pdfplumber` imports fail (`_cffi_backend`). Do **not** use it.
- LibreOffice/`soffice` exists but **cold-starts too slowly** — `recalc.py` times out even at 500s+.
  Do not block on it; see step 5.

## 1. Triage every file by type

```python
import pypdfium2 as pdfium
pdf = pdfium.PdfDocument(path)
for i in range(len(pdf)):
    t = pdf[i].get_textpage().get_text_range() or ""
    # <5 chars/page  => scanned image page, needs visual render (step 3)
```

- **Text PDF** → text extracts cleanly, parse directly.
- **Scanned PDF** (near-zero chars) → render to PNG and Read visually (step 3).
- **DOCX** → `python-docx`: pull `paragraphs` **and** every `tables` cell. Financial grids
  (rent roll, income proforma, concluded value, salient facts) are often **embedded EMF/PNG
  images**, not table cells — extract them from the zip (`word/media/`). PNGs are Read-able;
  **EMF vector images cannot be converted here** (no LibreOffice/inkscape/imagemagick) — capture
  whatever numbers the narrative text gives and flag the tables as "embedded, not text-extractable".
- **XLSX** → `openpyxl load_workbook(data_only=True)`; iterate rows. Watch for **hidden helper
  columns** off to the right (e.g. an income-approach block: NOI, cap rate, assigned value, $/SF).

Note: the harness's reported page count can be wrong (it once said 199 for a 16-page file).
Trust `len(pdf)` from pypdfium2/pypdf.

## 2. Extract text PDFs to .txt for reading

```python
txt = "\n".join(f"\n===== PAGE {i+1} =====\n{pdf[i].get_textpage().get_text_range()}"
                for i in range(len(pdf)))
```

Some Ryan/CoStar pages encode headers in a **symbol font** → labels come out as gibberish but
the **dollar values are usually still readable**. When a whole page is gibberish (CoStar comp
sheets), render it (step 3) instead.

## 3. Render scanned / image pages, then Read them

```python
img = pdf[page_index].render(scale=2.0).to_pil()   # scale 1.7–2.2 for full pages
img.save("DOC_pN.png")
```

Then use the **Read tool on the PNG** to read it visually (handwriting, stamps, checkboxes,
signatures all work). For a **dense grid** (rent roll, monthly P&L), render at `scale=5.0–6.0`
and **crop to quadrants** before Reading, or the text is too small:

```python
w,h = img.size
img.crop((0, int(h*0.05), int(w*0.34), int(h*0.24))).save("crop_rentroll.png")
```

Read page 1 of every doc first to identify the property (account #, address, owner), then only
render the pages that carry data (salient facts, proforma, rent roll, P&L) — skip photo/map/
boilerplate pages.

## 4. What to pull from each document type

- **Ryan "Valuation Protest / Property Tax Assessment Appeal"** (Adams County): cover = account #,
  address, owner/client, assessor value, taxpayer opinion, $/SF. p2 = property summary (assessee,
  parcel, site class, land AC, year built, SF, prior/current county value). p3 = Market Pro Forma
  (face rent, vacancy %, expense %, cap rate, PGI/EGI/OpEx/NOI, indicated value). Later pages =
  CoStar sale/rent comps — **skip these** (out of scope; see step 5).
- **County Income & Expense Survey**: owner-occupancy %, use, and the operating-expense grid
  (owner vs tenant columns). "in house" = owner self-performs, enter $0 and annotate.
- **P&L / Cash-Flow (Deerwoods, etc.)**: income accounts (rent, CAM escrows), expense accounts,
  totals, net income; note cash-basis and whether debt service/taxes are included (they usually are,
  so it's cash flow, not NOI). Parcel # often appears in the property-tax line memo.
- **IRS Form 8825**: gross rents, per-line expenses, total expenses, net income/(loss), EIN, per-property.
- **Restricted Appraisal (DOCX)**: owner of record, contact, site/GBA/NRA, year built, zoning,
  land:bldg, lease structure (modified gross/NNN), actual rent range, vacancy, OpEx $/SF & %EGI,
  sales-comp adjusted range, concluded values (often in EMF images — may be unrecoverable as text).
- **Rent roll**: unit, tenant, monthly rent, lease start/end, term/option, deposit, CAM.

## 5. Write into Master_Property_Database.xlsx

Structure (row-per-record so it scales as documents keep arriving):

1. **Source Documents** – index (file, type, property, period, preparer).
2. **Property Master** – one row per property: IDs, parcel, address, owner, assessee, type,
   tenancy, site/GBA/NRA, year built, land:bldg, occupancy, assessor value, taxpayer opinion, source.
3. **Tenants & Leases** – one row per tenant/lease: property, unit, tenant, monthly & annual rent,
   $/SF, term start/end, option, CAM, deposit, basis, notes.
4. **Assessments & Valuation** – historical assessed values + income-approach proformas.
5. **Income-Expense Summary** – total income / expense / net by property & year.
6. **Income-Expense Detail** – line items; totals via `=SUM` cross-checking printed totals.
7. **Related Parcels** – co-listed parcel schedules, as available.

**Scope: lease and income/expense only.** Do **not** store sales comparables, rent comps, or
other CoStar/appraisal market-support data — skip those pages during extraction. The database
holds subject-property identity, leases/rent rolls, income/expense, and valuation summaries.

Conventions: Arial; `$#,##0` money, `0.00%` percents stored as fractions; navy title bars;
blue header rows; green subtotal fills; cite the source file in each block; annotate every
assumption ("in house"→$0, value-in-image, cover-vs-schedule discrepancies).

**Recalc caveat:** `recalc.py`/LibreOffice will time out in this sandbox. Instead, **verify every
`=SUM`/`=AVERAGE` in Python** against the document's printed total before shipping (they must match
to the cent), and tell the user the formula cells populate on open in Excel/Sheets. Do not claim a
green recalc you couldn't run.

## 6. Deliver

Copy the xlsx into the repo, update `README.md`, commit to the working branch, push, and
`SendUserFile` the spreadsheet. Report new properties added and any data that couldn't be
extracted (e.g. EMF-embedded appraisal tables) so the user knows the gaps.

## Known entities in this portfolio (all Adams County, CO)

| Account | Address | Owner / Client |
|---|---|---|
| R0178308 | 7205 Gilpin Way, Denver | Center Land Properties / North Side Gardens LLC |
| R0169133 | 17608 E. 24th Dr, Aurora | ABC Supply Co. |
| R0103792 | 2011 E 58th Ave, Denver | TempTee Brand Steaks |
| — (0182511303018) | 5970 Marion Drive, Denver | Ankur Kumar / Deerwoods mgmt |
| R0103767 | 5650 Washington St, Denver | Washington Business Park Property LLC |
| R0024442 | 12260 Pennsylvania St, Thornton | Trinity Real Estate / Snyder Family Trust |
| R0002819 | 125 Bridge St, Brighton | Broadview LLC |
| R0092302 | 6770 E. 56th Ave, Aurora | 56th Ave. J Investments LLC |

Common players: Ryan LLC (Bennett Mecom, Jamie Hoffman) — tax agent; Deerwoods Management LLC —
property manager; Stevens & Associates — tax-reduction agent; Property Tax Advisors Inc — tax agent;
Jason Bennett, MAI — Adams County appraiser; Adams County parcels start `0182511...`.
