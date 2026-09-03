# R0087486 — 20700 Smith Road, Aurora, CO — Draft Appraisal Report

Source files:
- `Appraisal_Report_TEMPLATE.docx` — BAA restricted appraisal report template
- `Value_Analysis_R0087486.xlsm` — appraiser's value analysis workbook (29 tabs)
- `2025BAA4389_R00857486_20700_SMITH_ROAD.pdf` — Adams County Assessor property profile / cost breakdown / sale comps
- `CostBreakdownSheet.pdf` — Assessor's 3-improvement MVS cost card

Output: `Appraisal_Report_R0087486_20700_Smith_Road_DRAFT.docx`

## What was filled in

- Cover/letter identification fields (account #, parcel #, address, county, tax year,
  statutory date of value 6/30/2024, assessment date 1/1/2025).
- Summary of Salient Facts table (new).
- Site/improvement key-value tables that were directly supported by the workbook
  (zoning, land area, GBA, land-to-building ratio, year built, age, quality/condition,
  construction class by building, utilities, easements, etc.).
- Value Conclusion table (cover letter) with the three approach indications.
- Comparable Sales Grid + Sales Comparison Approach conclusion ($5,400,000 / $275.06 SF),
  built from the **Basic Sales Comparison** tab (4 sales).
- Lease/rent comparables table (subject has no in-place rent roll in the source data —
  used the 10 comps on the **Rent Roll** tab instead, labeled accordingly).
- Income Capitalization Approach pro forma, cap-rate sensitivity table, and conclusion
  ($4,900,000 / $249.59 SF @ 7.25% OAR), from the **Income Approach** tab.
- Cost Approach: per-building RCN/depreciation table and land value/land sales grid,
  from the **Cost Approach MB** (multi-building) and **Land Sales Grid** tabs
  ($6,390,000 / $325.54 SF).
- Reconciliation table with all three indications side by side.

## What was intentionally left blank (bracket placeholders untouched)

These require appraiser judgment, a site inspection, or exhibits/data this workbook
doesn't contain — filling them with invented values would misstate facts in what is
ultimately a legal filing:

- **Final reconciled/concluded value and approach weights** — the workbook's
  Reconciliation tab has all weights at 0% and a $0 concluded value. **This is the
  single most important open item.**
- Highest & Best Use narrative and conclusions (if vacant / as improved), most probable
  purchaser/buyer, valuation-methods-utilized statement.
- Maps, aerial/subject photos, building sketch, market rent/vacancy charts, cap-rate
  survey exhibits, comparable sales/rent maps.
- Site facts not present anywhere in the source files: soil type, flood map panel/date/
  zone, traffic counts, curb cuts, drainage, grade, easements detail beyond "none known",
  parking spaces/ratio.
- Building system detail beyond what the assessor cost card gave (foundation, framing,
  interior finishes, plumbing, electrical, roof) — only heating and general frame/exterior
  by building could be sourced.
- All narrative analysis paragraphs (market & neighborhood description, HBU legal/physical/
  financial feasibility discussion, adjustment rationale prose, site/improvement
  "Analysis/Comments" summaries).

## Data-quality issues found in the source workbook (flag to the appraiser)

1. **Reconciliation weights are unset** (see above) — no final value has been chosen.
2. **Market rent conflict**: Income Approach tab uses $20.00/SF market rent; the
   Comp Leases tab's own "Concluded Market Rent" cell computes $9.00/SF from its
   flagged in-set comps. These are far apart and not reconciled anywhere in the workbook.
3. **Basic Sales Comparison** tab's own note: "comparable data on this grid was recovered
   from links to a workbook that is no longer attached; the source formulas were one
   column out of register. Re-confirm each sale against the subject." Used as-is per the
   tab's stated methodology, but should be independently re-verified.
4. **Land Sales Grid** tab's own note: comparable data "was carried over from the previous
   version of this tab and should be re-confirmed against the subject."
5. **Building count/size mismatch**: the appraiser's Subject Property tab shows 4 buildings
   totaling 19,632 SF (6,832 / 6,400 / 4,000 / 2,400). The Assessor's Cost Breakdown Sheet
   PDF shows 3 improvements totaling 20,216 SF (4,000 / 13,400 / 2,816), one dated 1989
   (vs. 1979 for the others). The two building inventories do not map cleanly onto each
   other; the Cost Approach MB tab's own 4-building breakdown (used in this draft) does
   not obviously come from the assessor's 3-improvement card either. Worth reconciling
   against a field measurement before the report is finalized.
6. **Assessor's $/SF figure** on the Property Info tab ($252.87/SF) is computed off a
   21,365 SF basis, not the appraiser's own 19,632 SF GBA — a different building-size
   convention than everything else in the workbook. This draft recomputed $/SF against
   19,632 SF throughout instead of carrying the stale figure forward.
7. Cost Approach (single-building) tab flags itself: "STEP 2 INCOMPLETE — enter the
   average eave/wall height," which defaults its story-height multiplier to 1.000. This
   draft used the **Cost Approach MB** (multi-building) tab instead, which does have
   per-building heights entered and reconciles to the $6,390,000 figure already carried
   into the Reconciliation tab.
