<!-- To run: in this project, ask "Run the business ranker using franchise-agent/RUNBOOK.md". -->
# Bay Area independent business ranker

Produces a ranked Excel workbook of **independent (non-franchise)** businesses for sale in the San Francisco Bay Area. Default brief: food and beverage, entertainment, or any other sector, asking price up to $1,000,000. No franchises and no franchise resales. If the person gives a different budget, region or category for this run, use theirs.

Read-only research. Never contact brokers or sellers, never submit forms, NDAs or sign-ups. Label output as research, not financial advice.

## Steps

1. **Load last run.** Read the newest `data-business-*.json` in `/mnt/project-files/franchise-agent/` (or this repository) and use it as the baseline, so you can report new, sold or changed listings.
2. **Find listings.** Search bizben.com county and city pages for SF, San Mateo, Santa Clara, Alameda, Contra Costa, Marin, Sonoma, Napa and Solano (e.g. `bizben.com/business-for-sale/san-mateo-county-california-ca`, `.../san-francisco-county-california-ca`, `.../san-jose-city-california-ca-businesses-for-sale`, `.../fremont-city-california-ca-businesses-for-sale`), plus bizbuysell.com and bizquest.com Bay Area food and entertainment pages; note any site that blocks reading. Open individual listing pages when the card lacks cash flow.
3. **Filter.** Keep Bay Area listings at or under budget. Drop anything that is a franchise or franchise resale (named brands, "franchise" in the title or description, "branded" concepts unless clearly independent), anything marked In Contract or In Escrow, and duplicates. Aim for 20 to 30 listings, favoring ones that show cash flow.
4. **Record** for each: name as listed, category, city, asking, revenue, cash flow/SDE, monthly rent, lease remaining (in notes) and the exact listing URL. Set `"type": "Independent"`.
5. **Judge two scores** (1-5): `lease_score` (5 = rent under 8% of revenue with 5+ years left, or no storefront lease; 1 = high rent or short/unknown lease) and `info_score` (5 = full financials on the listing page; 1 = no financials or conflicting figures). The workbook computes payback (asking / cash flow), cash-flow margin and earnings size and ranks everything.
6. **Build.** Write `data-business-YYYY-MM-DD.json` (schema below) to this folder and run `python3 build_ranking.py data-business-YYYY-MM-DD.json business-ranking-YYYY-MM-DD.xlsx`. Recalculate with LibreOffice if available to confirm the formulas compute.
7. **Report.** Attach the workbook. In 3 to 5 sentences: top 3 listings with payback, new or removed listings since last run, and source gaps.

## Data JSON schema

```json
{
  "title": "Bay Area independent businesses for sale (up to $1M)",
  "business_title": "Bay Area independent (non-franchise) businesses for sale, ranked",
  "as_of": "YYYY-MM-DD",
  "disclaimer": "Research, not financial advice. Figures are sellers' claims from public listings; verify before any offer.",
  "businesses": [{"name": "", "type": "Independent", "category": "", "location": "", "asking": 0, "revenue": 0,
                  "cash_flow": 0, "monthly_rent": 0, "lease_score": 3, "info_score": 3, "notes": "", "url": "https://..."}],
  "screened_out": [{"name": "", "reason": "", "url": ""}],
  "notes": ["..."]
}
```
Omit numeric keys you don't know (never write 0). Optional `business_weights` overrides the default weights (payback 40%, margin 20%, earnings 15%, lease 15%, info 10%).
