# Bay Area business ranker

Ranks independent (non-franchise) San Francisco Bay Area businesses for sale under $1M into an Excel workbook.

- `RUNBOOK.md`: step-by-step research and scoring instructions (used by the Claude skill).
- `build_ranking.py`: builds the workbook from a data JSON file. Scores and ranks are live Excel formulas; weights are editable on the Criteria sheet.
- `data-business-YYYY-MM-DD.json`: one snapshot of listings per run.

```
pip install openpyxl
python3 build_ranking.py data-business-2026-11-22.json business-ranking-2026-11-22.xlsx
```

Research only, not financial advice. Listing figures are sellers' claims.
