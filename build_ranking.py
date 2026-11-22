#!/usr/bin/env python3
"""Build the franchise ranking workbook from a JSON data file.

Usage: python3 build_ranking.py data.json output.xlsx

Scores are live Excel formulas: change a weight on the Criteria sheet or a
1-5 factor score on the Ranking sheet and the totals and ranks recalculate.
"""
import json
import sys

from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

FACTORS = [
    ("ue", "Unit economics (Item 19)"),
    ("invest", "Investment fit"),
    ("fees", "Fees and royalties"),
    ("bay", "Bay Area market fit"),
    ("territory", "Territory availability"),
    ("health", "Franchisor health"),
]
HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")


def autosize(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def add_table(ws, ref, name):
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    ws.add_table(t)


def build(data, out):
    wb = Workbook()

    # Criteria sheet (weights referenced by formulas)
    crit = wb.active
    crit.title = "Criteria"
    if data.get("franchises"):
        crit.append(["Factor", "Weight", "What earns a 5"])
        style_header(crit, 1, 3)
        for key, label in FACTORS:
            w = data["weights"][key]
            crit.append([label, w["weight"], w["five"]])
        crit.append(["Total", "=SUM(B2:B7)", "Must equal 100%"])
        for r in range(2, 9):
            crit.cell(row=r, column=2).number_format = "0%"
        crit.cell(row=8, column=1).font = Font(bold=True)
        autosize(crit, [28, 10, 90])

        # Ranking sheet
        rk = wb.create_sheet("Ranking", 0)
        rk["A1"] = data["title"]
        rk["A1"].font = Font(bold=True, size=14)
        rk["A2"] = f"As of {data['as_of']}. {data['disclaimer']}"
        rk["A2"].font = Font(italic=True, color="595959")
        headers = [
            "Rank", "Franchise", "Where it's advertised", "Category", "Investment low ($)", "Investment high ($)",
            "Franchise fee ($)", "Ongoing fees (% of sales)", "Liquid capital ($)",
            "Item 19 avg sales ($)", "Item 19 period", "Sales / high investment",
        ] + [label + " (1-5)" for _, label in FACTORS] + ["Score /100", "Why", "Risks", "Research source"]
        hr = 4
        rk.append([])
        for c, h in enumerate(headers, start=1):
            rk.cell(row=hr, column=c, value=h)
        style_header(rk, hr, len(headers))
        rk.row_dimensions[hr].height = 45

        n = len(data["franchises"])
        first, last = hr + 1, hr + n
        fcol = 13  # first factor column (M)
        score_col = fcol + len(FACTORS)
        sc = get_column_letter(score_col)
        for i, f in enumerate(data["franchises"]):
            r = hr + 1 + i
            rk.cell(row=r, column=1, value=f"=RANK({sc}{r},${sc}${first}:${sc}${last})")
            rk.cell(row=r, column=2, value=f["name"])
            ad = rk.cell(row=r, column=3, value=f.get("ad_url"))
            if f.get("ad_url"):
                ad.hyperlink = f["ad_url"]
                ad.font = Font(color="0563C1", underline="single")
            rk.cell(row=r, column=4, value=f["category"])
            rk.cell(row=r, column=5, value=f.get("invest_low"))
            rk.cell(row=r, column=6, value=f.get("invest_high"))
            rk.cell(row=r, column=7, value=f.get("franchise_fee"))
            rk.cell(row=r, column=8, value=f.get("ongoing_fee_pct"))
            rk.cell(row=r, column=9, value=f.get("liquid_capital"))
            rk.cell(row=r, column=10, value=f.get("item19_avg"))
            rk.cell(row=r, column=11, value=f.get("item19_period", "Not disclosed"))
            rk.cell(row=r, column=12, value=f'=IF(OR(J{r}="",F{r}=""),"",J{r}/F{r})')
            for j, (key, _) in enumerate(FACTORS):
                rk.cell(row=r, column=fcol + j, value=f["scores"][key])
            parts = "+".join(
                f"{get_column_letter(fcol + j)}{r}*Criteria!$B${2 + j}" for j in range(len(FACTORS))
            )
            rk.cell(row=r, column=score_col, value=f"=ROUND(({parts})/5*100,0)")
            rk.cell(row=r, column=score_col + 1, value=f.get("why"))
            rk.cell(row=r, column=score_col + 2, value=f.get("risks"))
            src = rk.cell(row=r, column=score_col + 3, value=f.get("source"))
            if f.get("source"):
                src.hyperlink = f["source"]
                src.font = Font(color="0563C1", underline="single")
            for col in (5, 6, 7, 9, 10):
                rk.cell(row=r, column=col).number_format = "$#,##0"
            rk.cell(row=r, column=8).number_format = "0.0%"
            rk.cell(row=r, column=12).number_format = '0.00"x"'
            for col in (score_col + 1, score_col + 2):
                rk.cell(row=r, column=col).alignment = Alignment(wrap_text=True, vertical="top")
        rk.conditional_formatting.add(
            f"{sc}{first}:{sc}{last}",
            ColorScaleRule(start_type="min", start_color="F8696B", mid_type="percentile",
                           mid_value=50, mid_color="FFEB84", end_type="max", end_color="63BE7B"),
        )
        add_table(rk, f"A{hr}:{get_column_letter(len(headers))}{last}", "Ranking")
        rk.freeze_panes = rk.cell(row=hr + 1, column=4)
        autosize(rk, [7, 24, 40, 20, 13, 13, 12, 12, 13, 14, 16, 12] + [12] * len(FACTORS) + [10, 50, 45, 40])
        rk.cell(row=last + 2, column=1,
                value="Sort by Score (or the Rank column) after editing; scores and ranks are formulas.").font = Font(italic=True)

    # Businesses for sale (franchise resales + independent businesses), ranked
    bw = data.get("business_weights") or {
        "payback": {"weight": 0.40, "five": "Asking price / cash flow at or under 2 years (computed)"},
        "margin": {"weight": 0.20, "five": "Cash flow at or above 25% of revenue (computed)"},
        "earnings": {"weight": 0.15, "five": "Cash flow (SDE) at or above $250K a year (computed)"},
        "lease": {"weight": 0.15, "five": "Rent under 8% of revenue with 5+ years of lease left (manual 1-5)"},
        "info": {"weight": 0.10, "five": "Financials shown and verifiable, clear reason for sale, fits the buyer (manual 1-5)"},
    }
    bkeys = ["payback", "margin", "earnings", "lease", "info"]
    blabels = ["Payback", "Cash-flow margin", "Earnings size", "Lease / rent risk", "Info quality and fit"]
    crit.append([])
    crit.append(["Businesses for sale factor", "Weight", "What earns a 5"])
    style_header(crit, crit.max_row, 3)
    wrow = crit.max_row + 1
    for k, label in zip(bkeys, blabels):
        crit.append([label, bw[k]["weight"], bw[k]["five"]])
        crit.cell(row=crit.max_row, column=2).number_format = "0%"
    crit.append(["Total", f"=SUM(B{wrow}:B{wrow + 4})", "Must equal 100%"])
    crit.cell(row=crit.max_row, column=2).number_format = "0%"
    crit.cell(row=crit.max_row, column=1).font = Font(bold=True)

    biz = list(data.get("businesses", []))
    for x in data.get("resales", []):  # older data files
        biz.append({"name": x["brand"], "type": "Franchise resale", "category": x.get("category", ""),
                    "location": x["location"], "asking": x.get("asking"), "revenue": x.get("revenue"),
                    "cash_flow": x.get("cash_flow"), "notes": x.get("notes"), "url": x.get("url"),
                    "lease_score": x.get("lease_score", 3), "info_score": x.get("info_score", 3)})
    def _band(v, cuts, reverse=False):
        if v is None:
            return 1
        for score, cut in zip((5, 4, 3, 2), cuts):
            if (v <= cut) if reverse else (v >= cut):
                return score
        return 1

    def _py_score(x):
        a, rev, cf = x.get("asking"), x.get("revenue"), x.get("cash_flow")
        pb = a / cf if a and cf else None
        mg = cf / rev if cf and rev else None
        parts = {"payback": _band(pb, (2, 2.5, 3, 4), reverse=True),
                 "margin": _band(mg, (0.25, 0.2, 0.15, 0.1)),
                 "earnings": _band(cf, (250000, 175000, 120000, 75000)),
                 "lease": x.get("lease_score", 3), "info": x.get("info_score", 3)}
        return sum(parts[k] * bw[k]["weight"] for k in bkeys)

    biz.sort(key=_py_score, reverse=True)
    rs = wb.create_sheet("Businesses for sale", 1 if data.get("franchises") else 0)
    rs["A1"] = data.get("business_title", "Bay Area businesses for sale, ranked")
    rs["A1"].font = Font(bold=True, size=14)
    rs["A2"] = "Payback, margin and earnings scores are computed from the listing; lease and info scores are judgment calls you can edit."
    rs["A2"].font = Font(italic=True, color="595959")
    bh = ["Rank", "Business", "Type", "Category", "Location", "Listing (where it's advertised)",
          "Asking ($)", "Revenue ($)", "Cash flow / SDE ($)", "Monthly rent ($)",
          "Payback (years)", "Cash-flow margin", "Rent % of revenue"] + \
         [l + " (1-5)" for l in blabels] + ["Score /100", "Notes"]
    for c, h in enumerate(bh, start=1):
        rs.cell(row=4, column=c, value=h)
    style_header(rs, 4, len(bh))
    rs.row_dimensions[4].height = 45
    b_first, b_last = 5, 4 + max(len(biz), 1)
    for i, x in enumerate(biz):
        r = 5 + i
        rs.cell(row=r, column=1, value=f"=RANK(S{r},$S${b_first}:$S${b_last})")
        rs.cell(row=r, column=2, value=x["name"])
        rs.cell(row=r, column=3, value=x.get("type", "Independent"))
        rs.cell(row=r, column=4, value=x.get("category"))
        rs.cell(row=r, column=5, value=x.get("location"))
        u = rs.cell(row=r, column=6, value=x.get("url"))
        if x.get("url"):
            u.hyperlink = x["url"]
            u.font = Font(color="0563C1", underline="single")
        rs.cell(row=r, column=7, value=x.get("asking"))
        rs.cell(row=r, column=8, value=x.get("revenue"))
        rs.cell(row=r, column=9, value=x.get("cash_flow"))
        rs.cell(row=r, column=10, value=x.get("monthly_rent"))
        rs.cell(row=r, column=11, value=f'=IF(AND(N(G{r})>0,N(I{r})>0),G{r}/I{r},"")')
        rs.cell(row=r, column=12, value=f'=IF(AND(N(H{r})>0,N(I{r})>0),I{r}/H{r},"")')
        rs.cell(row=r, column=13, value=f'=IF(AND(N(H{r})>0,N(J{r})>0),J{r}*12/H{r},"")')
        rs.cell(row=r, column=14, value=f'=IF(K{r}="",1,IF(K{r}<=2,5,IF(K{r}<=2.5,4,IF(K{r}<=3,3,IF(K{r}<=4,2,1)))))')
        rs.cell(row=r, column=15, value=f'=IF(L{r}="",1,IF(L{r}>=0.25,5,IF(L{r}>=0.2,4,IF(L{r}>=0.15,3,IF(L{r}>=0.1,2,1)))))')
        rs.cell(row=r, column=16, value=f'=IF(N(I{r})=0,1,IF(I{r}>=250000,5,IF(I{r}>=175000,4,IF(I{r}>=120000,3,IF(I{r}>=75000,2,1)))))')
        rs.cell(row=r, column=17, value=x.get("lease_score", 3))
        rs.cell(row=r, column=18, value=x.get("info_score", 3))
        parts = "+".join(f"{get_column_letter(14 + j)}{r}*Criteria!$B${wrow + j}" for j in range(5))
        rs.cell(row=r, column=19, value=f"=ROUND(({parts})/5*100,0)")
        rs.cell(row=r, column=20, value=x.get("notes"))
        rs.cell(row=r, column=20).alignment = Alignment(wrap_text=True, vertical="top")
        for col in (7, 8, 9, 10):
            rs.cell(row=r, column=col).number_format = "$#,##0"
        rs.cell(row=r, column=11).number_format = "0.0"
        rs.cell(row=r, column=12).number_format = "0%"
        rs.cell(row=r, column=13).number_format = "0%"
    if biz:
        add_table(rs, f"A4:{get_column_letter(len(bh))}{b_last}", "BusinessesForSale")
        rs.conditional_formatting.add(
            f"S{b_first}:S{b_last}",
            ColorScaleRule(start_type="min", start_color="F8696B", mid_type="percentile",
                           mid_value=50, mid_color="FFEB84", end_type="max", end_color="63BE7B"),
        )
    rs.freeze_panes = "C5"
    autosize(rs, [7, 30, 16, 20, 22, 40, 12, 12, 14, 12, 10, 10, 10, 10, 10, 10, 10, 10, 10, 60])

    # Screened out
    if data.get("screened_out"):
        so = wb.create_sheet("Screened out")
        so.append(["Name", "Reason", "Source", "Link"])
        style_header(so, 1, 4)
        for x in data.get("screened_out", []):
            so.append([x["name"], x["reason"], x.get("source"), x.get("ad_url") or x.get("url")])
        autosize(so, [26, 90, 60, 50])

    # Notes
    nt = wb.create_sheet("Notes")
    for line in data.get("notes", []):
        nt.append([line])
    nt.column_dimensions["A"].width = 140

    wb.save(out)


if __name__ == "__main__":
    with open(sys.argv[1]) as fh:
        build(json.load(fh), sys.argv[2])
    print("wrote", sys.argv[2])
