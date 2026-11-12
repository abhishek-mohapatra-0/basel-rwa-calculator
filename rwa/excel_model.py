"""Write the RWA calculator as a live Excel model.

Every calculated cell is a formula pointing at the Parameters sheet, so changing
a risk weight, CCF or the capital ratio re-prices the whole book in Excel.
The Python engine's totals are written next to the Excel totals as a reconciliation.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .engine import summary
from .rules import BUCKETS, CAPITAL_RATIO, CCF, LTV_BANDS, RATED_RW, RATING_BUCKET, RETAIL_RW

ARIAL = "Arial"
BLUE = Font(name=ARIAL, color="0000FF")
BLACK = Font(name=ARIAL)
GREEN = Font(name=ARIAL, color="008000")
BOLD = Font(name=ARIAL, bold=True)
HDR_FILL = PatternFill("solid", fgColor="1F3864")
YELLOW = PatternFill("solid", fgColor="FFFF00")
thin = Side(style="thin", color="BFBFBF")
BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
CLASSES = ["Sovereign", "Bank", "Corporate", "Retail", "Residential Mortgage"]
AMT = "#,##0.00;(#,##0.00);-"
PCT = "0.0%;(0.0%);-"


def hdr(ws, row, col, labels):
    for j, h in enumerate(labels):
        c = ws.cell(row=row, column=col + j, value=h)
        c.font = Font(name=ARIAL, bold=True, color="FFFFFF"); c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = BOX


def put(ws, r, c, v, font=BLACK, fmt=None):
    cell = ws.cell(row=r, column=c, value=v); cell.font = font; cell.border = BOX
    if fmt:
        cell.number_format = fmt
    return cell


def parameters(wb):
    ws = wb.create_sheet("Parameters")
    ws["A1"] = "Parameters: Basel III standardised approach (simplified)"; ws["A1"].font = Font(name=ARIAL, bold=True, size=14)
    ws["A2"] = "Blue cells are inputs. Source: Basel Framework CRE20, https://www.bis.org/basel_framework/chapter/CRE/20.htm"
    ws["A2"].font = Font(name=ARIAL, italic=True, color="595959")
    hdr(ws, 4, 1, ["Rating", "Bucket"])
    for i, (r, b) in enumerate(RATING_BUCKET.items(), 5):
        put(ws, i, 1, r, BLUE); put(ws, i, 2, b, BLUE)
    hdr(ws, 4, 4, ["Bucket", "Sovereign", "Bank", "Corporate"])
    for i, b in enumerate(BUCKETS, 5):
        put(ws, i, 4, b, BLUE)
        for j, cls in enumerate(["Sovereign", "Bank", "Corporate"], 5):
            put(ws, i, j, RATED_RW[cls][BUCKETS.index(b)], BLUE, "0%")
    ws.cell(row=11, column=6).comment = Comment("Unrated banks taken as SCRA grade B (75%), CRE20.21.", "model")
    hdr(ws, 4, 9, ["Setting", "Value"])
    settings = [("Regulatory retail RW (CRE20.65)", RETAIL_RW, "0%"),
                ("Minimum capital ratio", CAPITAL_RATIO, "0.0%"),
                ("Defaulted: provision cover threshold", 0.20, "0%"),
                ("Defaulted RW, cover at or above threshold", 1.00, "0%"),
                ("Defaulted RW, cover below threshold (CRE20.88)", 1.50, "0%")]
    for i, (k, v, f) in enumerate(settings, 5):
        put(ws, i, 9, k); c = put(ws, i, 10, v, BLUE, f); c.fill = YELLOW
    hdr(ws, 4, 12, ["LTV up to", "Residential RW"])
    uppers = [lb for lb, _ in LTV_BANDS[1:]] + ["Above"]
    for i, (u, (_, w)) in enumerate(zip(uppers, LTV_BANDS), 5):
        put(ws, i, 12, u, BLUE, "0%"); put(ws, i, 13, w, BLUE, "0%")
    ws.cell(row=4, column=12).comment = Comment("CRE20.82, general residential real estate. A band applies when LTV is above the previous limit and up to this one.", "model")
    hdr(ws, 4, 15, ["Off-balance type", "CCF"])
    for i, (k, v) in enumerate(CCF.items(), 5):
        put(ws, i, 15, k, BLUE); put(ws, i, 16, v, BLUE, "0%")
    for col, w in {"A": 9, "B": 14, "D": 14, "I": 44, "L": 11, "M": 14, "O": 30}.items():
        ws.column_dimensions[col].width = w
    return ws


def exposures(wb, df: pd.DataFrame, title: str, prev_title: str | None, next_title: str | None):
    ws = wb.create_sheet(title)
    inputs = ["exposure_id", "counterparty", "asset_class", "rating", "drawn", "undrawn",
              "off_balance_type", "ltv", "defaulted", "specific_provision"]
    calc = ["CCF", "EAD", "Rating bucket", "Risk weight", "RWA", "Capital req."]
    extra = (["Prev EAD", "Prev RW", "Status", "New business RWA", "Volume effect", "RW effect"] if prev_title
             else ["Matured RWA"])
    hdr(ws, 1, 1, [c.replace("_", " ").title() for c in inputs] + calc + extra)
    n = len(df)
    for i, rec in enumerate(df[inputs].itertuples(index=False), 2):
        for j, v in enumerate(rec, 1):
            if isinstance(v, float) and pd.isna(v):
                v = None
            ws.cell(row=i, column=j, value=v).font = BLUE
        r = i
        P = "Parameters!"
        f = {
            11: f"=INDEX({P}$P$5:$P$9,MATCH(G{r},{P}$O$5:$O$9,0))",
            12: f"=E{r}+K{r}*F{r}",
            13: f"=INDEX({P}$B$5:$B$24,MATCH(D{r},{P}$A$5:$A$24,0))",
            14: (f"=IF(I{r}=\"Y\",IF(IFERROR(J{r}/E{r},0)<{P}$J$7,{P}$J$9,{P}$J$8),"
                 f"IF(C{r}=\"Retail\",{P}$J$5,IF(C{r}=\"Residential Mortgage\","
                 f"INDEX({P}$M$5:$M$10,COUNTIF({P}$L$5:$L$9,\"<\"&H{r})+1),"
                 f"INDEX({P}$E$5:$G$11,MATCH(M{r},{P}$D$5:$D$11,0),MATCH(C{r},{P}$E$4:$G$4,0)))))"),
            15: f"=L{r}*N{r}",
            16: f"=O{r}*{P}$J$6",
        }
        if prev_title:
            q = f"'{prev_title}'!"
            f.update({
                17: f"=IFERROR(INDEX({q}$L:$L,MATCH(A{r},{q}$A:$A,0)),\"\")",
                18: f"=IFERROR(INDEX({q}$N:$N,MATCH(A{r},{q}$A:$A,0)),\"\")",
                19: f"=IF(Q{r}=\"\",\"New\",\"Existing\")",
                20: f"=IF(S{r}=\"New\",O{r},0)",
                21: f"=IF(S{r}=\"New\",0,(L{r}-Q{r})*R{r})",
                22: f"=IF(S{r}=\"New\",0,L{r}*(N{r}-R{r}))",
            })
        if next_title:
            f[17] = f"=IF(ISNA(MATCH(A{r},'{next_title}'!$A:$A,0)),-O{r},0)"
        for c, formula in f.items():
            cell = ws.cell(row=r, column=c, value=formula)
            cell.font = GREEN if ("Parameters!" in formula or "'" in formula) and c in (11, 13, 17, 18) else BLACK
            cell.number_format = {11: "0%", 14: "0%", 18: "0%"}.get(c, AMT if c not in (13, 19) else "General")
    widths = [11, 11, 20, 8, 11, 11, 26, 7, 9, 11, 7, 12, 13, 9, 12, 11, 11, 9, 9, 13, 12, 11]
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{n + 1}"
    return n + 1


def build(q2: pd.DataFrame, q3: pd.DataFrame, path: Path, py_q2: pd.DataFrame, py_q3: pd.DataFrame) -> Path:
    wb = Workbook()
    notes = wb.active; notes.title = "Notes"
    lines = [
        ("Basel III RWA calculator: Q2 vs Q3 2025", Font(name=ARIAL, bold=True, size=16, color="1F3864")),
        ("Amounts in INR crore. Synthetic exposure book.", Font(name=ARIAL, italic=True)),
        ("", BLACK),
        ("How to use", BOLD),
        ("1. Inputs are in blue: exposure data on the Q2/Q3 sheets and every regulatory parameter on Parameters.", BLACK),
        ("2. Change a risk weight, CCF or the capital ratio on Parameters (yellow cells are the main levers) and the whole book reprices.", BLACK),
        ("3. RWA Summary shows EAD, RWA and capital by asset class; QoQ Walk explains why RWA moved.", BLACK),
        ("", BLACK),
        ("Method", BOLD),
        ("EAD = drawn + CCF x undrawn.  RWA = EAD x risk weight.  Capital requirement = RWA x minimum capital ratio.", BLACK),
        ("Walk: volume effect = (EAD now - EAD prev) x RW prev; risk-weight effect = EAD now x (RW now - RW prev).", BLACK),
        ("New business and maturities are booked at their own RWA, so the walk adds up exactly.", BLACK),
        ("", BLACK),
        ("Colour key: blue = input, black = formula, green = pulls from another sheet.", BLACK),
        ("Simplifications: no credit risk mitigation, no SA-CCR, no ECRA short-term ratings, unrated banks at SCRA grade B.", BLACK),
    ]
    for i, (t, f) in enumerate(lines, 1):
        notes.cell(row=i, column=1, value=t).font = f
    notes.column_dimensions["A"].width = 120

    parameters(wb)
    n2 = exposures(wb, q2, "Q2 Exposures", None, "Q3 Exposures")
    n3 = exposures(wb, q3, "Q3 Exposures", "Q2 Exposures", None)

    # Summary
    ws = wb.create_sheet("RWA Summary", 1)
    ws["A1"] = "RWA summary by asset class (INR crore)"; ws["A1"].font = Font(name=ARIAL, bold=True, size=14)
    cols = ["Asset class", "Q2 EAD", "Q2 RWA", "Q3 EAD", "Q3 RWA", "RWA change", "Change %",
            "Q3 RW density", "Q3 capital req.", "Python engine Q3 RWA", "Difference"]
    hdr(ws, 3, 1, cols)
    a2, a3 = "'Q2 Exposures'!", "'Q3 Exposures'!"
    for i, cls in enumerate(CLASSES + ["Total"], 4):
        put(ws, i, 1, cls, BOLD if cls == "Total" else BLACK)
        if cls == "Total":
            for c in range(2, 6):
                L = get_column_letter(c)
                put(ws, i, c, f"=SUM({L}4:{L}8)", BOLD, AMT)
            put(ws, i, 9, "=SUM(I4:I8)", BOLD, AMT)
        else:
            put(ws, i, 2, f"=SUMIFS({a2}$L:$L,{a2}$C:$C,$A{i})", GREEN, AMT)
            put(ws, i, 3, f"=SUMIFS({a2}$O:$O,{a2}$C:$C,$A{i})", GREEN, AMT)
            put(ws, i, 4, f"=SUMIFS({a3}$L:$L,{a3}$C:$C,$A{i})", GREEN, AMT)
            put(ws, i, 5, f"=SUMIFS({a3}$O:$O,{a3}$C:$C,$A{i})", GREEN, AMT)
            put(ws, i, 9, f"=SUMIFS({a3}$P:$P,{a3}$C:$C,$A{i})", GREEN, AMT)
        put(ws, i, 6, f"=E{i}-C{i}", BOLD if cls == "Total" else BLACK, AMT)
        put(ws, i, 7, f"=IFERROR(F{i}/C{i},0)", BLACK, PCT)
        put(ws, i, 8, f"=IFERROR(E{i}/D{i},0)", BLACK, PCT)
        put(ws, i, 10, round(float(py_q3.loc[cls, "rwa"]), 2), BLUE, AMT)
        put(ws, i, 11, f"=ROUND(E{i}-J{i},2)", BLACK, AMT)
    ws.cell(row=3, column=10).comment = Comment("Hardcoded from the Python engine (rwa/engine.py) on the same data, as an independent check.", "model")
    ws.cell(row=11, column=1, value="Difference column should read '-' on every row: Excel model and Python engine agree.").font = Font(name=ARIAL, italic=True)
    for j, w in enumerate([22, 14, 14, 14, 14, 13, 10, 12, 14, 16, 11], 1):
        ws.column_dimensions[get_column_letter(j)].width = w

    # Walk
    ws = wb.create_sheet("QoQ Walk", 2)
    ws["A1"] = "Why RWA moved: Q2 to Q3 2025 (INR crore)"; ws["A1"].font = Font(name=ARIAL, bold=True, size=14)
    cols = ["Asset class", "Q2 RWA", "New business", "Maturities", "Volume effect", "Risk-weight effect", "Q3 RWA (walk)", "Q3 RWA (actual)", "Check"]
    hdr(ws, 3, 1, cols)
    for i, cls in enumerate(CLASSES + ["Total"], 4):
        put(ws, i, 1, cls, BOLD if cls == "Total" else BLACK)
        if cls == "Total":
            for c in range(2, 9):
                L = get_column_letter(c); put(ws, i, c, f"=SUM({L}4:{L}8)", BOLD, AMT)
        else:
            put(ws, i, 2, f"=SUMIFS({a2}$O:$O,{a2}$C:$C,$A{i})", GREEN, AMT)
            put(ws, i, 3, f"=SUMIFS({a3}$T:$T,{a3}$C:$C,$A{i})", GREEN, AMT)
            put(ws, i, 4, f"=SUMIFS({a2}$Q:$Q,{a2}$C:$C,$A{i})", GREEN, AMT)
            put(ws, i, 5, f"=SUMIFS({a3}$U:$U,{a3}$C:$C,$A{i})", GREEN, AMT)
            put(ws, i, 6, f"=SUMIFS({a3}$V:$V,{a3}$C:$C,$A{i})", GREEN, AMT)
            put(ws, i, 7, f"=SUM(B{i}:F{i})", BLACK, AMT)
            put(ws, i, 8, f"=SUMIFS({a3}$O:$O,{a3}$C:$C,$A{i})", GREEN, AMT)
        put(ws, i, 9, f"=ROUND(G{i}-H{i},2)", BLACK, AMT)
    for j, w in enumerate([22, 14, 14, 13, 14, 16, 15, 15, 9], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ch = BarChart(); ch.type = "col"; ch.title = "RWA drivers by asset class"; ch.height, ch.width = 9, 20
    ch.add_data(Reference(ws, min_col=3, max_col=6, min_row=3, max_row=8), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=4, max_row=8))
    ws.add_chart(ch, "A12")

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path
