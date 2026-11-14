import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule

d = pd.read_csv("out/table324_clean.csv", keep_default_na=False, na_values=[""])
F = "Arial"
H = Font(name=F, bold=True, color="FFFFFF"); HF = PatternFill("solid", fgColor="1F4E79")
H2F = PatternFill("solid", fgColor="2E75B6")
B = Font(name=F); BB = Font(name=F, bold=True); T = Font(name=F, bold=True, size=16, color="1F4E79")
S = Font(name=F, italic=True, size=9, color="595959"); BLUE = Font(name=F, color="0000FF")
GREEN = Font(name=F, color="008000"); YEL = PatternFill("solid", fgColor="FFF2CC"); FIX = PatternFill("solid", fgColor="FFE699")
thin = Side(style="thin", color="BFBFBF"); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
NUM = "#,##0.0"
COLS = [c for c in d.columns if "|" in c]
L = "EFGHIJKLM"                     # numeric columns in the table sheets
FIRST, LAST = 3, 3 + len(d) - 1      # data rows 3..34
row_of = {o: FIRST + i for i, o in enumerate(d["Offense"])}

wb = Workbook()
ws = wb.active; ws.title = "Summary"

def table_sheet(name, values_formula):
    w = wb.create_sheet(name)
    for j, h in enumerate(["Offense charged", "Level", "Part of", "Note"], 1):
        c = w.cell(row=2, column=j, value=h); c.font = H; c.fill = HF; c.border = BOX
        w.merge_cells(start_row=1, start_column=j, end_row=1, end_column=j) if False else None
        w.cell(row=1, column=j).fill = HF
    for k, grp in enumerate(["Total", "Male", "Female"]):
        c0 = 5 + 3 * k
        w.merge_cells(start_row=1, start_column=c0, end_row=1, end_column=c0 + 2)
        c = w.cell(row=1, column=c0, value=f"{grp} (thousands)"); c.font = H; c.fill = HF; c.alignment = Alignment(horizontal="center")
        for a, age in enumerate(["All ages", "Under 18", "18 and over"]):
            cc = w.cell(row=2, column=c0 + a, value=age); cc.font = H; cc.fill = H2F; cc.border = BOX; cc.alignment = Alignment(horizontal="center")
    for i, r in d.iterrows():
        rr = FIRST + i
        note = {"1": "1 Buying, receiving, possessing stolen property", "2": "2 Except forcible rape and prostitution"}.get(str(r["Footnote"]).replace(".0", ""), "")
        for j, v in enumerate([("   " if r["Level"] == "Component" else "") + r["Offense"], r["Level"], r["Part of"] if isinstance(r["Part of"], str) else "", note], 1):
            c = w.cell(row=rr, column=j, value=v); c.font = BB if r["Level"] in ("Grand total", "Subtotal") else B; c.border = BOX
        for k, col in enumerate(COLS):
            c = w.cell(row=rr, column=5 + k, value=values_formula(rr, k, r[col])); c.font = B; c.number_format = NUM; c.border = BOX
    w.column_dimensions["A"].width = 40; w.column_dimensions["B"].width = 12; w.column_dimensions["C"].width = 15; w.column_dimensions["D"].width = 22
    for k in range(9): w.column_dimensions[L[k]].width = 11
    w.freeze_panes = "B3"
    return w

# 1) as printed
def printed(rr, k, v): return None if pd.isna(v) else float(v)
wc = table_sheet("Clean Table", printed)
wc.cell(row=LAST + 2, column=1, value="Values exactly as printed in the PDF. '–' (represents zero) = 0; '(X)' (not applicable) = blank. In thousands.").font = S

# 2) corrected
corr = {  # (offense, column index) -> (formula parts, explanation)
 ("Liquor laws", 0): ("H{r}+K{r}", "Total = Male + Female"),
 ("All other offenses (except traffic)", 0): ("H{r}+K{r}", "Total = Male + Female"),
 ("Driving under the influence", 1): ("I{r}+L{r}", "Under 18 = Male under 18 + Female under 18"),
 ("Property crime", 6): ("L{r}+M{r}", "Female = Under 18 + 18 and over"),
}
def corrected(rr, k, v):
    return f"='Clean Table'!{L[k]}{rr}"
wk = table_sheet("Corrected Table", corrected)
fixes = []
for (off, k), (expr, how) in corr.items():
    rr = row_of[off]
    f = "=" + "+".join(f"'Clean Table'!{p}" for p in expr.format(r=rr).split("+"))
    c = wk.cell(row=rr, column=5 + k, value=f); c.fill = FIX
    printed_v = d.loc[d["Offense"] == off, COLS[k]].iloc[0]
    c.comment = Comment(f"Printed in the PDF: {printed_v:,.1f}. Recomputed as {how} (same row).", "Stuard")
    fixes.append((off, COLS[k].replace(" | ", " - "), f"'Clean Table'!{L[k]}{rr}", f"'Corrected Table'!{L[k]}{rr}", how))
# grand-total helper rows (before / after)
for w, nm in ((wc, "Clean Table"), (wk, "Corrected Table")):
    hr = LAST + 4
    w.cell(row=hr, column=1, value="Check: Grand total - (subtotals + other offenses)").font = BB
    for k in range(9):
        col = L[k]
        f = (f"={col}{FIRST}-SUMIFS({col}{FIRST}:{col}{LAST},$B${FIRST}:$B${LAST},\"Subtotal\")"
             f"-SUMIFS({col}{FIRST}:{col}{LAST},$B${FIRST}:$B${LAST},\"Offense\")")
        c = w.cell(row=hr, column=5 + k, value=f); c.font = B; c.number_format = NUM
wk.cell(row=LAST + 2, column=1, value="Same as Clean Table, except the 4 highlighted cells (hover for the printed value and how it was recomputed).").font = S

# 3) checks
wch = wb.create_sheet("Checks", 1)
wch["A1"] = "Row checks (difference in thousands; tolerance in Summary!B11)"; wch["A1"].font = BB
heads = ["Offense charged", "Level", "Total - (Male + Female)", "Total - (Under 18 + 18+)", "Male - (Under 18 + 18+)",
         "Female - (Under 18 + 18+)", "Under 18: Total - (M + F)", "18+: Total - (M + F)", "Status", "Largest gap"]
for j, h in enumerate(heads, 1):
    c = wch.cell(row=2, column=j, value=h); c.font = H; c.fill = HF; c.border = BOX; c.alignment = Alignment(wrap_text=True, vertical="center")
CT = "'Clean Table'!"
for i in range(len(d)):
    r = FIRST + i
    wch.cell(row=r, column=1, value=f"={CT}A{r}").font = B
    wch.cell(row=r, column=2, value=f"={CT}B{r}").font = B
    fs = [f"={CT}E{r}-({CT}H{r}+{CT}K{r})", f"={CT}E{r}-({CT}F{r}+{CT}G{r})", f"={CT}H{r}-({CT}I{r}+{CT}J{r})",
          f"={CT}K{r}-({CT}L{r}+{CT}M{r})", f"={CT}F{r}-({CT}I{r}+{CT}L{r})", f"={CT}G{r}-({CT}J{r}+{CT}M{r})"]
    for j, f in enumerate(fs, 3):
        c = wch.cell(row=r, column=j, value=f); c.font = B; c.number_format = "0.0;-0.0;0.0"; c.border = BOX
    wch.cell(row=r, column=10, value=f"=MAX(ABS(C{r}),ABS(D{r}),ABS(E{r}),ABS(F{r}),ABS(G{r}),ABS(H{r}))").number_format = "0.0"
    wch.cell(row=r, column=9, value=f'=IF(J{r}<=Summary!$B$11,"OK",IF(J{r}<=Summary!$B$12,"SMALL GAP","WRONG VALUE"))').font = BB
    for j in (9, 10): wch.cell(row=r, column=j).border = BOX
red = PatternFill("solid", fgColor="F8CBAD"); green = PatternFill("solid", fgColor="E2EFDA")
wch.conditional_formatting.add(f"A{FIRST}:J{LAST}", FormulaRule(formula=[f'$I{FIRST}="WRONG VALUE"'], fill=red))
wch.conditional_formatting.add(f"A{FIRST}:J{LAST}", FormulaRule(formula=[f'$I{FIRST}="SMALL GAP"'], fill=YEL))
wch.conditional_formatting.add(f"I{FIRST}:I{LAST}", FormulaRule(formula=[f'$I{FIRST}="OK"'], fill=green))
wch.column_dimensions["A"].width = 40; wch.column_dimensions["B"].width = 12
for c in "CDEFGH": wch.column_dimensions[c].width = 14
wch.column_dimensions["I"].width = 14; wch.column_dimensions["J"].width = 10; wch.freeze_panes = "C3"
wch.cell(row=LAST + 2, column=1, value="SMALL GAP = a dash ('–', shown as zero) hides a small value, or rounding. WRONG VALUE = a printed number that does not match its own breakdown.").font = S

# 4) long format
long = []
dd = d.copy()
for (off, k), _ in corr.items():
    dd.loc[dd["Offense"] == off, COLS[k]] = None
for i, r in dd.iterrows():
    for k, col in enumerate(COLS):
        sex, age = col.split(" | ")
        long.append([r["Offense"], r["Level"], r["Part of"] if isinstance(r["Part of"], str) else "", sex, age, f"='Corrected Table'!{L[k]}{FIRST + i}"])
wl = wb.create_sheet("Long Format")
for j, h in enumerate(["Offense charged", "Level", "Part of", "Sex", "Age group", "Arrests (thousands)"], 1):
    c = wl.cell(row=1, column=j, value=h); c.font = H; c.fill = HF
for i, row in enumerate(long, 2):
    for j, v in enumerate(row, 1):
        c = wl.cell(row=i, column=j, value=v); c.font = B
        if j == 6: c.number_format = NUM
for col, w_ in zip("ABCDEF", [40, 12, 15, 9, 12, 18]): wl.column_dimensions[col].width = w_
wl.freeze_panes = "A2"; wl.auto_filter.ref = f"A1:F{len(long) + 1}"

# 5) source & method
wm = wb.create_sheet("Source & Method")
lines = [("Source & method", T),
 ("Source PDF", BB), ("U.S. Census Bureau, Statistical Abstract of the United States: 2012, Table 324 'Arrests by Sex and Age: 2009' (public domain, U.S. Government work). Original source: FBI, Uniform Crime Reports. Copy used: tabula-java test files (github.com/tabulapdf/tabula-java).", B),
 ("Why this PDF", BB), ("A typical messy table: two header levels, dot leaders, wrapped labels, footnote markers, '–' and '(X)' symbols, and a second table on the same page.", B),
 ("Steps", BB),
 ("1. Read every character with its position (pdfplumber), group them into lines, and assign numbers to the 9 columns by their right edge.", B),
 ("2. Join wrapped labels, strip dot leaders, move footnote markers to a Note column, mark subtotals and their components.", B),
 ("3. Convert '–' to 0 and '(X)' to blank, keep values in thousands as printed.", B),
 ("4. Run 6 arithmetic checks per row plus subtotal and grand-total checks, all as live formulas.", B),
 ("5. Recompute only the cells that fail the checks by a wide margin, keep the printed value in a comment, and confirm the grand total now adds up.", B),
 ("Tools", BB), ("Python (pdfplumber, pandas), Excel formulas (SUMIFS, IF, ABS, MAX), conditional formatting.", B)]
for i, (t, f) in enumerate(lines, 1):
    c = wm.cell(row=i, column=1, value=t); c.font = f; c.alignment = Alignment(wrap_text=True, vertical="top")
wm.column_dimensions["A"].width = 120

# Summary
ws["A1"] = "PDF table to Excel, with automatic checks"; ws["A1"].font = T
ws["A2"] = "U.S. Census Bureau, Statistical Abstract 2012, Table 324 (arrests by sex and age, 2009, in thousands)"; ws["A2"].font = S
items = [("Rows extracted", f"=COUNTA('Clean Table'!A{FIRST}:A{LAST})", "#,##0"),
         ("Numbers extracted", f"=COUNT('Clean Table'!E{FIRST}:M{LAST})", "#,##0"),
         ("Rows passing all 6 checks", f'=COUNTIF(Checks!I{FIRST}:I{LAST},"OK")', "#,##0"),
         ("Rows with a small gap (dash or rounding)", f'=COUNTIF(Checks!I{FIRST}:I{LAST},"SMALL GAP")', "#,##0"),
         ("Rows with a wrong printed value", f'=COUNTIF(Checks!I{FIRST}:I{LAST},"WRONG VALUE")', "#,##0")]
ws["A4"], ws["B4"] = "Extraction", "Count"
for c in "AB": ws[f"{c}4"].font = H; ws[f"{c}4"].fill = HF
for i, (a, f, nf) in enumerate(items, 5):
    ws.cell(row=i, column=1, value=a).font = B; c = ws.cell(row=i, column=2, value=f); c.font = B; c.number_format = nf
ws["A11"] = "Rounding tolerance (thousands) - edit"; ws["A11"].font = B; ws["B11"] = 0.15; ws["B11"].font = BLUE; ws["B11"].fill = YEL; ws["B11"].number_format = "0.00"
ws["A12"] = "Gap that counts as a wrong value (thousands) - edit"; ws["A12"].font = B; ws["B12"] = 1; ws["B12"].font = BLUE; ws["B12"].fill = YEL; ws["B12"].number_format = "0.00"
ws["A14"] = "Cells corrected"; ws["A14"].font = H; ws["A14"].fill = HF
for j, h in enumerate(["Column", "Printed", "Corrected", "How"], 2):
    c = ws.cell(row=14, column=j, value=h); c.font = H; c.fill = HF
for i, (off, col, pref, cref, how) in enumerate(fixes, 15):
    ws.cell(row=i, column=1, value=off).font = B; ws.cell(row=i, column=2, value=col).font = B
    c = ws.cell(row=i, column=3, value=f"={pref}"); c.font = GREEN; c.number_format = NUM
    c = ws.cell(row=i, column=4, value=f"={cref}"); c.font = GREEN; c.number_format = NUM; c.fill = FIX
    ws.cell(row=i, column=5, value=how).font = B
gr = 15 + len(fixes) + 1
hr = LAST + 4
ws.cell(row=gr, column=1, value="Grand total check - largest gap BEFORE corrections").font = BB
ws.cell(row=gr, column=2, value="=MAX(" + ",".join(f"ABS('Clean Table'!{L[k]}{hr})" for k in range(9)) + ")").number_format = NUM
ws.cell(row=gr + 1, column=1, value="Grand total check - largest gap AFTER corrections").font = BB
ws.cell(row=gr + 1, column=2, value="=MAX(" + ",".join(f"ABS('Corrected Table'!{L[k]}{hr})" for k in range(9)) + ")").number_format = NUM
ws.cell(row=gr + 3, column=1, value="Tabs: Checks (row by row), Clean Table (as printed), Corrected Table (4 fixes highlighted), Long Format (ready for pivot tables), Source & Method.").font = S
ws.column_dimensions["A"].width = 46; ws.column_dimensions["B"].width = 22; ws.column_dimensions["C"].width = 12
ws.column_dimensions["D"].width = 12; ws.column_dimensions["E"].width = 42
for w in (ws, wm): w.sheet_view.showGridLines = False
wb.save("PDF_Table_to_Excel_with_Checks.xlsx"); print("saved", len(fixes))
