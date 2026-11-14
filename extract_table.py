"""Extract Table 324 (U.S. Census Bureau, Statistical Abstract 2012) from PDF to a clean table.
Usage: python extract_table.py 12s0324.pdf out_dir
"""
import re, sys, json
from pathlib import Path
import pdfplumber, pandas as pd

pdf_path = Path(sys.argv[1]); out = Path(sys.argv[2]); out.mkdir(exist_ok=True)
page = pdfplumber.open(pdf_path).pages[0]

TOP, BOTTOM, LABEL_MAX_X = 88, 318, 137          # Table 324 area on the page (PDF points)
EDGES = [163, 190, 218, 245, 272, 300, 327, 354, 382]  # right edge of each numeric column
COLS = [(s, a) for s in ("Total", "Male", "Female") for a in ("All ages", "Under 18", "18 and over")]

chars = [c for c in page.chars if TOP < c["top"] < BOTTOM]
lines = []
for c in sorted(chars, key=lambda c: (round(c["top"], 0), c["x0"])):
    if lines and abs(lines[-1][0] - c["top"]) < 1.2:
        lines[-1][1].append(c)
    else:
        lines.append([c["top"], [c]])

def col_of(c):
    for k, e in enumerate(EDGES):
        if e - 26 < c["x1"] <= e + 1.5:
            return k
    return None

records, pending, issues = [], "", []
pend_indent = False
for top, cs in lines:
    label_cs = [c for c in cs if c["x1"] <= LABEL_MAX_X and c["size"] >= 5]
    foot_cs = [c for c in cs if c["x1"] <= LABEL_MAX_X and c["size"] < 5]
    num_cs = [c for c in cs if c["x1"] > LABEL_MAX_X and c["text"].strip()]
    label = ""
    for c in sorted(label_cs, key=lambda c: c["x0"]):
        if label and c["x0"] - prev_x1 > 1.0: label += " "
        label += c["text"]; prev_x1 = c["x1"]
    indent = min((c["x0"] for c in label_cs if c["text"].strip() not in ("", ".")), default=99) > 37.5
    label = re.sub(r"[\s.]+$", "", re.sub(r"(\s*\.\s*){2,}", " ", label)).strip()
    if label.endswith("etc"): label += "."
    vals = [""] * 9
    for c in sorted(num_cs, key=lambda c: c["x0"]):
        k = col_of(c)
        if k is None: issues.append(f"char '{c['text']}' at x={c['x1']:.0f} outside columns"); continue
        vals[k] += c["text"]
    vals = [v.replace(" ", "") for v in vals]
    if not any(vals):
        if label:
            if not pending: pend_indent = indent
            pending = (pending + " " + label).strip()
        continue
    full = (pending + " " + label).strip() if pending else label
    ind = pend_indent if pending else indent
    pending = ""
    foot = "".join(c["text"] for c in foot_cs)
    records.append({"Offense": full, "Indented": ind, "Footnote": foot, **{f"{s} | {a}": v for (s, a), v in zip(COLS, vals)}})

df = pd.DataFrame(records)
def parse(v):
    if v in ("–", "-"): return 0.0            # "–" = represents zero
    if v == "(X)": return None                # not applicable
    return float(v.replace(",", "")) if v else None
raw = df.copy()
for c in df.columns[3:]:
    df[c] = df[c].map(parse)

# hierarchy
df["Level"] = "Offense"
df.loc[df["Offense"].str.lower().eq("total"), "Level"] = "Grand total"
df.loc[df["Offense"].isin(["Violent crime", "Property crime"]), "Level"] = "Subtotal"
df.loc[df["Indented"], "Level"] = "Component"
grp, parent = None, []
for i, r in df.iterrows():
    if r["Level"] == "Subtotal": grp = r["Offense"]
    elif r["Level"] != "Component": grp = None
    parent.append(grp if r["Level"] == "Component" else "")
df["Part of"] = parent
df = df.drop(columns=["Indented"])
df.to_csv(out / "table324_clean.csv", index=False)
raw.to_csv(out / "table324_raw_text.csv", index=False)
json.dump({"rows": len(df), "issues": issues}, open(out / "extract_log.json", "w"), indent=2)
pd.set_option("display.width", 250); pd.set_option("display.max_columns", 20)
print(df[["Offense", "Level", "Part of", "Footnote"] + list(df.columns[2:11])].to_string())
print(issues[:10])
