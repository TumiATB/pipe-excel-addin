"""Recalculate the generated workbook and print the key figures and any formula errors."""
import json, sys, os, re, tempfile
import openpyxl, formulas

# Whole-column references are fine in Excel but slow in the Python engine; bound them for the check
COLS = re.compile(r"(?<![A-Z0-9$])(\$?)([A-Z]{1,2}):(\$?)([A-Z]{1,2})(?![0-9(])")
bound = lambda f: COLS.sub(lambda m: f"{m[1]}{m[2]}$1:{m[3]}{m[4]}$700", f)

book = json.load(open(sys.argv[1], encoding="utf-8"))
wb = openpyxl.Workbook()
wb.remove(wb.active)
for name in book["order"]:
    ws = wb.create_sheet(name)
    for addr, v in book["sheets"][name].items():
        ws[addr] = bound(v) if isinstance(v, str) and v.startswith("=") else v
tmp = os.path.join(tempfile.mkdtemp(), "pipe.xlsx")
wb.save(tmp)

xl = formulas.ExcelModel().loads(tmp).finish()
sol = xl.calculate()
vals = {}
for k, v in sol.items():
    # keys look like '[pipe.xlsx]A1 MATERIALITY'!J17
    try:
        sheet, cell = k.split("]", 1)[1].rsplit("!", 1)
        vals[(sheet.strip("'").upper(), cell.upper())] = v.value[0, 0] if hasattr(v, "value") else v
    except Exception:
        pass

def g(sheet, cell):
    return vals.get((sheet.upper(), cell))

checks = [
    ("TB", "C27", "TB total (nil)"),
    ("A1 Materiality", "I8", "Revenue"), ("A1 Materiality", "I9", "PBT"), ("A1 Materiality", "I10", "Equity"),
    ("A1 Materiality", "I11", "Assets"), ("A1 Materiality", "I12", "Liabilities"),
    ("A1 Materiality", "J17", "Materiality"), ("A1 Materiality", "J19", "PM"), ("A1 Materiality", "J21", "Trivial"),
    ("A1 Materiality", "B19", "PM label"),
    ("A3 Journals", "I8", "GL lines"), ("A3 Journals", "I9", "Manual lines"), ("A3 Journals", "I10", "Journals analysed"),
    ("A3 Journals", "N23", "JE-24117 score"), ("A3 Journals", "N24", "JE-23980 score"), ("A3 Journals", "N25", "JE-24203 score"),
    ("A3 Journals", "I55", "Met 1+"), ("A3 Journals", "I56", "Met 3+"), ("A3 Journals", "B59", "A3 conclusion"),
    ("A2 Risk assessment", "F22", "Sep vs avg"), ("A2 Risk assessment", "F24", "OM link"),
    ("A2 Risk assessment", "F29", "Revenue risk"), ("A2 Risk assessment", "F30", "Journals risk"), ("A2 Risk assessment", "F31", "Inventory risk"),
    ("A2 Risk assessment", "F36", "Payroll risk"), ("A2 Risk assessment", "G29", "Revenue rationale"), ("A2 Risk assessment", "B38", "A2 summary"),
    ("B1 Sampling", "D9", "Bank bal"), ("B1 Sampling", "E9", "Bank FS line"), ("B1 Sampling", "I24", "B1 PM"),
    ("B1 Sampling", "C29", "REV risk"), ("B1 Sampling", "E29", "REV pop"), ("B1 Sampling", "E30", "INV pop"), ("B1 Sampling", "E31", "PAY pop"),
    ("B1 Sampling", "E32", "PPE pop"), ("B1 Sampling", "E33", "JE pop"), ("B1 Sampling", "F29", "REV conf"), ("B1 Sampling", "H29", "REV interval"),
    ("B1 Sampling", "I29", "REV sample value"), ("B1 Sampling", "J29", "REV coverage"), ("B1 Sampling", "K29", "REV min"), ("B1 Sampling", "L29", "REV check"),
    ("B1 Sampling", "L31", "PAY check"), ("B1 Sampling", "L33", "JE check"),
    ("D3.1 Revenue", "I11", "D3.1 pop"), ("D3.1 Revenue", "I14", "D3.1 PM"), ("D3.1 Revenue", "L11", "D3.1 agreed"), ("D3.1 Revenue", "L12", "D3.1 exc"),
    ("D3.1 Revenue", "L14", "D3.1 known"), ("D3.1 Revenue", "L15", "D3.1 projected"), ("D3.1 Revenue", "L16", "D3.1 result"),
    ("E2.1 Inventory", "L14", "E2.1 known"), ("E2.1 Inventory", "L15", "E2.1 projected"), ("E2.1 Inventory", "L16", "E2.1 result"),
    ("F4.1 Payables", "L14", "F4.1 known"), ("F4.1 Payables", "L16", "F4.1 result"),
    ("C2.1 PPE", "L13", "C2.1 no support"), ("J1.1 Journals", "I11", "J1.1 pop"), ("J1.1 Journals", "L12", "J1.1 exc"),
    ("S1 Misstatements", "E13", "S1 known"), ("S1 Misstatements", "F13", "S1 projected"), ("S1 Misstatements", "F19", "Headroom"), ("S1 Misstatements", "B21", "S1 conclusion"),
    ("S1 Misstatements", "G8", "S1 REV status"),
]
for s, c, label in checks:
    print(f"{label:22} {s}!{c} = {g(s, c)}")

errs = [(k, v) for k, v in vals.items() if isinstance(v, str) and v.startswith("#")] + \
       [(k, str(v)) for k, v in vals.items() if type(v).__name__ in ("XlError",)]
print("errors:", len(errs))
for e in errs[:30]:
    print("  ", e)
