# Pipe for Excel (demo)

An Office.js task-pane add-in that writes audit working papers into the open workbook. It is the Excel companion to the Pipe web mock-up. All AI output is simulated, and the client data is fictional.

## Install (Excel on the web)

1. Download `manifest.xml` from https://tumiatb.github.io/pipe-excel-addin/
2. In Excel on the web: **Home › Add-ins › More add-ins › My add-ins › Upload my add-in**.
3. Click **Open Pipe** on the Home tab.

## What it does

- **Setup** writes an `Index` sheet holding the engagement settings. Every working-paper header links to it.
- **Audit tab** runs the three workflows in order: GL analysis (A1, A3, A2), Sampling (B1), then Detail testing (D3.1–J1.1, S1). Each one has a plan you approve and a sign-off. Sign-off writes the reviewer's initials and the date into each paper.
- **Processes tab** runs any one paper on its own. Pipe checks the workbook for the inputs it needs (for example, A1 for materiality). It asks for anything missing and writes those values into shaded cells.
- Calculations are live formulas linked to the `TB` and `GL` sheets and to the other papers. Judgements go in shaded input cells.

## Files

| Path | Purpose |
| --- | --- |
| `manifest.xml` | Add-in manifest (Pages URLs) |
| `taskpane.html`, `taskpane.css` | Pane |
| `js/data.js` | Demo engagement data and sample selection |
| `js/wp.js` | Working-paper builders (cells, formulas, formats) |
| `js/xl.js` | Applies builder output to Excel, or to a preview workbook outside Office |
| `js/app.js` | Pane UI, workflow engine, chat |
| `test/` | Builds the workbook in Node and recalculates it with the Python `formulas` package |

## Develop

Open `taskpane.html` from any static server (for example `python -m http.server`). Outside Excel it runs in preview mode with a simulated workbook beside the pane.

Formula check:

```bash
node test/build-book.js book.json
python test/check_formulas.py book.json
```
