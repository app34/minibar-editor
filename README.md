# Minibar Workbook Editor

Edit the existing **Minibar Consumption August 2026.xlsx** without building a new workbook.

The app opens the real file, changes only cell **values**, then writes the **same .xlsx** back. Colour fills, fonts, merged headers, column widths and formulas stay in the file.

## Why not “export a new Excel”?

This workbook is a template, not a data dump:

- Daily sheets (`1 August` … `31 August`) store villa × item consumption
- `Pick up` stores request / received quantities
- `Rooms` stores room par stock
- `Summary` already pulls daily totals with formulas such as `='1 August'!EJ3` and receiving with `='Pick up'!X4`

If you generate a fresh spreadsheet, those colours and links disappear. This editor never recreates sheets.

## How staff use it

1. Open `index.html` (browser or Replit / GitHub Pages)
2. Tap **Open XLSX** and choose the current month file
3. **Daily use** → pick the date sheet → search villa number → tap + / − on items
4. Optional: edit **Receiving** or **Room stock**
5. Tap **Save same file** — the download keeps the original file name

Empty cells stay empty (same as the current file). Only real quantities are written.

## Colour codes in the original file (kept)

| Colour | Typical use |
| --- | --- |
| Orange `#FFC000` | Food snacks (S.no 1–5) |
| Green `#92D050` | Soft drinks |
| Red `#FF0000` | Flagged / low-use items |
| Gold `#FFD966` / Yellow `#FFFF00` | Beers |
| Blue `#78BDFC` | Some wines |
| Dark green `#00B050` | Room stock grouping |

The UI swatches next to items follow the same groups. The Excel fills themselves are not rewritten.

## Sheet-name quirk (already handled)

Some daily tabs have two spaces: `2  August`, `3  August` … `7  August`. Others have one space. The app matches sheets by the day number, not a hard-coded name.

## For the developer

### Frontend (this folder)

- `index.html` + `styles.css` + `app.js`
- ExcelJS 4.4 from CDN
- Important rule in `setValueKeepStyle()`: assign `cell.value` only. Do not clone sheets or restyle cells.

Host as static files. No backend required.

### In-place Python (server / VPS)

If you want to overwrite the file on disk instead of downloading:

```python
from openpyxl import load_workbook

path = "Minibar Consumption August 2026.xlsx"
wb = load_workbook(path)          # keep_vba=False is fine for .xlsx
ws = wb["1 August"]               # or the exact tab name
ws.cell(row=3, column=4).value = 2   # villa 102, salted peanuts
wb.save(path)                     # same path = same file
```

`openpyxl.load_workbook()` + `save()` preserves fills, fonts, merges and formulas. Do **not** use pandas `to_excel()` for this file — pandas rebuilds a new book and drops colour.

Never open with `data_only=True` and then save. That replaces formulas with cached numbers.

### After save

Open the downloaded file in Excel / LibreOffice so Summary formulas recalculate. ExcelJS and openpyxl store the formulas; they do not recalculate them in the browser.

## Typical daily cell map

On a day sheet:

- Row 2 = villa numbers (101 … 606)
- Column B = article name
- Data grid starts at C3
- Column `EJ` = `=SUM(C3:EI3)` item total — leave this formula alone
- Summary day columns already point at those `EJ` totals
