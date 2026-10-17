# Inclement Weather Analysis (React)

React migration of `INCLEMENT_WEATHER_ANALYSIS_TEMPLATE.xlsx`. The workbook's formulas run live in the
browser through a small Excel-compatible engine (`src/engine`). Before any edits, it reproduces all 4,900 of the
workbook's saved results exactly (`npm test`).

## Run

```bash
npm install
npm run dev      # development server
npm run build    # static build in dist/
npm test         # engine vs Excel cached values, category/variable propagation, export
```

## Pages

| Menu | Mirrors | Notes |
|---|---|---|
| See climate charts | Climate sheet charts | The 4 charts only (Mean max temp, Lowest temp, Mean monthly rainfall, Mean wind speeds), fed by `Climate!B4:M4`, `B16:M16`, `B29:M29`, `B55:M55`/`B62:M62`. |
| Modify variables | Rain+Other Factors, Weather Calculation E5:E6 & H7:S12, Climate thresholds | Direct/Consequential per category (Total = SUM), category names, add/remove categories, hot/cold direct impact, monthly allowances, Low/Med/High thresholds. |
| View weather impact | Weather Calculation | Only the Weather Analysis Impact Table is shown; the whole sheet is calculated. **Download weather calculation** exports the full sheet (formatting, merges, conditional formats, current values) to `.xlsx`. |
| Critical Site Summary | Critical Site Works Summary | Full layout; input cells are editable, calculated cells are locked. Activity rows 12–31 (hidden in the workbook) appear with **Show activity rows**. No planner name is pre-filled. |
| Project Calendar | Project Site Calendar | Full layout; input cells are editable, calculated cells are locked. |

Users never see spreadsheet mechanics: there are no formula bars, row/column headers or gridlines; formula cells cannot be
edited and entries starting with `=` are rejected. The template's spreadsheet instructions (Weather Calculation columns U–V,
the Project Site Calendar notes in columns W–AB and the "click the + sign" note on the summary) are removed.

## Starting state

The app starts with **no climate data**. The Climate charts, Weather impact and Critical Site Summary pages show an upload
prompt until a BoM csv is uploaded; the uploaded file becomes the BOM csv data sheet. Modify variables and Project Calendar
can be filled in beforehand. **Reset all** resets variables, categories and edits but keeps the uploaded csv.

## CSV upload

Upload the BoM *Monthly climate statistics* csv as downloaded (first line = "Monthly Climate Statistics for …").
The row containing `Statistic Element` is placed on **row 16** of the `BOM csv data` sheet, exactly as in the template,
so every `='BOM csv data'!B17`-style link reads the right statistic. Each statistic label is checked against the
template's layout and any mismatched rows are listed as a warning.

Current BoM exports are rewritten to the sheet's exact format on upload:
- `(°C)` / `°C` become `(Degrees C)` / `Degrees C`, and the unit is dropped from "Date of …" rows, as in the sheet;
- BoM's mislabelled second "Decile 5 (median) monthly rainfall" row (really Decile 9) is relabelled to row 47's name;
- dates such as `27-Jan-26` become real dates (two-digit years up to the current year are 20xx, older ones 19xx);
- trailing empty columns are dropped.

Use **View BOM csv data** to see the uploaded file as placed on the sheet. `samples/laverton_bom.csv` (current BoM format) is an example
for testing; it is not loaded automatically.

## Categories

Adding a category does what inserting it in Excel would: it adds Rain+Other Factors columns and a summary row, a
Rain Exposure block in Weather Calculation and a row in the Weather Analysis Impact Table. It also updates references
in other sheets, including the Critical Site Works Summary `HLOOKUP` row numbers for Heat/Cold/Wind/etc., so those
still point at the right rows. A new level can be used as a Rain Exposure Rating in row 11 of the Critical Site Works Summary.
The impact table shows each category's name in column B (`=A37`, `=A44`, …).

All inputs and the uploaded csv persist in the browser (localStorage).

## Updating the template

Replace `public/INCLEMENT_WEATHER_ANALYSIS_TEMPLATE.xlsx` and run `npm run extract` (Python 3 + openpyxl) to
regenerate `src/template/workbook.json`.
