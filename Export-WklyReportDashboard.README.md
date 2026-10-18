# Export-WklyReportDashboard.ps1 — ITSM executive dashboard from `Wkly_Report`

Turns the `Wkly_Report` worksheet of the weekly ITSM workbook into an interactive HTML
dashboard. It is **read-only**: no worksheet, formula, chart, PivotTable or value is changed.

## 1. How it works (hybrid approach)

| Part of the report | How it reaches the dashboard | Exact? |
|---|---|---|
| Numbers, %, dates, durations, `-`, `#DIV/0!` | Read as Excel *displays* them (cell `.Text`, text-box text) | Yes — never recalculated |
| Charts / PivotCharts | Exported as PNG by Excel itself (`Chart.Export`) | Yes — Excel's own image |
| KPI cards, bands, text boxes, colours | Re-drawn in HTML at the same position, colour, font, corner radius and shadow | Layout and colours match; fonts render as on your PC |
| Cell tables (e.g. assignment groups) | Re-drawn with the cells' displayed fill, font and borders | Values exact |
| Pictures / logos | Copied out of the workbook file (no clipboard) | Yes |
| Whole page, for checking | **Excel view** button = Excel's own PDF of that page | Yes |
| Slicers, form controls, SmartArt, data bars, icon sets, sparklines | **Not** re-drawn: a labelled placeholder is shown and the item is listed in **Data check** | Use Excel view |

**Pages are detected on every run**, so the dashboard follows the workbook when pages are
added or removed. A new page starts at each large-font page title (≥ 20 pt, left side of the
sheet). If none are found, manual page breaks are used. You can also set the start rows
yourself (see section 3).

## 2. Run it

Prerequisites: Windows, Microsoft Excel 2010 or later, and Windows PowerShell 5.1 or
PowerShell 7. No modules or internet access are needed.

```powershell
# one-time, if scripts are blocked on this PC (current user only)
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned

cd C:\Temp
.\Export-WklyReportDashboard.ps1 -OpenWhenDone
```

Defaults: the workbook is `C:\Temp\Bi_Wkly_Report-Daily Data.xlsx`; output goes to
`C:\Temp\ITSM_Dashboard\`. Open the dashboard there in Edge or Chrome; it has the workbook's name, e.g. `Bi_Wkly_Report-Daily Data.html`.

The first line printed shows the build, e.g. `Export-WklyReportDashboard build 2026.10.18-r1`,
so you can confirm which version ran.

Weekly routine: update the workbook as usual → save → run the script → open `<workbook name>.html`.
The workbook can stay open in Excel while the script runs.

## 3. Options

| Parameter | Default | Purpose |
|---|---|---|
| `-WorkbookPath` | `C:\Temp\Bi_Wkly_Report-Daily Data.xlsx` | Workbook to read |
| `-OutputDirectory` | `C:\Temp\ITSM_Dashboard` | Where `<workbook name>.html` + `assets\` go (only its own earlier output is replaced) |
| `-PageMode` | `Auto` | `Auto`, `Titles`, `PageBreaks`, `Manual` |
| `-PageStartRows` | – | With `-PageMode Manual`: e.g. `1,40,72,105` |
| `-TitleMinFontSize` | `20` | Font size that counts as a page title |
| `-MinPageHeightPt` | `250` | Title candidates closer than this belong to the same page |
| `-DashboardTitle` | `ITSM Executive Dashboard` | Top-bar text |
| `-Classification` | `CONFIDENTIAL` | Marking on top bar and every page footer |
| `-SkipExcelView` | off | Skip the per-page Excel PDFs (faster) |
| `-SkipSheetPdf` | off | Skip `Wkly_Report (Excel export).pdf` |
| `-ExportDashboardPdf` | off | Also save `<workbook name> (Dashboard).pdf` using Edge (headless, local) |
| `-OpenWhenDone` | off | Open the dashboard at the end |

Output folder contents: `<workbook name>.html`, `assets\` (charts, pictures, CSS/JS, `excel-view\page-NN.pdf`),
`Wkly_Report (Excel export).pdf`, `extraction-log.txt`, and optionally `<workbook name> (Dashboard).pdf`.

## 4. Using the dashboard

- **☰ (top left)**: opens the page list as a pop-out; pages always use the whole window width.
- **− / + / Fit** (or keys `-` `+`): zoom the pages; all pages share one scale and left edge so they line up.
- **Sidebar**: one entry per detected page. ⚠ shows how many values on that page need checking.
- **Data check**: source file, SHA-256 before/after (proves the file was not changed), the
  page-detection method, and every Excel error, empty formula result, `####` cell and object
  that was not reproduced, each with its cell address or object name.
- **Inspect: on/off**: click any text box to see what Excel reported for it (fonts, colours, spacing,
  cell link). Useful when something looks different from Excel — a photo of the panel is enough to diagnose it.
- **Markers: on/off**: dashed outlines on Excel errors and empty formula results. The value
  itself is never changed.
- **Excel view** (per page): Excel's own rendering, for side-by-side comparison.
- **Present** (or key `P`): full screen, one page at a time; ← → / PageUp/PageDown, Esc to exit.
- **Print / PDF**: one report page per A4 landscape sheet.
- Hover any cell value to see its address, e.g. `Wkly_Report!E46`.

## 5. Validation checklist (first run, then spot-check weekly)

1. The console's first line shows the expected build.
2. **Data check** → *Original file SHA-256 before / after: unchanged*.
3. Sidebar page count equals the number of pages in Wkly_Report this week, and each page's
   `rows a–b` range starts at that page's header.
4. For each page, click **Excel view** and compare it with the dashboard view. Check every KPI
   card value, the period text, each chart, and table values.
5. Every `#DIV/0!` you see in Excel appears in the dashboard and in Data check — not as 0.
6. Every empty KPI in Excel is empty in the dashboard (dashed marker), not 0.
7. Anything listed under *Needs attention* as "not reproduced" is visible in Excel view.
8. Open `Wkly_Report (Excel export).pdf` and confirm it matches what you see in Excel.

## 6. Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `Workbook not found` | Check `-WorkbookPath`; quote paths that contain spaces. |
| `Microsoft Excel is not available` | Excel desktop is not installed or not registered (COM). Run on a PC with Excel. |
| `Worksheet 'Wkly_Report' was not found` | The tab was renamed; the error lists the tabs in the file. |
| `Could not copy the workbook` | Another program locks the file exclusively; close it or wait for sync (OneDrive). |
| Wrong number of pages / a page split in two | Titles in this week's file are smaller or larger than usual. Try `-TitleMinFontSize 18` or `-MinPageHeightPt 350`, or use `-PageMode Manual -PageStartRows 1,40,...`. The method used is shown in Data check. |
| A chart shows a red "could not be exported" box | The script tries three export methods (direct, after activating the chart, and with the private Excel window moved off-screen). If all fail, Excel's own error text is shown in Data check and the log; send `extraction-log.txt`. |
| "Exported chart image is very small and may be blank" | Compare that chart with Excel view; send the log if it really is blank. |
| A cell shows `####` | The column is too narrow in Excel itself; widen it in the workbook. |
| A shape looks rectangular but is a different shape in Excel | Uncommon outlines are drawn as rectangles (text and colours kept); listed under *Info*. |
| Excel view is empty or the button is disabled | Excel's PDF export needs a printer driver (the built-in *Microsoft Print to PDF* is enough). Check Data check for the error. |
| Values differ from what you saw in Excel | The script shows values as last **saved**. Save the workbook (after any refresh) before running. |
| An `EXCEL.EXE` process remains | The script closes only the Excel it started (and stops it by its own process ID if it hangs). Your other Excel windows are never touched. |

Always attach `extraction-log.txt` (in the output folder) when reporting a problem.

## 7. Known limitations

- Fonts are rendered by the browser, so text wrapping can differ slightly from Excel. Use
  Excel view when exact pixel layout matters.
- Charts are images: they look identical but are not hover-interactive.
- Calculation is set to Manual before the copy is opened, so formulas (including `TODAY()`)
  show their last saved values. PivotTables or queries set to "refresh on open" may still try
  to refresh in the hidden copy; save the workbook after your normal refresh so the saved
  values are current.
- Picture cropping set in Excel is not applied (listed under *Info*). `.xlsb`/`.xls` files:
  pictures are shown as placeholders (charts and values are unaffected).
