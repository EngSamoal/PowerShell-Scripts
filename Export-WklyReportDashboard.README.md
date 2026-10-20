# Export-WklyReportDashboard.ps1 — ITSM executive dashboard from `Wkly_Report`

Turns the `Wkly_Report` worksheet of the weekly ITSM workbook into an HTML dashboard.
It is **read-only**: no worksheet, formula, chart, PivotTable or value is changed.

## 1. How it works

- The workbook is copied to a temporary folder and only that copy is opened, read-only, in a
  separate hidden Excel instance. Your own Excel windows are not touched, and the copy is closed
  without saving and deleted.
- Each report page is shown as **Excel's own picture of that page**, so every value, chart, colour,
  text box and table looks exactly as Excel draws it.
- Formulas are calculated as when you open the file in Excel, so time-based values (`NOW()`,
  `TODAY()`, ages) are current. External links are not updated and PivotTables/queries are not
  refreshed — save the workbook after your normal refresh.
- **Pages are detected on every run**, so the dashboard follows the workbook when pages are added
  or removed. A new page starts at each large-font page title on the left of the sheet. You can
  also set the start rows yourself (see section 3).
- If Excel's picture of a page can't be captured, that page is shown re-drawn in HTML instead and
  the reason is written to `extraction-log.txt`.

## 2. Run it

Prerequisites: Windows, Microsoft Excel 2010 or later, and Windows PowerShell 5.1 (PowerShell ISE)
or PowerShell 7. No modules or internet access are needed.

```powershell
# one-time, if scripts are blocked on this PC (current user only)
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned

cd C:\Temp
.\Export-WklyReportDashboard.ps1 -OpenWhenDone
```

Defaults: the workbook is `C:\Temp\Bi_Wkly_Report-Daily Data.xlsx`; output goes to
`C:\Temp\ITSM_Dashboard\`. Open the dashboard there in Edge or Chrome; it has the workbook's name,
e.g. `Bi_Wkly_Report-Daily Data.html`.

The first line printed shows the build, e.g. `Export-WklyReportDashboard build 2026.10.20-r15`,
so you can confirm which version ran.

Weekly routine: update the workbook as usual → **save** → run the script → open `<workbook name>.html`.
The workbook can stay open in Excel while the script runs. Avoid copying/pasting in other programs
while it runs (it uses the clipboard briefly for the page pictures and restores it afterwards).

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
| `-SkipSheetPdf` | off | Skip `Wkly_Report (Excel export).pdf` |
| `-ExportDashboardPdf` | off | Also save `<workbook name> (Dashboard).pdf` using Edge (headless, local) |
| `-OpenWhenDone` | off | Open the dashboard at the end |

Output folder contents: `<workbook name>.html`, `assets\` (page pictures, charts, CSS/JS),
`Wkly_Report (Excel export).pdf`, `extraction-log.txt`, and optionally `<workbook name> (Dashboard).pdf`.

## 4. Using the dashboard

- **☰ (top left)**: opens the page list; pages always use the whole window width.
- **− / + / Fit** (or keys `-` `+`): zoom the pages; all pages share one scale and left edge.
- **Present** (or key `P`): full screen, one page at a time; ← → / PageUp/PageDown, Esc to exit.
- **Print / PDF**: one report page per A4 landscape sheet.

## 5. Validation checklist (first run, then spot-check weekly)

1. The console's first line shows the expected build.
2. The log says *Original workbook unchanged (SHA-256 identical before and after)*.
3. The log says *Excel page pictures: N of N pages captured*.
4. The page list has one entry per page of Wkly_Report this week.
5. Open `Wkly_Report (Excel export).pdf` or the workbook and compare a few KPI values, a chart and
   a table with the dashboard.

## 6. Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `Workbook not found` | Check `-WorkbookPath`; quote paths that contain spaces. |
| `Microsoft Excel is not available` | Excel desktop is not installed or not registered (COM). Run on a PC with Excel. |
| `Worksheet 'Wkly_Report' was not found` | The tab was renamed; the error lists the tabs in the file. |
| `Could not copy the workbook` | Another program locks the file exclusively; close it or wait for sync (OneDrive). |
| Wrong number of pages / a page split in two | Titles in this week's file are smaller or larger than usual. Try `-TitleMinFontSize 18` or `-MinPageHeightPt 350`, or use `-PageMode Manual -PageStartRows 1,40,...`. The log lists every title it found. |
| A page looks re-drawn instead of like Excel | Its picture could not be captured; the log line for that page gives the reason (often another program using the clipboard). Run again. |
| Values differ from what you saw in Excel | Time-based values change minute by minute; refreshed data (PivotTables, queries) appears only after you save the workbook. Save before running. |
| `Fingerprint could not be compared (file locked)` | Informational: the file was locked by another program; the script only reads it. |
| An `EXCEL.EXE` process remains | The script closes only the Excel it started (and stops it by its own process ID if it hangs). Your other Excel windows are never touched. |

Always attach `extraction-log.txt` (in the output folder) when reporting a problem.

## 7. Known limitations

- Page pictures and charts are images: they look identical to Excel but are not hover-interactive.
- `.xlsb`/`.xls` files: pictures inside the sheet may be shown as placeholders on pages that fall
  back to the re-drawn view.
