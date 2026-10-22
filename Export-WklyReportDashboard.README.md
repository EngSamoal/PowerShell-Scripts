# Export-WklyReportDashboard.ps1 — ITSM executive dashboard from `Wkly_Report`

Turns the `Wkly_Report` worksheet of the weekly ITSM workbook into **one self-contained HTML
file**: styling, buttons and every page picture are embedded, so the file works on its own when
sent by email or Teams, offline, in Edge or Chrome. From build r32 it also writes a **PDF** and a
**PowerPoint (.pptx)** copy next to it, with the same name (no assets folder is produced).
It is **read-only**: no worksheet, formula, chart, PivotTable or value is changed.

## 1. How it works

- The workbook is copied to a temporary folder and only that copy is opened, read-only, in a
  separate hidden Excel instance. Your own Excel windows are not touched, and the copy is closed
  without saving and deleted. Only the `Wkly_Report` tab is read.
- **Pages are detected on every run** from their large titles, so the dashboard follows the workbook
  when pages are added or removed.
- **Each page is rebuilt from what it contains** (default `-RenderMode Rebuilt`):
  - every text box, KPI card, cell, colour, font, accent bar and panel is read with its exact text
    and formatting (values exactly as Excel displays them - nothing is recalculated);
  - every object is placed at **Excel's exact position and size** (build r29 and later), and each page
    is cleaned up first:
    - every object belongs to one page only (no repeats across pages, no empty frames);
    - each page is cropped to its own content (no shared width, so no empty side margins);
    - helper cells outside the report panels (e.g. lookup values next to a card) are not shown;
    - empty and duplicate text boxes are skipped;
    - figures that sit as separate boxes on top of a sentence ("Team responds in ___ min") are joined
      into that sentence, so they can no longer overlap or drift;
    - text uses Excel's line spacing; wrapped text that still needs slightly more room than its box is
      shrunk by at most 15% instead of running over the next object;
  - **charts are Excel's own pictures of them** (default `-ChartRender Excel`, build r31 and later),
    taken with the temporary copy zoomed to 300% so they stay sharp. They match Excel exactly: data
    labels and callouts where Excel puts them, 3D, smoothed lines, rotated axis labels, colours. If that
    picture cannot be taken on a PC, Excel's normal chart export is used (the log says which).
    `-ChartRender Native` redraws charts from their data instead (sharp, values on hover, but not
    identical to Excel).
- `-RenderMode Picture` shows each page as a high-resolution picture of Excel's rendering instead
  (the previous behaviour).
- **PDF and PowerPoint** (build r32 and later): after the HTML is written, Microsoft Edge (or Chrome)
  opens that local file in a hidden window and takes one picture of each finished page at twice the
  screen resolution. Nothing is sent anywhere.
  - the PDF has one page per dashboard page, each exactly that page's size, so nothing is cut off;
  - the PowerPoint file has one 16:9 slide per page, with the whole page fitted in. It is written
    directly, so PowerPoint does not need to be installed;
  - both are written only when every page was captured; otherwise the log says why, and the HTML is
    still complete. Close an open copy of the PDF/PPTX before running, or it cannot be replaced.
- Each run writes a **page inventory** (`page-inventory <date>.json`, next to the log) listing what
  was read from every page - attach it when something differs from Excel.

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
| `-OutputDirectory` | `C:\Temp\ITSM_Dashboard` | Folder for the single `<workbook name>.html` |
| `-RenderMode` | `Rebuilt` | `Rebuilt` (pages rebuilt from content, native charts) or `Picture` (Excel's rendering as pictures) |
| `-ChartRender` | `Excel` | `Excel` (Excel's own chart pictures, identical to Excel) or `Native` (redrawn from data, values on hover) |
| `-PageMode` | `Auto` | `Auto`, `Titles`, `PageBreaks`, `Manual` |
| `-PageStartRows` | – | With `-PageMode Manual`: e.g. `1,40,72,105` |
| `-TitleMinFontSize` | `17` | Font size that counts as a page title |
| `-MinPageHeightPt` | `250` | Title candidates closer than this belong to the same page |
| `-DashboardTitle` | `ITSM Executive Dashboard` | Top-bar text |
| `-Classification` | `CONFIDENTIAL` | Marking on top bar and every page footer |
| `-SkipPdf` | off | Do not write the PDF |
| `-SkipPptx` | off | Do not write the PowerPoint file |
| `-OpenWhenDone` | off | Open the dashboard at the end |

Output: `<workbook name>.html`, `.pdf` and `.pptx`. Working files go to a private temporary folder that
is deleted at the end. The log is written to `%TEMP%\ITSM_Dashboard_logs\` (its path is printed at the end).

## 4. Using the dashboard

- **☰ (top left)**: opens the page list; pages always use the whole window width.
- **− / + / Fit** (or keys `-` `+`): zoom the pages; all pages share one scale and left edge.
- **Present** (or key `P`): full screen, one page at a time; ← → / PageUp/PageDown, Esc to exit.
- **Print / PDF**: one report page per A4 landscape sheet.

## 5. Validation checklist (first run, then spot-check weekly)

1. The console's first line shows the expected build.
2. The log says *Original workbook unchanged (SHA-256 identical before and after)*.
3. The log says *Excel page pictures: N of N pages captured (Excel PDF: N)* and
   *Validation: N page(s) in the file (expected N) … 0 external reference(s)*.
4. The output folder contains only the `.html`. Copy it alone to another PC (or send it by email/Teams),
   download it and open it in Edge or Chrome: pages, zoom, page list and Present all work.
5. The page list has one entry per page of Wkly_Report this week; compare a few KPI values, a chart
   and a table with the workbook.

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

Always attach the log (path printed at the end of the run) when reporting a problem.

## 7. Known limitations

- Rebuilt pages use the same text, colours and positions as Excel, but text is drawn by the browser,
  so line breaks can differ slightly; uncommon shape outlines are drawn as rectangles. Text that
  overflows its box in Excel itself can still touch the object below it (shrinking stops at 15%).
- The page inventory written by r29 also records text alignment and inner margins. If any page still
  differs from Excel, send the new `page-inventory` file and the log from the same run.
- With `-ChartRender Excel` (default) charts are pictures: they do not show values on hover.
- With `-ChartRender Native`, charts reproduce data, colours, scale, labels and legend, but Excel's
  special effects (3D, bevels, glow, gradients, label callouts) are drawn flat and simplified.
- Linked/camera pictures (e.g. a picture of a table on another tab) are shown as Excel's picture.
- The file includes all page content and is typically a few MB; Outlook/Teams previews may not run the
  page's script - download the file and open it in Edge or Chrome.
