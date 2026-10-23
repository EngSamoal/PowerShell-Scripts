# Export-WklyReportDashboardB.ps1: Design B dashboard from `Wkly_Report`

This is a separate script from `Export-WklyReportDashboard.ps1`. It reads the same `Wkly_Report` worksheet in
the same read-only way, but presents it in **Design B**: charcoal, green and gold, a full-screen cover, one
panel per page with a coloured header bar, and large text for presenting.

The output is **one self-contained HTML file**, named after the workbook, for example
`Bi_Wkly_Report-Daily Data.html`. It works offline and can be sent by email or Teams. From build B4 a
**PDF** and a **PowerPoint (.pptx)** copy with the same name are written next to it:

- each page is shown as in Present mode, 1920 px wide and at least 1080 px high, so most pages fill a
  16:9 slide exactly; a longer page is kept whole (taller PDF page, fitted whole on its slide);
- the pictures are taken by Microsoft Edge (or Chrome) in a hidden window from the local HTML file, with no
  network access; PowerPoint does not need to be installed;
- both files are written only when every page was captured; otherwise the log says why, and the HTML is
  still complete. Close an open copy of the PDF/PPTX before running, or it cannot be replaced.
- `-SkipPdf` / `-SkipPptx` turn them off.

## Safety (same as the original script)

- The original workbook is never opened. A temporary copy is opened read-only in a separate hidden Excel
  instance, closed without saving, and deleted.
- Only the `Wkly_Report` tab is read. Nothing in the workbook is changed.
- Your own Excel windows are not touched.
- The log confirms the original file is unchanged (SHA-256 before and after).
- Values are shown exactly as Excel displays them. Nothing is recalculated or invented, and Excel errors
  such as `#DIV/0!` are shown as-is with a red outline.

## Run it

```powershell
cd C:\Temp
.\Export-WklyReportDashboardB.ps1 -OpenWhenDone
```

The first line printed shows the build, for example `Export-WklyReportDashboardB (Design B) build 2026.10.23-B5`.

| Parameter | Default |
|---|---|
| `-WorkbookPath` | `C:\Temp\Bi_Wkly_Report-Daily Data.xlsx` |
| `-OutputDirectory` | `C:\Temp\ITSM_Dashboard_DesignB` (a different folder from the original script, so the two never overwrite each other) |
| `-DashboardTitle` | `IT Service Management` |
| `-Classification` | `CONFIDENTIAL` |
| `-PageMode`, `-PageStartRows`, `-TitleMinFontSize`, `-MinPageHeightPt` | Same as the original script |

The log is written to `%TEMP%\ITSM_Dashboard_logs\designB-log <date>.txt`. Next to it is
`page-inventory <date>.json`, which lists everything read from each page. For every page, the log
also lists what it became, for example `page 3 'Incident Summary': 2 chart, 5 kpi`. Attach both
files if something looks wrong.

## How each page is arranged

Pages are detected on every run, so the dashboard follows the workbook when pages are added or removed.

- **Page 1 becomes the cover.** It shows:
  - the title, very large;
  - the text above the title, as the gold eyebrow line;
  - the text below the title, as the tagline;
  - every `Label: value` line (Reporting Period, Report Date, Prepared By, ...);
  - any KPI cards on the page, as large tiles on the right.
- **Every other page gets a coloured header bar.** The bar shows the page number and title, plus the
  text from Excel's header band. Bar colours repeat in the order charcoal, green, gold.
- **The page content becomes cards**, kept in Excel's arrangement: side-by-side cards stay side by side
  in Excel's width proportions, and stacked cards stay stacked.

  | Content in Excel | Shown as |
  |---|---|
  | A card whose largest text is a figure | KPI tile: label, figure, details |
  | A chart | Chart redrawn from its own data, using Excel's axis scale and number formats |
  | A block of cells | Table |
  | Other text | Text card |

  Short text directly above a chart or table becomes its title.
- **Colours follow the dashboard palette.** Excel's chart colours are mapped by hue:

  | Excel colour | Dashboard colour |
  |---|---|
  | Red | Red (it marks problems) |
  | Green | Green |
  | Yellow or orange | Gold |
  | Grey | Grey |
  | Any other colour | The next free colour from the dashboard palette |

  Status words (Met, Breached, Open, Resolved, High, Med, Low, Monitoring, ...) get a coloured pill. The
  word itself is unchanged.
- **Blank chart values stay as gaps**, never 0.
- **A chart type that is not redrawn is shown as Excel's own picture of that chart.**

## Using the dashboard

- Dark mode is the default (build B5); **Light / Dark** switches. In light mode each page panel, its KPI
  tiles and its footer are solid tints of the page colour (charcoal, green or gold); cards stay white.
- KPI tiles side by side share one label height, so their numbers line up on one line.

- **☰ Pages** (top left) opens the page list. Pages always use the full width.
- **− / + / Fit** (or the keys `-` and `+`) zoom. Charts redraw sharply at every zoom level.
- **Light / Dark** switches the theme.
- **Present** (or `P`) shows one page at a time. Use ← and → to move, and Esc to exit.
- **Print** prints one page per A4 landscape sheet.

## Known limitations

- Build B2 follows the approved sample page by page: each page of the sample is filled with the
  workbook's own labels, figures and charts. Wording follows the workbook, so a label can read slightly
  differently from the sample.
- Helper cells outside the report panels are used only to fill tiles (e.g. totals); they are not shown.
- Rows of the weekly update log that contain only `[]` placeholders are skipped (the log says how many).
- The assignment-groups table is shown as Excel's own picture of it.
- A chart that has no place in the sample is added at the end of its page, so nothing is dropped.

- The arrangement comes from Excel's positions. A page with unusual overlapping objects may group cards
  differently from the sample design. The log line for each page shows what was found.
- The sample design was hand-arranged from screenshots. This script builds each page automatically from
  what `Wkly_Report` actually contains, so wording, order and grouping follow the workbook.
