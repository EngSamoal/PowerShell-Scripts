# Export-WklyReportDashboardC.ps1: Style C "Briefing" dashboard from `Wkly_Report`

This is a separate script from `Export-WklyReportDashboard.ps1` (Style A) and
`Export-WklyReportDashboardB.ps1` (Design B). It reads the same `Wkly_Report` worksheet in the same
read-only way and presents it in the approved **Style C "Briefing"** design
(`samples/ITSM_Dashboard_Style_C_sample.html`): charcoal, green, gold and stone, solid colours only.

Each run writes three files named after the workbook, for example in `C:\Temp\ITSM_Dashboard_StyleC`:

| File | What it is |
|---|---|
| `Bi_Wkly_Report-Daily Data.html` | The dashboard: one self-contained file; works offline and can be sent by email or Teams |
| `Bi_Wkly_Report-Daily Data.pdf` | One PDF page per dashboard page, 1920 × 1080 as designed (a longer page is kept whole) |
| `Bi_Wkly_Report-Daily Data.pptx` | One 16:9 slide per page; a 1920 × 1080 page fills its slide exactly. PowerPoint is not needed |

## Safety (same as the other two scripts)

- The original workbook is never opened. A temporary copy is opened read-only in a separate hidden Excel
  instance, closed without saving, and deleted. Only the `Wkly_Report` tab is read.
- Your own Excel windows are not touched. The log confirms the original file is unchanged (SHA-256).
- Values are shown exactly as Excel displays them. Nothing is recalculated or invented. Excel errors
  such as `#DIV/0!` are shown as-is with a dashed outline.
- No internet: the PDF/PowerPoint pictures are taken by Edge (or Chrome) from the local file, in a hidden
  window with no network access.

## Run it

```powershell
cd C:\Temp
.\Export-WklyReportDashboardC.ps1 -OpenWhenDone
```

The first line printed shows the build, for example `Export-WklyReportDashboardC (Style C) build 2026.10.23-C1`.

| Parameter | Default |
|---|---|
| `-WorkbookPath` | `C:\Temp\Bi_Wkly_Report-Daily Data.xlsx` |
| `-OutputDirectory` | `C:\Temp\ITSM_Dashboard_StyleC` |
| `-DashboardTitle` | `IT Service Management Briefing` |
| `-Classification` | `CONFIDENTIAL` |
| `-SkipPdf` / `-SkipPptx` | off: both files are written |
| `-OpenWhenDone` | off |

## How each page is filled

- **Cover**: the workbook's title, the text above it (eyebrow), the text under it, every `Label: value`
  line (Reporting Period, Report Date, Prepared By) and three big numbers (total tickets, resolution SLA,
  MTTR resolution, found as in Design B).
- **Every other page**: a solid colour panel (green, gold, charcoal in turn) with the page title, the
  page's **own subtitle from Excel as its key message**, and the page's KPI cards as big numbers.
  The charts, tables and text are on the right.
- **Charts** are redrawn from their own data in the sample's forms:

  | Workbook chart | Shown as |
  |---|---|
  | Monthly volumes (columns or line) | Stacked area / area, peak and latest month labelled |
  | Status doughnut | Waffle (each square = 1%) |
  | Met / breached SLA doughnut | Progress ring |
  | Priority pie, priority per record type | Proportion bar(s) |
  | Categories, classification (bars, top 10) | Lollipop ranking, with Excel's % labels |
  | Top agents | Ranked list |
  | Weekly created vs closed | Grouped columns, "n open" notes |
  | Ageing distribution | Columns in Excel's colours, "count · %" labels |

- **No invented text or targets**: a target line is drawn only where the workbook gives one; the MTTR
  bars use the workbook's own "(57%)" style figures.
- A workbook chart that the sample has no block for is still shown, in the best sample form. A page
  added to the workbook later is shown in the same style.

## Using the dashboard

- One page at a time: page buttons at the bottom, ◀ ▶, arrow keys, Home / End, swipe on touch screens.
- **− / % / + / Fit** zoom (also the `-` and `+` keys), remembered on this PC.
- **All pages** shows every page one after another; **Light / Dark**; **Full screen** (or `F`); **Print**.
- Pages are designed for 1920 × 1080 and scale down to fit smaller screens.

## Known limitations

- The key message of each page is the page's subtitle from Excel; where a page has no subtitle, no
  message is shown.
- Linked/camera pictures (the assignment-groups table) are shown as Excel's picture of them.
- Rows of the weekly update log that contain only `[]` placeholders are skipped (the log says how many).
- PDF and PowerPoint pages are pictures of the dashboard pages: text in them cannot be selected or edited.
