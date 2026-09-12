# SCM Inventory & Distribution Planning Executive Control Tower

A responsive Python + HTML + Tailwind dashboard built around the supplied `MC Dashboard IMPORT.xlsx` structure.

## What is included

- A **YTD / Weekly dropdown selector** so only one KPI period is displayed at a time, with Actual and Trend together in the **same chart**, with a **straight directional regression Trend line** kept visually above Actual to prevent overlap. A **Show Data / Hide Data** toggle controls rounded value labels above every Actual KPI point, for:
  - MC Class A DoI
  - MUTI MC DoI
  - Overall Class A Stock-Out Rate
  - MUTI MC Stock-Out Rate – Per Branch
  - MUTI MC Stock-Out Rate – Overall After PO Balance
- Area dashboard with Overall / Area filter, Class A/B/C stock-out rates, average stock-out, descending area ranking, and **Branch Class A Stock-Out Rate** ranking filtered by the selected Area.
- Branch dashboard placed **below Area Performance**, with Class A/B/C metrics and a responsive Model Intelligence matrix showing Rank, Model, Brand, Stock Status, Inventory, Suggested Transfer, and DoI.
- Branch Request Simulator with projected DoI calculation:
  - `(Inventory + Requested Quantity) / Avg. Daily Sale (Qty)`
- Admin-only **Admin Actions dropdown** containing Excel Import and PowerPoint Export, plus request-line add and branch request Excel export.
- Guest mode is display-only.
- Responsive layouts for desktop, tablet, and mobile.
- All displayed numeric values use **half-up whole-number rounding**: `8.4 → 8`, `8.5 → 9`. Full source precision is retained internally for calculations.
- No Excel installation is required on the server; the import reader uses standard XLSX XML internals.

## Stock-Out formula implemented

For each Area or Branch and each class:

`Class Stock-Out Rate = Stockout rows / rows with a nonblank Stock Status (branch) in that same class`

`Overall Average Stock-Out = (Class A rate + Class B rate + Class C rate) / 3`

This intentionally uses the **same class** as denominator, resolving the apparent Class A / Class B denominator typo in the original written instruction and matching the later explicit Area Class A formula.

## Run locally

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

pip install -r requirements.txt

# Recommended: change the admin password
# Windows PowerShell
$env:SCM_ADMIN_PASSWORD="YourStrongPassword"
$env:SCM_SECRET_KEY="YourLongRandomSecret"
# macOS/Linux
export SCM_ADMIN_PASSWORD="YourStrongPassword"
export SCM_SECRET_KEY="YourLongRandomSecret"

python app.py
```

The app prints and opens the selected local URL automatically. Preferred URL: `http://127.0.0.1:5055` (with automatic fallback if that port is unavailable).

Demo default admin password (only when `SCM_ADMIN_PASSWORD` is not set): `admin123`

## Production start

```bash
waitress-serve --host=0.0.0.0 --port=5000 app:app
```

## Import file contract

The imported workbook must contain:

- `Raw`
- `KPI_YTD_Input`
- `KPI_WEEKLY_Input`

The `Raw` sheet must include the key headers used by the supplied file, including `Standard Description`, `Branch`, `Area`, `RANK`, `CLASS`, `Avg. Daily Sale (Qty)`, `Inv. Qty Total`, `DoI (Branch)`, `Stock Status (branch)`, and `Suggested Transfer`.

## PowerPoint export

Admin users can export a presentation-ready deck containing:

- Executive KPI summary
- YTD KPI trends
- Weekly KPI trends
- Area performance and rankings
- Selected branch performance
- Management action frame

The current Area and Branch filters are passed into the presentation export.

## Deployment note

Tailwind CSS and Chart.js are loaded from their public CDNs. For a fully offline corporate deployment, download/pin those assets into `static/` and update the two `<script>` tags in `templates/index.html`.


## Windows socket error 10013 fix (v1.1+)

Version 1.1 no longer forces `0.0.0.0:5000`. It binds to `127.0.0.1` and automatically selects the first usable port from 5055, 8050, 8088, 8765, 9000, or 5000. If all are unavailable it requests a free Windows ephemeral port. The selected browser URL is printed in the console and opened automatically.

You may still force a port before launch, for example: `set PORT=8050` then `python app.py`. To allow access from other PCs on your LAN, set `SCM_HOST=0.0.0.0` only after configuring Windows Firewall appropriately.


## Version 1.2 dashboard changes

- Whole-number half-up rounding across dashboard displays and exports.
- Executive KPI charts now show **YTD or Weekly**, selected from one global dropdown, instead of both at the same time.
- Area Performance includes a **Branch Class A Stock-Out Rate** ranking. The list follows the selected Area; Overall shows all branches.
- Branch Performance has been moved below Area Performance for a clearer management review flow.
- PowerPoint and Branch Request Excel exports use the same whole-number display rule.

## Version 1.3 dashboard changes

- KPI **Trend** is now rendered in a dedicated strip **above Actual** so the trend never overlaps the source data line.
- **YTD x-axis displays month only** (`Jan`, `Feb`, `Mar`, etc.); Weekly retains weekly/date labels.
- Added persistent **Dark / Light Mode** toggle. The selected theme is remembered in the browser and is also available on the Admin login screen.
- Redesigned Branch Model presentation into a unified **Model Intelligence** matrix with Class A/B/C selector tabs. Columns: Rank, Model, Brand, Stock Status, Inventory, Suggested Transfer, and DoI. Mobile view automatically changes to compact responsive rows.
- PowerPoint KPI charts follow the same separated Trend-above-Actual presentation, and YTD uses month-only labels.
- PowerPoint Branch Performance uses a unified model matrix rather than three dense class columns.
- `BRANCH REQUEST STATUS REPORT` Excel export is configured for **A4 portrait**, one-page-wide fit, tighter margins, smaller professional fonts, print area, repeating table header, and wrapped body text for cleaner printing.


## Version 1.4 dashboard changes

- Corrected the KPI visualization so **Actual and Trend are in the same line chart**. Trend is visually separated just above Actual rather than placed in a second chart/strip.
- Trend tooltips retain the calculated trend value even when the line is visually offset for readability.
- Consolidated **Import Data** and **Export PowerPoint** into one **Admin Actions** dropdown menu.
- YTD charts are constrained to the available card width with no horizontal chart scrolling. All available YTD months remain visible as month-only labels.
- The PowerPoint KPI charts use the same combined Actual + Trend presentation.


## Version 1.5 dashboard changes

- Finalized KPI Trend behavior as a **straight directional regression line**. It no longer bends or follows the Actual KPI curve.
- Actual and Trend remain in the **same line chart**.
- One constant visual offset is applied to the entire Trend line only when required so it stays above Actual without changing its slope/direction.
- Dashboard Trend tooltips continue to show the mathematically calculated regression Trend value, not the visual offset value.
- PowerPoint KPI charts use the same straight directional Trend presentation.
- Added `FULL_MASTER_PROMPT.md` containing the consolidated final system specification and all revisions from the project.

## Version 1.6 trend presentation correction

Only the Trend presentation was changed; the dashboard layout and visual design remain unchanged.

- Actual and Trend Direction remain inside the same KPI chart.
- Trend Direction is one straight dashed line with no markers.
- The trend does not bend, smooth, or follow the Actual series.
- A single constant vertical display offset keeps the whole trend line above Actual while preserving its straight slope.
- Legend wording automatically shows `Trend Direction ↑`, `Trend Direction ↓`, or `Trend Direction →`.
- YTD/Weekly data, formulas, filters, layout, colors, Dark/Light Mode, Model Intelligence, Area/Branch sections, Admin Actions, and exports are otherwise unchanged.


## v1.7 — Point Data Label Toggle

- Added a **Show Data / Hide Data** button beside the KPI Period selector.
- The toggle applies to all Executive KPI line graphs at once.
- **Show Data** draws the rounded Actual value directly above every Actual point.
- **Hide Data** removes only the visible point values; Actual/Trend lines, point markers, tooltips, filters, and calculations remain unchanged.
- Trend Direction remains marker-free and does not receive point-value labels.
- The setting is remembered in the browser using `localStorage`.
- No other dashboard layout or design was changed.
