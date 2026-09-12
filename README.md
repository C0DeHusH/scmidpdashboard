# SCM Inventory & Distribution Planning Executive Control Tower

## Release: v2.4

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
- **Brand & Model Status Summary** with filters for Area, Branch, Brand, Model, Class, and Stock Status. It recalculates total Models, Branches, Inventory, Stock-Out Rate, Average DoI, status mix, Brand-level status exposure, and Model-level status exposure from `Stock Status (branch)`.
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

For Render / Linux production:

```bash
gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --threads 4 --timeout 120
```

Render Build Command:

```bash
pip install -r requirements.txt
```

`gunicorn` is included in `requirements.txt` in v1.8. For local Windows service hosting you may still use Waitress if preferred:

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


## v1.8 — Brand & Model Status Summary

- Added a new **Brand & Model Status Summary** section below Branch Performance.
- Added six live filters: **Area, Branch, Brand, Model, Class, and Stock Status**.
- Summary cards display filtered **Models, Branches, Inventory, Stock-Out Rate, and Average DoI**.
- Added a dynamic **Status Mix** using the exact values from `Stock Status (branch)` (currently including Stockout, Re-order, and Over in the supplied workbook).
- Added **Brand Status Summary**, ranked by highest stock-out exposure, showing model count, branch coverage, inventory, suggested transfer, stock-out rate, and status mix per brand.
- Added **Model Status Summary**, grouped by Brand + Model across the selected scope, showing Class, branch coverage, inventory, suggested transfer, Average DoI, stock-out rate, and status mix.
- All visible figures keep the approved half-up whole-number display rule while calculations retain source precision.
- Existing KPI design, straight dashed Trend Direction, Show/Hide Data, Area/Branch sections, Model Intelligence, Dark/Light Mode, simulator, and exports are unchanged.
- Added `gunicorn>=22.0` to `requirements.txt` for Render deployment.

## v1.9 — Brand & Model Network by Area

- Added a new **Brand & Model Network by Area** section without redesigning the approved dashboard.
- Added filters for **Area, Brand, Model, Class, and Stock Status**, plus a Reset Filters button.
- Added a selectable **Network Metric** with five views: **Branch Coverage, Inventory, Stock-Out Rate, Average DoI, and Suggested Transfer**.
- Added network summary cards for **Areas, Brands, Models, Branches, Inventory, and Stock-Out Rate**.
- Added an **Area Network Summary** showing branch/model/brand coverage, inventory, suggested transfer, Average DoI, stock-out exposure, and dynamic status mix per Area.
- Added a **Brand / Model × Area Network Matrix**. Rows are Brand + Model; columns are Areas. Each Area cell displays the selected network metric plus branch coverage and inventory context.
- Selecting a specific Area reduces the matrix to that Area; selecting Overall shows the full network across all Areas.
- All values keep the approved whole-number half-up display rule while calculations retain source precision.
- Existing KPI, Trend Direction, Show/Hide Data, Area/Branch Performance, Model Intelligence, Status Summary, simulator, Dark/Light Mode, Admin/Guest, Render deployment, and exports remain unchanged.

## v2.0 — Brand/Model Performance + Delivery Plan Intelligence

### KPI point labels

- **Actual point data labels are hidden by default** every time the dashboard opens.
- The existing **Show Data / Hide Data** control remains available and affects all Executive KPI charts at once.
- Trend Direction remains unchanged: one straight dashed regression direction line above Actual.

### Separate Brand & Model Performance

A new **Brand & Model Performance** section is separate from both the Status Summary and the Network-by-Area matrix.

Filters:

- Area
- Branch
- Brand
- Model
- Class

Performance outputs include:

- Brand count / Model count / Branch coverage
- Class A Stock-Out Rate
- Overall Stock-Out Rate
- Inventory
- Average DoI
- Suggested Transfer
- Brand Performance ranking
- Model Performance ranking

Rows are prioritized by **Class A Stock-Out exposure first**, then overall Stock-Out exposure and Suggested Transfer requirement.

### Delivery Plan tab

A third main tab, **Delivery Plan**, was added. The supplied `Delivery_Capacity_Simulator_V13- Final.xlsm` was used as the business/master-data reference, while the web presentation was redesigned independently.

The default delivery masterlist includes the reference truck capacities, item/model index sizes, and branch-to-area mapping. It remains fully editable by Admin.

#### Allocation import

Admin may import `.xlsx`, `.xlsm`, or `.csv`. The only business fields required are:

```text
Model
Branch
Quantity
Remarks
```

A ready-to-use import workbook is included at:

```text
data/Delivery_Allocation_Import_Template.xlsx
```

and can also be downloaded directly from the Delivery Plan tab.

Imported allocation rows can be reviewed, manually adjusted, added, or removed before saving.

#### Daily Trip Schedule

Admin can create and modify the daily schedule using:

```text
Day
Branch
```

The same branch may be scheduled on different days. Duplicate Day + Branch rows are removed on save.

Truck assignment is generated by the analysis engine rather than stored in the schedule itself.

#### Daily fleet analysis

For the selected day, the engine:

1. Finds allocation rows for scheduled branches.
2. Converts each model quantity to **truck index load** using the editable Model Index master.
3. Pulls Branch + Model **Class A/B/C** and Stock Status from the current SCM Dashboard import when available.
4. Prioritizes branches with the highest Class A allocation first, then risk status and required index.
5. Assigns one or two branches per truck depending on available capacity.
6. Prefers efficient two-branch pairing when the combined load fits a truck; same-area pairing receives preference.
7. Uses the smallest available truck that safely fits the proposed load so larger trucks remain available for larger allocations.
8. If one branch itself exceeds truck capacity, it assigns the largest available truck, marks **OVERLOAD**, and creates a model-level loading priority.
9. Within constrained loads, **Class A models are planned first**, then risk-status items, then remaining models.
10. Reports planned units and deferred units when full allocation cannot be carried.

Truck statuses are driven by editable thresholds:

```text
OVERLOAD                  > 100%
FULL / HIGH UTILIZATION   >= Full threshold (default 90%)
OPTIMIZED                 >= Underutilized threshold (default 75%)
UNDERUTILIZED             < Underutilized threshold
IDLE                      No assigned branch
```

The Delivery tab shows:

- Daily fleet utilization summary
- Per-truck capacity and utilization
- One/two-branch assignment
- Class A units per truck
- Deferred units
- Priority Loading Detail by Branch + Model
- Unassigned branches when fleet capacity is exhausted
- Branch Loading Priority ranking
- Missing Model Index warnings

### Editable Delivery Masterlists

Admin can update:

- **Trucks:** Plate, capacity index, description, active/inactive
- **Models:** Model, index size
- **Branches:** Branch, Area, active/inactive
- **Planning thresholds:** Underutilized %, Full %, Maximum branches per truck

The dashboard automatically adds Branches discovered in the SCM Raw/Distribution import if they are missing from the Delivery Branch master, without overwriting existing master data.

### Render persistence

Delivery allocations, schedule edits, and delivery masterlist changes are written to the application state directory.

For Render production, use a Persistent Disk and set for example:

```text
SCM_DATA_DIR=/var/data/scm
```

Mount the Render Persistent Disk at `/var/data`. Without a persistent disk, Render's ephemeral filesystem can lose imported/edited delivery state after a restart or redeploy.

## v2.1 — Executive Cleanup + Scheduled Truck Delivery Logic

### Executive Dashboard

- Removed **Brand & Model Network by Area** from the Executive Dashboard.
- Removed the **Brand / Model × Area Network Matrix** from the Executive Dashboard.
- The separate **Brand & Model Performance** section was later removed in v2.2.
- Enhanced **Brand & Model Status Summary** with a responsive stacked **Status Distribution Graph**.
- The Status graph can switch between **By Brand** and **By Model** and follows the active Area / Branch / Brand / Model / Class / Stock Status filters.
- Executive KPI point-value labels remain **Hidden by default**. The user may press **Show Data** to reveal rounded Actual point values.

### Delivery Plan — schedule-first architecture

The Delivery Plan was rebuilt so the user's saved dispatch schedule is the authoritative rule.

Each schedule row now contains:

```text
Day
Truck / Plate
Branch
```

Example:

```text
Monday | MAD 2439 | MUTI 3S
Monday | MAD 2439 | HONDA KORONADAL
```

A Truck may carry up to the configured maximum branches per day (default: 2).

Imported allocation data still requires only:

```text
Model
Branch
Quantity
Remarks
```

The system automatically maps each imported Branch to its saved **Day + Truck** schedule. The import file does not need Day or Truck columns.

#### Scenario 1 — two scheduled Branches overload the Truck

If Monday / MAD 2439 is scheduled for MUTI 3S + HONDA KORONADAL and both Branches receive high allocation, the system keeps both on MAD 2439 Monday and evaluates the combined load.

If the Truck is overloaded, the loading recommendation is capacity-aware and prioritizes:

```text
1. Class A models
2. Within Class A: Stockout first, then Re-order / risk status
3. Class B
4. Class C
```

The Truck card displays Planned Units and Deferred Units per Model so the planner can see what should load first.

#### Scenario 2 — one scheduled Branch has no allocation

If Monday / MAD 2439 is scheduled for MUTI 3S + HONDA KORONADAL but no allocation exists for HONDA KORONADAL, this is accepted as a valid schedule.

The system does **not** force another Branch into the open slot. The Truck analyzes only the allocation that exists and visibly marks HONDA KORONADAL as:

```text
Scheduled / No Allocation
```

#### Unscheduled imported Branch

If an imported Branch has no saved weekly schedule, it is shown as **UNSCHEDULED** and is not silently assigned to another Truck.

### Delivery Masterlists

The large Delivery Masterlist Manager is now hidden from the normal Delivery Plan screen.

Admin sees one button:

```text
Update Masterlists
```

This opens a modal for:

- Trucks / capacities
- Model index sizes
- Branch / Area master
- Utilization thresholds
- Maximum Branches per Truck

This keeps the operational Delivery Plan focused on Schedule → Allocation → Truck Analysis → Priority Loading.

## v2.2 — Delivery Plan Export + Weekly Schedule Layout Fix

### Executive Dashboard cleanup

- Removed the separate **Brand & Model Performance** section from the Executive Dashboard.
- **Brand & Model Status Summary** and its filterable Status Distribution Graph remain unchanged.

### Weekly Truck Schedule responsive redesign

- Rebuilt the Weekly Truck Schedule edit rows to prevent overlapping fields.
- Day, Truck, Branch, and Remove controls now use responsive minimum widths and stack cleanly on narrower screens.
- The Add Trip controls use the same responsive layout and no longer squeeze long Branch names into overlapping fields.

### Delivery Plan Excel export

Admin now has an **Export Delivery Plan** button beside **Update Masterlists**.

The export follows the currently selected Analyze Day and creates a highly formatted `.xlsx` workbook with:

1. **Daily Delivery Plan**
   - Executive delivery summary KPIs
   - Truck Dispatch & Capacity Summary
   - Truck utilization and overload status
   - Model Priority Loading Recommendation
   - Requested / Planned / Deferred quantities
   - Class A and stock-risk highlighting
   - Branch Loading Priority
   - Unscheduled imported allocations when present

2. **Weekly Schedule**
   - Day + Truck + Branch
   - Area
   - Allocation / No Allocation state
   - Units, Class A units, risk units and required load index
   - Truck utilization and Truck Status
   - Recommended action

Both worksheets are configured for **A4 Landscape**, **fit to one page wide**, compact print fonts, repeating print titles, tight margins, and print-ready conditional emphasis for overload / high utilization / underutilized / no-allocation conditions.


## v2.3 — Status Graph Filter + Area-Aware Searchable Weekly Schedule

### Status Distribution Graph
- Added a dedicated **Stock Status** filter inside the Status Distribution Graph.
- The graph can now be viewed **By Brand** or **By Model** and independently narrowed to one Stock Status.
- The graph-specific filter does not change the other Brand & Model Status Summary cards/tables unless the main summary Stock Status filter is also changed.

### Weekly Truck Schedule
- Added an **Area** filter to the Add Trip controls.
- After an Area is selected, the searchable Branch list contains **only active branches under that Area**.
- Truck selection is now **type-to-search** using the active Truck Master.
- Branch selection is now **type-to-search** using the active Branch Master.
- Existing saved schedule rows show Area; Truck uses the full dropdown list while Branch remains searchable and Area-filtered.
- Changing an existing row's Area immediately rebuilds that row's Branch suggestions and clears an invalid branch selection.
- The saved dispatch rule remains **Day + Truck + Branch**; Area is a selection/filter aid derived from the Branch Master.


## v2.4 — Brand/Model Status Logic + Truck Dropdown + Delivery Layout

### How Brand & Model Status Summary is computed
The summary uses the `Raw` / Distribution worksheet at **branch-model row level**. A usable record requires both `Standard Description` (displayed as Model) and `Branch`.

Source fields:
- Brand = `BRAND`
- Model = `Standard Description`
- Class = `CLASS`
- Inventory = `Inv. Qty Total`
- DoI = `DoI (Branch)`
- Stock Status = `Stock Status (branch)`
- Suggested Transfer = `Suggested Transfer`

All selected filters are applied first: Area, Branch, Brand, Model, Class and Stock Status.

Summary formulas after filtering:
- **Models** = unique `(Brand, Model)` combinations.
- **Branches** = unique Branches represented by the filtered records.
- **Inventory** = sum of `Inv. Qty Total` across filtered rows.
- **Suggested Transfer** = sum of `Suggested Transfer` across filtered rows.
- **Average DoI** = arithmetic average of `DoI (Branch)` across filtered branch-model rows. It is not inventory-weighted or sales-weighted.
- **Stock-Out Count** = count of filtered rows whose `Stock Status (branch)` contains `Stockout` or `Stock Out`.
- **Stock-Out Rate** = Stock-Out Count ÷ count of filtered rows with a nonblank `Stock Status (branch)` × 100.
- **Status Mix** = count of filtered branch-model rows for each exact Stock Status value.

Brand Summary groups the filtered rows by Brand. Model Summary groups them by `(Brand, Model)`. Both compute their own Inventory, Suggested Transfer, Average DoI, Stock-Out Count/Rate and Status Mix. Rows are ranked highest Stock-Out Rate first, then highest Stock-Out Count.

The Status Distribution Graph is **record-count based**, not unit-quantity based. Each bar segment counts how many branch-model records carry that Stock Status. Its Stock Status graph filter can narrow the visualization to one status.

### Weekly Truck Schedule
- Truck selection is now a **full dropdown list** populated from Truck Master instead of type-to-search.
- All Truck Master records are visible in the dropdown; inactive trucks are labeled `(Inactive)` and disabled for new assignment.
- Existing saved schedules retain their currently saved truck even if that truck is later made inactive.
- Area-filtered searchable Branch behavior remains unchanged.

### Delivery Plan presentation
`Imported Allocation + Scheduled Dispatch` is now presented first. `Branch Loading Priority` is placed directly below it as a full-width responsive priority section for a clearer planning sequence.
