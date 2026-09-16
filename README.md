# SCM Inventory & Distribution Planning Dashboard

## Release: v2.37


### v2.37 Operations Excellence Presentation Export Redesign
- Executive PowerPoint export now uses the approved dark-navy / gold **Operations Excellence** cover and KPI design references.
- Cover artwork keeps the globe/network visual and updates the presentation source date dynamically.
- Every KPI is exported as its own full executive slide for **YTD** and **Weekly** with Current Value, movement direction, Average/High/Low, Actual + Trend graph and management action strip.
- All six KPIs remain included, including **Stock-Out · Before PO Balance**.
- KPI direction rules remain business-aware: stock-out decreases are Positive; MC Class A DoI decreases are Negative because higher coverage is the goal.
- ABC Model Position, Area, Branch and model-detail sections remain in the export after the redesigned KPI pages.

### v2.35 Protected Saved Weekly Delivery Plan
- **Clear Board** clears only the current working/on-screen Weekly Truck Schedule and Allocation view.
- The saved weekly **Schedule + Allocation** is protected and remains persisted.
- **View Weekly Schedule** and **View Allocation Import** explicitly reload the protected saved plan.
- A blank board immediately after Clear Board cannot overwrite the saved weekly plan through **Save Schedule + Allocation**.
- **Clear Allocations** remains a separate, explicitly destructive saved-data action.




### v2.36 Missing KPI Line Graph Restored
- Added **MUTI MC : Stock Outrate - Overall (Before PO Balance)** as the sixth Executive KPI.
- Reads YTD and Weekly values directly from the matching KPI blocks in `KPI_YTD_Input` and `KPI_WEEKLY_Input`.
- Uses **lower-is-better** direction logic, consistent with the other stock-out rate KPIs.
- Included automatically in Executive Dashboard KPI cards/line graphs and presentation export.
- Bundled dashboard import workbook updated from the user-provided `MC Dashbord IMPORT(3).xlsx` reference.

### v2.34 DoI Round-Up + ABC Export + Class A DoI Direction
- All Days of Inventory (DoI) displays and calculated New DoI values now follow the approved **round-up / ceiling rule** (example: 6.01 → 7 days).
- Executive **Class A / B / C Model Position** now includes a **Brand filter**; the summary cards and rows react to the selected Brand.
- **Export Presentation** now includes the imported Class A/B/C Model Position with Class Rank, Brand, Model, rounded-up DoI and Stock Status. The current Brand filter is carried into the exported deck.
- **MC Class A DoI** is now treated as **higher-is-better**. If Current DoI is lower than Previous DoI, the movement is labeled **Negative**; a higher Current DoI is **Positive**.
- KPI cards, chart legends and presentation summaries now use KPI-specific favorable-direction logic instead of assuming every decrease is favorable.

### v2.32 Executive Reorder Card + Unified Weekly Plan Save
- Executive Dashboard now includes a **Class A / B / C Model Position** card sourced directly from the latest imported Re-order sheet.
- The card displays **Class, imported Rank, Brand, Model, DoI and Stock Status**, with A/B/C filters and class summary counts / average DoI.
- Delivery Control Board title is now **Delivery Plan for the Week**.
- Added **Save Schedule + Allocation** so the current Weekly Truck Schedule and Allocation plan can be committed together in one action.
- Existing Save Schedule, Allocation Control, direct Edit / Transfer, Split & Transfer and load-rebalance behavior remain available.

### v2.31 Allocation Edit / Truck Rebalance
- **Edit / Transfer** on any model now opens that allocation directly in a focused popup instead of opening the full Allocation Control Center.
- The popup includes Model, Branch, Quantity, Class, Remarks and a **Target Truck / Trip** selector.
- Selecting another saved Day + Truck creates a manual rebalance override; the model is planned on that truck even when it was originally mapped to another route.
- **Save & Rebalance** can edit the allocation and move it to another truck in one action. Selecting **AUTO** removes the manual override and returns the allocation to its normal Branch schedule.
- Added **Split & Transfer**: move only part of an allocation quantity to another truck while retaining the balance on the original allocation.
- Target truck trips are validated against the saved Weekly Truck Schedule and the configured maximum route-stop limit.
- Manual transfers are clearly marked in the Delivery Plan and remain capacity-aware; any residual quantity still follows auto-rollover/backorder rules.

### v2.30 Delivery Plan Control-Tower Redesign
- Audited the Delivery Plan workflow and removed the redundant full-page **Imported Allocation + Scheduled Dispatch** section.
- Replaced it with an on-demand **Allocation Control Center** for correction/maintenance only.
- Main workflow is now: **Weekly Truck Schedule → Allocation Import → Weekly Delivery Plan & Capacity Control → Branch Dispatch Priority**.
- Weekly Delivery Plan now has executive plan-health, completion/exception pulse cards, clickable Monday-Saturday day cards, and stronger truck/branch visual hierarchy.
- **Every model allocation** in the capacity plan can be Adjusted, Edited or Deleted — not only overloaded/carryover/backorder models.
- Allocation Control Center supports Model, Branch, Quantity, Class, Remarks and schedule-mapping review without cluttering the main planning page.
- Added **BURGMAN15** to Model Master with **Index Size 1.5**. A one-time persisted-master migration applies this to existing Render persistent data without repeatedly overwriting future user edits.
- Audit fix: **MUTI SURIGAO / AREA VI** is added to Branch Master when absent so the supplied weekly schedule captures all 37 source trip rows instead of rejecting that route.
- Branch Dispatch Priority remains because it is not redundant: truck cards answer **capacity/control**, while Branch Dispatch Priority answers **execution sequence by branch**.





### v2.29 Weekly Truck Schedule Compact Workflow
- **Weekly Truck Schedule** can now be minimized and expanded without affecting saved schedule data. The user preference is retained in the browser.
- Added a **+ New Trip Assignment** button. Day, Truck, Area and Branch entry fields now open in a focused popup instead of permanently occupying the Delivery Control Board.
- Adding, editing, or removing a trip marks the schedule as **Pending Save**. The **Save Schedule** button changes to `Save Schedule • Pending` and is visually emphasized until the route is committed.
- After **Add Trip**, the popup closes, the Weekly Truck Schedule expands so the newly staged trip is visible, and the user is prompted to trigger **Save Schedule**.
- **Save Schedule** persists the full route, clears only the New Trip Assignment popup fields, removes the pending-save indicator, and retains all saved trips until **Clear Board**.
- Imported schedules remain persisted immediately by the import workflow; after import the schedule is shown as saved/clean and can be extended through **New Trip Assignment**.


### v2.28 Weekly Truck Schedule Save Behavior
- **Save Schedule** persists the current Weekly Truck Schedule and then clears only the **New Trip Assignment** entry fields (Day, Truck, Area, Branch).
- Saved schedule rows remain visible and active after saving or importing.
- The persisted Weekly Truck Schedule is cleared only by **Clear Board**.
- The separate **Clear Schedule** control was removed to prevent accidental deletion of the saved delivery rhythm.
- **Clear Allocations** remains available and does not affect the saved Weekly Truck Schedule.

### v2.26 Delivery Control Board — Frequency-Aware Whole-Week Auto-Rollover

- Delivery planning is now explicitly **Monday through Saturday**, with **Whole Week** as the default operational view.
- Weekly Truck Schedule accepts repeated Branch delivery slots. A Branch may be scheduled once, twice, or more during the week; the board displays its **weekly delivery frequency**.
- A Day + Truck route is treated as one capacity pool shared by its scheduled Branch stops. Multi-stop routes from the supplied schedule are supported (default maximum 6 route stops, configurable up to 12).
- Allocation is consumed only once across the week. Loading priority remains **Class A → Class B → Class C**, with stock-risk status used inside each Class.
- If a trip exceeds capacity and that Branch has another saved delivery slot later in the week, residual quantity is **automatically carried forward to the next scheduled trip**.
- If no later Branch trip is available, remaining quantity is marked **BACKORDER • FOLLOWING WEEK / ROUTE ADJUSTMENT REQUIRED** and is shown in the board and Excel remarks/action fields.
- Whole Week view now includes a six-day ribbon for Monday-Saturday with trips, planned units, planned utilization, carryover and backorder.
- Truck cards show requested vs planned units, planned utilization, carryover, backorder and per-Branch frequency.
- Branch Loading Priority shows Delivery Day, trip number, Truck, Area, weekly frequency, requested/planned units, carryover, backorder and next-trip routing.
- **Clear Board** clears all operational Delivery Control Board data: saved Weekly Truck Schedule, saved/imported allocation batches, entry fields and derived analysis. Truck/Model/Branch masterlists remain reference configuration.
- The supplied Weekly Truck Schedule and Unit Allocation samples are bundled in `examples/` for testing.


### v2.25 Delivery Week View, Batch Entry Reset & Export Sheets

- **Analyze Day** now includes **Whole Week**. Whole Week aggregates all scheduled delivery days and shows every Branch Loading Priority record with its scheduled delivery day.
- **Export Delivery Plan** now creates separate **Monday, Tuesday, Wednesday, Thursday, Friday, Saturday** worksheets plus **Weekly Schedule**. Each daily sheet clearly displays its Delivery Day.
- All generated Excel exports are intentionally **unfrozen**; no frozen panes are applied.
- After **Save Schedule**, the New Trip Assignment entry controls are cleared for the next schedule batch while saved trips remain visible.
- After **Save** in Imported Allocation + Scheduled Dispatch, saved allocations remain persisted for analysis while the entry grid clears for the next allocation batch; subsequent batches are appended to the saved baseline.
- Branch Loading Priority displays the scheduled delivery day in both single-day and Whole Week analysis.


### v2.21 Motion + Status Graph Workspace

- Removed the `Filtered Status Mix` card and the `How to read` block from Brand & Model Status Summary.
- Status Distribution Graph now uses the full available section width and a taller responsive plotting area.
- Added refined transitions for sidebar expansion, workspace switching, theme changes, filters, cards, rows, forms, and charts.
- Added subtle entrance/refresh animations and Chart.js easing.
- Added `prefers-reduced-motion` support so accessibility settings disable nonessential motion.
- Added a subtle login-card entrance/theme transition.

### v2.19 Presentation Detail + Management Table Theme Fix

- PowerPoint KPI line graphs now show rounded data labels on each Actual point.
- PowerPoint Branch Model Intelligence now separates Class A, Class B and Class C into separate branch/class slides instead of mixing all classes in one model table.
- ORDER REVIEW & PLANNING table was redesigned into a stronger grouped management grid with clearer row cards, line number chips, and Reference / Order Plan / Calculated Result separation.
- Management Model Order Table now responds correctly to Dark and Light Mode.


### v2.18 Branded Executive Deck + Order Table Revamp

- Revamped **Scheduled Truck Delivery Plan** into a cleaner command-board layout with container-responsive sections that remain readable with the sidebar expanded or collapsed.
- Weekly schedule rows are now explicit **Trip cards** with Day / Truck / Area / Branch separated cleanly instead of visually mixing controls.
- Imported Allocation is visually separated from its **Scheduled Dispatch mapping** (Day / Truck / Status).
- Delivery truck-analysis and branch-priority grids now respond to actual workspace width rather than viewport width.
- Management Order Plan Excel removes `Source: Sheet2` and the generated timestamp. The workbook keeps only the Management/Re-order title.
- Management Excel printable columns are now: **Line No., Brand, Model, Unit Cost, Current Inventory, DoI, Stock Status, PO Balance, Order Qty, New DoI, Total Amount, Remarks**.
- **Class and Rank are not exported**; `Line No.` is the sequence column.
- Management Excel column widths, short headers, shrink-to-fit rules and whole-peso number formats were retuned to eliminate unreadable `#######` cells while retaining Letter Portrait / one-page-wide printing.

### v2.12 presentation and export cleanup

- Removed explanatory lock/edit formula text from the Management Order Plan UI.
- Renamed sidebar actions to **Data Import** and **Export Deck**.
- Replaced KPI interpretation prose with period **Average / High / Low** statistics.
- Management Order Plan Excel uses narrower print margins and keeps ordinary text cells on one line; Remarks remains wrapped.
- Branch Request Excel no longer includes the Requested Items / Total Request Qty / Risk Items summary block.


A responsive Python + HTML + Tailwind SCM control tower for Inventory & Distribution Planning, Management Order Planning, Branch Request simulation, and schedule-first Delivery Planning.

### v2.11 design and usability changes

- Restored the approved **dark charcoal + gold/yellow** visual theme across the dashboard, Admin login, Management Excel styling, and PowerPoint KPI Actual-line accent.
- Upgraded typography with **Inter** for interface/body text and **Manrope** for headings and major labels.
- Delivery Plan is rebuilt as five separate full-width workflow sections so fields remain readable whether the sidebar is expanded or collapsed.
- Delivery layouts use **container-aware responsiveness** based on the real workspace width rather than browser width alone.
- Imported Allocation + Scheduled Dispatch now uses responsive cards/grids instead of a forced wide horizontal table.
- Management Order Plan is redesigned into clearly separated **Model Reference / Current Position / Planning Inputs / Calculated Result** groups.
- Locked workbook-reference fields in Management Order Plan: **Class, Brand, Model, DoI, Stock Status**.
- Admin-editable Management fields: **Unit Cost, Current Inventory, PO Balance, Order Quantity, Remarks**.
- **New DoI** and **Total Amount** stay automatic and cannot be overwritten.
- Legacy saved overrides for locked fields are ignored, preserving the imported workbook as source of truth.
- Management Excel download continues to include only rows where **Order Quantity > 0** and stays Letter Portrait / print-ready.

## Management Order Plan formulas

- `Total Amount = Unit Cost × Order Quantity`
- `New DoI = (Current Inventory + PO Balance + Order Quantity) ÷ Avg. Daily Sale (Qty)`

The updated import workbook may contain a Management/Re-order worksheet. The application auto-detects the sheet by its required headers, so it may be named `Sheet2`, `Management`, or another name.

## What is included

- YTD / Weekly KPI selector with Actual + straight directional Trend in the same chart.
- Actual point labels hidden by default with Show Data / Hide Data toggle.
- Area Performance and Branch Performance dashboards.
- Area-first Brand & Model Status Summary.
- Collapsible left sidebar with Dark/Light Mode, Admin Login/Logout, Import, PowerPoint Export, Branch Request Simulator, Management Order Plan, and Delivery Plan.
- Schedule-first Delivery Plan with Truck/Branch schedule, Allocation Import, overload analysis, Class A loading priority, and landscape Delivery Plan export.
- Management Order Plan with filtered planning, saved inputs, New DoI, order-value calculation, and portrait Excel export.
- Guest mode remains view-only.
- Half-up whole-number display rounding: `8.4 → 8`, `8.5 → 9` while calculations retain source precision.

## Stock-Out formula implemented

For each Area or Branch and each class:

`Class Stock-Out Rate = Stockout rows / rows with a nonblank Stock Status (branch) in that same class`

`Overall Average Stock-Out = (Class A rate + Class B rate + Class C rate) / 3`

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
- `BRANCH REQUEST STATUS REPORT` Excel export is configured for **Letter portrait**, one-page-wide fit, tighter margins, smaller professional fonts, print area, repeating table header, and wrapped body text for cleaner printing.


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

Imported allocation data uses the approved template:

```text
Model
Branch
Quantity
CLASS
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

Both worksheets are configured for **Letter Landscape**, **fit to one page wide**, compact print fonts, repeating print titles, tight margins, and print-ready conditional emphasis for overload / high utilization / underutilized / no-allocation conditions.


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


## v2.5 — Collapsible Sidebar + Area-First Brand/Model Status Intelligence

- Replaced top navigation tabs with a **collapsible command sidebar**. Desktop users can expand/collapse it; mobile users get an off-canvas drawer.
- Moved **Dark/Light Mode, Admin Login/Logout, Import Dashboard Data, and Export PowerPoint** into the sidebar.
- Added a **minimize / hide control for the top header**. A small Show Header control restores it, maximizing dashboard viewing space. Sidebar and header preferences persist in browser local storage.
- Revamped **Brand & Model Status Summary** into an Area-first analysis:
  - Area × Brand × Class rows show which Brand has Stockout / Re-order / Over / other statuses in each Area.
  - Area × Brand × Model × Class rows show the exact Model stock-status exposure within each Area.
  - Class A / B / C remains explicit in both views.
  - Status Distribution Graph can be switched between **By Area, By Brand, and By Model** and retains the Stock Status filter.
  - Summary cards now include Areas, Brands, Models, Branches, Inventory, and Stock-Out Rate.
- No changes to KPI formulas, Delivery Plan logic, request simulator logic, rounding rules, or export calculations.


## v2.6 — Command Rail Cleanup + Simplified Status Intelligence

- Removed the duplicate top-header branding/title (`MUTI MC • SUPPLY CHAIN MANAGEMENT` and `Inventory & Distribution Planning Control Tower`) to maximize dashboard viewing space.
- Reworked the collapsible sidebar into a cleaner command rail with consistent inline SVG icons, a stronger active-state indicator, compact spacing, and icon-only collapsed mode.
- Dark/Light Mode, Admin Login/Logout, Import Dashboard Data, and Export PowerPoint remain inside the sidebar with the new icon system.
- On mobile, a compact hamburger/access bar replaces the removed top header and opens the sidebar drawer.
- Removed **Brand Stock Status by Area & Class** from Brand & Model Status Summary.
- Simplified **Status Distribution Graph**: removed its separate View and Stock Status controls. The graph is now always Area-based and automatically inherits the main Area / Branch / Brand / Model / Class / Stock Status filters above it.
- **Model Stock Status by Area** remains the detailed drill-down table.

## v2.11 Design & Usability Update

- Restored the approved dark charcoal + gold/yellow visual theme across the dashboard, login, PowerPoint chart accent, and Management Order Plan Excel styling.
- Upgraded typography using Inter for interface/body copy and Manrope for headings and major labels, with clearer hierarchy and tabular-number treatment.
- Revamped Delivery Plan into five separated full-width workflow sections: Weekly Truck Schedule, Allocation Import, Daily Truck Analysis, Imported Allocation + Scheduled Dispatch, and Branch Loading Priority.
- Delivery Plan responsiveness now uses container-aware layout rules so schedule fields and allocation fields remain readable whether the sidebar is expanded or collapsed.
- Imported Allocation + Scheduled Dispatch no longer depends on a wide horizontal table; allocation lines render as responsive cards/grids.
- Revamped Management Order Plan with grouped headers and a stronger distinction between locked reference data, planning inputs, and calculated results.
- Locked/read-only in Management Order Plan: Class, Brand, Model, DoI, Stock Status.
- Editable for Admin: Unit Cost, Current Inventory, PO Balance, Order Quantity, Remarks.
- New DoI and Total Amount remain formula-driven.
- Legacy saved overrides for locked fields are ignored so workbook reference fields stay authoritative.


## v2.16 Management Order Plan readability update
- Renamed Current Inventory display to `Inv.` in the Management Order Plan.
- Replaced `Filtered Grand Total` with `Grand Total`.
- Widened Stock Status, Order Qty, and New DoI columns in the print-ready Management Excel export.
- Retained Letter Portrait, fit-to-one-page-width, narrow print margins, and order-only export behavior.

## v2.16 refinements
- Redesigned the **Order Review & Planning / Management Model Order Table** header and grouped table treatment for clearer Reference / Planning Input / Calculated Result separation.
- Management Excel export now creates **one worksheet per Brand** whenever ordered items span multiple Brands. Single-brand exports contain one Brand-named worksheet.
- Management Excel column widths follow the approved print specification: Line No. 10, Brand 13, Model 18, Unit Cost 15, Stock Status 13, PO Bal. 18, Total Amount 19, Remarks 30; key operational numeric columns are centered.
- Branch Performance Area label now refreshes directly from the selected Branch and displays `Area • Branch`.
- Replaced formula/style instructional copy under Area Performance and Executive KPI Trends with operational decision-support guidance.


## Version 2.16 Update

- PowerPoint Export now captures both YTD and Weekly KPI trends in the same deck.
- Area Performance is exported using the Overall / All-Area view.
- Branch Class A Stock-Out Ranking exports all branches.
- Branch Performance exports all branches across paginated ranking slides.
- A model priority snapshot is included to show all-branch Class A and stock-risk exceptions.


## v2.18 Branded presentation and Order Review table

- Added the Brilliant4 logo asset to the dashboard package.
- PowerPoint Export Deck now uses the Brilliant4 logo on the cover and every section slide.
- Presentation styling was upgraded with a stronger executive control-tower layout while preserving YTD, Weekly, Area, Branch Class A, all-Branch and model priority coverage.
- ORDER REVIEW & PLANNING / Management Model Order Table was visually revamped into a stronger three-part review structure: Reference, Inventory Position, Order Plan and Financial Impact.


## v2.21 — Weekly Truck Schedule Bulk Import / Export
- Weekly Truck Schedule now has **Import Schedule**, **Export Template**, and **Save Schedule** actions.
- Export Template downloads the current schedule in an import-compatible Excel layout using columns **Day, Truck, Area, Branch**.
- The template can be edited in Excel and imported back in bulk.
- Import accepts `.xlsx`, `.xlsm`, and `.csv`.
- Import validates Day, active Truck, active Branch, duplicate Branch assignments, and the configured maximum branches per truck/day.
- Area is reconciled against Branch Master; mismatched Area values are corrected with a warning.
- Import replaces the saved weekly schedule only when at least one valid row is found; invalid files do not erase the current schedule.

## v2.22 — Delivery Control Board Clear Actions

Admin now has three protected clear actions in **DELIVERY CONTROL BOARD → Scheduled Truck Delivery Plan**:

- **Clear Board** — clears the saved Weekly Truck Schedule and all imported/manual allocation rows.
- Weekly Truck Schedule is retained after Save Schedule and is removed only by **Clear Board**.
- **Clear Allocations** — clears only imported/manual allocation rows.

Every clear action requires confirmation, saves the cleared state immediately, refreshes Daily Truck Analysis / Scheduled Dispatch / Branch Loading Priority, and **never clears Truck, Model or Branch masterlists**.


## v2.23 — Allocation Model & Branch Dropdowns

- Imported Allocation + Scheduled Dispatch now uses dropdown lists for **Model** and **Branch** in Admin mode.
- Model choices come from the Delivery Model masterlist.
- Branch choices come from active Delivery Branch master records and show the Area in the option label.
- Changing Branch immediately recalculates the mapped Day, Truck and SCHEDULED / UNSCHEDULED status on the same allocation line.
- The existing responsive card layout is retained so Day, Truck and schedule status remain visible without horizontal page scrolling.


## v2.25 — Unit Allocation CLASS Input
- The approved Unit Allocation import template is now: **Model | Branch | Quantity | CLASS | Remarks**.
- The bundled `/delivery/template` download uses the newly supplied `Unit_Allocation_Template.xlsx`.
- Delivery allocation rows persist `class` as A/B/C. Imported CLASS takes priority for delivery loading priority; if blank/invalid, the system falls back to the Dashboard Branch+Model Class when available.
- Imported Allocation + Scheduled Dispatch includes an Admin Class dropdown (A/B/C) and shows Class in guest/read-only mode.
- Class A/B/C priority analysis and Delivery Plan export now use the allocation CLASS value when supplied.
- Older allocation files without CLASS remain importable for backward compatibility and receive a warning/fallback behavior.


## v2.28 — Overload Allocation Correction
- Overloaded/carryover/backorder allocation lines now expose Admin controls: Adjust Qty, Edit Allocation, Delete.
- Quantity adjustment and delete actions persist immediately and recalculate the Delivery Plan.
- Saved allocation batches can be reopened for full Model/Branch/Quantity/Class/Remarks correction, then Save & Recalculate.
- Existing save-and-clear batch behavior is preserved.


## v2.35 Clear Board protection
- Clear Board now resets only the current on-screen working Delivery Board.
- The saved Weekly Truck Schedule and saved Allocation for the week are preserved.
- Use **View Weekly Schedule** or **View Allocation Import** to reload the saved weekly plan after clearing the working board.
- The explicit **Clear Allocations** action remains destructive for saved allocation rows and shows a permanent-action warning.
