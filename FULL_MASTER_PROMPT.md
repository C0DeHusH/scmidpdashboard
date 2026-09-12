# MASTER PROMPT — SCM INVENTORY & DISTRIBUTION PLANNING EXECUTIVE CONTROL TOWER

## ROLE
You are a high-level Python programmer, supply-chain analytics architect, UI/UX designer, and reporting automation expert. Build a production-ready, responsive Supply Chain Management — Inventory & Distribution Planning Executive Control Tower using Python for the backend and HTML + Tailwind CSS + JavaScript/Chart.js for the frontend.

Use the supplied Excel import workbook as the source of truth for sheet names, headers, KPI series, branch/area/model data, formulas, and data refresh behavior. Use the supplied Branch Request workbook as the design/reference for the downloadable branch-request status report.

The application must be management-ready, visually powerful, responsive on PC/tablet/mobile, easy to refresh, and safe for Admin and Guest roles.

---

## 1. TECHNOLOGY AND APPLICATION ARCHITECTURE

Build the application with:

- Python backend using Flask.
- HTML templates.
- Tailwind CSS for responsive styling.
- JavaScript for interactivity.
- Chart.js for KPI charts.
- XLSX reader/import logic for the supplied workbook.
- XLSX export for Branch Request Status Report.
- PowerPoint export for a ready-to-present executive deck.
- Responsive design for desktop, tablet, and mobile.
- Dark Mode and Light Mode.
- Local Windows launch support with safe automatic port fallback.

The app should run locally from `python app.py` or a Windows launcher such as `run_windows.bat`.

Default local binding should be `127.0.0.1`, not forced `0.0.0.0`, to avoid Windows socket permission error 10013. Prefer port 5055 and automatically fall back to other usable ports such as 8050, 8088, 8765, 9000, then 5000. If all are unavailable, request a free ephemeral port.

---

## 2. USER ROLES AND SECURITY

Support two user modes:

### Guest
Guest is view-only.

Guest may:
- View KPI Dashboard.
- Switch YTD / Weekly period.
- Use Area filter.
- Use Branch filter.
- View Class A/B/C Model Intelligence.
- Use Dark / Light Mode.

Guest must not be able to:
- Import data.
- Export PowerPoint.
- Add/submit branch request lines.
- Download branch-request Excel.
- Perform Admin data-management actions.

### Admin
Admin can perform all Guest functions plus:
- Import/replace the latest Excel workbook.
- Export PowerPoint presentation.
- Build branch requests.
- Export Branch Request Status Report Excel.

Use environment variables for production credentials, including admin password and application secret. A demo default password may exist for local testing only.

---

## 3. DATA IMPORT CONTRACT

The imported workbook must contain these sheets:

- `Raw`
- `KPI_YTD_Input`
- `KPI_WEEKLY_Input`

The `Raw` sheet represents Distribution-level data and should include the supplied workbook's headers such as:

- Branch
- Area
- Standard Description
- RANK
- CLASS
- Avg. Daily Sale (Qty)
- Inv. Qty Total
- DoI (Branch)
- Stock Status (branch)
- Suggested Transfer
- Brand or equivalent source field when available

Do not hard-code branch names, areas, models, KPI dates, or KPI values. The dashboard must refresh dynamically from the imported workbook.

Validate the workbook before accepting an Admin import. If required sheets/columns are missing, reject the import with a useful error message and keep the existing active dataset.

---

## 4. GLOBAL NUMERIC DISPLAY / ROUNDING RULE

All displayed numbers must use standard half-up whole-number rounding:

- 8.4 → 8
- 8.49 → 8
- 8.5 → 9
- 8.6 → 9

Apply this display rule consistently to:

- KPI cards.
- Chart values/tooltips.
- Percentages.
- DoI.
- Inventory.
- Suggested Transfer.
- Rank where numeric.
- Branch Request Simulator outputs.
- PowerPoint values.
- Branch Request Excel display values.

Important: retain full source precision internally for calculations. Round only for display/export presentation unless a business formula explicitly requires otherwise.

---

## 5. EXECUTIVE KPI DASHBOARD

Create powerful management KPI visuals for the following five KPIs:

1. MC Class A DoI
2. MUTI MC : DoI
3. Overall Class A Stock Out Rate
4. MUTI MC : Stock Outrate - Per Branch
5. MUTI MC : Stock Outrate - Overall after PO Balance

Source KPI values from:

- `KPI_YTD_Input`
- `KPI_WEEKLY_Input`

### KPI Period Selector
Do not display YTD and Weekly simultaneously.

Add one global dropdown:

- YTD
- Weekly

Selecting a period refreshes all KPI cards/charts to the chosen period.

### YTD X-Axis
For YTD charts:
- Display month labels only: Jan, Feb, Mar, Apr, etc.
- Show every available YTD month.
- Fit the full chart into the available card/screen width.
- Do not require horizontal scrolling.
- Do not hide months merely to solve responsiveness.
- Use responsive font/tick sizing so the graph stays readable on desktop/tablet/mobile.

Weekly charts may show compact weekly/date labels.

---

## 6. KPI ACTUAL + TREND LINE DESIGN — FINAL RULE

Actual and Trend must be shown in the SAME chart.

### Actual Series
- Display Actual KPI values as the main line.
- Actual may use point markers and light area fill.
- Actual line may naturally rise/fall with source values.

### Trend Series
The Trend must be a true STRAIGHT DIRECTIONAL LINE.

Do NOT:
- Make the Trend follow the Actual curve.
- Use per-period max logic that bends the Trend around Actual.
- Place Trend in a separate chart or strip.

Required behavior:

1. Calculate a normal linear regression over the KPI period:
   `Trend(i) = intercept + slope × i`
2. This creates one straight line showing the overall direction of the KPI.
3. Render it as a dashed Trend line in the SAME chart as Actual.
4. The Trend should appear visually above the Actual line and must not overlap it.
5. To achieve that without bending the Trend, calculate ONE constant visual vertical offset for the entire regression line when needed.
6. Apply the same offset to every Trend point. Never offset each Trend point differently.
7. Because the offset is constant, the displayed Trend remains perfectly straight and preserves its slope/direction.
8. Dashboard tooltip should report the mathematically calculated regression Trend value, not the visual-offset value.
9. The PowerPoint export must use the same straight directional Trend behavior.

The Trend exists to communicate direction, not to mimic the Actual line.

---

## 7. KPI CARD PRESENTATION

For each KPI display:

- KPI name.
- Latest value.
- Unit (% or Days).
- Change versus previous period.
- Direction indicator.
- One combined Actual + straight Trend chart.
- Short interpretation text.

Where lower values are favorable, clearly communicate this in the interpretation.

Keep charts elegant, uncluttered, and presentation-ready.

---

## 8. AREA PERFORMANCE DASHBOARD

Place Area Performance immediately below the Executive KPI section.

Add an Area filter with:

- Overall
- Individual Area selections

### Area Metrics
Calculate and display:

#### Class A Stock Out Rate
`Total Class A Stock Out Count ÷ Total Class A Stock Status Count`

#### Class B Stock Out Rate
`Total Class B Stock Out Count ÷ Total Class B Stock Status Count`

#### Class C Stock Out Rate
`Total Class C Stock Out Count ÷ Total Class C Stock Status Count`

#### Overall Average Stock Out Rate
`(Class A Stock Out Rate + Class B Stock Out Rate + Class C Stock Out Rate) ÷ 3`

Use only rows with a nonblank `Stock Status (branch)` in the denominator for the same class.

Important: use SAME-CLASS denominators. Do not divide Class A stockouts by Class B stock-status counts.

### Area Ranking
Display every Area's Average Stock Out Rate sorted from highest to lowest.

For each Area show:
- Rank.
- Area.
- Class A Stock Out Rate.
- Class B Stock Out Rate.
- Class C Stock Out Rate.
- Overall Average Stock Out Rate.

Visually emphasize the highest-risk Areas.

### Branch Class A Stock Out Rate inside Area Performance
Add a Branch Class A Stock Out Rate ranking within the Area Performance section.

Formula per branch:
`Branch Class A Stock Out Count ÷ Branch Class A Stock Status Count`

Behavior:
- When Area filter = Overall, display all branches sorted highest Class A Stock Out Rate to lowest.
- When a specific Area is selected, display only branches belonging to that Area.
- Show Branch, Area, stockout/status counts, and Class A Stock Out Rate.

---

## 9. BRANCH PERFORMANCE DASHBOARD

Place the entire Branch Performance section BELOW Area Performance.

Add a Branch filter.

For the selected Branch calculate:

### Class A Stock Out Rate
`Branch Class A Stock Out Count ÷ Branch Class A Stock Status Count`

### Class B Stock Out Rate
`Branch Class B Stock Out Count ÷ Branch Class B Stock Status Count`

### Class C Stock Out Rate
`Branch Class C Stock Out Count ÷ Branch Class C Stock Status Count`

### Overall Average Stock Out Rate
`(Branch Class A Rate + Branch Class B Rate + Branch Class C Rate) ÷ 3`

Display KPI cards for Class A, B, C, and Overall Average.

---

## 10. MODEL INTELLIGENCE — REDESIGNED CLASS A/B/C PRESENTATION

Do not use three cramped Class A/B/C card columns.

Create one unified `Model Intelligence` section with selector tabs/buttons:

- Class A
- Class B
- Class C

When the user selects a class, show only models belonging to that class for the selected Branch.

Display these fields:

- Rank
- Model (`Standard Description` renamed/displayed as `Model`)
- Brand
- Stock Status (from `Stock Status (branch)`)
- Inventory
- Suggested Transfer
- DoI (from `DoI (Branch)`)

Desktop:
- Use a clean management matrix/table.
- Keep columns aligned and readable.
- Use compact status badges.

Tablet/Mobile:
- Automatically reflow each model into a compact responsive layout/cards.
- No horizontal page scrolling.
- No text overlap.

Highlight Class A as highest priority without making the design visually noisy.

---

## 11. SECOND TAB — BRANCH REQUEST SIMULATOR

Create a second major application tab called Branch Request Simulator.

### Branch Selection
User selects a Branch.

### Request Builder
User selects a Model from a dropdown containing only models for the selected Branch.

Automatically show:

- Model
- Current Inventory
- Stock Status (`Stock Status (branch)`)
- Suggested Transfer
- Current DoI (`DoI (Branch)`)
- Avg. Daily Sale (Qty)

Add editable fields:

- Requested Quantity
- Remarks

### Projected New DoI
Calculate:

`New DoI = (Current Inventory + Requested Quantity) ÷ Avg. Daily Sale (Qty)`

If Avg. Daily Sale is zero, return a clear N/A/no-ADS result instead of division-by-zero.

Allow Admin to add multiple model request lines before export.

Guest users must not be allowed to create/export requests.

---

## 12. BRANCH REQUEST STATUS REPORT — EXCEL EXPORT

Admin can download the current request as a formatted Excel report.

Report must show:

- Requesting Branch
- Area
- Model
- Class
- Current Inventory
- Requested Quantity
- Stock Status
- Current DoI
- New DoI
- Remarks

### Printing Requirements
The exported `BRANCH REQUEST STATUS REPORT` must be optimized for printing:

- A4 paper.
- Portrait orientation.
- Fit to one page wide.
- Reasonable automatic page height/multiple pages when needed.
- Professional, compact font sizes; no oversized title/body text.
- Tight but readable margins.
- Wrapped text for Model/Remarks.
- Defined print area.
- Table header row repeats on succeeding printed pages.
- Consistent borders/alignment.
- Clear title and branch metadata.

The report should look professional when printed directly from Excel without manual formatting.

---

## 13. ADMIN ACTIONS DROPDOWN

Do not show separate large Import and Export buttons.

Create one compact dropdown button labeled similar to:

`Admin Actions ▾`

Inside include:

- Import Data
- Export PowerPoint

Display this menu only to Admin users.

Keep the header clean and executive-looking.

---

## 14. POWERPOINT EXPORT

Admin can export a ready-to-present `.pptx` management deck.

The PowerPoint must be presentation-ready without manual cleanup.

Include:

1. Executive cover / control-tower summary.
2. Latest KPI summary cards.
3. YTD KPI charts.
4. Weekly KPI charts.
5. Area Performance metrics/ranking.
6. Branch Class A Stock Out Rate exposure.
7. Selected Branch Performance.
8. Model Intelligence / priority model summary.
9. Management action frame or recommendations.

For KPI PowerPoint charts:
- Actual and Trend must be in the same chart.
- Trend must be a straight directional regression line.
- Use one constant visual vertical offset when necessary to prevent Trend/Actual overlap.
- Do not bend Trend around Actual.
- YTD axes display month only.
- Apply whole-number half-up display formatting.

Use a polished dark executive theme with strong typography and restrained accent colors.

---

## 15. DARK MODE / LIGHT MODE

Add a visible Dark / Light Mode toggle.

Requirements:
- Dashboard supports both themes.
- Login page supports both themes.
- Charts adapt their axis/grid/text colors to the selected theme.
- Tables/cards adapt cleanly.
- Persist the selected mode in browser local storage so the user's preference remains on refresh/reopen.

---

## 16. RESPONSIVE DESIGN

The entire system must work correctly on:

- Desktop/PC.
- Laptop.
- Tablet.
- Mobile phone.

Requirements:
- No overlapping text.
- No overlapping cards.
- No horizontal page scrolling for normal dashboard use.
- KPI charts fit card width.
- YTD displays all months within the screen/card.
- Model Intelligence transforms gracefully for mobile.
- Filters/buttons remain finger-friendly on touch devices.

---

## 17. MANAGEMENT VISUAL DESIGN

Use a world-class SCM control-tower style:

- Executive visual hierarchy.
- Strong KPI cards.
- Clean spacing.
- Minimal clutter.
- Dark/light theme support.
- Meaningful risk colors only.
- Highlight Class A stockout exposure.
- High-to-low rankings must be immediately understandable.
- Avoid oversized fonts, decorative excess, and dense layouts.

---

## 18. DATA INTEGRITY / FORMULA RULES

Use the imported workbook as source of truth.

Do not silently invent source values.

Stock-out formulas must use:

`Stock Out Rate = Stockout Count ÷ Same-Class Valid Stock Status Count`

Where Valid Stock Status Count means rows with a nonblank branch stock-status value in that same class and selected scope.

Overall Average:

`(Class A Rate + Class B Rate + Class C Rate) ÷ 3`

Simulator New DoI:

`(Inventory + Requested Quantity) ÷ Avg. Daily Sale (Qty)`

Trend:

Use least-squares linear regression over the ordered KPI periods.

Do not use a moving line that follows every Actual point for the trend presentation.

---

## 19. IMPORT/REFRESH BEHAVIOR

Admin imports a refreshed workbook.

After successful validation/import:
- Replace active dashboard data.
- Refresh all KPI, Area, Branch, model, and simulator data.
- Do not require code modification.
- Preserve current application functionality.
- Show a success/error notification.

---

## 20. ACCEPTANCE CRITERIA

The system is considered complete only when all of the following are true:

- All five KPIs read dynamically from YTD/Weekly sheets.
- YTD/Weekly is selected from one dropdown and not shown together.
- YTD month labels fit the screen without horizontal scrolling.
- Actual and Trend appear in the same chart.
- Trend is a straight regression direction line, not a curve following Actual.
- Trend is visually above Actual without overlap through one constant display offset.
- Dark and Light Mode both work and persist.
- All displayed numeric data follows half-up whole-number rounding.
- Area Performance supports Overall and Per Area filters.
- Area stock-out ranking is highest-to-lowest.
- Branch Class A Stock Out Rate is included under Area Performance.
- Branch Performance is placed below Area Performance.
- Branch Model Intelligence uses Class A/B/C selector tabs and the redesigned matrix/mobile layout.
- Branch Request Simulator calculates projected DoI correctly.
- Admin Actions dropdown contains Import Data and Export PowerPoint.
- Guest is view-only.
- PowerPoint is ready to present.
- Branch Request Excel is A4 portrait, fit-one-page-wide, compact-font, and print-ready.
- Dashboard remains usable on PC, tablet, and mobile with no overlapping elements.
- Windows local startup avoids the fixed `0.0.0.0:5000` socket problem and uses automatic port fallback.

---

## 21. FINAL VERSION CHANGE HISTORY

### Initial Build
- Python + Flask + HTML + Tailwind SCM dashboard.
- KPI, Area, Branch, Simulator, Admin/Guest, Excel import, Excel request export, PowerPoint export.

### v1.1
- Fixed Windows socket error 10013.
- Local binding changed to 127.0.0.1 with automatic port fallback.

### v1.2
- Added half-up whole-number display rounding.
- Replaced simultaneous YTD/Weekly charts with a period dropdown.
- Added Branch Class A Stock Out Rate under Area Performance.
- Moved Branch Performance below Area Performance.

### v1.3
- Added YTD month-only labels.
- Added persistent Dark/Light Mode.
- Redesigned Class A/B/C model presentation into Model Intelligence.
- Changed Branch Request export to A4 portrait with compact print-ready formatting.

### v1.4
- Returned Actual + Trend to the same chart.
- Combined Import and PowerPoint Export into Admin Actions dropdown.
- Made YTD charts fit the screen with no horizontal scroll.

### v1.5 — FINAL TREND RULE
- Trend is a straight linear-regression direction line.
- Trend does not bend around or follow the Actual curve.
- Actual and Trend stay in one chart.
- If separation is required, apply one constant visual vertical offset to the whole Trend line so it stays straight and above Actual.
- Tooltips retain the real calculated regression Trend value.
- PowerPoint uses the same straight directional Trend behavior.

---

## FINAL DEVELOPMENT INSTRUCTION

Deliver a clean, modular, maintainable project with meaningful Python modules, robust workbook validation, responsive HTML/Tailwind templates, safe Admin authorization checks on backend routes, polished charts, professional exports, and clear README/startup instructions.

Do not merely mock the interface. The calculations, filters, imports, exports, simulator, Excel report, PowerPoint generation, responsive behavior, and role permissions must be functional end-to-end using the actual supplied Excel data structure.

### v1.6 — APPROVED TREND PRESENTATION REFERENCE

Do not redesign the dashboard. Change only the way the Trend Direction line is presented.

The KPI chart must visually follow this rule:

- Keep Actual and Trend Direction in the same line graph.
- Actual retains the existing solid line, markers, data labels, and existing dashboard colors/design.
- Trend Direction is a light/muted dashed straight line positioned above Actual.
- Trend Direction has no markers and no area fill.
- Trend Direction is a single straight linear-regression direction line; it must never trace, bend around, or follow the Actual curve.
- Apply one constant vertical presentation offset to the entire regression line only when needed to keep it visually above Actual. Never offset individual trend points independently.
- Use a clean dashed pattern similar to the approved reference image.
- In the legend use `Trend Direction ↑`, `Trend Direction ↓`, or `Trend Direction →` based on the regression slope.
- Do not change any other dashboard presentation, layout, colors, spacing, filters, sections, or component design as part of this correction.
