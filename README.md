# SCM Inventory & Distribution Planning Control Tower — v2.47.6



## v2.47.6 Component-Level Filter Refresh

Filtering is now a non-navigation interaction across the control tower. The main SCM dashboard already updated its Area, Branch, Status, Management, Reorder, Simulator and Delivery sections in place; v2.47.6 brings Motorcycle Aging to the same standard. Global Aging filters refresh only the affected Aging summary/cards/charts/tables, while Unit-Level Traceability refreshes only its own card and table. Requests are cancellable, stale responses cannot overwrite newer filter choices, URL state is updated with `history.replaceState`, and scroll/focus are preserved.

## v2.47.5 Light-Mode Visibility + Trend Legend + Unit Filter Position

- Fixed Aging KPI card values and helper text that could disappear or become too faint in Light mode.
- Replaced the Chart.js KPI legend with a high-contrast dashboard legend that stays readable in both themes.
- Trend Direction now uses clear directional icons: ↗ upward, ↘ downward, → flat; Actual remains a solid blue line and Trend remains dashed.
- Unit-Level Traceability remembers the filter control's viewport position before a server-side filter refresh and restores it after the page reloads, so the user stays in the same working section instead of jumping upward.
- No inventory, KPI, Aging, Management, Branch Request or Delivery business rules were changed.

## v2.47.4 Theme Readability + Current YTD Date

- Strengthens text, table, modal, control, badge, sidebar and chart contrast in both Light and Dark themes.
- YTD charts keep prior periods as month-only labels while the newest imported point shows the actual day (for example `Sep 21`).
- Each YTD chart now displays `Latest data · September 21, 2026` using the exact newest imported period.
- The latest-point tooltip also shows the complete date.

## v2.47.3 Aging Export Data Reliability Fix

- Fixed a stale Area/Branch filter bug that could make **Area Intelligence**, **Model Intelligence**, and **Unit Detail** export with no rows even while the Aging dashboard visibly contained data.
- Dashboard, CSV, and Excel export now use one canonical filter-normalization path.
- Export links are generated from the **validated active filter scope**, never from raw browser query parameters.
- If an Area changes and an old Branch no longer belongs to it, the Branch is safely reset to **All Branches** before export.
- Invalid stale Area/Brand values are also normalized against the current Aging dataset.
- The Excel report now records **Dashboard Rows** and any automatic **Scope Repair** in Executive Summary → Applied Filters.
- Empty export sections now show a clear diagnostic message instead of appearing silently blank.


## v2.47.2 Motorcycle Aging Per-Area Intelligence + Export

The Motorcycle Aging workspace now uses an explicit risk-first workflow. Model-Level Intelligence defaults to **91+ Units · Highest → Lowest**, provides a visible Highest/Lowest direction control, sort-by selector, model search, quick views for Risk First / Capital Risk / Oldest First, live row ranking, exposure badges, and a compact percentage meter. Area and Branch rankings show their rank order clearly, while the unit table includes an age-band legend without adding back redundant Age Group/Age Detail columns.


### What changed in v2.47.2
- Added a dedicated **Area Intelligence** worksheet to the Aging Excel report.
- Per Area reporting now includes Units, Inventory Value, Average Age, 91+ Units, 91+ %, Aged Value, 180+ Units, 366+ Units, Oldest Unit, Branch Count and Exposure level.
- Area report is sorted Highest → Lowest by aging exposure and remains responsive to the active Aging filters.
- Unit Detail still carries the Area on every motorcycle row for traceability.

## Runtime foundation: Universal Import + Dynamic KPI Periods

v2.46.9 restores deployment-independent Unified Data Refresh behavior. A missing Vercel Blob store is no longer a normal import/save failure: Local and Render use writable filesystem state, while Vercel can operate with `/tmp` runtime fallback. Private Vercel Blob remains optional for durability across cold starts and redeployments.

The KPI engine is now period-column dynamic. Add the next YTD or Weekly period to the right of the existing periods and the dashboard will capture it automatically when the date/value cells are valid. Excel serial dates and common text/formula dates are supported. The supplied `MC Dashbord IMPORT(6).xlsx` is bundled as the baseline for this release.


### What changed in v2.47.1
- Added Unit-Level Traceability filters: exact-unit search, aging-band filter, and operational sort options.
- Added a professional multi-sheet Aging Excel export (Executive Summary, Model Intelligence, Unit Detail, Data Dictionary).
- Added a global Light/Dark contrast guard so text, controls, semantic colors, tables, charts, and Brilliant4 logos remain readable in either theme.
- Preserved v2.46.9 universal import behavior and dynamic YTD/Weekly KPI columns.

See `RELEASE_NOTES_v2.47.1.md` for this release and `VERCEL_DEPLOYMENT.md` for deployment guidance.

A unified local Flask control tower for:

- Executive SCM KPI monitoring
- Management Order Planning
- Branch Request simulation
- Weekly Delivery Planning and truck capacity control
- Motorcycle Aging Decision Intelligence

## v2.46 unified source architecture

```text
One Consolidated Excel Workbook
  ├─ Raw
  ├─ KPI_YTD_Input
  ├─ KPI_WEEKLY_Input
  ├─ Reorder / Management
  └─ Aging
          ↓
Waitress / Flask Control Tower
  ├─ Executive Dashboard
  ├─ Management Order Plan
  ├─ Branch Request Simulator
  ├─ Delivery Plan
  └─ Motorcycle Aging
```

There is **no Supabase requirement** and **no separate Aging import**.

### Local state

- Unified saved import: `uploads/active_import.xlsx`
- Management edits: `uploads/management_allocations.json`
- Delivery state: `uploads/delivery/`
- Motorcycle Aging runtime database: `uploads/aging/aging.db`
- Global cleared/no-data marker: `uploads/.scm_no_data`
- Generated exports are built on demand and downloaded; they are not archived by the app.

## Historical v2.46.7 Vercel import + persistence

Vercel deployments always use the system temporary area for workbook/SQLite processing, regardless of any stale `SCM_DATA_DIR` value left in project settings. Durable state remains in **Private Vercel Blob**. v2.46.7 supports both the current per-request Vercel OIDC model (`x-vercel-oidc-token` + `BLOB_STORE_ID`) and older `BLOB_READ_WRITE_TOKEN` connections. OIDC-only stores hydrate lazily on the first request, then the Unified Import verifies Blob write/read access before touching the active dashboard state.

See `VERCEL_DEPLOYMENT.md` before redeploying to Vercel.

## v2.45 highlights

- One **Data Import** dialog refreshes SCM + Reorder/Management + Motorcycle Aging together.
- Aging Area is automatically mapped from the Raw Branch/Area master when the Aging sheet does not contain Area.
- `M1`, `M2`, and `MUTI` Aging branch aliases are standardized to the Raw master.
- The supplied consolidated workbook is the shipped baseline and the Aging seed was regenerated from it.
- The old Aging-only import screen/reference workbook were removed.
- Import commit is rollback-protected: a failed module does not intentionally leave SCM and Aging on different sources.
- Navigation remains visually continuous when moving into Motorcycle Aging.
- New professional motion system: smooth workspace entry, sidebar feedback, modal transitions, loading progress, and restrained chart animation.
- Reduced-motion accessibility remains supported.

## Windows quick start

```cmd
python -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
powershell -ExecutionPolicy Bypass -File .\setup_windows.ps1
run_windows.bat
```

See `SCM_LOCAL_SETUP_GUIDE.md` for the full setup and unified-import guide.


## v2.45 update

Sidebar actions are consolidated under **Data Operations**. The PowerPoint export now includes a report-ready flow with KPI trends, Area/Branch Class A-B-C Stock-Out summaries, Branch model action pages, Management Order intelligence and Aging Unit slides.

## GitHub + Render deployment

This repository includes a production `render.yaml` configured for Gunicorn, `/health`, and a persistent local-data disk. See `RENDER_DEPLOYMENT.md` before deploying. Because this no-Supabase edition stores runtime state locally, a Render persistent disk is required for durable imports and Aging/Delivery/Management state.


## v2.46.0 navigation reliability
Data Operations is now an in-sidebar accordion. Hover on desktop or tap/click on touch/keyboard expands Data Import, Export Deck, and Clear Data inside the sidebar. It never floats over workspace content. The Windows launcher also auto-installs missing Python dependencies.

## v2.46.0 verification

Run the dependency-free core smoke suite after code changes:

```bash
python -m unittest discover -s tests -v
```

For the detailed findings and intentionally deferred refactors, see `CODE_AUDIT_v2.46.0.md`.
## v2.46.3 Unified Data Refresh recovery

The consolidated workbook import now uses a staged transaction: workbook signature check → SCM/KPI/Management preflight → Aging preflight → full in-memory dashboard parse → rollback snapshot → Aging refresh → atomic workbook commit → publish. Core data is not replaced until validation and parsing complete. Unexpected server errors are returned as JSON with an import reference and are written to `uploads/logs/unified_import_errors.log` for local diagnostics. Worksheet names tolerate spaces/underscores/hyphens and Raw/Aging header rows can appear within the first eight rows.

See `RELEASE_NOTES_v2.46.3.md` for details.

## v2.46.2 Admin access reliability

The local Admin session no longer uses Flask's generic `session` cookie. The dashboard now uses `scm_idp_admin`, which prevents other local Flask applications from overwriting SCM authentication on the same hostname. A per-user signing key also survives version-folder upgrades, protected browser requests explicitly send same-origin credentials, and stale Admin pages redirect to `/login` instead of showing misleading authorization failures.

See `RELEASE_NOTES_v2.46.2.md` for the full fix.