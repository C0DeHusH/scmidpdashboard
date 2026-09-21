
## Vercel production readiness

Online Admin imports require **both** a stable Admin signing key and durable Blob storage.

1. Set `SCM_SECRET_KEY` and `SCM_ADMIN_PASSWORD` in **Vercel Project Settings → Environment Variables** for the environment you deploy (Production and Preview if both are used).
2. In **Project → Storage**, create or connect a **Vercel Blob** store to that same project/environment, then redeploy. An OIDC token alone does not mean Blob storage is connected.
3. Open `/health`. A ready deployment must report `serverless_session_ready: true`, `blob_store_id_present: true`, and `persistent_storage: true`.

v2.46.7 deliberately blocks Admin writes instead of accepting a fragile serverless login or pretending an ephemeral `/tmp` save is durable.

> **Current release:** v2.46.7 — Vercel Import/OIDC Fix. See `RELEASE_NOTES_v2.46.7.md`.

> **v2.46.3:** Unified Data Refresh recovery release. Imports are now fully preflighted before commit, tolerate harmless worksheet/header placement differences, preserve prior data transactionally, and always return a readable stage/reference when an import cannot be completed.
>
> **v2.46.2:** Admin access reliability release. Management Order Plan, Branch Request Simulator, Delivery Plan, Data Operations, and Aging Admin actions use an application-specific session cookie, durable local signing key, explicit browser credentials, and automatic re-authentication handling.

# SCM Inventory & Distribution Planning Control Tower — v2.46.7


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

## v2.46.7 Vercel import + persistence

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