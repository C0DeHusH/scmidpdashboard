# Vercel Deployment Checklist — SCM IDP Dashboard v2.46.7

## 1. Configure stable Admin authentication

Open **Vercel → Project → Settings → Environment Variables**.

Create:

- `SCM_SECRET_KEY` = a long random secret (recommended 64+ random characters)
- `SCM_ADMIN_PASSWORD` = your Admin password
- `SCM_UTC_OFFSET_HOURS` = `8`

Apply them to **Production**. Apply them to **Preview** too if you use Preview deployments.

Do not upload or commit a `.env` file containing production secrets.

## 2. Connect persistent Blob storage

Open **Vercel → Project → Storage**.

Either create a Blob store from the project or connect an existing Blob store to this project. Use a **Private** store for dashboard operational data.

For current OIDC Blob connections, Vercel provides the request OIDC token and the Blob project connection supplies `BLOB_STORE_ID`. The presence of `x-vercel-oidc-token` by itself is not proof that Blob is connected.

After changing Storage connections, **redeploy** the project.

## 3. Check `/health` before importing

Open:

`https://YOUR-DOMAIN/health`

Do not run Unified Data Refresh until all of these are true:

- `status` = `ok`
- `deployment_ready` = `true`
- `serverless_session_ready` = `true`
- `blob_store_id_present` = `true`
- `persistent_storage` = `true`
- `blocking_issues` = `[]`

For an OIDC Blob connection, `storage_auth_mode` should normally be `oidc-request`.

## 4. Sign in again after redeploy

Old cookies may have been signed by the previous deployment key. Sign out, clear the old SCM cookie if necessary, and sign in once using the configured Admin password.

## 5. Run Unified Data Refresh

The uploaded `.xlsx` is staged under Vercel `/tmp`, validated, used to refresh Aging, then published as a durable Blob revision. `/tmp` is only working storage; Blob is the durable source of truth.

## What your v2.46.6 health output meant

- `oidc_request_token_present: true`: Vercel OIDC is available to the Function.
- `blob_store_id_present: false`: no Blob store is connected to this deployment/environment.
- `persistent_storage: false`: the dashboard correctly refused to claim that `/tmp` was durable.
- `data_source: bundled-baseline`: no successful durable Unified Import is active online yet.

This is why Local works while Vercel import fails: Local has a persistent filesystem; Vercel Functions do not.
