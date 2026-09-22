from __future__ import annotations

from flask import Flask, jsonify, render_template, request, session, redirect, url_for, send_file, flash
from pathlib import Path
from io import BytesIO
from datetime import datetime, timezone, timedelta
from decimal import Decimal, ROUND_HALF_UP
from html import escape as html_escape
from copy import deepcopy
from functools import wraps
from dotenv import load_dotenv
import hmac
import math
import hashlib
import os
import shutil
import json
import secrets
import socket
import tempfile
import threading
import time
import webbrowser
import zipfile
import traceback

from dashboard.metrics import DashboardStore
from dashboard.xlsx_export import build_branch_request_xlsx, build_delivery_plan_xlsx, build_management_order_xlsx
from dashboard.ppt_export import build_presentation
from dashboard.delivery import DeliveryStore, DAYS
from dashboard.cloud_state import VercelBlobState

BASE = Path(__file__).resolve().parent
load_dotenv(BASE / ".env")
DATA_FILE = BASE / "data" / "MC_Dashboard_IMPORT.xlsx"
APP_VERSION = "2.48.2"
IS_VERCEL = bool(os.environ.get("VERCEL") or os.environ.get("VERCEL_ENV"))

# Vercel's deployed project filesystem is read-only except for the function's
# writable system temp area. Runtime working state therefore lives under /tmp on
# Vercel. v2.46.9 treats external Blob persistence as optional: when connected it
# is used for cross-instance durability; when absent, imports/saves continue using
# the runtime fallback instead of failing. SCM_DATA_DIR remains supported on Local/Render.
if IS_VERCEL:
    os.environ["SCM_DATA_DIR"] = str(Path(tempfile.gettempdir()) / "scm-idp-dashboard")
STATE_ROOT = Path(os.environ.get("SCM_DATA_DIR", str(BASE / "uploads")))
STATE_ROOT.mkdir(parents=True, exist_ok=True)
UPLOAD_DIR = STATE_ROOT
# Incoming multipart files are isolated from dashboard state.  This prevents a
# stale/custom SCM_DATA_DIR from ever affecting the receive-upload stage.
IMPORT_TMP_DIR = Path(tempfile.gettempdir()) / "scm-idp-dashboard-imports"
IMPORT_TMP_DIR.mkdir(parents=True, exist_ok=True)
ACTIVE_IMPORT = UPLOAD_DIR / "active_import.xlsx"
EMPTY_STATE_MARKER = STATE_ROOT / ".scm_no_data"
MANAGEMENT_ALLOCATIONS = STATE_ROOT / "management_allocations.json"
DELIVERY_STATE_DIR = STATE_ROOT / "delivery"

cloud_state = VercelBlobState()
CLOUD_BOOT_ERROR = ""
CLOUD_BOOT_HYDRATED = False
try:
    # Legacy/static Blob credentials are available at module import time and can
    # hydrate immediately. Modern Vercel OIDC is delivered on each Request, so
    # OIDC-only deployments hydrate lazily in @app.before_request below.
    if cloud_state.enabled:
        cloud_state.hydrate_core(STATE_ROOT)
        cloud_state.hydrate_named_files(STATE_ROOT, [
            "management_allocations.json",
            "delivery/delivery_master.json",
            "delivery/delivery_allocations.json",
            "delivery/delivery_schedule.json",
        ])
        CLOUD_BOOT_HYDRATED = True
except Exception as exc:
    CLOUD_BOOT_ERROR = str(exc)

from dashboard.aging.routes import aging_bp
from dashboard.aging.db import (
    clear_data as clear_aging_data,
    record_count as aging_record_count,
    DB_PATH as AGING_DB_PATH,
    backup_database as backup_aging_database,
    restore_database as restore_aging_database,
)
from dashboard.aging.analytics import executive_summary as aging_executive_summary
from dashboard.aging.importer import import_excel as import_aging_excel, validate_unified_aging


def _shared_session_secret_path() -> Path:
    """Return a per-user secret path that survives extracting a newer dashboard version."""
    explicit = str(os.environ.get("SCM_SESSION_SECRET_FILE") or "").strip()
    if explicit:
        return Path(explicit).expanduser()

    # Windows cookies are shared by hostname rather than TCP port, so keeping the
    # signing key outside a versioned extraction folder avoids unnecessary session
    # invalidation when the user upgrades the local dashboard package.
    local_app_data = str(os.environ.get("LOCALAPPDATA") or "").strip()
    if local_app_data:
        return Path(local_app_data) / "SCM_IDP_Dashboard" / ".session_secret"

    xdg_state = str(os.environ.get("XDG_STATE_HOME") or "").strip()
    if xdg_state:
        return Path(xdg_state) / "scm-idp-dashboard" / ".session_secret"
    return Path.home() / ".scm-idp-dashboard" / ".session_secret"


def _read_secret(path: Path) -> str:
    try:
        value = path.read_text(encoding="utf-8").strip() if path.exists() else ""
        return value if len(value) >= 32 else ""
    except OSError:
        return ""


def _write_secret(path: Path, value: str) -> None:
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_name(path.name + ".tmp")
        tmp.write_text(value, encoding="utf-8")
        try:
            os.chmod(tmp, 0o600)
        except OSError:
            pass
        tmp.replace(path)
    except OSError:
        # A read-only profile or state directory should not prevent local startup.
        pass


SESSION_KEY_SOURCE = "unknown"


def _load_or_create_secret_key() -> str:
    """Return a stable Flask signing key and record where it came from.

    A Vercel Function can be served by different isolated instances. A key generated
    on the local filesystem therefore cannot be trusted for Admin sessions online.
    v2.46.9 keeps that situation as an explicit session configuration error
    instead of allowing a login that later appears to "randomly" log out.
    """
    global SESSION_KEY_SOURCE
    configured = str(os.environ.get("SCM_SECRET_KEY") or "").strip()
    if configured:
        SESSION_KEY_SOURCE = "SCM_SECRET_KEY"
        return configured

    if IS_VERCEL:
        # Legacy Blob credentials are stable across instances. A configured Admin
        # password is also acceptable as key material when combined with the stable
        # Vercel project id. SCM_SECRET_KEY remains the preferred production option.
        blob_secret = str(os.environ.get("BLOB_READ_WRITE_TOKEN") or "").strip()
        admin_secret = str(os.environ.get("SCM_ADMIN_PASSWORD") or "").strip()
        project_id = str(os.environ.get("VERCEL_PROJECT_ID") or "scm-idp-dashboard").strip()
        if blob_secret:
            SESSION_KEY_SOURCE = "BLOB_READ_WRITE_TOKEN-derived"
            return hashlib.sha256(f"scm-idp-session|{blob_secret}|{project_id}".encode("utf-8")).hexdigest()
        if admin_secret:
            SESSION_KEY_SOURCE = "SCM_ADMIN_PASSWORD-derived"
            return hashlib.sha256(f"scm-idp-session|{admin_secret}|{project_id}".encode("utf-8")).hexdigest()

        # Keep Flask able to render public/diagnostic pages, but Admin endpoints are
        # blocked below until a stable deployment secret is configured. Never claim
        # this ephemeral key is safe for cross-instance authentication.
        SESSION_KEY_SOURCE = "ephemeral-unconfigured"
        return secrets.token_urlsafe(48)

    shared_path = _shared_session_secret_path()
    local_path = STATE_ROOT / ".session_secret"
    existing = _read_secret(shared_path) or _read_secret(local_path)
    if not existing:
        existing = secrets.token_urlsafe(48)
    _write_secret(shared_path, existing)
    _write_secret(local_path, existing)
    SESSION_KEY_SOURCE = "local-session-file"
    return existing


app = Flask(__name__)
app.secret_key = _load_or_create_secret_key()
SERVERLESS_SESSION_READY = (not IS_VERCEL) or SESSION_KEY_SOURCE != "ephemeral-unconfigured"
app.config["MAX_CONTENT_LENGTH"] = (4 * 1024 * 1024) if IS_VERCEL else (25 * 1024 * 1024)
# IMPORTANT: use an application-specific cookie name. Browser cookies are scoped
# by host/path, not by port, so Flask's default `session` cookie can be overwritten
# by another local Flask app or another SCM version running on 127.0.0.1.
app.config["SESSION_COOKIE_NAME"] = str(os.environ.get("SCM_SESSION_COOKIE_NAME") or "scm_idp_admin")
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = os.environ.get("SCM_COOKIE_SECURE", "1" if IS_VERCEL else "0") == "1"
app.config["SESSION_COOKIE_PATH"] = "/"
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(hours=max(1, int(os.environ.get("SCM_ADMIN_SESSION_HOURS", "12"))))
app.config["SESSION_REFRESH_EACH_REQUEST"] = True
app.config["JSON_SORT_KEYS"] = False
app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 3600
ADMIN_PASSWORD = str(os.environ.get("SCM_ADMIN_PASSWORD") or "admin123")
if ADMIN_PASSWORD == "admin123":
    app.logger.warning("SCM_ADMIN_PASSWORD is not configured; local fallback password is active. Set SCM_ADMIN_PASSWORD before production use.")
if CLOUD_BOOT_ERROR:
    app.logger.error("Cloud state hydration failed: %s", CLOUD_BOOT_ERROR)
APP_TZ = timezone(timedelta(hours=int(os.environ.get("SCM_UTC_OFFSET_HOURS", "8"))))

app.register_blueprint(aging_bp)


def _now_local() -> datetime:
    return datetime.now(APP_TZ)


def _persist_small_state(path: Path, payload: bytes) -> None:
    """Best-effort cloud mirror for small mutable JSON state.

    Local/runtime state is always the primary compatibility path. A transient or
    misconfigured cloud backend must not turn a valid user save into an error
    unless strict durable mode was explicitly requested.
    """
    global _cloud_runtime_error
    if not cloud_state.enabled:
        return
    try:
        cloud_state.persist_named_file(STATE_ROOT, path, payload=payload)
        _cloud_runtime_error = ""
    except Exception as exc:
        _cloud_runtime_error = str(exc)
        app.logger.warning("Optional cloud state mirror failed for %s", path, exc_info=True)
        if cloud_state.durable_required:
            raise


store = DashboardStore(
    None if EMPTY_STATE_MARKER.exists()
    else (ACTIVE_IMPORT if ACTIVE_IMPORT.exists() else DATA_FILE)
)
delivery_store = DeliveryStore(
    BASE / "data" / "delivery_master.json",
    DELIVERY_STATE_DIR,
    persist_callback=_persist_small_state if cloud_state.enabled else None,
)
delivery_store.sync_dashboard_branches(store.raw_records)
_management_allocations_cache = None
_bootstrap_cache = {}
_presentation_input_cache = {}
_management_lock = threading.RLock()
_cache_lock = threading.RLock()
_import_lock = threading.Lock()
_cloud_runtime_lock = threading.RLock()
_cloud_runtime_hydrated = CLOUD_BOOT_HYDRATED
_cloud_runtime_error = CLOUD_BOOT_ERROR

_MANAGEMENT_NUMERIC_FIELDS = frozenset({"quantity", "unit_cost", "inventory", "po_balance"})
_MANAGEMENT_TEXT_FIELDS = frozenset({"remarks"})


def _json_payload() -> dict:
    payload = request.get_json(silent=True)
    return payload if isinstance(payload, dict) else {}


def _kpi_period_diagnostics() -> dict:
    """Report how many imported period columns are active in each KPI view."""
    ytd_counts = [len((payload.get("ytd") or {}).get("labels") or []) for payload in store.kpis.values()]
    weekly_counts = [len((payload.get("weekly") or {}).get("labels") or []) for payload in store.kpis.values()]
    ytd_labels = next(((payload.get("ytd") or {}).get("labels") or [] for payload in store.kpis.values() if (payload.get("ytd") or {}).get("labels")), [])
    weekly_labels = next(((payload.get("weekly") or {}).get("labels") or [] for payload in store.kpis.values() if (payload.get("weekly") or {}).get("labels")), [])
    return {
        "ytd_periods": max(ytd_counts, default=0),
        "weekly_periods": max(weekly_counts, default=0),
        "latest_ytd": ytd_labels[-1] if ytd_labels else None,
        "latest_weekly": weekly_labels[-1] if weekly_labels else None,
    }


def _send_bytes(payload: bytes, filename: str, mimetype: str):
    return send_file(BytesIO(payload), mimetype=mimetype, as_attachment=True, download_name=filename)


def _persistent_storage_error(*, allow_recovery: bool = False):
    """Block mutations only when strict durable storage was explicitly required.

    v2.46.9 restores universal deployment compatibility: Local, Render and Vercel
    can import/save without external storage. Vercel Blob is an optional durability
    layer unless SCM_REQUIRE_DURABLE_STORAGE=1 is configured.
    """
    if not IS_VERCEL or not cloud_state.durable_required:
        return None
    if _cloud_runtime_error and not (allow_recovery and cloud_state.enabled):
        return jsonify({
            "error": "Strict durable storage could not be initialized for this request.",
            "detail": _cloud_runtime_error,
            "storage_required": True,
            "storage": cloud_state.status().provider,
            "auth_mode": cloud_state.auth_mode,
        }), 503
    if not cloud_state.enabled:
        status = cloud_state.status()
        return jsonify({
            "error": status.detail,
            "storage_required": True,
            "storage": status.provider,
            "auth_mode": status.auth_mode,
            "blob_store_id_present": bool(cloud_state.store_id),
            "oidc_request_token_present": cloud_state.request_oidc_present,
            "legacy_blob_token_present": bool(cloud_state.read_write_token),
        }), 503
    return None


def _save_temp_upload(file_storage, *, prefix: str, allowed_extensions: set[str]) -> Path:
    """Save an upload only to the OS temp area, never the deployed project tree."""
    ext = Path(file_storage.filename or "").suffix.lower()
    if ext not in allowed_extensions:
        allowed = ", ".join(sorted(allowed_extensions))
        raise ValueError(f"Unsupported file type. Allowed: {allowed}.")
    fd, raw_path = tempfile.mkstemp(prefix=prefix, suffix=ext, dir=IMPORT_TMP_DIR)
    os.close(fd)
    path = Path(raw_path)
    try:
        # Workbooks are capped by MAX_CONTENT_LENGTH, so buffering the upload here
        # is safe and avoids FileStorage.save() differences across serverless hosts.
        payload = file_storage.stream.read()
        if not payload:
            raise ValueError("The selected upload is empty.")
        path.write_bytes(payload)
    except Exception:
        path.unlink(missing_ok=True)
        raise
    return path


def _safe_unlink(path: Path | None) -> None:
    """Best-effort cleanup that must never turn a handled import error into HTTP 500."""
    if path is None:
        return
    try:
        path.unlink(missing_ok=True)
    except OSError:
        app.logger.warning("Could not remove temporary import artifact: %s", path, exc_info=True)


def _write_import_error_log(reference: str, stage: str, exc: BaseException) -> None:
    """Persist technical diagnostics locally without exposing workbook contents in the UI."""
    try:
        log_dir = STATE_ROOT / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        log_path = log_dir / "unified_import_errors.log"
        stamp = _now_local().isoformat(timespec="seconds")
        detail = traceback.format_exc()
        with log_path.open("a", encoding="utf-8") as handle:
            handle.write(f"\n[{stamp}] reference={reference} stage={stage}\n")
            handle.write(f"{type(exc).__name__}: {exc}\n")
            handle.write(detail.rstrip() + "\n")
    except OSError:
        app.logger.warning("Could not write unified import diagnostics", exc_info=True)


def _invalidate_bootstrap_cache():
    with _cache_lock:
        _bootstrap_cache.clear()
        _presentation_input_cache.clear()


@app.before_request
def _bind_vercel_oidc_and_hydrate_runtime():
    """Bind per-request Vercel OIDC and lazily hydrate durable state.

    Vercel exposes the runtime OIDC token in the x-vercel-oidc-token request
    header, so an OIDC-only Blob store cannot be hydrated safely at module import
    time. The first request performs hydration once per warm Python process.
    """
    global _cloud_runtime_hydrated, _cloud_runtime_error, _management_allocations_cache
    if not IS_VERCEL:
        return None

    cloud_state.bind_request_oidc(request.headers.get("x-vercel-oidc-token", ""))
    if not cloud_state.enabled:
        return None

    # Ensure writes performed by DeliveryStore during this request are durable,
    # including OIDC-only deployments where persistence was unavailable at import.
    delivery_store.set_persist_callback(_persist_small_state)

    if _cloud_runtime_hydrated:
        return None

    with _cloud_runtime_lock:
        if _cloud_runtime_hydrated:
            return None
        try:
            manifest = cloud_state.hydrate_core(STATE_ROOT)
            cloud_state.hydrate_named_files(STATE_ROOT, [
                "management_allocations.json",
                "delivery/delivery_master.json",
                "delivery/delivery_allocations.json",
                "delivery/delivery_schedule.json",
            ])

            if manifest:
                if EMPTY_STATE_MARKER.exists():
                    store.clear()
                elif ACTIVE_IMPORT.exists():
                    store.load(ACTIVE_IMPORT, prevalidated=True)

            # Named Delivery files may have just been restored from Blob. Reload
            # them before serving the request so the UI sees durable state.
            delivery_store.load()
            delivery_store.sync_dashboard_branches(store.raw_records)
            with _management_lock:
                _management_allocations_cache = None
            _invalidate_bootstrap_cache()
            _cloud_runtime_error = ""
            _cloud_runtime_hydrated = True
        except Exception as exc:
            _cloud_runtime_error = str(exc)
            app.logger.exception("Vercel Blob lazy hydration failed")
    return None


def _presentation_inputs():
    """Cache expensive Area/Branch drilldown preparation until the workbook changes."""
    key = store.generated_at.isoformat()
    with _cache_lock:
        cached = _presentation_input_cache.get(key)
        if cached is not None:
            return cached

    payload = {
        "area": store.area_dashboard("Overall"),
        "branches": [store.branch_dashboard(branch) for branch in store.branches],
    }
    with _cache_lock:
        _presentation_input_cache.clear()
        _presentation_input_cache[key] = payload
    return payload


def role():
    return "admin" if session.get("is_admin") else "guest"


def _load_management_allocations():
    """Load persisted Management planning edits with backward compatibility."""
    global _management_allocations_cache
    with _management_lock:
        if _management_allocations_cache is not None:
            return deepcopy(_management_allocations_cache)
        try:
            data = json.loads(MANAGEMENT_ALLOCATIONS.read_text(encoding="utf-8")) if MANAGEMENT_ALLOCATIONS.exists() else {}
            if not isinstance(data, dict):
                data = {}
            out = {}
            numeric_fields = {"quantity", "unit_cost", "inventory", "doi", "po_balance"}
            text_fields = {"remarks", "class", "brand", "model", "stock_status"}
            for key, value in data.items():
                try:
                    if isinstance(value, dict):
                        edit = {}
                        if "quantity" in value or "allocation" in value:
                            edit["quantity"] = max(0.0, float(value.get("quantity", value.get("allocation", 0)) or 0))
                        for field in numeric_fields - {"quantity"}:
                            if field in value:
                                edit[field] = max(0.0, float(value.get(field) or 0))
                        for field in text_fields:
                            if field in value:
                                edit[field] = str(value.get(field, "") or "").strip()
                        if "class" in edit:
                            edit["class"] = edit["class"].upper().replace("CLASS ", "").strip()
                    else:
                        edit = {"quantity": max(0.0, float(value or 0))}
                    out[str(key)] = edit
                except (TypeError, ValueError):
                    continue
            _management_allocations_cache = out
            return deepcopy(out)
        except (OSError, json.JSONDecodeError):
            _management_allocations_cache = {}
            return {}


def _save_management_allocations(data):
    global _management_allocations_cache
    snapshot = deepcopy(data)
    serialized = json.dumps(snapshot, indent=2, sort_keys=True).encode("utf-8")
    with _management_lock:
        # Always commit the writable local/runtime copy first. Cloud persistence is
        # an optional mirror so a storage-provider outage cannot break Order Plan.
        MANAGEMENT_ALLOCATIONS.parent.mkdir(parents=True, exist_ok=True)
        tmp = MANAGEMENT_ALLOCATIONS.with_suffix(".tmp")
        tmp.write_bytes(serialized)
        tmp.replace(MANAGEMENT_ALLOCATIONS)
        _management_allocations_cache = snapshot
        if cloud_state.enabled:
            _persist_small_state(MANAGEMENT_ALLOCATIONS, serialized)


def _merge_management_edits(base: dict, incoming: dict, valid_keys: set[str]) -> tuple[dict, int]:
    """Merge user-editable Management fields consistently for save and export."""
    if not isinstance(incoming, dict):
        raise ValueError("Management edits must be a key/value object.")
    merged = deepcopy(base)
    updated = 0
    allowed = _MANAGEMENT_NUMERIC_FIELDS | _MANAGEMENT_TEXT_FIELDS
    for raw_key, value in incoming.items():
        key = str(raw_key)
        if key not in valid_keys or not isinstance(value, dict):
            continue
        edit = {k: v for k, v in dict(merged.get(key) or {}).items() if k in allowed}
        try:
            for field in _MANAGEMENT_NUMERIC_FIELDS:
                if field in value:
                    edit[field] = max(0.0, float(value.get(field) or 0))
        except (TypeError, ValueError) as exc:
            raise ValueError(f"Invalid numeric Management value for {key}.") from exc
        for field in _MANAGEMENT_TEXT_FIELDS:
            if field in value:
                edit[field] = str(value.get(field, "") or "").strip()
        merged[key] = edit
        updated += 1
    return merged, updated


def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if IS_VERCEL and not SERVERLESS_SESSION_READY:
            return jsonify({
                "error": "Vercel Admin sessions are not configured for multi-instance use.",
                "detail": "Set SCM_SECRET_KEY (preferred) or SCM_ADMIN_PASSWORD in Vercel Project Settings > Environment Variables, then redeploy.",
                "configuration_required": True,
            }), 503
        if role() != "admin":
            return jsonify({
                "error": "Your Admin session is not active. Please sign in again.",
                "reauth_required": True,
                "login_url": "/login",
            }), 401
        return fn(*args, **kwargs)
    return wrapper


@app.get("/")
def index():
    return render_template("index.html", role=role())


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        if IS_VERCEL and not SERVERLESS_SESSION_READY:
            flash("Vercel deployment setup is incomplete. Add SCM_SECRET_KEY (preferred) or SCM_ADMIN_PASSWORD in Project Settings > Environment Variables, then redeploy.", "error")
            return render_template("login.html", deployment_ready=False), 503
        password = request.form.get("password", "")
        if hmac.compare_digest(password, ADMIN_PASSWORD):
            # Start a fresh authenticated session and keep it alive through the
            # configured work-session window. This also avoids carrying stale keys
            # from another local Flask application into the SCM admin session.
            session.clear()
            session.permanent = True
            session["is_admin"] = True
            session["authenticated_at"] = _now_local().isoformat()
            next_url = str(request.args.get("next") or request.form.get("next") or "").strip()
            if not next_url.startswith("/") or next_url.startswith("//"):
                next_url = url_for("index")
            return redirect(next_url)
        flash("Invalid admin password.", "error")
    return render_template("login.html", deployment_ready=SERVERLESS_SESSION_READY)


@app.post("/logout")
def logout():
    session.clear()
    return redirect(url_for("index"))


@app.get("/api/session")
def api_session():
    """Lightweight browser/session handshake used by admin workspaces."""
    is_admin = role() == "admin"
    return jsonify({
        "role": "admin" if is_admin else "guest",
        "is_admin": is_admin,
        "authenticated_at": session.get("authenticated_at") if is_admin else None,
        "deployment_ready": SERVERLESS_SESSION_READY,
        "configuration_required": bool(IS_VERCEL and not SERVERLESS_SESSION_READY),
        "session_key_source": SESSION_KEY_SOURCE,
    })


@app.get("/api/bootstrap")
def api_bootstrap():
    no_data = EMPTY_STATE_MARKER.exists()
    cache_key = (role(), store.generated_at.isoformat(), ACTIVE_IMPORT.exists(), no_data)
    with _cache_lock:
        cached = _bootstrap_cache.get(cache_key)
    if cached is None:
        cached = store.bootstrap(role())
        cached["data_source"] = "No Data" if no_data or not cached.get("has_data") else ("Saved Import" if ACTIVE_IMPORT.exists() else "Bundled Baseline")
        cached["has_saved_import"] = ACTIVE_IMPORT.exists()
        with _cache_lock:
            _bootstrap_cache.clear()
            _bootstrap_cache[cache_key] = cached
    return jsonify(cached)


@app.get("/api/area")
def api_area():
    return jsonify(store.area_dashboard(request.args.get("area", "Overall")))


@app.get("/api/branch")
def api_branch():
    branch = request.args.get("branch", "")
    return jsonify(store.branch_dashboard(branch))


@app.get("/api/status-summary")
def api_status_summary():
    return jsonify(store.status_summary(
        area=request.args.get("area", "Overall"),
        branch=request.args.get("branch", "All Branches"),
        brand=request.args.get("brand", "All Brands"),
        model=request.args.get("model", "All Models"),
        class_key=request.args.get("class", "All Classes"),
        status=request.args.get("status", "All Statuses"),
    ))


@app.get("/api/management")
def api_management():
    return jsonify(store.management_dashboard(
        brand=request.args.get("brand", "All Brands"),
        class_key=request.args.get("class", "All Classes"),
        status=request.args.get("status", "All Statuses"),
        model=request.args.get("model", "All Models"),
        allocations=_load_management_allocations(),
    ))


@app.post("/admin/management/allocations")
@admin_required
def admin_management_allocations():
    storage_error = _persistent_storage_error()
    if storage_error:
        return storage_error
    payload = _json_payload()
    incoming = payload.get("allocations") or {}
    try:
        current, updated = _merge_management_edits(
            _load_management_allocations(),
            incoming,
            {r["key"] for r in store.management_records},
        )
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    try:
        _save_management_allocations(current)
    except Exception as exc:
        app.logger.exception("Unable to persist Management Order Plan")
        return jsonify({"error": f"Unable to save Management Order Plan to persistent storage: {exc}"}), 503
    return jsonify({"ok": True, "updated": updated, "message": f"Saved {updated} Management planning line(s)."})


def _prepare_management_output(payload: dict) -> dict:
    incoming = payload.get("orders") or {}
    combined, _ = _merge_management_edits(
        _load_management_allocations(),
        incoming,
        {r["key"] for r in store.management_records},
    )

    visible_keys = payload.get("keys") or []
    if visible_keys and isinstance(visible_keys, list):
        # Build the same effective row set for Excel and direct print. This keeps
        # on-screen edits aligned with the user's current filtered view.
        data = store.management_dashboard(allocations=combined)
        by_key = {r.get("key"): r for r in data.get("rows", [])}
        data["rows"] = [by_key[k] for k in visible_keys if k in by_key]
        data["selected"] = {
            "brand": str(payload.get("brand") or "All Brands"),
            "class": str(payload.get("class") or "All Classes"),
            "status": str(payload.get("status") or "All Statuses"),
            "model": str(payload.get("model") or "All Models"),
        }
    else:
        data = store.management_dashboard(
            brand=str(payload.get("brand") or "All Brands"),
            class_key=str(payload.get("class") or "All Classes"),
            status=str(payload.get("status") or "All Statuses"),
            model=str(payload.get("model") or "All Models"),
            allocations=combined,
        )

    # Output only true order lines, matching the downloaded Excel workbook.
    ordered_rows = [r for r in data.get("rows", []) if float(r.get("allocation", 0) or 0) > 0]
    if not ordered_rows:
        raise ValueError("No models with Order Quantity greater than zero to output.")
    data["rows"] = ordered_rows
    data["summary"] = {
        "models": len(ordered_rows),
        "current_inventory": round(sum(float(r.get("inventory", 0) or 0) for r in ordered_rows), 4),
        "po_balance": round(sum(float(r.get("po_balance", 0) or 0) for r in ordered_rows), 4),
        "allocation_order": round(sum(float(r.get("allocation", 0) or 0) for r in ordered_rows), 4),
        "grand_total": round(sum(float(r.get("total_amount", 0) or 0) for r in ordered_rows), 4),
    }
    return data


def _whole_print(value) -> str:
    try:
        number = Decimal(str(value or 0)).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
        return f"{int(number):,}"
    except Exception:
        return html_escape(str(value or ""))


def _doi_print(value) -> str:
    try:
        return f"{int(math.ceil(max(0.0, float(value or 0)))):,}"
    except Exception:
        return html_escape(str(value or ""))


def _money_print(value) -> str:
    try:
        number = Decimal(str(value or 0)).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
        return f"₱{int(number):,}"
    except Exception:
        return "₱0"


def _direct_print_document(
    title: str,
    body: str,
    *,
    orientation: str = "portrait",
    page_margin: str = ".25in",
    preview_width: str = "8.5in",
) -> str:
    """Return a browser-print document styled to mirror the XLSX report.

    The direct print path deliberately uses the same palette, typography, row
    heights and proportional column widths as the downloadable workbook.  The
    browser only renders the final report; no XLSX is created or stored.
    """
    safe_title = html_escape(title)
    orientation = "portrait" if orientation == "portrait" else "landscape"
    safe_margin = page_margin if page_margin in {".08in", ".25in", ".28in"} else ".25in"
    safe_preview_width = "11in" if preview_width == "11in" else "8.5in"
    return f'''<!doctype html>
<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{safe_title}</title>
<style>
@page{{size:letter {orientation};margin:{safe_margin}}}
*{{box-sizing:border-box}}
html,body{{margin:0;padding:0;background:#fff;color:#0f172a;font-family:Aptos,"Segoe UI",Arial,sans-serif}}
body{{font-size:8pt;-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}}
.print-sheet{{width:100%;page-break-after:always;break-after:page}}
.print-sheet:last-child{{page-break-after:auto;break-after:auto}}
.xls-row{{display:flex;align-items:center;width:100%}}
.xls-title{{height:24pt;background:#0f172a;color:#fff;font-family:"Aptos Display",Aptos,"Segoe UI",Arial,sans-serif;font-size:14pt;font-weight:800;padding:0 4pt;letter-spacing:0}}
.xls-title.management{{height:23pt}}
.xls-subtitle{{height:18pt;background:#1e293b;color:#fbbf24;font-size:8pt;font-weight:800;padding:0 4pt}}
.xls-subtitle.management{{height:17pt;font-size:7.5pt}}
.xls-spacer-8{{height:8pt}}
.xls-spacer-5{{height:5pt}}
.xls-filter{{height:17pt;display:flex;align-items:center;color:#475569;background:#fff;font-size:7pt;padding:0 2pt}}
.xls-table{{width:100%;border-collapse:collapse;table-layout:fixed;border-spacing:0}}
.xls-table thead{{display:table-header-group}}
.xls-table tr{{page-break-inside:avoid;break-inside:avoid}}
.xls-table th,.xls-table td{{font-variant-numeric:tabular-nums;vertical-align:middle}}
.xls-table th{{height:25pt;background:#0f172a;color:#fff;border:1px solid #e2e8f0;padding:2pt 2.5pt;font-size:7pt;font-weight:800;text-align:center;white-space:normal;line-height:1.12}}
.branch-table th{{height:22pt;font-size:8pt}}
.xls-table tbody tr.data-row{{height:24pt}}
.xls-table tbody tr.data-row td{{border:1px solid #e2e8f0;padding:2.5pt 2.5pt;font-size:7.5pt;font-weight:700;background:#f8fafc;line-height:1.15;overflow-wrap:anywhere}}
.xls-table tbody tr.data-row.even td{{background:#f1f5f9}}
.branch-table tbody tr.data-row td{{font-size:8.5pt;font-weight:800}}
.xls-table td.text-left{{text-align:left}}
.xls-table td.text-center{{text-align:center}}
.xls-table td.text-right{{text-align:right}}
.xls-table td.wrap{{white-space:normal;overflow-wrap:anywhere}}
.xls-table td.status-cell{{text-align:center;background:#fffbeb!important;color:#0f172a;font-weight:800}}
.xls-table td.remarks-cell{{font-size:7pt!important;font-weight:400!important;white-space:normal;overflow-wrap:anywhere}}
.xls-table tr.grand-total{{height:22pt}}
.xls-table tr.grand-total td{{font-size:7pt;font-weight:800}}
.xls-table td.grand-label{{background:#1e293b;color:#fff;text-align:right;border:1px solid #e2e8f0;padding:2.5pt}}
.xls-table td.grand-qty{{background:#1e293b;color:#fff;text-align:center;border:1px solid #e2e8f0;padding:2.5pt}}
.xls-table td.grand-amount{{background:#1e293b;color:#fff;text-align:right;border:1px solid #e2e8f0;padding:2.5pt}}
.xls-table td.grand-blank{{background:#fff;border:0;padding:0}}
.branch-meta{{width:100%;border-collapse:collapse;table-layout:fixed;border-spacing:0}}
.branch-meta td{{height:18pt;border:1px solid #e2e8f0;background:#1e293b;padding:2pt 3pt;vertical-align:middle}}
.branch-meta .meta-label{{color:#fbbf24;font-size:8pt;font-weight:800;text-transform:uppercase}}
.branch-meta .meta-value{{color:#fff;font-size:8.5pt;font-weight:800}}
@media screen{{
 body{{padding:18px;background:#e2e8f0}}
 .print-sheet{{background:#fff;max-width:{safe_preview_width};margin:0 auto 18px;box-shadow:0 12px 30px rgba(15,23,42,.14)}}
}}
@media print{{body{{background:#fff}}.print-sheet{{box-shadow:none;margin:0}}}}
</style></head><body>{body}<script>window.addEventListener("load",()=>setTimeout(()=>window.print(),160));</script></body></html>'''


def _management_print_html(data: dict) -> str:
    selected = data.get("selected") or {}
    title = html_escape(str(data.get("title") or "Management Order Plan"))
    grouped = {}
    for row in data.get("rows", []):
        brand = str(row.get("brand") or "Unspecified").strip() or "Unspecified"
        grouped.setdefault(brand, []).append(row)

    # Same relative widths as build_management_order_xlsx().
    excel_widths = [10.0, 13.0, 18.0, 15.0, 9.0, 8.0, 13.0, 18.0, 12.0, 10.0, 19.0, 30.0]
    width_total = sum(excel_widths)
    widths = [f"{(w / width_total) * 100:.4f}%" for w in excel_widths]
    headers = ["LINE NO.", "BRAND", "MODEL", "UNIT COST", "INV.", "DoI", "STOCK STATUS", "PO BAL.", "ORDER QTY", "NEW DoI", "TOTAL AMOUNT", "REMARKS"]
    colgroup = '<colgroup>' + ''.join(f'<col style="width:{w}">' for w in widths) + '</colgroup>'

    sections = []
    for brand in sorted(grouped, key=str.lower):
        rows = grouped[brand]
        filters = [f"Brand: {brand}"]
        if selected.get("class") and selected.get("class") != "All Classes":
            filters.append(f"Class: {selected.get('class')}")
        if selected.get("status") and selected.get("status") != "All Statuses":
            filters.append(f"Status: {selected.get('status')}")
        if selected.get("model") and selected.get("model") != "All Models":
            filters.append(f"Model: {selected.get('model')}")

        body_rows = []
        qty_total = 0.0
        amount_total = 0.0
        for idx, row in enumerate(rows, 1):
            qty = float(row.get("allocation", 0) or 0)
            amount = float(row.get("total_amount", 0) or 0)
            qty_total += qty
            amount_total += amount
            values = [
                _whole_print(idx),
                html_escape(str(row.get("brand") or "")),
                html_escape(str(row.get("model") or "")),
                _money_print(row.get("unit_cost", 0)),
                _whole_print(row.get("inventory", 0)),
                _doi_print(row.get("doi", 0)),
                html_escape(str(row.get("stock_status") or "")),
                _whole_print(row.get("po_balance", 0)),
                _whole_print(qty),
                _doi_print(row.get("new_doi", 0)),
                _money_print(amount),
                html_escape(str(row.get("remarks") or "")),
            ]
            aligns = [
                "text-center", "text-left", "text-left", "text-right",
                "text-center", "text-center", "status-cell", "text-center",
                "text-center", "text-center", "text-right", "text-left remarks-cell",
            ]
            parity = " even" if idx % 2 == 0 else ""
            body_rows.append(
                f'<tr class="data-row{parity}">'
                + ''.join(f'<td class="{align}">{value}</td>' for align, value in zip(aligns, values))
                + '</tr>'
            )

        total_row = (
            f'<tr class="grand-total">'
            f'<td colspan="8" class="grand-label">GRAND TOTAL</td>'
            f'<td class="grand-qty">{_whole_print(qty_total)}</td>'
            f'<td class="grand-blank"></td>'
            f'<td class="grand-amount">{_money_print(amount_total)}</td>'
            f'<td class="grand-blank"></td>'
            f'</tr>'
        )
        filter_line = html_escape(" | ".join(filters))
        header_html = ''.join(f'<th>{html_escape(h)}</th>' for h in headers)
        sections.append(
            '<section class="print-sheet">'
            '<div class="xls-row xls-title management">MANAGEMENT ORDER PLAN</div>'
            f'<div class="xls-row xls-subtitle management">{title}</div>'
            '<div class="xls-spacer-5"></div>'
            f'<div class="xls-filter">{filter_line}</div>'
            '<div class="xls-spacer-5"></div>'
            f'<table class="xls-table management-table">{colgroup}<thead><tr>{header_html}</tr></thead>'
            f'<tbody>{"".join(body_rows)}{total_row}</tbody></table>'
            '</section>'
        )

    # Management workbook is configured as Letter portrait + fit-to-width.
    return _direct_print_document(
        "Management Order Plan",
        ''.join(sections),
        orientation="portrait",
        page_margin=".08in",
        preview_width="8.5in",
    )


@app.post("/admin/export/management")
@admin_required
def admin_export_management():
    try:
        data = _prepare_management_output(_json_payload())
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    xlsx = build_management_order_xlsx(data)
    stamp = _now_local().strftime("%Y%m%d_%H%M%S")
    filename = f"Management_Order_Plan_{stamp}.xlsx"
    return _send_bytes(xlsx, filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.post("/admin/print/management")
@admin_required
def admin_print_management():
    try:
        data = _prepare_management_output(_json_payload())
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    response = app.response_class(_management_print_html(data), mimetype="text/html")
    response.headers["Cache-Control"] = "no-store"
    return response


@app.get("/api/delivery")
def api_delivery():
    day = request.args.get("day", "Whole Week")
    return jsonify(delivery_store.bootstrap(day, store.raw_records))


@app.get("/delivery/template")
def delivery_template():
    path = BASE / "data" / "Delivery_Allocation_Import_Template.xlsx"
    return send_file(path, as_attachment=True, download_name="Unit_Allocation_Template.xlsx")


@app.get("/api/branch-models")
def api_branch_models():
    branch = request.args.get("branch", "")
    return jsonify({"branch": branch, "models": store.branch_models(branch)})


@app.get("/api/model")
def api_model():
    branch = request.args.get("branch", "")
    model = request.args.get("model", "")
    item = store.model_lookup(branch, model)
    if not item:
        return jsonify({"error": "Model not found for selected branch."}), 404
    return jsonify(item)


@app.post("/admin/import")
@admin_required
def admin_import():
    """Transactionally refresh SCM, Management and Aging from one workbook.

    v2.46.9 makes deployment storage provider-independent. The workbook is
    validated and committed to the writable runtime first. If durable cloud
    storage is available it is mirrored as an additional layer; if it is absent
    or temporarily unavailable, the import still completes in runtime-fallback
    mode unless strict durable storage was explicitly requested.
    """
    global _management_allocations_cache, _cloud_runtime_error, _cloud_runtime_hydrated

    reference = f"IMP-{_now_local().strftime('%Y%m%d-%H%M%S')}-{secrets.token_hex(2).upper()}"
    stage = "receiving upload"
    tmp: Path | None = None
    active_backup: Path | None = None
    aging_backup: Path | None = None
    aging_cloud_snapshot: Path | None = None
    warnings: list[str] = []
    cloud_ready = bool(cloud_state.enabled)

    def fail(message: str, status: int = 400, *, error_stage: str | None = None):
        return jsonify({
            "error": message,
            "stage": error_stage or stage,
            "reference": reference,
            "previous_data_active": True,
        }), status

    try:
        # Only strict deployments block on missing external persistence. Normal
        # deployments continue with the universal writable runtime fallback.
        storage_error = _persistent_storage_error(allow_recovery=True)
        if storage_error:
            response, status = storage_error
            payload = response.get_json(silent=True) or {}
            detail = str(payload.get("detail") or "").strip()
            message = str(payload.get("error") or "Strict durable storage is unavailable.")
            if detail and detail not in message:
                message = f"{message} {detail}"
            return fail(message, status, error_stage="checking strict durable storage")

        # Blob is an optional mirror in the standard deployment mode. Probe it
        # when present, but never reject a valid workbook because a provider is
        # missing or temporarily unavailable.
        if IS_VERCEL and cloud_ready:
            stage = "checking optional cloud persistence"
            try:
                cloud_state.probe()
                _cloud_runtime_error = ""
            except Exception as exc:
                _write_import_error_log(reference, stage, exc)
                _cloud_runtime_error = str(exc)
                if cloud_state.durable_required:
                    return fail(f"Strict durable storage check failed: {exc}", 503)
                cloud_ready = False
                warnings.append("Cloud persistence was unavailable, so this refresh was committed to the Vercel runtime fallback.")

        stage = "receiving upload"
        uploaded = request.files.get("file")
        if not uploaded:
            return fail("Please select the consolidated .xlsx import workbook.")
        original_name = uploaded.filename or "SCM_Import.xlsx"

        try:
            tmp = _save_temp_upload(uploaded, prefix="unified_import_", allowed_extensions={".xlsx"})
        except ValueError as exc:
            return fail(str(exc))
        except OSError as exc:
            _write_import_error_log(reference, stage, exc)
            return fail(f"The workbook could not be staged in the writable runtime area: {exc}", 500)

        if not zipfile.is_zipfile(tmp):
            return fail("The selected file is not a valid .xlsx workbook. Re-save it as Excel Workbook (*.xlsx) and retry.")

        with _import_lock:
            token = reference.replace(":", "-")
            active_backup = UPLOAD_DIR / f".active_import_backup_{token}.xlsx"
            aging_backup = AGING_DB_PATH.with_name(f".aging_backup_{token}.db")
            aging_cloud_snapshot = AGING_DB_PATH.with_name(f".aging_cloud_{token}.db")
            previous_cloud_manifest = None
            if cloud_ready:
                try:
                    previous_cloud_manifest = cloud_state.get_json(cloud_state.core_manifest_key)
                except Exception as exc:
                    _cloud_runtime_error = str(exc)
                    app.logger.warning("Previous cloud manifest could not be read during recovery import [%s]", reference, exc_info=True)
                    # New cloud publication may still repair the manifest, so this
                    # warning alone does not disable the mirror.
                    warnings.append("The previous cloud revision could not be read; the new import will attempt to repair it.")

            stage = "validating SCM workbook"
            check = DashboardStore.validate(tmp)
            if not check.ok:
                return fail(check.message)

            stage = "validating Motorcycle Aging"
            try:
                aging_preflight = validate_unified_aging(tmp, "Aging")
            except Exception as exc:
                return fail(f"Aging validation failed: {exc}")

            stage = "staging SCM, KPI and Management data"
            candidate_store = DashboardStore(None)
            try:
                candidate_store.load(tmp, prevalidated=True)
            except Exception as exc:
                return fail(f"SCM/KPI staging failed: {exc}")

            active = ACTIVE_IMPORT
            had_active = active.exists()
            had_aging_db = AGING_DB_PATH.exists()
            workbook_committed = False
            aging_committed = False
            cloud_published = False

            try:
                stage = "creating rollback snapshot"
                if had_active:
                    shutil.copy2(active, active_backup)
                if had_aging_db:
                    backup_aging_database(aging_backup)

                stage = "refreshing Motorcycle Aging"
                aging_result = import_aging_excel(
                    tmp,
                    original_name,
                    mode="replace",
                    sheet_name=aging_preflight.get("sheet") or "Aging",
                )
                aging_committed = True

                # A consistent SQLite snapshot is needed only when the optional
                # cloud mirror is actually available for this request.
                if cloud_ready:
                    stage = "preparing optional cloud snapshot"
                    try:
                        backup_aging_database(aging_cloud_snapshot)
                    except Exception as exc:
                        _cloud_runtime_error = str(exc)
                        if cloud_state.durable_required:
                            raise
                        app.logger.warning("Cloud Aging snapshot skipped [%s]", reference, exc_info=True)
                        warnings.append("Cloud Aging snapshot was skipped; runtime data remains active.")
                        cloud_ready = False

                stage = "committing unified workbook"
                os.replace(tmp, active)
                tmp = None
                workbook_committed = True

                # Cloud publication is additive, not a prerequisite for a valid
                # import. The manifest is still written last by publish_core().
                if cloud_ready:
                    stage = "mirroring control-tower state to cloud"
                    try:
                        cloud_state.publish_core(active, aging_cloud_snapshot, reference)
                        cloud_published = True
                        _cloud_runtime_error = ""
                    except Exception as exc:
                        _cloud_runtime_error = str(exc)
                        _write_import_error_log(reference, stage, exc)
                        if cloud_state.durable_required:
                            raise
                        app.logger.warning("Optional cloud publication failed [%s]", reference, exc_info=True)
                        warnings.append("Cloud persistence failed, but the validated import remains active in runtime storage.")
                        cloud_ready = False

                stage = "publishing dashboard state"
                store.adopt_from(candidate_store, active)
                _safe_unlink(EMPTY_STATE_MARKER)
                _cloud_runtime_hydrated = True

            except Exception as exc:
                app.logger.exception("Unified import failed at %s [%s]", stage, reference)
                _write_import_error_log(reference, stage, exc)

                try:
                    if workbook_committed:
                        if had_active and active_backup and active_backup.exists():
                            shutil.copy2(active_backup, active)
                        else:
                            _safe_unlink(active)
                except Exception:
                    app.logger.exception("SCM workbook rollback failed [%s]", reference)

                try:
                    if aging_committed:
                        if had_aging_db and aging_backup and aging_backup.exists():
                            restore_aging_database(aging_backup)
                        else:
                            clear_aging_data()
                except Exception:
                    app.logger.exception("Aging database rollback failed [%s]", reference)

                if cloud_published:
                    try:
                        if previous_cloud_manifest:
                            cloud_state.put_json(cloud_state.core_manifest_key, previous_cloud_manifest)
                        else:
                            cloud_state.delete(cloud_state.core_manifest_key)
                    except Exception:
                        app.logger.exception("Cloud core manifest rollback failed [%s]", reference)

                _invalidate_bootstrap_cache()
                return fail(f"Unified refresh stopped during {stage}: {exc}", 400)

            stage = "finalizing workspace"
            try:
                _safe_unlink(MANAGEMENT_ALLOCATIONS)
                with _management_lock:
                    _management_allocations_cache = {}
                if cloud_published:
                    try:
                        cloud_state.delete_named_file(STATE_ROOT, MANAGEMENT_ALLOCATIONS)
                    except Exception as exc:
                        _cloud_runtime_error = str(exc)
                        app.logger.warning("Optional Management cloud reset warning [%s]", reference, exc_info=True)
                        warnings.append("Management planning cloud reset could not be mirrored; runtime reset succeeded.")
            except Exception:
                app.logger.warning("Management planning reset warning [%s]", reference, exc_info=True)

            try:
                delivery_store.sync_dashboard_branches(store.raw_records)
            except Exception:
                app.logger.warning("Delivery branch synchronization warning [%s]", reference, exc_info=True)

            _invalidate_bootstrap_cache()

            persistence_mode = "local-filesystem"
            persistence_durable = not IS_VERCEL
            if IS_VERCEL:
                if cloud_published:
                    persistence_mode = "durable-cloud"
                    persistence_durable = True
                else:
                    persistence_mode = "runtime-fallback"
                    persistence_durable = False
                    if not any("runtime" in item.lower() for item in warnings):
                        warnings.append("Vercel runtime fallback is active. Import succeeded without external storage; data may reset after a cold start or redeployment.")

            message = (
                "Unified import complete — Executive/KPI, Reorder/Management and Motorcycle Aging refreshed from one workbook. "
                f"Aging: {aging_result['rows']:,} units · {aging_result['branches']} branches · "
                f"{aging_result['areas']} areas · as of {aging_result['as_of_date']}."
            )
            return jsonify({
                "ok": True,
                "message": message,
                "reference": reference,
                "generated_at": store.generated_at.isoformat(),
                "persistence": {
                    "mode": persistence_mode,
                    "durable": persistence_durable,
                    "provider": cloud_state.status().provider if IS_VERCEL else "local-filesystem",
                },
                "warnings": warnings,
                "modules": {
                    "scm": {"status": "updated", "records": len(store.raw_records)},
                    "kpi": {"status": "updated", **_kpi_period_diagnostics()},
                    "management": {"status": "updated", "source": store.management_sheet_name or "Not detected"},
                    "aging": {"status": "updated", **aging_result},
                },
            })

    except Exception as exc:
        app.logger.exception("Unexpected unified import error at %s [%s]", stage, reference)
        _write_import_error_log(reference, stage, exc)
        return fail(
            "The unified refresh encountered an unexpected server error. "
            "The previous data remains active; use the reference below for diagnostics.",
            500,
        )
    finally:
        _safe_unlink(tmp)
        _safe_unlink(active_backup)
        _safe_unlink(aging_backup)
        _safe_unlink(aging_cloud_snapshot)


@app.get("/admin/export/pptx")
@admin_required
def admin_export_pptx():
    # v2.18 PowerPoint export is a branded KPI review deck: YTD, Weekly,
    # All-Area Performance, Branch rankings and per-Branch A/B/C model details.
    data = store.bootstrap(role())
    data["export_generated_at"] = _now_local().isoformat()
    # Include the integrated Motorcycle Aging executive snapshot only when
    # Aging data is actually loaded. A global Clear Data therefore produces
    # no Aging slides and no stale values in the management deck.
    data["aging_summary"] = aging_executive_summary() if aging_record_count() > 0 else None
    prepared = _presentation_inputs()
    all_branch_dashboards = prepared["branches"]
    reorder_brand = request.args.get("reorder_brand", "All Brands")
    first_branch = all_branch_dashboards[0] if all_branch_dashboards else {}
    ppt = build_presentation(data, prepared["area"], first_branch, all_branch_dashboards, reorder_brand=reorder_brand)
    stamp = _now_local().strftime("%Y%m%d_%H%M%S")
    filename = f"SCM_IDP_Executive_Control_Tower_{stamp}.pptx"
    return _send_bytes(ppt, filename, "application/vnd.openxmlformats-officedocument.presentationml.presentation")


def _prepare_branch_request_output(payload: dict) -> tuple[str, str, list]:
    branch = str(payload.get("branch", "")).strip()
    items = payload.get("items") or []
    if not branch or not items:
        raise ValueError("Select a Branch and include at least one requested item.")
    verified = []
    area = ""
    for item in items:
        model = str(item.get("model", "")).strip()
        base = store.model_lookup(branch, model)
        if not base:
            continue
        try:
            qty = float(item.get("requested_qty", 0) or 0)
        except (TypeError, ValueError):
            continue
        if qty <= 0:
            continue
        new_doi = ((base["inventory"] + qty) / base["avg_daily_sale"]) if base["avg_daily_sale"] > 0 else "N/A"
        area = base["area"]
        verified.append({**base, "requested_qty": qty, "remarks": str(item.get("remarks", "")), "new_doi": round(new_doi, 2) if isinstance(new_doi, float) else new_doi})
    if not verified:
        raise ValueError("No valid request lines were found for the selected Branch.")
    return branch, area, verified


def _branch_request_print_html(branch: str, area: str, items: list) -> str:
    # Direct print uses the same worksheet structure as build_branch_request_xlsx(),
    # except the requested print-only simplifications: date only and no signature block.
    generated = _now_local().strftime("%d %b %Y")
    headers = ["NO.", "MODEL", "CLASS", "INVENTORY", "REQUEST QTY", "STOCK STATUS", "CURRENT DoI", "NEW DoI", "REMARKS"]

    # Same relative widths as the downloadable Excel workbook.
    excel_widths = [5.5, 20.0, 8.0, 9.0, 10.0, 13.0, 10.0, 10.0, 22.0]
    width_total = sum(excel_widths)
    widths = [f"{(w / width_total) * 100:.4f}%" for w in excel_widths]
    colgroup = '<colgroup>' + ''.join(f'<col style="width:{w}">' for w in widths) + '</colgroup>'

    body_rows = []
    for idx, item in enumerate(items, 1):
        new_doi = item.get("new_doi", "N/A")
        values = [
            _whole_print(idx),
            html_escape(str(item.get("model") or "")),
            html_escape(str(item.get("class") or "")),
            _whole_print(item.get("inventory", 0)),
            _whole_print(item.get("requested_qty", 0)),
            html_escape(str(item.get("stock_status") or "")),
            _doi_print(item.get("doi", 0)),
            _doi_print(new_doi) if isinstance(new_doi, (int, float)) else html_escape(str(new_doi)),
            html_escape(str(item.get("remarks") or "")),
        ]
        aligns = [
            "text-right", "text-left", "text-left", "text-right", "text-right",
            "text-left", "text-right", "text-right", "text-left wrap",
        ]
        parity = " even" if idx % 2 == 0 else ""
        body_rows.append(
            f'<tr class="data-row{parity}">'
            + ''.join(f'<td class="{align}">{value}</td>' for align, value in zip(aligns, values))
            + '</tr>'
        )

    # Match Excel merged-cell geometry: A:B label, C:E value, F:G label, H:I value.
    meta = f'''<table class="branch-meta">{colgroup}<tbody>
<tr>
<td colspan="2" class="meta-label">REQUESTING BRANCH</td>
<td colspan="3" class="meta-value">{html_escape(branch)}</td>
<td colspan="2" class="meta-label">REPORT GENERATED</td>
<td colspan="2" class="meta-value">{html_escape(generated)}</td>
</tr>
<tr>
<td colspan="2" class="meta-label">AREA</td>
<td colspan="3" class="meta-value">{html_escape(area)}</td>
<td colspan="4" style="background:#fff;border:0"></td>
</tr>
</tbody></table>'''
    header_html = ''.join(f'<th>{html_escape(h)}</th>' for h in headers)
    section = (
        '<section class="print-sheet">'
        '<div class="xls-row xls-title">BRANCH REQUEST STATUS REPORT</div>'
        '<div class="xls-row xls-subtitle">MUTI MC SCM Executive Control Tower • Branch Request Report</div>'
        '<div class="xls-spacer-8"></div>'
        f'{meta}'
        '<div class="xls-spacer-8"></div>'
        f'<table class="xls-table branch-table">{colgroup}<thead><tr>{header_html}</tr></thead>'
        f'<tbody>{"".join(body_rows)}</tbody></table>'
        '</section>'
    )
    return _direct_print_document(
        f"Branch Request - {branch}",
        section,
        orientation="portrait",
        page_margin=".25in",
        preview_width="8.5in",
    )


@app.post("/admin/export/request")
@admin_required
def admin_export_request():
    try:
        branch, area, verified = _prepare_branch_request_output(_json_payload())
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    xlsx = build_branch_request_xlsx(branch, area, verified)
    stamp = _now_local().strftime("%Y%m%d_%H%M%S")
    safe_branch = "_".join(branch.split())
    filename = f"Branch_Request_{safe_branch}_{stamp}.xlsx"
    return _send_bytes(xlsx, filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.post("/admin/print/request")
@admin_required
def admin_print_request():
    try:
        branch, area, verified = _prepare_branch_request_output(_json_payload())
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    response = app.response_class(_branch_request_print_html(branch, area, verified), mimetype="text/html")
    response.headers["Cache-Control"] = "no-store"
    return response


@app.post("/admin/delivery/import")
@admin_required
def admin_delivery_import():
    storage_error = _persistent_storage_error()
    if storage_error:
        return storage_error
    uploaded = request.files.get("file")
    if not uploaded:
        return jsonify({"error": "Select an allocation file."}), 400
    try:
        tmp = _save_temp_upload(uploaded, prefix="delivery_allocation_", allowed_extensions={".xlsx", ".xlsm", ".csv"})
    except ValueError:
        return jsonify({"error": "Allocation import accepts .xlsx, .xlsm or .csv."}), 400
    try:
        rows, warnings = delivery_store.import_allocations(tmp)
        return jsonify({"ok": True, "rows": len(rows), "warnings": warnings[:20], "message": f"Imported {len(rows)} allocation rows."})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400
    finally:
        tmp.unlink(missing_ok=True)


@app.post("/admin/delivery/master")
@admin_required
def admin_delivery_master():
    storage_error = _persistent_storage_error()
    if storage_error:
        return storage_error
    payload = _json_payload()
    try:
        master = delivery_store.update_master(str(payload.get("section", "")), payload.get("rows"))
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400
    return jsonify({"ok": True, "master": master})


@app.post("/admin/delivery/schedule/import")
@admin_required
def admin_delivery_schedule_import():
    storage_error = _persistent_storage_error()
    if storage_error:
        return storage_error
    uploaded = request.files.get("file")
    if not uploaded:
        return jsonify({"error": "Select a Weekly Truck Schedule file."}), 400
    try:
        tmp = _save_temp_upload(uploaded, prefix="weekly_schedule_", allowed_extensions={".xlsx", ".xlsm", ".csv"})
    except ValueError:
        return jsonify({"error": "Weekly Schedule import accepts .xlsx, .xlsm or .csv."}), 400
    try:
        schedule, warnings = delivery_store.import_schedule(tmp)
        return jsonify({"ok": True, "rows": len(schedule), "warnings": warnings[:30], "message": f"Imported and saved {len(schedule)} Weekly Truck Schedule row(s)."})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400
    finally:
        tmp.unlink(missing_ok=True)


@app.post("/admin/delivery/schedule")
@admin_required
def admin_delivery_schedule():
    storage_error = _persistent_storage_error()
    if storage_error:
        return storage_error
    payload = _json_payload()
    try:
        schedule = delivery_store.update_schedule(payload.get("rows") or [])
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400
    return jsonify({"ok": True, "schedule": schedule})


@app.post("/admin/delivery/plan")
@admin_required
def admin_delivery_plan_save():
    storage_error = _persistent_storage_error()
    if storage_error:
        return storage_error
    """Persist the current weekly truck schedule and allocation plan together."""
    payload = _json_payload()
    try:
        schedule = delivery_store.update_schedule(payload.get("schedule") or [])
        allocations = delivery_store.replace_allocations(payload.get("allocations") or [])
    except Exception as exc:
        return jsonify({"error": f"Unable to save Delivery Plan: {exc}"}), 400
    return jsonify({
        "ok": True,
        "schedule": schedule,
        "allocations": allocations,
        "message": f"Delivery Plan saved: {len(schedule)} schedule trip(s) and {len(allocations)} allocation row(s).",
    })


@app.post("/admin/delivery/allocations")
@admin_required
def admin_delivery_allocations():
    storage_error = _persistent_storage_error()
    if storage_error:
        return storage_error
    payload = _json_payload()
    try:
        allocations = delivery_store.replace_allocations(payload.get("rows") or [])
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400
    return jsonify({"ok": True, "allocations": allocations})


@app.patch("/admin/delivery/allocation/<int:index>")
@admin_required
def admin_delivery_allocation_update(index: int):
    storage_error = _persistent_storage_error()
    if storage_error:
        return storage_error
    payload = _json_payload()
    try:
        allocations = delivery_store.update_allocation(index, payload)
    except (ValueError, IndexError) as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": f"Unable to update allocation: {exc}"}), 400
    return jsonify({"ok": True, "allocations": allocations, "message": "Allocation updated. Delivery capacity has been recalculated."})


@app.post("/admin/delivery/allocation/<int:index>/transfer")
@admin_required
def admin_delivery_allocation_transfer(index: int):
    storage_error = _persistent_storage_error()
    if storage_error:
        return storage_error
    payload = _json_payload()
    try:
        allocations = delivery_store.transfer_allocation(
            index,
            payload.get("target_day"),
            payload.get("target_plate"),
            payload.get("quantity"),
        )
    except (ValueError, IndexError) as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": f"Unable to transfer allocation: {exc}"}), 400
    return jsonify({"ok": True, "allocations": allocations, "message": "Allocation transferred. Delivery capacity has been rebalanced."})


@app.delete("/admin/delivery/allocation/<int:index>")
@admin_required
def admin_delivery_allocation_delete(index: int):
    storage_error = _persistent_storage_error()
    if storage_error:
        return storage_error
    try:
        allocations = delivery_store.delete_allocation(index)
    except (ValueError, IndexError) as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"error": f"Unable to delete allocation: {exc}"}), 400
    return jsonify({"ok": True, "allocations": allocations, "message": "Allocation deleted. Delivery capacity has been recalculated."})



@app.post("/admin/clear-data")
@admin_required
def admin_clear_data():
    storage_error = _persistent_storage_error()
    if storage_error:
        return storage_error
    """Destructively clear all user-loaded SCM and Motorcycle Aging data.

    Runtime state is always cleared first. If optional cloud persistence is
    connected, the empty marker is mirrored there; cloud failure is non-blocking
    unless strict durable storage was explicitly requested.
    """
    global _management_allocations_cache, _cloud_runtime_error
    payload = _json_payload()
    confirmation = str(payload.get("confirmation") or "").strip().upper()
    if confirmation != "CLEAR DATA":
        return jsonify({"error": "Type CLEAR DATA exactly to confirm the reset."}), 400

    ACTIVE_IMPORT.unlink(missing_ok=True)
    (UPLOAD_DIR / "candidate_import.xlsx").unlink(missing_ok=True)
    MANAGEMENT_ALLOCATIONS.unlink(missing_ok=True)
    _management_allocations_cache = {}

    store.clear()
    EMPTY_STATE_MARKER.write_text(_now_local().isoformat(), encoding="utf-8")

    delivery_store.reset_to_defaults()
    delivery_store.sync_dashboard_branches([])
    clear_aging_data()
    _invalidate_bootstrap_cache()

    for legacy_dir in (STATE_ROOT / "exports", BASE / "exports"):
        if legacy_dir.exists() and legacy_dir.is_dir():
            shutil.rmtree(legacy_dir, ignore_errors=True)

    warnings = []
    reset_reference = f"CLEAR-{_now_local().strftime('%Y%m%d-%H%M%S')}-{secrets.token_hex(2).upper()}"
    if cloud_state.enabled:
        try:
            cloud_state.publish_empty_core(reset_reference)
            cloud_state.delete_named_file(STATE_ROOT, MANAGEMENT_ALLOCATIONS)
            _cloud_runtime_error = ""
        except Exception as exc:
            _cloud_runtime_error = str(exc)
            app.logger.warning("Optional cloud Clear Data mirror failed", exc_info=True)
            if cloud_state.durable_required:
                return jsonify({"error": f"Runtime data was cleared, but strict durable storage could not be updated: {exc}"}), 503
            warnings.append("Runtime data was cleared successfully; the optional cloud mirror could not be updated.")

    return jsonify({
        "ok": True,
        "message": "All imported and saved runtime data was cleared. Dashboards are now empty until new files are imported.",
        "records": 0,
        "aging_records": 0,
        "warnings": warnings,
    })


@app.post("/admin/delivery/clear")
@admin_required
def admin_delivery_clear():
    """Clear transient delivery editing data without deleting a saved weekly plan.

    v2.35 deliberately treats the top-level Clear Board action as a UI reset only.
    Persisted Weekly Truck Schedule + Allocation rows are the saved weekly recovery point
    and must survive Clear Board. The explicit allocation-only action remains available
    for users who intentionally want to delete saved allocation rows.
    """
    payload = _json_payload()
    target = str(payload.get("target", "all")).strip().lower()
    if target not in {"all", "allocations"}:
        return jsonify({"error": "Clear target must be all or allocations."}), 400
    if target == "all":
        return jsonify({
            "ok": True,
            "target": "all",
            "preserved": True,
            "saved_schedule_rows": len(delivery_store.schedule),
            "saved_allocation_rows": len(delivery_store.allocations),
            "message": "Working Delivery Board cleared. The saved Weekly Truck Schedule and saved Allocation for the week were preserved and can be opened again.",
        })
    storage_error = _persistent_storage_error()
    if storage_error:
        return storage_error
    try:
        delivery_store.replace_allocations([])
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400
    return jsonify({
        "ok": True,
        "target": "allocations",
        "message": "All saved delivery allocation rows cleared. Weekly Truck Schedule and masterlists were preserved.",
    })


@app.get("/admin/export/delivery")
@admin_required
def admin_export_delivery():
    day = request.args.get("day", "Whole Week")
    if day not in DAYS and day != "Whole Week":
        day = "Whole Week"
    plan = delivery_store.analysis_bundle(store.raw_records)
    weekly = plan["daily"]
    current = plan["weekly"] if day == "Whole Week" else weekly[day]
    xlsx = build_delivery_plan_xlsx(day, current, weekly, delivery_store.schedule)
    stamp = _now_local().strftime("%Y%m%d_%H%M%S")
    safe_day = day.replace(" ", "_")
    filename = f"Delivery_Plan_{safe_day}_{stamp}.xlsx"
    return _send_bytes(xlsx, filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/health")
def health():
    storage_status = cloud_state.status()
    blocking_issues = []
    warnings = []

    if IS_VERCEL and not SERVERLESS_SESSION_READY:
        blocking_issues.append("stable-admin-session-secret-missing")

    if IS_VERCEL and cloud_state.durable_required and not storage_status.enabled:
        blocking_issues.append("strict-durable-storage-not-connected")
    elif IS_VERCEL and not storage_status.enabled:
        warnings.append("runtime-storage-fallback-active")

    if _cloud_runtime_error:
        if cloud_state.durable_required:
            blocking_issues.append("cloud-runtime-hydration-error")
        else:
            warnings.append("optional-cloud-runtime-error")

    deployment_ready = (not IS_VERCEL) or not blocking_issues
    mutation_ready = deployment_ready
    return {
        "status": "ok" if deployment_ready else "degraded",
        "deployment_ready": deployment_ready,
        "mutation_ready": mutation_ready,
        "import_ready": mutation_ready,
        "blocking_issues": blocking_issues,
        "warnings": warnings,
        "version": APP_VERSION,
        "role": role(),
        "serverless_session_ready": SERVERLESS_SESSION_READY,
        "session_key_source": SESSION_KEY_SOURCE,
        "scm_secret_key_configured": bool(str(os.environ.get("SCM_SECRET_KEY") or "").strip()),
        "admin_password_configured": bool(str(os.environ.get("SCM_ADMIN_PASSWORD") or "").strip()),
        "records": len(store.raw_records),
        "management_models": len(store.management_records),
        "kpi_periods": _kpi_period_diagnostics(),
        "aging_records": aging_record_count(),
        "delivery_allocations": len(delivery_store.allocations),
        "runtime": storage_status.runtime,
        "storage": storage_status.provider,
        "storage_mode": cloud_state.persistence_mode,
        "persistent_storage": storage_status.enabled,
        "strict_durable_storage": cloud_state.durable_required,
        "storage_detail": storage_status.detail,
        "storage_auth_mode": storage_status.auth_mode,
        "storage_runtime_error": _cloud_runtime_error or None,
        "cloud_runtime_hydrated": _cloud_runtime_hydrated,
        "blob_store_id_present": bool(cloud_state.store_id),
        "oidc_request_token_present": cloud_state.request_oidc_present,
        "legacy_blob_token_present": bool(cloud_state.read_write_token),
        "state_root": str(STATE_ROOT),
        "import_temp_root": str(IMPORT_TMP_DIR),
        "saved_import": ACTIVE_IMPORT.exists(),
        "data_source": "no-data" if EMPTY_STATE_MARKER.exists() else ("saved-import" if ACTIVE_IMPORT.exists() else "bundled-baseline"),
        "export_storage": "in-memory-download-only",
    }


@app.errorhandler(413)
def _request_too_large(_exc):
    message = (
        "Upload exceeds the 4 MB Vercel Function limit for this server-side import."
        if IS_VERCEL else "Upload exceeds the 25 MB dashboard limit."
    )
    if request.path.startswith("/admin/") or request.path.startswith("/api/"):
        return jsonify({"error": message}), 413
    return message, 413


@app.after_request
def _response_headers(response):
    # Static assets may be cached; API/data responses must stay fresh after imports.
    if request.path.startswith("/static/"):
        response.headers["Cache-Control"] = "public, max-age=3600"
    elif request.path == "/" or request.path.startswith("/login") or request.path.startswith("/api/") or request.path.startswith("/admin/") or request.path.startswith("/aging/"):
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "SAMEORIGIN"
    response.headers["Referrer-Policy"] = "no-referrer"
    response.headers["Permissions-Policy"] = "camera=(), microphone=(), geolocation=()"
    return response


def _port_is_available(host: str, port: int) -> bool:
    """Return True only when Windows/Linux allows us to bind this host/port."""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            sock.bind((host, port))
        return True
    except OSError:
        return False


def _pick_port(host: str) -> int:
    """Choose a usable port, avoiding Windows reserved/excluded ports."""
    candidates = []
    env_port = os.environ.get("PORT")
    if env_port:
        try:
            candidates.append(int(env_port))
        except ValueError:
            print(f"WARNING: Ignoring invalid PORT={env_port!r}")

    # 5055 is the preferred local dashboard port. Alternatives are tried automatically.
    candidates.extend([5055, 8050, 8088, 8765, 9000, 5000])

    seen = set()
    for port in candidates:
        if port in seen or not (1 <= port <= 65535):
            continue
        seen.add(port)
        if _port_is_available(host, port):
            return port

    # Ask Windows for any free ephemeral port as a final fallback.
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind((host, 0))
        return sock.getsockname()[1]


def _open_browser(url: str) -> None:
    if os.environ.get("SCM_NO_BROWSER") == "1":
        return

    def opener():
        time.sleep(1.2)
        try:
            webbrowser.open(url)
        except Exception:
            pass

    threading.Thread(target=opener, daemon=True).start()


if __name__ == "__main__":
    # Local-only by default: safer and avoids Windows firewall/reserved-socket issues.
    host = os.environ.get("SCM_HOST", "127.0.0.1")
    port = _pick_port(host)
    url = f"http://{host}:{port}"

    print("=" * 68)
    print(" SCM Inventory & Distribution Planning Dashboard")
    print(f" Open in browser: {url}")
    print(" Admin login: /login")
    print(" Press CTRL+C to stop the dashboard.")
    print("=" * 68)

    _open_browser(url)
    # Waitress is used locally instead of Flask's development server. It is more
    # stable on Windows and can keep lightweight dashboard requests responsive
    # while an export request is running.
    from waitress import serve
    threads = max(4, int(os.environ.get("SCM_SERVER_THREADS", "8")))
    serve(app, host=host, port=port, threads=threads, channel_timeout=240)
