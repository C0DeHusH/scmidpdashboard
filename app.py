from __future__ import annotations

from flask import Flask, jsonify, render_template, request, session, redirect, url_for, send_file, flash
from pathlib import Path
from io import BytesIO
from datetime import datetime, timezone, timedelta
from dotenv import load_dotenv
import os
import shutil
import json
import socket
import threading
import time
import webbrowser

from dashboard.metrics import DashboardStore
from dashboard.xlsx_export import build_branch_request_xlsx, build_delivery_plan_xlsx, build_management_order_xlsx
from dashboard.ppt_export import build_presentation
from dashboard.delivery import DeliveryStore, DAYS

BASE = Path(__file__).resolve().parent
load_dotenv(BASE / ".env")
DATA_FILE = BASE / "data" / "MC_Dashboard_IMPORT.xlsx"
STATE_ROOT = Path(os.environ.get("SCM_DATA_DIR", str(BASE / "uploads")))
STATE_ROOT.mkdir(parents=True, exist_ok=True)
UPLOAD_DIR = STATE_ROOT

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

app = Flask(__name__)
app.secret_key = os.environ.get("SCM_SECRET_KEY", "change-this-secret-before-production")
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["JSON_SORT_KEYS"] = False
app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 3600
ADMIN_PASSWORD = os.environ.get("SCM_ADMIN_PASSWORD", "admin123")
ACTIVE_IMPORT = UPLOAD_DIR / "active_import.xlsx"
EMPTY_STATE_MARKER = STATE_ROOT / ".scm_no_data"
MANAGEMENT_ALLOCATIONS = STATE_ROOT / "management_allocations.json"
APP_TZ = timezone(timedelta(hours=int(os.environ.get("SCM_UTC_OFFSET_HOURS", "8"))))

app.register_blueprint(aging_bp)


def _now_local() -> datetime:
    return datetime.now(APP_TZ)


store = DashboardStore(
    None if EMPTY_STATE_MARKER.exists()
    else (ACTIVE_IMPORT if ACTIVE_IMPORT.exists() else DATA_FILE)
)
delivery_store = DeliveryStore(
    BASE / "data" / "delivery_master.json",
    STATE_ROOT / "delivery",
)
delivery_store.sync_dashboard_branches(store.raw_records)
_management_allocations_cache = None
_bootstrap_cache = {}
_presentation_input_cache = {}


def _invalidate_bootstrap_cache():
    _bootstrap_cache.clear()
    _presentation_input_cache.clear()


def _presentation_inputs():
    """Cache expensive Area/Branch drilldown preparation until the workbook changes."""
    key = store.generated_at.isoformat()
    cached = _presentation_input_cache.get(key)
    if cached is not None:
        return cached
    payload = {
        "area": store.area_dashboard("Overall"),
        "branches": [store.branch_dashboard(branch) for branch in store.branches],
    }
    _presentation_input_cache.clear()
    _presentation_input_cache[key] = payload
    return payload


def role():
    return "admin" if session.get("is_admin") else "guest"


def _load_management_allocations():
    """Load persisted Management planning edits with backward compatibility.

    Management planning edits are cached in memory after the first local read.
    Writes update the cache and an atomic local JSON file.
    """
    global _management_allocations_cache
    if _management_allocations_cache is not None:
        return json.loads(json.dumps(_management_allocations_cache))
    try:
        data = None
        if MANAGEMENT_ALLOCATIONS.exists():
            data = json.loads(MANAGEMENT_ALLOCATIONS.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            _management_allocations_cache = {}
            return {}
        out = {}
        numeric_fields = {"quantity", "unit_cost", "inventory", "doi", "po_balance"}
        text_fields = {"remarks", "class", "brand", "model", "stock_status"}
        for key, value in data.items():
            try:
                if isinstance(value, dict):
                    edit = {}
                    # Backward compatibility: allocation -> quantity.
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
        return json.loads(json.dumps(out))
    except Exception:
        _management_allocations_cache = {}
        return {}


def _save_management_allocations(data):
    global _management_allocations_cache
    _management_allocations_cache = json.loads(json.dumps(data))
    MANAGEMENT_ALLOCATIONS.parent.mkdir(parents=True, exist_ok=True)
    tmp = MANAGEMENT_ALLOCATIONS.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, indent=2, sort_keys=True), encoding="utf-8")
    tmp.replace(MANAGEMENT_ALLOCATIONS)


def admin_required(fn):
    from functools import wraps
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if role() != "admin":
            return jsonify({"error": "Admin access required."}), 403
        return fn(*args, **kwargs)
    return wrapper


@app.get("/")
def index():
    return render_template("index.html", role=role())


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        password = request.form.get("password", "")
        if password == ADMIN_PASSWORD:
            session["is_admin"] = True
            return redirect(url_for("index"))
        flash("Invalid admin password.", "error")
    return render_template("login.html")


@app.post("/logout")
def logout():
    session.clear()
    return redirect(url_for("index"))


@app.get("/api/bootstrap")
def api_bootstrap():
    no_data = EMPTY_STATE_MARKER.exists()
    cache_key = (role(), store.generated_at.isoformat(), ACTIVE_IMPORT.exists(), no_data)
    if cache_key not in _bootstrap_cache:
        payload = store.bootstrap(role())
        payload["data_source"] = "No Data" if no_data or not payload.get("has_data") else ("Saved Import" if ACTIVE_IMPORT.exists() else "Bundled Baseline")
        payload["has_saved_import"] = ACTIVE_IMPORT.exists()
        _bootstrap_cache.clear()
        _bootstrap_cache[cache_key] = payload
    return jsonify(_bootstrap_cache[cache_key])


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
    payload = request.get_json(force=True, silent=True) or {}
    incoming = payload.get("allocations") or {}
    if not isinstance(incoming, dict):
        return jsonify({"error": "Management edits must be a key/value object."}), 400
    current = _load_management_allocations()
    valid_keys = {r["key"] for r in store.management_records}
    numeric_fields = {"quantity", "unit_cost", "inventory", "po_balance"}
    text_fields = {"remarks"}
    updated = 0
    for key, value in incoming.items():
        key = str(key)
        if key not in valid_keys or not isinstance(value, dict):
            continue
        edit = {k: v for k, v in dict(current.get(key) or {}).items() if k in (numeric_fields | text_fields)}
        try:
            for field in numeric_fields:
                if field in value:
                    edit[field] = max(0.0, float(value.get(field) or 0))
            for field in text_fields:
                if field in value:
                    edit[field] = str(value.get(field, "") or "").strip()
        except (TypeError, ValueError):
            return jsonify({"error": f"Invalid numeric Management value for {key}."}), 400
        current[key] = edit
        updated += 1
    _save_management_allocations(current)
    return jsonify({"ok": True, "updated": updated, "message": f"Saved {updated} Management planning line(s)."})


@app.post("/admin/export/management")
@admin_required
def admin_export_management():
    payload = request.get_json(force=True, silent=True) or {}
    incoming = payload.get("orders") or {}
    if not isinstance(incoming, dict):
        return jsonify({"error": "Management edits must be a key/value object."}), 400
    combined = _load_management_allocations()
    valid_keys = {r["key"] for r in store.management_records}
    numeric_fields = {"quantity", "unit_cost", "inventory", "po_balance"}
    text_fields = {"remarks"}
    for key, value in incoming.items():
        key = str(key)
        if key not in valid_keys or not isinstance(value, dict):
            continue
        edit = {k: v for k, v in dict(combined.get(key) or {}).items() if k in (numeric_fields | text_fields)}
        try:
            for field in numeric_fields:
                if field in value:
                    edit[field] = max(0.0, float(value.get(field) or 0))
            for field in text_fields:
                if field in value:
                    edit[field] = str(value.get(field, "") or "").strip()
        except (TypeError, ValueError):
            return jsonify({"error": f"Invalid numeric Management value for {key}."}), 400
        combined[key] = edit

    visible_keys = payload.get("keys") or []
    if visible_keys and isinstance(visible_keys, list):
        # Rebuild all effective rows first, then keep the exact rows currently
        # visible on screen. This preserves unsaved Brand/Model/Status edits even
        # when those edits would no longer match the pre-edit filter value.
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
    # Export only actual order lines. Rows with zero Order Quantity stay on the
    # dashboard for planning but are intentionally excluded from the order file.
    ordered_rows = [r for r in data.get("rows", []) if float(r.get("allocation", 0) or 0) > 0]
    if not ordered_rows:
        return jsonify({"error": "No models with Order Quantity greater than zero to export."}), 400
    data["rows"] = ordered_rows
    data["summary"] = {
        "models": len(ordered_rows),
        "current_inventory": round(sum(float(r.get("inventory", 0) or 0) for r in ordered_rows), 4),
        "po_balance": round(sum(float(r.get("po_balance", 0) or 0) for r in ordered_rows), 4),
        "allocation_order": round(sum(float(r.get("allocation", 0) or 0) for r in ordered_rows), 4),
        "grand_total": round(sum(float(r.get("total_amount", 0) or 0) for r in ordered_rows), 4),
    }
    xlsx = build_management_order_xlsx(data)
    stamp = _now_local().strftime("%Y%m%d_%H%M%S")
    filename = f"Management_Order_Plan_{stamp}.xlsx"
    return send_file(
        BytesIO(xlsx),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


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
    """Import one consolidated workbook into every analytical module.

    v2.44 treats Raw/KPI/Reorder/Aging as one source-of-truth refresh.  The
    previous active workbook and Aging database are backed up before commit so
    a failed Aging parse cannot leave SCM and Aging on different source files.
    """
    global _management_allocations_cache
    f = request.files.get("file")
    if not f or not f.filename.lower().endswith(".xlsx"):
        return jsonify({"error": "Please select the consolidated .xlsx import workbook."}), 400

    tmp = UPLOAD_DIR / "candidate_import.xlsx"
    active = ACTIVE_IMPORT
    active_backup = UPLOAD_DIR / ".active_import_before_unified_refresh.xlsx"
    aging_backup = AGING_DB_PATH.with_name(".aging_before_unified_refresh.db")
    tmp.unlink(missing_ok=True)
    active_backup.unlink(missing_ok=True)
    aging_backup.unlink(missing_ok=True)
    f.save(tmp)

    check = DashboardStore.validate(tmp)
    if not check.ok:
        tmp.unlink(missing_ok=True)
        return jsonify({"error": check.message}), 400
    try:
        validate_unified_aging(tmp, "Aging")
    except Exception as exc:
        tmp.unlink(missing_ok=True)
        return jsonify({"error": f"Aging validation failed: {exc}"}), 400

    had_active = active.exists()
    had_aging_db = AGING_DB_PATH.exists()
    try:
        if had_active:
            shutil.copy2(active, active_backup)
        if had_aging_db:
            backup_aging_database(aging_backup)

        # Load Aging first while the candidate still exists, then commit the
        # same file as the active SCM workbook.  Any exception rolls both back.
        aging_result = import_aging_excel(tmp, f.filename, mode="replace", sheet_name="Aging")
        os.replace(tmp, active)
        store.load(active, prevalidated=True)
        EMPTY_STATE_MARKER.unlink(missing_ok=True)

        # A new consolidated workbook is a new planning baseline.  Management
        # edits tied to the previous source are intentionally cleared.
        MANAGEMENT_ALLOCATIONS.unlink(missing_ok=True)
        _management_allocations_cache = {}
        delivery_store.sync_dashboard_branches(store.raw_records)
        _invalidate_bootstrap_cache()

        message = (
            "Unified import complete — Executive/KPI, Reorder/Management and Motorcycle Aging refreshed from one workbook. "
            f"Aging: {aging_result['rows']:,} units · {aging_result['branches']} branches · "
            f"{aging_result['areas']} areas · as of {aging_result['as_of_date']}."
        )
        return jsonify({
            "ok": True,
            "message": message,
            "generated_at": store.generated_at.isoformat(),
            "modules": {
                "scm": {"status": "updated", "records": len(store.raw_records)},
                "management": {"status": "updated", "source": "Reorder / Management worksheet"},
                "aging": {"status": "updated", **aging_result},
            },
        })
    except Exception as exc:
        # Roll back the saved workbook.
        try:
            if had_active and active_backup.exists():
                shutil.copy2(active_backup, active)
                store.load(active)
            else:
                active.unlink(missing_ok=True)
                if EMPTY_STATE_MARKER.exists():
                    store.clear()
                elif DATA_FILE.exists():
                    store.load(DATA_FILE)
        except Exception:
            pass

        # Roll back the Aging database to the previous refresh.
        try:
            if had_aging_db and aging_backup.exists():
                restore_aging_database(aging_backup)
            elif AGING_DB_PATH.exists():
                clear_aging_data()
        except Exception:
            pass
        _invalidate_bootstrap_cache()
        return jsonify({"error": f"Unified import failed and the previous data was restored: {exc}"}), 400
    finally:
        tmp.unlink(missing_ok=True)
        active_backup.unlink(missing_ok=True)
        aging_backup.unlink(missing_ok=True)


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
    return send_file(BytesIO(ppt), mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation", as_attachment=True, download_name=filename)


@app.post("/admin/export/request")
@admin_required
def admin_export_request():
    payload = request.get_json(force=True, silent=True) or {}
    branch = str(payload.get("branch", "")).strip()
    items = payload.get("items") or []
    if not branch or not items:
        return jsonify({"error": "Branch and at least one requested item are required."}), 400
    verified = []
    area = ""
    for item in items:
        model = str(item.get("model", "")).strip()
        base = store.model_lookup(branch, model)
        if not base:
            continue
        qty = float(item.get("requested_qty", 0) or 0)
        new_doi = ((base["inventory"] + qty) / base["avg_daily_sale"]) if base["avg_daily_sale"] > 0 else "N/A"
        area = base["area"]
        verified.append({**base, "requested_qty": qty, "remarks": str(item.get("remarks", "")), "new_doi": round(new_doi, 2) if isinstance(new_doi, float) else new_doi})
    if not verified:
        return jsonify({"error": "No valid request lines."}), 400
    xlsx = build_branch_request_xlsx(branch, area, verified)
    stamp = _now_local().strftime("%Y%m%d_%H%M%S")
    safe_branch = "_".join(branch.split())
    filename = f"Branch_Request_{safe_branch}_{stamp}.xlsx"
    return send_file(BytesIO(xlsx), mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name=filename)


@app.post("/admin/delivery/import")
@admin_required
def admin_delivery_import():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "Select an allocation file."}), 400
    ext = Path(f.filename or "").suffix.lower()
    if ext not in {".xlsx", ".xlsm", ".csv"}:
        return jsonify({"error": "Allocation import accepts .xlsx, .xlsm or .csv."}), 400
    tmp = UPLOAD_DIR / f"delivery_allocation_candidate{ext}"
    f.save(tmp)
    try:
        rows, warnings = delivery_store.import_allocations(tmp)
    except Exception as exc:
        tmp.unlink(missing_ok=True)
        return jsonify({"error": str(exc)}), 400
    tmp.unlink(missing_ok=True)
    return jsonify({"ok": True, "rows": len(rows), "warnings": warnings[:20], "message": f"Imported {len(rows)} allocation rows."})


@app.post("/admin/delivery/master")
@admin_required
def admin_delivery_master():
    payload = request.get_json(force=True, silent=True) or {}
    try:
        master = delivery_store.update_master(str(payload.get("section", "")), payload.get("rows"))
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400
    return jsonify({"ok": True, "master": master})


@app.post("/admin/delivery/schedule/import")
@admin_required
def admin_delivery_schedule_import():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "Select a Weekly Truck Schedule file."}), 400
    ext = Path(f.filename or "").suffix.lower()
    if ext not in {".xlsx", ".xlsm", ".csv"}:
        return jsonify({"error": "Weekly Schedule import accepts .xlsx, .xlsm or .csv."}), 400
    tmp = UPLOAD_DIR / f"weekly_schedule_candidate{ext}"
    f.save(tmp)
    try:
        schedule, warnings = delivery_store.import_schedule(tmp)
    except Exception as exc:
        tmp.unlink(missing_ok=True)
        return jsonify({"error": str(exc)}), 400
    tmp.unlink(missing_ok=True)
    return jsonify({"ok": True, "rows": len(schedule), "warnings": warnings[:30], "message": f"Imported and saved {len(schedule)} Weekly Truck Schedule row(s)."})


@app.post("/admin/delivery/schedule")
@admin_required
def admin_delivery_schedule():
    payload = request.get_json(force=True, silent=True) or {}
    try:
        schedule = delivery_store.update_schedule(payload.get("rows") or [])
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400
    return jsonify({"ok": True, "schedule": schedule})


@app.post("/admin/delivery/plan")
@admin_required
def admin_delivery_plan_save():
    """Persist the current weekly truck schedule and allocation plan together."""
    payload = request.get_json(force=True, silent=True) or {}
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
    payload = request.get_json(force=True, silent=True) or {}
    try:
        allocations = delivery_store.replace_allocations(payload.get("rows") or [])
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400
    return jsonify({"ok": True, "allocations": allocations})


@app.patch("/admin/delivery/allocation/<int:index>")
@admin_required
def admin_delivery_allocation_update(index: int):
    payload = request.get_json(force=True, silent=True) or {}
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
    payload = request.get_json(force=True, silent=True) or {}
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
    """Destructively clear all user-loaded SCM and Motorcycle Aging data.

    The application remains installed, but dashboards stay empty until fresh
    files are imported again. Bundled templates/reference files are preserved.
    """
    global _management_allocations_cache
    payload = request.get_json(force=True, silent=True) or {}
    confirmation = str(payload.get("confirmation") or "").strip().upper()
    if confirmation != "CLEAR DATA":
        return jsonify({"error": "Type CLEAR DATA exactly to confirm the reset."}), 400

    ACTIVE_IMPORT.unlink(missing_ok=True)
    (UPLOAD_DIR / "candidate_import.xlsx").unlink(missing_ok=True)
    MANAGEMENT_ALLOCATIONS.unlink(missing_ok=True)
    _management_allocations_cache = {}

    # Main SCM dashboard must remain empty after reset, including across restarts.
    store.clear()
    EMPTY_STATE_MARKER.write_text(_now_local().isoformat(), encoding="utf-8")

    # Clear saved planning rows while keeping application configuration/templates.
    delivery_store.reset_to_defaults()
    delivery_store.sync_dashboard_branches([])

    # The integrated Motorcycle Aging module is part of the same system reset.
    clear_aging_data()

    _invalidate_bootstrap_cache()

    for legacy_dir in (STATE_ROOT / "exports", BASE / "exports"):
        if legacy_dir.exists() and legacy_dir.is_dir():
            shutil.rmtree(legacy_dir, ignore_errors=True)

    return jsonify({
        "ok": True,
        "message": "All imported and saved system data was cleared. Dashboards are now empty until new files are imported.",
        "records": 0,
        "aging_records": 0,
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
    payload = request.get_json(force=True, silent=True) or {}
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
    return send_file(
        BytesIO(xlsx),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


@app.get("/health")
def health():
    return {
        "status": "ok",
        "role": role(),
        "records": len(store.raw_records),
        "management_models": len(store.management_records),
        "aging_records": aging_record_count(),
        "delivery_allocations": len(delivery_store.allocations),
        "storage": "local",
        "saved_import": ACTIVE_IMPORT.exists(),
        "data_source": "no-data" if EMPTY_STATE_MARKER.exists() else ("saved-import" if ACTIVE_IMPORT.exists() else "bundled-baseline"),
        "export_storage": "in-memory-download-only",
    }


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
