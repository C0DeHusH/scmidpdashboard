from __future__ import annotations

from flask import Flask, jsonify, render_template, request, session, redirect, url_for, send_file, flash
from pathlib import Path
from io import BytesIO
from datetime import datetime
import os
import shutil
import json
import socket
import threading
import time
import webbrowser

from dashboard.metrics import DashboardStore
from dashboard.xlsx_export import build_branch_request_xlsx, build_delivery_plan_xlsx, build_management_order_xlsx, build_weekly_schedule_template_xlsx
from dashboard.ppt_export import build_presentation
from dashboard.delivery import DeliveryStore, DAYS

BASE = Path(__file__).resolve().parent
DATA_FILE = BASE / "data" / "MC_Dashboard_IMPORT.xlsx"
STATE_ROOT = Path(os.environ.get("SCM_DATA_DIR", str(BASE / "uploads")))
STATE_ROOT.mkdir(parents=True, exist_ok=True)
UPLOAD_DIR = STATE_ROOT

app = Flask(__name__)
app.secret_key = os.environ.get("SCM_SECRET_KEY", "change-this-secret-before-production")
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
ADMIN_PASSWORD = os.environ.get("SCM_ADMIN_PASSWORD", "admin123")
ACTIVE_IMPORT = UPLOAD_DIR / "active_import.xlsx"
MANAGEMENT_ALLOCATIONS = STATE_ROOT / "management_allocations.json"
store = DashboardStore(ACTIVE_IMPORT if ACTIVE_IMPORT.exists() else DATA_FILE)
delivery_store = DeliveryStore(BASE / "data" / "delivery_master.json", STATE_ROOT / "delivery")
delivery_store.sync_dashboard_branches(store.raw_records)


def role():
    return "admin" if session.get("is_admin") else "guest"


def _load_management_allocations():
    """Load persisted Management planning edits with backward compatibility.

    Older releases stored only Order Quantity (and later Remarks). v2.9 allows
    the visible source fields to be overridden without modifying the imported
    workbook itself. The original row key remains the stable identity even when
    Brand or Model is edited.
    """
    if not MANAGEMENT_ALLOCATIONS.exists():
        return {}
    try:
        data = json.loads(MANAGEMENT_ALLOCATIONS.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
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
        return out
    except Exception:
        return {}


def _save_management_allocations(data):
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
    return jsonify(store.bootstrap(role()))


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
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    return send_file(
        BytesIO(xlsx),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"Management_Order_Plan_{stamp}.xlsx",
    )


@app.get("/api/delivery")
def api_delivery():
    day = request.args.get("day", "Monday")
    return jsonify(delivery_store.bootstrap(day, store.raw_records))


@app.get("/delivery/template")
def delivery_template():
    path = BASE / "data" / "Delivery_Allocation_Import_Template.xlsx"
    return send_file(path, as_attachment=True, download_name="Delivery_Allocation_Import_Template.xlsx")


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
    f = request.files.get("file")
    if not f or not f.filename.lower().endswith(".xlsx"):
        return jsonify({"error": "Please select an .xlsx import workbook."}), 400
    tmp = UPLOAD_DIR / "candidate_import.xlsx"
    f.save(tmp)
    check = DashboardStore.validate(tmp)
    if not check.ok:
        tmp.unlink(missing_ok=True)
        return jsonify({"error": check.message}), 400
    active = UPLOAD_DIR / "active_import.xlsx"
    shutil.move(tmp, active)
    store.load(active)
    delivery_store.sync_dashboard_branches(store.raw_records)
    return jsonify({"ok": True, "message": "Import completed. Dashboard data refreshed.", "generated_at": store.generated_at.isoformat()})


@app.get("/admin/export/pptx")
@admin_required
def admin_export_pptx():
    # v2.18 PowerPoint export is a branded KPI review deck: YTD, Weekly,
    # All-Area Performance, Branch rankings and per-Branch A/B/C model details.
    data = store.bootstrap(role())
    all_branch_dashboards = [store.branch_dashboard(branch) for branch in store.branches]
    ppt = build_presentation(data, store.area_dashboard("Overall"), store.branch_dashboard(None), all_branch_dashboards)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    return send_file(BytesIO(ppt), mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation", as_attachment=True, download_name=f"SCM_IDP_Executive_Control_Tower_{stamp}.pptx")


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
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    safe_branch = "_".join(branch.split())
    return send_file(BytesIO(xlsx), mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name=f"Branch_Request_{safe_branch}_{stamp}.xlsx")


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


@app.get("/admin/export/delivery-schedule-template")
@admin_required
def admin_export_delivery_schedule_template():
    xlsx = build_weekly_schedule_template_xlsx(delivery_store.schedule, delivery_store.master)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    return send_file(
        BytesIO(xlsx),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"Weekly_Truck_Schedule_Template_{stamp}.xlsx",
    )


@app.post("/admin/delivery/schedule")
@admin_required
def admin_delivery_schedule():
    payload = request.get_json(force=True, silent=True) or {}
    try:
        schedule = delivery_store.update_schedule(payload.get("rows") or [])
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400
    return jsonify({"ok": True, "schedule": schedule})


@app.post("/admin/delivery/allocations")
@admin_required
def admin_delivery_allocations():
    payload = request.get_json(force=True, silent=True) or {}
    try:
        allocations = delivery_store.replace_allocations(payload.get("rows") or [])
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400
    return jsonify({"ok": True, "allocations": allocations})


@app.post("/admin/delivery/clear")
@admin_required
def admin_delivery_clear():
    payload = request.get_json(force=True, silent=True) or {}
    target = str(payload.get("target", "all")).strip().lower()
    if target not in {"all", "schedule", "allocations"}:
        return jsonify({"error": "Clear target must be all, schedule or allocations."}), 400
    try:
        if target in {"all", "schedule"}:
            delivery_store.update_schedule([])
        if target in {"all", "allocations"}:
            delivery_store.replace_allocations([])
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400
    messages = {
        "all": "Delivery Control Board cleared. Schedule and allocations are now empty; masterlists were preserved.",
        "schedule": "Weekly Truck Schedule cleared. Delivery masterlists and allocations were preserved.",
        "allocations": "All delivery allocation rows cleared. Weekly Truck Schedule and masterlists were preserved.",
    }
    return jsonify({"ok": True, "target": target, "message": messages[target]})


@app.get("/admin/export/delivery")
@admin_required
def admin_export_delivery():
    day = request.args.get("day", "Whole Week")
    if day not in DAYS and day != "Whole Week":
        day = "Whole Week"
    weekly = {d: delivery_store.analyze(d, store.raw_records) for d in DAYS}
    current = delivery_store.analyze_week(store.raw_records) if day == "Whole Week" else weekly[day]
    xlsx = build_delivery_plan_xlsx(day, current, weekly, delivery_store.schedule)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    safe_day = day.replace(" ", "_")
    return send_file(
        BytesIO(xlsx),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"Delivery_Plan_{safe_day}_{stamp}.xlsx",
    )


@app.get("/health")
def health():
    return {"status": "ok", "role": role(), "records": len(store.raw_records), "management_models": len(store.management_records), "delivery_allocations": len(delivery_store.allocations)}


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
    app.run(
        host=host,
        port=port,
        debug=os.environ.get("FLASK_DEBUG") == "1",
        use_reloader=False,
    )
