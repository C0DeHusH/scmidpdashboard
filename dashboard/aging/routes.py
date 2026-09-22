from __future__ import annotations

import csv
import io
import os
from datetime import date, datetime
from pathlib import Path

from flask import Blueprint, Response, flash, jsonify, redirect, render_template, request, send_file, session, url_for

from .analytics import filtered_dataset, summarize, normalize_filters
from .db import connect, init_db
from .exporter import build_aging_report_xlsx, build_aging_source_format_xlsx

PROJECT_ROOT = Path(__file__).resolve().parents[2]
STATE_ROOT = Path(os.environ.get("SCM_DATA_DIR", str(PROJECT_ROOT / "uploads")))
aging_bp = Blueprint("aging", __name__, url_prefix="/aging")
init_db()


def _is_admin() -> bool:
    return bool(session.get("is_admin"))


def _require_admin():
    if not _is_admin():
        flash("Admin access is required to import the consolidated SCM workbook.", "error")
        return redirect(url_for("login"))
    return None


def _fmt_qty(v):
    try:
        f = float(v)
        return f"{int(f):,}" if f.is_integer() else f"{f:,.1f}"
    except Exception:
        return "0"


@aging_bp.app_template_filter("qty")
def qty_filter(v):
    return _fmt_qty(v)


@aging_bp.app_template_filter("money")
def money_filter(v):
    try:
        return f"₱{float(v):,.2f}"
    except Exception:
        return "₱0.00"


def current_as_of() -> str:
    with connect() as conn:
        row = conn.execute("SELECT as_of_date FROM imports ORDER BY id DESC LIMIT 1").fetchone()
    return row[0] if row else date.today().isoformat()


def get_filters():
    return {
        "branch": request.args.get("branch", "").strip().upper(),
        "area": request.args.get("area", "").strip(),
        "brand": request.args.get("brand", "").strip(),
        "std": request.args.get("std", "").strip(),
        "q": request.args.get("q", "").strip(),
    }


def parse_as_of(value: str) -> date:
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except Exception:
        return date.today()


UNIT_AGE_BANDS = {"all", "healthy", "action", "high", "critical"}
UNIT_SORTS = {"oldest", "newest", "value_desc", "value_asc", "branch", "model"}


def get_unit_filters() -> dict[str, str]:
    age = request.args.get("unit_age", "all").strip().lower()
    sort = request.args.get("unit_sort", "oldest").strip().lower()
    return {
        "age": age if age in UNIT_AGE_BANDS else "all",
        "sort": sort if sort in UNIT_SORTS else "oldest",
        "q": request.args.get("unit_q", "").strip(),
    }


def _export_args(filters: dict[str, str], as_of: str, basis: str, unit_filters: dict[str, str]) -> dict[str, str]:
    """Return a canonical query string for export/download links.

    Never reuse raw request.args because it may still contain a stale Branch
    after the Area dropdown has been changed.
    """
    return {
        "as_of": as_of,
        "basis": basis,
        "area": filters.get("area", ""),
        "branch": filters.get("branch", ""),
        "brand": filters.get("brand", ""),
        "std": filters.get("std", ""),
        "q": filters.get("q", ""),
        "unit_age": unit_filters.get("age", "all"),
        "unit_sort": unit_filters.get("sort", "oldest"),
        "unit_q": unit_filters.get("q", ""),
    }


def _unit_query_match(row: dict, query: str) -> bool:
    if not query:
        return True
    q = query.casefold()
    fields = (
        "branch_name", "area", "standard_description", "description", "brand",
        "engine_no", "chassis", "barcode", "location",
    )
    return any(q in str(row.get(field) or "").casefold() for field in fields)


def _unit_age_match(days: int, band: str) -> bool:
    if band == "healthy":
        return days <= 90
    if band == "action":
        return 91 <= days <= 180
    if band == "high":
        return 181 <= days <= 365
    if band == "critical":
        return days >= 366
    return True


def unit_band_counts(rows: list[dict], query: str = "") -> dict[str, int]:
    counts = {"all": 0, "healthy": 0, "action": 0, "high": 0, "critical": 0}
    for row in rows:
        if not _unit_query_match(row, query):
            continue
        days = int(row.get("age_days") or 0)
        counts["all"] += 1
        if days <= 90:
            counts["healthy"] += 1
        elif days <= 180:
            counts["action"] += 1
        elif days <= 365:
            counts["high"] += 1
        else:
            counts["critical"] += 1
    return counts


def apply_unit_filters(rows: list[dict], unit_filters: dict[str, str]) -> list[dict]:
    band = unit_filters.get("age", "all")
    query = unit_filters.get("q", "")
    filtered = [
        row for row in rows
        if _unit_query_match(row, query) and _unit_age_match(int(row.get("age_days") or 0), band)
    ]
    sort_mode = unit_filters.get("sort", "oldest")
    if sort_mode == "newest":
        filtered.sort(key=lambda r: (int(r.get("age_days") or 0), str(r.get("standard_description") or "").casefold()))
    elif sort_mode == "value_desc":
        filtered.sort(key=lambda r: (float(r.get("inventory_value") or 0), int(r.get("age_days") or 0)), reverse=True)
    elif sort_mode == "value_asc":
        filtered.sort(key=lambda r: (float(r.get("inventory_value") or 0), int(r.get("age_days") or 0)))
    elif sort_mode == "branch":
        filtered.sort(key=lambda r: (str(r.get("branch_name") or "").casefold(), -int(r.get("age_days") or 0)))
    elif sort_mode == "model":
        filtered.sort(key=lambda r: (str(r.get("standard_description") or "").casefold(), -int(r.get("age_days") or 0)))
    else:
        filtered.sort(key=lambda r: (int(r.get("age_days") or 0), float(r.get("inventory_value") or 0)), reverse=True)
    return filtered


def _resolve_view_state() -> dict:
    """Resolve and normalize the current Aging filter scope once per request."""
    filters = get_filters()
    as_of_str = request.args.get("as_of") or current_as_of()
    as_of = parse_as_of(as_of_str)
    basis = request.args.get("basis", "branch")
    if basis not in {"branch", "company"}:
        basis = "branch"
    filters, options, filter_warnings = normalize_filters(filters)
    rows = filtered_dataset(filters, as_of, basis)
    return {
        "filters": filters,
        "options": options,
        "filter_warnings": filter_warnings,
        "as_of": as_of.isoformat(),
        "as_of_date": as_of,
        "basis": basis,
        "base_rows": rows,
        "row_count": len(rows),
    }


def _unit_view_context(view: dict) -> dict:
    """Build only the Unit Trace state so unit-level filtering stays lightweight."""
    unit_filters = get_unit_filters()
    rows = view["base_rows"]
    band_counts = unit_band_counts(rows, unit_filters.get("q", ""))
    unit_rows = apply_unit_filters(rows, unit_filters)
    export_args = _export_args(view["filters"], view["as_of"], view["basis"], unit_filters)
    return {
        "rows": unit_rows[:250],
        "unit_filters": unit_filters,
        "unit_row_count": len(unit_rows),
        "unit_band_counts": band_counts,
        "export_args": export_args,
    }


def _aging_chart_data(summary: dict) -> dict:
    return {
        "buckets": summary.get("bucket_labels", []),
        "bucketValues": summary.get("bucket_values", []),
        "areas": summary.get("area_labels", []),
        "areaValues": summary.get("area_values", []),
        "areaPctValues": summary.get("area_pct_values", []),
        "brands": summary.get("brand_labels", []),
        "brandValues": summary.get("brand_values", []),
        "models": summary.get("aged_model_labels", []),
        "modelValues": summary.get("aged_model_values", []),
    }


def _dashboard_context() -> dict:
    view = _resolve_view_state()
    summary = summarize(view["base_rows"])
    unit = _unit_view_context(view)
    with connect() as conn:
        latest_import = conn.execute("SELECT * FROM imports ORDER BY id DESC LIMIT 1").fetchone()
    return {
        **view,
        **unit,
        "summary": summary,
        "latest_import": latest_import,
        "aging_chart_data": _aging_chart_data(summary),
        "role": "admin" if _is_admin() else "guest",
    }


@aging_bp.get("/")
def dashboard():
    return render_template("aging/dashboard.html", **_dashboard_context())


@aging_bp.get("/partial")
def dashboard_partial():
    """Return only the Aging regions affected by the global Aging filters.

    The page shell, navigation and filter controls stay mounted in the browser.
    This prevents full-page navigation and keeps the user's scroll/focus stable.
    """
    ctx = _dashboard_context()
    return jsonify({
        "summary_html": render_template("aging/_summary_region.html", **ctx),
        "results_html": render_template("aging/_results_region.html", **ctx),
        "filters": ctx["filters"],
        "options": ctx["options"],
        "as_of": ctx["as_of"],
        "basis": ctx["basis"],
        "row_count": ctx["row_count"],
        "filter_warnings": ctx["filter_warnings"],
        "chart_data": ctx["aging_chart_data"],
        "export_xlsx": url_for("aging.export_xlsx", **ctx["export_args"]),
        "export_csv": url_for("aging.export_csv", **ctx["export_args"]),
        "view_url": url_for("aging.dashboard", **ctx["export_args"]),
    })


@aging_bp.get("/partial/units")
def unit_trace_partial():
    """Refresh only Unit-Level Traceability without recalculating/rendering all cards."""
    view = _resolve_view_state()
    unit = _unit_view_context(view)
    ctx = {
        **view,
        **unit,
        "role": "admin" if _is_admin() else "guest",
    }
    return jsonify({
        "html": render_template("aging/_unit_trace.html", **ctx),
        "view_url": url_for("aging.dashboard", **unit["export_args"]),
        "unit_row_count": unit["unit_row_count"],
    })


@aging_bp.get("/units")
def units():
    return redirect(url_for("aging.dashboard", **request.args))


@aging_bp.route("/import", methods=["GET", "POST"])
def import_data():
    """Legacy URL retained only to open the unified system import dialog."""
    denied = _require_admin()
    if denied:
        return denied
    flash("Motorcycle Aging now refreshes from the same consolidated workbook as the SCM dashboard.", "success")
    return redirect("/?open_import=1")


@aging_bp.get("/export.csv")
def export_csv():
    filters = get_filters()
    as_of_str = request.args.get("as_of") or current_as_of()
    as_of = parse_as_of(as_of_str)
    basis = request.args.get("basis", "branch")
    if basis not in {"branch", "company"}:
        basis = "branch"
    filters, _options, _filter_warnings = normalize_filters(filters)
    rows = filtered_dataset(filters, as_of, basis)
    rows = apply_unit_filters(rows, get_unit_filters())
    sio = io.StringIO()
    writer = csv.writer(sio)
    writer.writerow([
        "Branch", "Area", "Standard Description", "Description", "Brand", "Engine No.", "Chassis",
        "Incoming Date", "Created On", "Qty", "Amount", "Inventory Value", "Age Basis", "Age Days",
        "Location", "Color", "Barcode", "As Of Date"
    ])
    for r in rows:
        writer.writerow([
            r.get("branch_name"), r.get("area"), r.get("standard_description"), r.get("description"),
            r.get("brand"), r.get("engine_no"), r.get("chassis"), r.get("incoming_date"), r.get("created_on"),
            r.get("qty"), r.get("amount"), r.get("inventory_value"), basis, r.get("age_days"),
            r.get("location"), r.get("color"), r.get("barcode"), as_of.isoformat()
        ])
    return Response(
        sio.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": f"attachment; filename=motorcycle_aging_{as_of.isoformat()}.csv"},
    )


@aging_bp.get("/export.xlsx")
def export_xlsx():
    filters = get_filters()
    unit_filters = get_unit_filters()
    as_of_str = request.args.get("as_of") or current_as_of()
    as_of = parse_as_of(as_of_str)
    basis = request.args.get("basis", "branch")
    if basis not in {"branch", "company"}:
        basis = "branch"

    filters, _options, filter_warnings = normalize_filters(filters)
    rows = filtered_dataset(filters, as_of, basis)
    summary = summarize(rows)
    unit_rows = apply_unit_filters(rows, unit_filters)
    with connect() as conn:
        latest_import = conn.execute("SELECT * FROM imports ORDER BY id DESC LIMIT 1").fetchone()
    source_filename = latest_import["filename"] if latest_import else ""
    detail_view = request.args.get("detail", "").strip() == "1"
    if detail_view:
        payload = build_aging_report_xlsx(
            summary=summary,
            unit_rows=unit_rows,
            as_of=as_of.isoformat(),
            basis=basis,
            filters=filters,
            unit_filters=unit_filters,
            source_filename=source_filename,
            filter_warnings=filter_warnings,
            base_row_count=len(rows),
            active_sheet="Unit Detail",
        )
        filename = f"Unit_Trace_Aging_Action_List_{as_of.isoformat()}.xlsx"
    else:
        # Main Aging export intentionally mirrors the imported Aging report:
        # one visible row per unit, same 21-column operational layout, with the
        # current Area / Branch / Brand / Model / unit filters already applied.
        payload = build_aging_source_format_xlsx(
            unit_rows=unit_rows,
            as_of=as_of.isoformat(),
            source_workbook_path=STATE_ROOT / "active_import.xlsx",
        )
        filename = f"Aging_Report_Filtered_{as_of.isoformat()}.xlsx"
    return send_file(
        io.BytesIO(payload),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
        max_age=0,
    )


@aging_bp.get("/download-template")
def download_template():
    """Backward-compatible link: the consolidated workbook is now the only import source."""
    denied = _require_admin()
    if denied:
        return denied
    flash("A separate Aging template is no longer required. Use the consolidated SCM workbook.", "success")
    return redirect("/?open_import=1")
