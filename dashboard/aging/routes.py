from __future__ import annotations

import csv
import io
import os
from datetime import date, datetime
from pathlib import Path

from flask import Blueprint, Response, flash, redirect, render_template, request, send_file, session, url_for

from .analytics import filter_options, filtered_dataset, summarize
from .db import connect, init_db
from .exporter import build_aging_report_xlsx

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


@aging_bp.get("/")
def dashboard():
    filters = get_filters()
    as_of_str = request.args.get("as_of") or current_as_of()
    as_of = parse_as_of(as_of_str)
    basis = request.args.get("basis", "branch")
    if basis not in {"branch", "company"}:
        basis = "branch"
    options = filter_options(filters.get("area", ""))
    valid_branch_keys = {b["branch_key"] for b in options["branches"]}
    if filters.get("branch") and filters["branch"] not in valid_branch_keys:
        filters["branch"] = ""
    rows = filtered_dataset(filters, as_of, basis)
    summary = summarize(rows)
    unit_filters = get_unit_filters()
    band_counts = unit_band_counts(rows, unit_filters.get("q", ""))
    unit_rows = apply_unit_filters(rows, unit_filters)
    display_rows = unit_rows[:250]
    with connect() as conn:
        latest_import = conn.execute("SELECT * FROM imports ORDER BY id DESC LIMIT 1").fetchone()
    return render_template(
        "aging/dashboard.html",
        rows=display_rows,
        row_count=len(rows),
        summary=summary,
        options=options,
        filters=filters,
        as_of=as_of.isoformat(),
        basis=basis,
        latest_import=latest_import,
        unit_filters=unit_filters,
        unit_row_count=len(unit_rows),
        unit_band_counts=band_counts,
        role="admin" if _is_admin() else "guest",
    )


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

    rows = filtered_dataset(filters, as_of, basis)
    summary = summarize(rows)
    unit_rows = apply_unit_filters(rows, unit_filters)
    with connect() as conn:
        latest_import = conn.execute("SELECT * FROM imports ORDER BY id DESC LIMIT 1").fetchone()
    source_filename = latest_import["filename"] if latest_import else ""
    payload = build_aging_report_xlsx(
        summary=summary,
        unit_rows=unit_rows,
        as_of=as_of.isoformat(),
        basis=basis,
        filters=filters,
        unit_filters=unit_filters,
        source_filename=source_filename,
    )
    filename = f"Motorcycle_Aging_Intelligence_{as_of.isoformat()}.xlsx"
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
