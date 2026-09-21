from __future__ import annotations

from collections import Counter, defaultdict
from datetime import date, datetime
from typing import Any

from .db import connect

DETAILED_BUCKETS = ["0-30 DAYS", "31-60 DAYS", "61-90 DAYS", "91-180 DAYS", "181-365 DAYS", "366+ DAYS"]
STANDARD_BUCKETS = ["1-30 DAYS", "31-60 DAYS", "61-90 DAYS", "91 DAYS UP"]


def parse_date(value: str | None) -> date | None:
    if not value:
        return None
    try:
        return datetime.strptime(value[:10], "%Y-%m-%d").date()
    except (TypeError, ValueError):
        return None


def age_days(row: dict[str, Any], as_of: date, basis: str) -> int:
    source = row.get("incoming_date") if basis == "company" else row.get("created_on")
    dt = parse_date(source)
    if not dt:
        return 0
    return max((as_of - dt).days, 0)


def detailed_bucket(days: int) -> str:
    if days <= 30:
        return "0-30 DAYS"
    if days <= 60:
        return "31-60 DAYS"
    if days <= 90:
        return "61-90 DAYS"
    if days <= 180:
        return "91-180 DAYS"
    if days <= 365:
        return "181-365 DAYS"
    return "366+ DAYS"


def standard_bucket(days: int) -> str:
    if days <= 30:
        return "1-30 DAYS"
    if days <= 60:
        return "31-60 DAYS"
    if days <= 90:
        return "61-90 DAYS"
    return "91 DAYS UP"


def _area_expr() -> str:
    return "COALESCE(NULLIF(u.area,''), NULLIF(b.area,''), 'UNMAPPED')"


def _base_rows(filters: dict[str, str]) -> list[dict[str, Any]]:
    conditions = []
    params: list[Any] = []
    area_expr = _area_expr()
    if filters.get("branch"):
        conditions.append("u.branch_key = ?")
        params.append(filters["branch"])
    if filters.get("area"):
        conditions.append(f"{area_expr} = ?")
        params.append(filters["area"])
    if filters.get("brand"):
        conditions.append("COALESCE(u.brand,'') = ?")
        params.append(filters["brand"])
    if filters.get("std"):
        conditions.append("LOWER(COALESCE(u.standard_description,'')) LIKE ?")
        params.append(f"%{filters['std'].lower()}%")
    if filters.get("q"):
        q = f"%{filters['q'].lower()}%"
        conditions.append("(LOWER(COALESCE(u.engine_no,'')) LIKE ? OR LOWER(COALESCE(u.chassis,'')) LIKE ? OR LOWER(COALESCE(u.barcode,'')) LIKE ? OR LOWER(COALESCE(u.description,'')) LIKE ?)")
        params.extend([q, q, q, q])
    where = "WHERE " + " AND ".join(conditions) if conditions else ""
    sql = f"""
        SELECT u.*,
               COALESCE(NULLIF(u.branch_original,''), NULLIF(b.branch_name,''), u.branch_key) AS branch_name,
               {area_expr} AS area_display
        FROM units u
        LEFT JOIN branch_area b ON b.branch_key = u.branch_key
        {where}
        ORDER BY u.id DESC
    """
    with connect() as conn:
        result = []
        for record in conn.execute(sql, params).fetchall():
            row = dict(record)
            row["area"] = row.pop("area_display")
            result.append(row)
        return result


def filtered_dataset(filters: dict[str, str], as_of: date, basis: str) -> list[dict[str, Any]]:
    rows = _base_rows(filters)
    out = []
    for row in rows:
        days = age_days(row, as_of, basis)
        row["age_days"] = days
        # Aging bands remain internal for charts/analytics even though they are no longer shown as unit-table columns.
        row["age_bucket"] = detailed_bucket(days)
        row["age_bucket_standard"] = standard_bucket(days)
        row["inventory_value"] = float(row.get("qty") or 0) * float(row.get("amount") or 0)
        out.append(row)
    return out


def _risk_level(pct: float) -> tuple[str, str]:
    if pct >= 40:
        return "High", "A large share of the selected inventory is already beyond 90 days. Prioritize selling, transfer, and liquidation review."
    if pct >= 20:
        return "Watch", "A meaningful share of inventory is beyond 90 days. Focus action on the highest-risk areas, branches, and models."
    return "Controlled", "Overall 91+ day exposure is relatively contained. Continue monitoring the concentrated risk pockets shown below."


def summarize(rows: list[dict[str, Any]]) -> dict[str, Any]:
    total_qty = sum(float(r.get("qty") or 0) for r in rows)
    total_value = sum(float(r.get("inventory_value") or 0) for r in rows)
    weighted_age = sum(float(r.get("qty") or 0) * int(r.get("age_days") or 0) for r in rows)
    avg_age = weighted_age / total_qty if total_qty else 0
    aged_90 = sum(float(r.get("qty") or 0) for r in rows if int(r.get("age_days") or 0) > 90)
    aged_180 = sum(float(r.get("qty") or 0) for r in rows if int(r.get("age_days") or 0) > 180)
    aged_365 = sum(float(r.get("qty") or 0) for r in rows if int(r.get("age_days") or 0) > 365)
    aged_90_value = sum(float(r.get("inventory_value") or 0) for r in rows if int(r.get("age_days") or 0) > 90)
    aged_180_value = sum(float(r.get("inventory_value") or 0) for r in rows if int(r.get("age_days") or 0) > 180)
    healthy_qty = max(total_qty - aged_90, 0)
    healthy_pct = (healthy_qty / total_qty * 100) if total_qty else 0
    oldest = max((int(r.get("age_days") or 0) for r in rows), default=0)
    branch_count = len({r.get("branch_name") for r in rows if r.get("branch_name")})
    area_count = len({r.get("area") for r in rows if r.get("area")})
    aged_90_pct = (aged_90 / total_qty * 100) if total_qty else 0
    aged_365_pct = (aged_365 / total_qty * 100) if total_qty else 0
    aged_value_pct = (aged_90_value / total_value * 100) if total_value else 0
    risk_level, risk_message = _risk_level(aged_90_pct)

    standard_counts = Counter()
    detailed_counts = Counter()
    brand_counts = Counter()
    aged_models = defaultdict(float)
    model_stats: dict[str, dict[str, Any]] = {}
    area_stats: dict[str, dict[str, Any]] = {}
    branch_stats: dict[str, dict[str, Any]] = {}

    for r in rows:
        qty = float(r.get("qty") or 0)
        value = float(r.get("inventory_value") or 0)
        age = int(r.get("age_days") or 0)
        aged = age > 90
        area = r.get("area") or "UNMAPPED"
        branch = r.get("branch_name") or r.get("branch_key") or "UNSPECIFIED"
        brand = r.get("brand") or "UNSPECIFIED"
        model = r.get("standard_description") or r.get("description") or "UNSPECIFIED"

        standard_counts[r.get("age_bucket_standard") or "UNKNOWN"] += qty
        detailed_counts[r.get("age_bucket") or "UNKNOWN"] += qty
        brand_counts[brand] += qty
        if aged:
            aged_models[model] += qty

        a = area_stats.setdefault(area, {"name": area, "qty": 0.0, "value": 0.0, "aged_90": 0.0, "aged_value": 0.0, "oldest": 0, "branches": set()})
        a["qty"] += qty
        a["value"] += value
        a["aged_90"] += qty if aged else 0
        a["aged_value"] += value if aged else 0
        a["oldest"] = max(a["oldest"], age)
        a["branches"].add(branch)

        b = branch_stats.setdefault(branch, {"name": branch, "area": area, "qty": 0.0, "value": 0.0, "aged_90": 0.0, "aged_value": 0.0, "oldest": 0})
        b["qty"] += qty
        b["value"] += value
        b["aged_90"] += qty if aged else 0
        b["aged_value"] += value if aged else 0
        b["oldest"] = max(b["oldest"], age)

        s = model_stats.setdefault(model, {
            "standard_description": model,
            "qty": 0.0,
            "value": 0.0,
            "age_weight": 0.0,
            "aged_90": 0.0,
            "aged_value": 0.0,
            "oldest": 0,
            "branches": set(),
            "areas": set(),
        })
        s["qty"] += qty
        s["value"] += value
        s["age_weight"] += qty * age
        s["aged_90"] += qty if aged else 0
        s["aged_value"] += value if aged else 0
        s["oldest"] = max(s["oldest"], age)
        s["branches"].add(branch)
        s["areas"].add(area)

    area_ranking = []
    for s in area_stats.values():
        qty = s["qty"]
        area_ranking.append({
            "name": s["name"], "qty": qty, "value": s["value"], "aged_90": s["aged_90"],
            "aged_90_pct": (s["aged_90"] / qty * 100) if qty else 0,
            "aged_value": s["aged_value"], "oldest": s["oldest"], "branch_count": len(s["branches"]),
        })
    area_ranking.sort(key=lambda x: (x["aged_90"], x["aged_90_pct"], x["aged_value"]), reverse=True)

    branch_ranking = []
    for s in branch_stats.values():
        qty = s["qty"]
        branch_ranking.append({
            "name": s["name"], "area": s["area"], "qty": qty, "value": s["value"], "aged_90": s["aged_90"],
            "aged_90_pct": (s["aged_90"] / qty * 100) if qty else 0,
            "aged_value": s["aged_value"], "oldest": s["oldest"],
        })
    branch_ranking.sort(key=lambda x: (x["aged_90"], x["aged_90_pct"], x["aged_value"]), reverse=True)

    model_summary = []
    for s in model_stats.values():
        qty = s["qty"]
        aged_90_pct_model = (s["aged_90"] / qty * 100) if qty else 0
        model_risk_level, _ = _risk_level(aged_90_pct_model)
        model_summary.append({
            "standard_description": s["standard_description"],
            "qty": qty,
            "value": s["value"],
            "avg_age": s["age_weight"] / qty if qty else 0,
            "aged_90": s["aged_90"],
            "aged_90_pct": aged_90_pct_model,
            "aged_value": s["aged_value"],
            "oldest": s["oldest"],
            "branch_count": len(s["branches"]),
            "area_count": len(s["areas"]),
            "risk_level": model_risk_level,
            "risk_rank": {"High": 3, "Watch": 2, "Controlled": 1}.get(model_risk_level, 0),
        })
    model_summary.sort(key=lambda x: (x["aged_90"], x["aged_90_pct"], x["aged_value"]), reverse=True)
    model_risk_counts = Counter(x["risk_level"] for x in model_summary)

    top_area = area_ranking[0] if area_ranking else None
    top_branch = branch_ranking[0] if branch_ranking else None
    top_model = model_summary[0] if model_summary else None

    return {
        "total_qty": total_qty,
        "total_value": total_value,
        "avg_age": avg_age,
        "aged_90": aged_90,
        "aged_180": aged_180,
        "aged_365": aged_365,
        "aged_90_value": aged_90_value,
        "aged_180_value": aged_180_value,
        "aged_value_pct": aged_value_pct,
        "healthy_qty": healthy_qty,
        "healthy_pct": healthy_pct,
        "oldest": oldest,
        "branch_count": branch_count,
        "area_count": area_count,
        "aged_90_pct": aged_90_pct,
        "aged_365_pct": aged_365_pct,
        "risk_level": risk_level,
        "risk_message": risk_message,
        "bucket_labels": STANDARD_BUCKETS,
        "bucket_values": [standard_counts.get(x, 0) for x in STANDARD_BUCKETS],
        "detailed_bucket_labels": DETAILED_BUCKETS,
        "detailed_bucket_values": [detailed_counts.get(x, 0) for x in DETAILED_BUCKETS],
        "area_labels": [x["name"] for x in area_ranking[:10]],
        "area_values": [x["aged_90"] for x in area_ranking[:10]],
        "area_pct_values": [x["aged_90_pct"] for x in area_ranking[:10]],
        "brand_labels": [x for x, _ in brand_counts.most_common(8)],
        "brand_values": [v for _, v in brand_counts.most_common(8)],
        "aged_model_labels": [x for x, _ in sorted(aged_models.items(), key=lambda kv: kv[1], reverse=True)[:10]],
        "aged_model_values": [v for _, v in sorted(aged_models.items(), key=lambda kv: kv[1], reverse=True)[:10]],
        "area_ranking": area_ranking[:12],
        "area_ranking_all": area_ranking,
        "branch_ranking": branch_ranking[:15],
        "branch_ranking_all": branch_ranking,
        "model_summary": model_summary[:75],
        "model_summary_all": model_summary,
        "model_risk_counts": {
            "High": model_risk_counts.get("High", 0),
            "Watch": model_risk_counts.get("Watch", 0),
            "Controlled": model_risk_counts.get("Controlled", 0),
        },
        "top_area": top_area,
        "top_branch": top_branch,
        "top_model": top_model,
    }


def filter_options(selected_area: str = "") -> dict[str, Any]:
    area_expr = _area_expr()
    branch_where = ""
    params: list[Any] = []
    if selected_area:
        branch_where = f"WHERE {area_expr} = ?"
        params.append(selected_area)

    with connect() as conn:
        branches = [dict(r) for r in conn.execute(f"""
            SELECT u.branch_key,
                   COALESCE(NULLIF(MAX(u.branch_original),''), u.branch_key) AS branch_name,
                   {area_expr} AS area
            FROM units u
            LEFT JOIN branch_area b ON b.branch_key=u.branch_key
            {branch_where}
            GROUP BY u.branch_key, {area_expr}
            ORDER BY branch_name
        """, params).fetchall()]
        areas = [r[0] for r in conn.execute(f"""
            SELECT DISTINCT {area_expr} AS area
            FROM units u LEFT JOIN branch_area b ON b.branch_key=u.branch_key
            WHERE {area_expr} IS NOT NULL AND {area_expr}<>''
            ORDER BY area
        """).fetchall() if r[0]]
        brands = [r[0] for r in conn.execute("SELECT DISTINCT brand FROM units WHERE brand IS NOT NULL AND brand<>'' ORDER BY brand").fetchall() if r[0]]
    return {"branches": branches, "areas": areas, "brands": brands}


def executive_summary() -> dict[str, Any] | None:
    """Return an unfiltered management summary for the latest Aging import."""
    with connect() as conn:
        row = conn.execute("SELECT as_of_date, filename, imported_at FROM imports ORDER BY id DESC LIMIT 1").fetchone()
    if not row:
        return None
    as_of = parse_date(row["as_of_date"])
    if as_of is None:
        return None
    rows = filtered_dataset({}, as_of, "branch")
    if not rows:
        return None
    out = summarize(rows)
    out["as_of_date"] = as_of.isoformat()
    out["source_filename"] = row["filename"]
    out["imported_at"] = row["imported_at"]
    return out
