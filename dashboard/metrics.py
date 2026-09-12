from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from threading import RLock
from typing import Any, Dict, Iterable, List, Optional, Tuple
import math

from .xlsx_reader import XlsxReader, excel_serial_to_date

TARGET_KPIS = [
    "MC Class A Doi",
    "MUTI MC : DoI",
    "Overall Class A Stock Out Rate",
    "MUTI MC : Stock Outrate - Per Branch",
    "MUTI MC : Stock Outrate - Overall after PO Balance",
]

KPI_META = {
    "MC Class A Doi": {"unit": "days", "label": "MC Class A DoI", "good": "low"},
    "MUTI MC : DoI": {"unit": "days", "label": "MUTI MC DoI", "good": "balanced"},
    "Overall Class A Stock Out Rate": {"unit": "percent", "label": "Overall Class A Stock-Out", "good": "low"},
    "MUTI MC : Stock Outrate - Per Branch": {"unit": "percent", "label": "Stock-Out Rate · Per Branch", "good": "low"},
    "MUTI MC : Stock Outrate - Overall after PO Balance": {"unit": "percent", "label": "Stock-Out · After PO Balance", "good": "low"},
}


def _num(value: Any, default: float = 0.0) -> float:
    if value is None or value == "":
        return default
    try:
        x = float(value)
        return x if math.isfinite(x) else default
    except (TypeError, ValueError):
        return default


def _clean(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _is_stockout(value: Any) -> bool:
    s = _clean(value).lower().replace("-", " ").replace("_", " ")
    return "stockout" in s.replace(" ", "") or "stock out" in s


def _class_key(value: Any) -> Optional[str]:
    s = _clean(value).upper()
    if s.endswith("A") or s == "A" or "CLASS A" in s:
        return "A"
    if s.endswith("B") or s == "B" or "CLASS B" in s:
        return "B"
    if s.endswith("C") or s == "C" or "CLASS C" in s:
        return "C"
    return None


def _trend(values: List[float]) -> List[float]:
    n = len(values)
    if n <= 1:
        return values[:]
    sx = n * (n - 1) / 2
    sy = sum(values)
    sxx = sum(i * i for i in range(n))
    sxy = sum(i * y for i, y in enumerate(values))
    denom = n * sxx - sx * sx
    slope = 0.0 if denom == 0 else (n * sxy - sx * sy) / denom
    intercept = (sy - slope * sx) / n
    return [intercept + slope * i for i in range(n)]


def _pct(rate: float) -> float:
    return round(rate * 100.0, 4)


@dataclass
class ImportValidation:
    ok: bool
    message: str


class DashboardStore:
    def __init__(self, workbook_path: str | Path):
        self._lock = RLock()
        self.workbook_path = Path(workbook_path)
        self.raw_records: List[Dict[str, Any]] = []
        self.kpis: Dict[str, Dict[str, Any]] = {}
        self.areas: List[str] = []
        self.branches: List[str] = []
        self.generated_at = datetime.now()
        self.load(self.workbook_path)

    @staticmethod
    def validate(path: str | Path) -> ImportValidation:
        try:
            r = XlsxReader(path)
            required = {"Raw", "KPI_YTD_Input", "KPI_WEEKLY_Input"}
            missing = required - set(r.sheet_names)
            if missing:
                return ImportValidation(False, "Missing worksheet(s): " + ", ".join(sorted(missing)))
            raw = r.read_sheet("Raw").rows
            if len(raw) < 3:
                return ImportValidation(False, "Raw sheet does not contain enough rows.")
            headers = [str(x).strip() if x is not None else "" for x in raw[1]]
            needed = {
                "Standard Description", "Branch", "Area", "RANK", "CLASS",
                "Avg. Daily Sale (Qty)", "Inv. Qty Total", "DoI (Branch)",
                "Stock Status (branch)", "Suggested Transfer"
            }
            missing_cols = needed - set(headers)
            if missing_cols:
                return ImportValidation(False, "Raw sheet missing column(s): " + ", ".join(sorted(missing_cols)))
            return ImportValidation(True, "Valid SCM dashboard import workbook.")
        except Exception as exc:
            return ImportValidation(False, f"Unable to read workbook: {exc}")

    def load(self, path: str | Path) -> None:
        path = Path(path)
        valid = self.validate(path)
        if not valid.ok:
            raise ValueError(valid.message)
        reader = XlsxReader(path)
        raw_rows = reader.read_sheet("Raw").rows
        ytd_rows = reader.read_sheet("KPI_YTD_Input").rows
        weekly_rows = reader.read_sheet("KPI_WEEKLY_Input").rows

        records = self._parse_raw(raw_rows)
        kpis = {}
        for name in TARGET_KPIS:
            kpis[name] = {
                "meta": KPI_META[name],
                "ytd": self._extract_kpi(ytd_rows, name),
                "weekly": self._extract_kpi(weekly_rows, name),
            }

        with self._lock:
            self.workbook_path = path
            self.raw_records = records
            self.kpis = kpis
            self.areas = sorted({r["area"] for r in records if r["area"]})
            self.branches = sorted({r["branch"] for r in records if r["branch"]})
            self.generated_at = datetime.now()

    def _parse_raw(self, rows: List[List[Any]]) -> List[Dict[str, Any]]:
        headers = [_clean(x) for x in rows[1]]
        first_idx: Dict[str, int] = {}
        for i, h in enumerate(headers):
            if h and h not in first_idx:
                first_idx[h] = i

        def get(row: List[Any], header: str) -> Any:
            i = first_idx.get(header)
            return row[i] if i is not None and i < len(row) else None

        records: List[Dict[str, Any]] = []
        for row in rows[2:]:
            model = _clean(get(row, "Standard Description"))
            branch = _clean(get(row, "Branch"))
            if not model or not branch:
                continue
            records.append({
                "model": model,
                "branch": branch,
                "area": _clean(get(row, "Area")),
                "abc": _clean(get(row, "ABC")),
                "rank": int(_num(get(row, "RANK"), 999999)),
                "class": _class_key(get(row, "CLASS")),
                "class_label": _clean(get(row, "CLASS")),
                "brand": _clean(get(row, "BRAND")),
                "avg_daily_sale": _num(get(row, "Avg. Daily Sale (Qty)")),
                "inventory": _num(get(row, "Inv. Qty Total")),
                "doi": _num(get(row, "DoI (Branch)")),
                "stock_status": _clean(get(row, "Stock Status (branch)")),
                "suggested_transfer": _num(get(row, "Suggested Transfer")),
            })
        return records

    def _extract_kpi(self, rows: List[List[Any]], target: str) -> Dict[str, Any]:
        idx = None
        for i, row in enumerate(rows):
            if row and _clean(row[0]).lower() == target.lower():
                idx = i
                break
        if idx is None:
            return {"labels": [], "values": [], "trend": [], "latest": None, "delta": None}

        value_row = rows[idx]
        date_row = None
        # Block headers place dates on the next row. Child KPI rows inherit the
        # closest preceding date row in the same block.
        if idx + 1 < len(rows) and self._looks_like_dates(rows[idx + 1]):
            date_row = rows[idx + 1]
        else:
            for j in range(idx - 1, max(-1, idx - 14), -1):
                if self._looks_like_dates(rows[j]):
                    date_row = rows[j]
                    break
        if date_row is None:
            return {"labels": [], "values": [], "trend": [], "latest": None, "delta": None}

        labels: List[str] = []
        values: List[float] = []
        max_len = max(len(value_row), len(date_row))
        for c in range(2, max_len):
            d = date_row[c] if c < len(date_row) else None
            v = value_row[c] if c < len(value_row) else None
            dt = excel_serial_to_date(d)
            if dt is None or v is None or v == "":
                continue
            try:
                fv = float(v)
            except (TypeError, ValueError):
                continue
            labels.append(dt.strftime("%b %d, %Y"))
            values.append(fv)

        meta = KPI_META[target]
        display_values = [_pct(v) for v in values] if meta["unit"] == "percent" else [round(v, 3) for v in values]
        trend = _trend(display_values)
        latest = display_values[-1] if display_values else None
        delta = (display_values[-1] - display_values[-2]) if len(display_values) > 1 else None
        return {
            "labels": labels,
            "values": display_values,
            "trend": [round(x, 3) for x in trend],
            "latest": latest,
            "delta": round(delta, 3) if delta is not None else None,
        }

    @staticmethod
    def _looks_like_dates(row: List[Any]) -> bool:
        vals = row[2:] if len(row) > 2 else []
        hits = 0
        for v in vals:
            try:
                if 30000 <= float(v) <= 80000:
                    hits += 1
            except (TypeError, ValueError):
                pass
        return hits >= 2

    def _summary(self, records: Iterable[Dict[str, Any]]) -> Dict[str, Any]:
        data = list(records)
        by_class: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
        for r in data:
            if r["class"] in ("A", "B", "C"):
                by_class[r["class"]].append(r)
        rates = {}
        counts = {}
        for cls in ("A", "B", "C"):
            valid = [r for r in by_class[cls] if r["stock_status"]]
            stockouts = sum(1 for r in valid if _is_stockout(r["stock_status"]))
            rate = stockouts / len(valid) if valid else 0.0
            rates[cls] = _pct(rate)
            counts[cls] = {"stockout": stockouts, "status_count": len(valid)}
        return {
            "class_rates": rates,
            "counts": counts,
            "overall_avg": round(sum(rates.values()) / 3.0, 4),
            "record_count": len(data),
        }

    def area_dashboard(self, area: str = "Overall") -> Dict[str, Any]:
        with self._lock:
            records = list(self.raw_records)
        if area and area != "Overall":
            filtered = [r for r in records if r["area"] == area]
        else:
            filtered = records

        ranking = []
        for a in self.areas:
            s = self._summary(r for r in records if r["area"] == a)
            ranking.append({
                "area": a,
                "overall_avg": s["overall_avg"],
                "class_a": s["class_rates"]["A"],
                "class_b": s["class_rates"]["B"],
                "class_c": s["class_rates"]["C"],
            })
        ranking.sort(key=lambda x: x["overall_avg"], reverse=True)

        # Branch-level Class A stock-out exposure for the selected Area.
        # "Overall" returns all branches; a specific Area returns only its branches.
        branch_class_a = []
        scoped_branches = sorted({r["branch"] for r in filtered if r["branch"]})
        for branch in scoped_branches:
            bs = self._summary(r for r in filtered if r["branch"] == branch)
            branch_area = next((r["area"] for r in filtered if r["branch"] == branch), "")
            branch_class_a.append({
                "branch": branch,
                "area": branch_area,
                "class_a": bs["class_rates"]["A"],
                "overall_avg": bs["overall_avg"],
                "class_a_counts": bs["counts"]["A"],
            })
        branch_class_a.sort(key=lambda x: (-x["class_a"], x["branch"]))

        return {
            "selected": area or "Overall",
            "summary": self._summary(filtered),
            "ranking": ranking,
            "branch_class_a": branch_class_a,
        }

    def branch_dashboard(self, branch: Optional[str] = None) -> Dict[str, Any]:
        with self._lock:
            records = list(self.raw_records)
            branches = list(self.branches)
        if not branch or branch not in branches:
            branch = branches[0] if branches else ""
        filtered = [r for r in records if r["branch"] == branch]
        area = filtered[0]["area"] if filtered else ""
        classes: Dict[str, List[Dict[str, Any]]] = {"A": [], "B": [], "C": []}
        seen: Dict[str, set] = {"A": set(), "B": set(), "C": set()}
        for cls in ("A", "B", "C"):
            candidates = sorted(
                (r for r in filtered if r["class"] == cls),
                key=lambda x: (x["rank"], x["model"]),
            )
            for r in candidates:
                if r["model"] in seen[cls]:
                    continue
                seen[cls].add(r["model"])
                classes[cls].append({
                    "rank": r["rank"],
                    "model": r["model"],
                    "stock_status": r["stock_status"],
                    "inventory": round(r["inventory"], 3),
                    "suggested_transfer": round(r["suggested_transfer"], 3),
                    "doi": round(r["doi"], 3),
                    "avg_daily_sale": round(r["avg_daily_sale"], 6),
                    "brand": r["brand"],
                })
        return {
            "branch": branch,
            "area": area,
            "summary": self._summary(filtered),
            "classes": classes,
        }

    def model_lookup(self, branch: str, model: str) -> Optional[Dict[str, Any]]:
        with self._lock:
            recs = [r for r in self.raw_records if r["branch"] == branch and r["model"] == model]
        if not recs:
            return None
        r = sorted(recs, key=lambda x: x["rank"])[0]
        return {
            "branch": r["branch"], "area": r["area"], "model": r["model"],
            "class": f"Class {r['class']}" if r["class"] else r["class_label"],
            "rank": r["rank"], "inventory": round(r["inventory"], 3),
            "stock_status": r["stock_status"],
            "suggested_transfer": round(r["suggested_transfer"], 3),
            "doi": round(r["doi"], 3),
            "avg_daily_sale": round(r["avg_daily_sale"], 6),
        }

    def branch_models(self, branch: str) -> List[str]:
        with self._lock:
            return sorted({r["model"] for r in self.raw_records if r["branch"] == branch})


    def status_summary(
        self,
        area: str = "Overall",
        branch: str = "All Branches",
        brand: str = "All Brands",
        model: str = "All Models",
        class_key: str = "All Classes",
        status: str = "All Statuses",
    ) -> Dict[str, Any]:
        """Filterable Brand + Model stock-status summary from Raw/Distribution records."""
        with self._lock:
            records = list(self.raw_records)

        all_brands = sorted({r["brand"] or "Unspecified" for r in records})
        all_models = sorted({r["model"] for r in records if r["model"]})
        all_statuses = sorted({r["stock_status"] for r in records if r["stock_status"]})

        def matches(r: Dict[str, Any]) -> bool:
            r_brand = r["brand"] or "Unspecified"
            if area and area != "Overall" and r["area"] != area:
                return False
            if branch and branch != "All Branches" and r["branch"] != branch:
                return False
            if brand and brand != "All Brands" and r_brand != brand:
                return False
            if model and model != "All Models" and r["model"] != model:
                return False
            if class_key and class_key != "All Classes" and r["class"] != class_key:
                return False
            if status and status != "All Statuses" and r["stock_status"].lower() != status.lower():
                return False
            return True

        filtered = [r for r in records if matches(r)]
        valid_status = [r for r in filtered if r["stock_status"]]
        stockout_count = sum(1 for r in valid_status if _is_stockout(r["stock_status"]))
        avg_doi = sum(r["doi"] for r in filtered) / len(filtered) if filtered else 0.0
        summary = {
            "records": len(filtered),
            "models": len({(r["brand"] or "Unspecified", r["model"]) for r in filtered}),
            "branches": len({r["branch"] for r in filtered if r["branch"]}),
            "inventory": round(sum(r["inventory"] for r in filtered), 4),
            "suggested_transfer": round(sum(r["suggested_transfer"] for r in filtered), 4),
            "avg_doi": round(avg_doi, 4),
            "stockout_count": stockout_count,
            "stockout_rate": _pct(stockout_count / len(valid_status)) if valid_status else 0.0,
        }

        status_breakdown: Dict[str, int] = defaultdict(int)
        for r in valid_status:
            status_breakdown[r["stock_status"]] += 1

        brand_groups: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
        model_groups: Dict[Tuple[str, str], List[Dict[str, Any]]] = defaultdict(list)
        for r in filtered:
            b = r["brand"] or "Unspecified"
            brand_groups[b].append(r)
            model_groups[(b, r["model"])].append(r)

        brand_rows = []
        for b, rows in brand_groups.items():
            statuses: Dict[str, int] = defaultdict(int)
            valid = [r for r in rows if r["stock_status"]]
            for r in valid:
                statuses[r["stock_status"]] += 1
            so = sum(1 for r in valid if _is_stockout(r["stock_status"]))
            brand_rows.append({
                "brand": b,
                "models": len({r["model"] for r in rows}),
                "branches": len({r["branch"] for r in rows if r["branch"]}),
                "records": len(rows),
                "inventory": round(sum(r["inventory"] for r in rows), 4),
                "suggested_transfer": round(sum(r["suggested_transfer"] for r in rows), 4),
                "avg_doi": round(sum(r["doi"] for r in rows) / len(rows), 4) if rows else 0.0,
                "stockout_count": so,
                "stockout_rate": _pct(so / len(valid)) if valid else 0.0,
                "statuses": dict(sorted(statuses.items())),
            })
        brand_rows.sort(key=lambda x: (-x["stockout_rate"], -x["stockout_count"], x["brand"]))

        model_rows = []
        for (b, m), rows in model_groups.items():
            statuses: Dict[str, int] = defaultdict(int)
            valid = [r for r in rows if r["stock_status"]]
            for r in valid:
                statuses[r["stock_status"]] += 1
            so = sum(1 for r in valid if _is_stockout(r["stock_status"]))
            classes = sorted({r["class"] for r in rows if r["class"]})
            model_rows.append({
                "brand": b,
                "model": m,
                "class": "/".join(classes) if classes else "—",
                "branches": len({r["branch"] for r in rows if r["branch"]}),
                "records": len(rows),
                "inventory": round(sum(r["inventory"] for r in rows), 4),
                "suggested_transfer": round(sum(r["suggested_transfer"] for r in rows), 4),
                "avg_doi": round(sum(r["doi"] for r in rows) / len(rows), 4) if rows else 0.0,
                "stockout_count": so,
                "stockout_rate": _pct(so / len(valid)) if valid else 0.0,
                "statuses": dict(sorted(statuses.items())),
            })
        model_rows.sort(key=lambda x: (-x["stockout_rate"], -x["stockout_count"], x["brand"], x["model"]))

        return {
            "selected": {
                "area": area or "Overall",
                "branch": branch or "All Branches",
                "brand": brand or "All Brands",
                "model": model or "All Models",
                "class": class_key or "All Classes",
                "status": status or "All Statuses",
            },
            "filters": {
                "areas": ["Overall"] + self.areas,
                "branches": ["All Branches"] + self.branches,
                "brands": ["All Brands"] + all_brands,
                "models": ["All Models"] + all_models,
                "classes": ["All Classes", "A", "B", "C"],
                "statuses": ["All Statuses"] + all_statuses,
            },
            "summary": summary,
            "status_breakdown": dict(sorted(status_breakdown.items())),
            "brands": brand_rows,
            "models": model_rows,
        }


    def brand_model_performance(
        self,
        area: str = "Overall",
        branch: str = "All Branches",
        brand: str = "All Brands",
        model: str = "All Models",
        class_key: str = "All Classes",
    ) -> Dict[str, Any]:
        """Separate performance view for Brand and Model execution, distinct from status/network views."""
        with self._lock:
            records = list(self.raw_records)

        all_brands = sorted({r["brand"] or "Unspecified" for r in records})
        all_models = sorted({r["model"] for r in records if r["model"]})

        def matches(r: Dict[str, Any]) -> bool:
            r_brand = r["brand"] or "Unspecified"
            if area and area != "Overall" and r["area"] != area:
                return False
            if branch and branch != "All Branches" and r["branch"] != branch:
                return False
            if brand and brand != "All Brands" and r_brand != brand:
                return False
            if model and model != "All Models" and r["model"] != model:
                return False
            if class_key and class_key != "All Classes" and r["class"] != class_key:
                return False
            return True

        filtered = [r for r in records if matches(r)]

        def metrics(rows: List[Dict[str, Any]]) -> Dict[str, Any]:
            valid = [r for r in rows if r["stock_status"]]
            so = sum(1 for r in valid if _is_stockout(r["stock_status"]))
            class_a = [r for r in rows if r["class"] == "A" and r["stock_status"]]
            class_a_so = sum(1 for r in class_a if _is_stockout(r["stock_status"]))
            return {
                "records": len(rows),
                "branches": len({r["branch"] for r in rows if r["branch"]}),
                "areas": len({r["area"] for r in rows if r["area"]}),
                "inventory": round(sum(r["inventory"] for r in rows), 4),
                "suggested_transfer": round(sum(r["suggested_transfer"] for r in rows), 4),
                "avg_doi": round(sum(r["doi"] for r in rows) / len(rows), 4) if rows else 0.0,
                "stockout_count": so,
                "stockout_rate": _pct(so / len(valid)) if valid else 0.0,
                "class_a_records": len(class_a),
                "class_a_stockout_count": class_a_so,
                "class_a_stockout_rate": _pct(class_a_so / len(class_a)) if class_a else 0.0,
            }

        summary = metrics(filtered)
        summary.update({
            "brands": len({r["brand"] or "Unspecified" for r in filtered}),
            "models": len({(r["brand"] or "Unspecified", r["model"]) for r in filtered}),
        })

        brand_groups: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
        model_groups: Dict[Tuple[str, str], List[Dict[str, Any]]] = defaultdict(list)
        for r in filtered:
            b = r["brand"] or "Unspecified"
            brand_groups[b].append(r)
            model_groups[(b, r["model"])].append(r)

        brand_rows = []
        for b, rows in brand_groups.items():
            m = metrics(rows)
            m.update({"brand": b, "models": len({r["model"] for r in rows})})
            brand_rows.append(m)
        brand_rows.sort(key=lambda x: (-x["class_a_stockout_rate"], -x["stockout_rate"], -x["suggested_transfer"], x["brand"]))

        model_rows = []
        for (b, mod), rows in model_groups.items():
            m = metrics(rows)
            classes = sorted({r["class"] for r in rows if r["class"]})
            stock_statuses = sorted({r["stock_status"] for r in rows if r["stock_status"]})
            m.update({
                "brand": b,
                "model": mod,
                "class": "/".join(classes) if classes else "—",
                "statuses": stock_statuses,
            })
            model_rows.append(m)
        model_rows.sort(key=lambda x: (-x["class_a_stockout_rate"], -x["stockout_rate"], -x["suggested_transfer"], x["brand"], x["model"]))

        return {
            "selected": {"area": area, "branch": branch, "brand": brand, "model": model, "class": class_key},
            "filters": {
                "areas": ["Overall"] + self.areas,
                "branches": ["All Branches"] + self.branches,
                "brands": ["All Brands"] + all_brands,
                "models": ["All Models"] + all_models,
                "classes": ["All Classes", "A", "B", "C"],
            },
            "summary": summary,
            "brands": brand_rows,
            "models": model_rows,
        }

    def network_summary(
        self,
        area: str = "Overall",
        brand: str = "All Brands",
        model: str = "All Models",
        class_key: str = "All Classes",
        status: str = "All Statuses",
    ) -> Dict[str, Any]:
        """Brand / Model network view aggregated per Area for SCM network analysis."""
        with self._lock:
            records = list(self.raw_records)

        all_brands = sorted({r["brand"] or "Unspecified" for r in records})
        all_models = sorted({r["model"] for r in records if r["model"]})
        all_statuses = sorted({r["stock_status"] for r in records if r["stock_status"]})

        def matches(r: Dict[str, Any]) -> bool:
            r_brand = r["brand"] or "Unspecified"
            if area and area != "Overall" and r["area"] != area:
                return False
            if brand and brand != "All Brands" and r_brand != brand:
                return False
            if model and model != "All Models" and r["model"] != model:
                return False
            if class_key and class_key != "All Classes" and r["class"] != class_key:
                return False
            if status and status != "All Statuses" and r["stock_status"].lower() != status.lower():
                return False
            return True

        filtered = [r for r in records if matches(r)]
        scoped_areas = [area] if area and area != "Overall" else list(self.areas)
        present_areas = sorted({r["area"] for r in filtered if r["area"]})
        valid_status = [r for r in filtered if r["stock_status"]]
        stockouts = sum(1 for r in valid_status if _is_stockout(r["stock_status"]))

        summary = {
            "areas": len(present_areas),
            "brands": len({r["brand"] or "Unspecified" for r in filtered}),
            "models": len({(r["brand"] or "Unspecified", r["model"]) for r in filtered}),
            "branches": len({r["branch"] for r in filtered if r["branch"]}),
            "inventory": round(sum(r["inventory"] for r in filtered), 4),
            "suggested_transfer": round(sum(r["suggested_transfer"] for r in filtered), 4),
            "avg_doi": round(sum(r["doi"] for r in filtered) / len(filtered), 4) if filtered else 0.0,
            "stockout_count": stockouts,
            "stockout_rate": _pct(stockouts / len(valid_status)) if valid_status else 0.0,
        }

        area_rows = []
        for a in scoped_areas:
            rows = [r for r in filtered if r["area"] == a]
            if not rows:
                continue
            valid = [r for r in rows if r["stock_status"]]
            so = sum(1 for r in valid if _is_stockout(r["stock_status"]))
            statuses: Dict[str, int] = defaultdict(int)
            for r in valid:
                statuses[r["stock_status"]] += 1
            area_rows.append({
                "area": a,
                "brands": len({r["brand"] or "Unspecified" for r in rows}),
                "models": len({(r["brand"] or "Unspecified", r["model"]) for r in rows}),
                "branches": len({r["branch"] for r in rows if r["branch"]}),
                "records": len(rows),
                "inventory": round(sum(r["inventory"] for r in rows), 4),
                "suggested_transfer": round(sum(r["suggested_transfer"] for r in rows), 4),
                "avg_doi": round(sum(r["doi"] for r in rows) / len(rows), 4),
                "stockout_count": so,
                "stockout_rate": _pct(so / len(valid)) if valid else 0.0,
                "statuses": dict(sorted(statuses.items())),
            })
        area_rows.sort(key=lambda x: (-x["stockout_rate"], x["area"]))

        grouped: Dict[Tuple[str, str], List[Dict[str, Any]]] = defaultdict(list)
        for r in filtered:
            grouped[(r["brand"] or "Unspecified", r["model"])].append(r)

        matrix_rows = []
        for (b, m), rows in grouped.items():
            classes = sorted({r["class"] for r in rows if r["class"]})
            cells = {}
            for a in scoped_areas:
                ar = [r for r in rows if r["area"] == a]
                if not ar:
                    continue
                valid = [r for r in ar if r["stock_status"]]
                so = sum(1 for r in valid if _is_stockout(r["stock_status"]))
                statuses: Dict[str, int] = defaultdict(int)
                for r in valid:
                    statuses[r["stock_status"]] += 1
                cells[a] = {
                    "branches": len({r["branch"] for r in ar if r["branch"]}),
                    "records": len(ar),
                    "inventory": round(sum(r["inventory"] for r in ar), 4),
                    "suggested_transfer": round(sum(r["suggested_transfer"] for r in ar), 4),
                    "avg_doi": round(sum(r["doi"] for r in ar) / len(ar), 4),
                    "stockout_count": so,
                    "stockout_rate": _pct(so / len(valid)) if valid else 0.0,
                    "statuses": dict(sorted(statuses.items())),
                }
            valid_all = [r for r in rows if r["stock_status"]]
            so_all = sum(1 for r in valid_all if _is_stockout(r["stock_status"]))
            matrix_rows.append({
                "brand": b,
                "model": m,
                "class": "/".join(classes) if classes else "—",
                "areas": len({r["area"] for r in rows if r["area"]}),
                "branches": len({r["branch"] for r in rows if r["branch"]}),
                "inventory": round(sum(r["inventory"] for r in rows), 4),
                "suggested_transfer": round(sum(r["suggested_transfer"] for r in rows), 4),
                "avg_doi": round(sum(r["doi"] for r in rows) / len(rows), 4) if rows else 0.0,
                "stockout_rate": _pct(so_all / len(valid_all)) if valid_all else 0.0,
                "cells": cells,
            })
        matrix_rows.sort(key=lambda x: (-x["stockout_rate"], x["brand"], x["model"]))

        return {
            "selected": {
                "area": area or "Overall",
                "brand": brand or "All Brands",
                "model": model or "All Models",
                "class": class_key or "All Classes",
                "status": status or "All Statuses",
            },
            "filters": {
                "areas": ["Overall"] + self.areas,
                "brands": ["All Brands"] + all_brands,
                "models": ["All Models"] + all_models,
                "classes": ["All Classes", "A", "B", "C"],
                "statuses": ["All Statuses"] + all_statuses,
            },
            "areas": scoped_areas,
            "summary": summary,
            "area_rows": area_rows,
            "matrix": matrix_rows,
        }

    def bootstrap(self, role: str) -> Dict[str, Any]:
        branch = self.branches[0] if self.branches else ""
        return {
            "role": role,
            "generated_at": self.generated_at.isoformat(),
            "areas": ["Overall"] + self.areas,
            "branches": self.branches,
            "kpis": self.kpis,
            "area": self.area_dashboard("Overall"),
            "branch": self.branch_dashboard(branch),
            "status_summary": self.status_summary(),
        }
