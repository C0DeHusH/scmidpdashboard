from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime, timezone, timedelta
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
    "MUTI MC : Stock Outrate - Overall (Before PO Balance)",
]

KPI_META = {
    "MC Class A Doi": {"unit": "days", "label": "MC Class A DoI", "good": "high"},
    "MUTI MC : DoI": {"unit": "days", "label": "MUTI MC DoI", "good": "balanced"},
    "Overall Class A Stock Out Rate": {"unit": "percent", "label": "Overall Class A Stock-Out", "good": "low"},
    "MUTI MC : Stock Outrate - Per Branch": {"unit": "percent", "label": "Stock-Out Rate · Per Branch", "good": "low"},
    "MUTI MC : Stock Outrate - Overall after PO Balance": {"unit": "percent", "label": "Stock-Out · After PO Balance", "good": "low"},
    "MUTI MC : Stock Outrate - Overall (Before PO Balance)": {"unit": "percent", "label": "Stock-Out · Before PO Balance", "good": "low"},
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


def _ceil_doi(value: Any) -> int:
    """Approved DoI display rule: always round fractional days upward."""
    return int(math.ceil(max(0.0, _num(value))))


@dataclass
class ImportValidation:
    ok: bool
    message: str


class DashboardStore:
    def __init__(self, workbook_path: str | Path | None):
        self._lock = RLock()
        self.workbook_path: Optional[Path] = None
        self.raw_records: List[Dict[str, Any]] = []
        self.management_records: List[Dict[str, Any]] = []
        self.management_sheet_name: str = ""
        self.management_title: str = "Management Order Planning"
        self.kpis: Dict[str, Dict[str, Any]] = {}
        self.areas: List[str] = []
        self.branches: List[str] = []
        self.generated_at = datetime.now(timezone(timedelta(hours=8)))
        if workbook_path is None:
            self.clear()
        else:
            self.load(workbook_path)

    @staticmethod
    def _empty_kpis() -> Dict[str, Dict[str, Any]]:
        # Build fresh lists per KPI/period; shared empty lists can leak state if a
        # renderer or future transformation mutates one series in place.
        def empty_series() -> Dict[str, Any]:
            return {"labels": [], "values": [], "trend": [], "latest": None, "delta": None}

        return {
            name: {
                "meta": KPI_META[name],
                "ytd": empty_series(),
                "weekly": empty_series(),
            }
            for name in TARGET_KPIS
        }

    def clear(self) -> None:
        """Put the dashboard into a persistent no-data presentation state."""
        with self._lock:
            self.workbook_path = None
            self.raw_records = []
            self.management_records = []
            self.management_sheet_name = ""
            self.management_title = "Management Order Planning"
            self.kpis = self._empty_kpis()
            self.areas = []
            self.branches = []
            self.generated_at = datetime.now(timezone(timedelta(hours=8)))

    @staticmethod
    def _find_management_sheet(reader: XlsxReader) -> Tuple[Optional[str], Optional[int]]:
        """Find the Management/Re-order sheet by its headers, not its worksheet name.

        This lets users rename Sheet2 to Management (or another name) without
        breaking the dashboard import, as long as the required columns remain.
        """
        needed = {
            "Standard Description", "Class", "Brand", "Cost",
            "Avg. Daily Sale (Qty)", "Inv. Qty Total", "DoI",
            "Stock Status", "PO Balance", "Allocation", "DoI after PO Bal",
        }
        for sheet_name in reader.sheet_names:
            if sheet_name in {"Raw", "KPI_YTD_Input", "KPI_WEEKLY_Input"}:
                continue
            try:
                rows = reader.read_sheet(sheet_name).rows
            except Exception:
                continue
            for header_idx in range(min(6, len(rows))):
                headers = {_clean(x) for x in rows[header_idx] if _clean(x)}
                if needed.issubset(headers):
                    return sheet_name, header_idx
        return None, None

    @staticmethod
    def validate(path: str | Path) -> ImportValidation:
        try:
            r = XlsxReader(path)
            # v2.44 uses one consolidated workbook for the full control tower.
            # Aging is now part of the same import contract rather than a separate upload.
            required = {"Raw", "KPI_YTD_Input", "KPI_WEEKLY_Input", "Aging"}
            missing = required - set(r.sheet_names)
            if missing:
                return ImportValidation(False, "Unified import is missing worksheet(s): " + ", ".join(sorted(missing)))
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

            aging_rows = r.read_sheet("Aging").rows
            if len(aging_rows) < 2:
                return ImportValidation(False, "Aging sheet does not contain inventory rows.")
            aging_headers = {_clean(x).upper() for x in aging_rows[0] if _clean(x)}
            aging_needed = {"BRANCH", "INCOMING DATE", "CREATED ON", "STANDARD DESCRIPTION"}
            aging_missing = aging_needed - aging_headers
            if aging_missing:
                return ImportValidation(False, "Aging sheet missing column(s): " + ", ".join(sorted(aging_missing)))

            management_sheet, management_header = DashboardStore._find_management_sheet(r)
            if management_sheet is None or management_header is None:
                return ImportValidation(True, "Valid unified SCM + Aging workbook. Management source not found; Management Order Plan will remain empty.")
            return ImportValidation(True, f"Valid unified SCM + Aging workbook. Management source: {management_sheet}.")
        except Exception as exc:
            return ImportValidation(False, f"Unable to read workbook: {exc}")

    def load(self, path: str | Path, *, prevalidated: bool = False) -> None:
        path = Path(path)
        if not prevalidated:
            valid = self.validate(path)
            if not valid.ok:
                raise ValueError(valid.message)
        reader = XlsxReader(path)
        raw_rows = reader.read_sheet("Raw").rows
        ytd_rows = reader.read_sheet("KPI_YTD_Input").rows
        weekly_rows = reader.read_sheet("KPI_WEEKLY_Input").rows
        management_sheet, management_header = self._find_management_sheet(reader)
        management_rows = reader.read_sheet(management_sheet).rows if management_sheet else []

        records = self._parse_raw(raw_rows)
        management_records, management_title = self._parse_management(management_rows, management_header or 0)
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
            self.management_records = management_records
            self.management_sheet_name = management_sheet or ""
            self.management_title = management_title
            self.kpis = kpis
            self.areas = sorted({r["area"] for r in records if r["area"]})
            self.branches = sorted({r["branch"] for r in records if r["branch"]})
            self.generated_at = datetime.now(timezone(timedelta(hours=8)))

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
                "doi": _ceil_doi(get(row, "DoI (Branch)")),
                "stock_status": _clean(get(row, "Stock Status (branch)")),
                "suggested_transfer": _num(get(row, "Suggested Transfer")),
            })
        return records

    def _parse_management(self, rows: List[List[Any]], header_idx: int) -> Tuple[List[Dict[str, Any]], str]:
        if not rows or header_idx >= len(rows):
            return [], "Management Order Planning"
        headers = [_clean(x) for x in rows[header_idx]]
        first_idx: Dict[str, int] = {}
        for i, h in enumerate(headers):
            if h and h not in first_idx:
                first_idx[h] = i

        def get(row: List[Any], header: str) -> Any:
            i = first_idx.get(header)
            return row[i] if i is not None and i < len(row) else None

        title = "Management Order Planning"
        if header_idx > 0 and rows[header_idx - 1]:
            candidate = _clean(rows[header_idx - 1][0])
            if candidate:
                title = candidate

        out: List[Dict[str, Any]] = []
        for row in rows[header_idx + 1:]:
            model = _clean(get(row, "Standard Description"))
            brand = _clean(get(row, "Brand"))
            if not model:
                continue
            cls = _class_key(get(row, "Class"))
            key = f"{brand or 'Unspecified'}||{model}"
            out.append({
                "key": key,
                "model": model,
                "abc": _clean(get(row, "ABC")),
                "rank": int(_num(get(row, "Rank"), 999999)),
                "class": cls,
                "class_label": _clean(get(row, "Class")) or (f"Class {cls}" if cls else ""),
                "brand": brand or "Unspecified",
                "unit_cost": _num(get(row, "Cost")),
                "avg_daily_sale": _num(get(row, "Avg. Daily Sale (Qty)")),
                "inventory": _num(get(row, "Inv. Qty Total")),
                "doi": _ceil_doi(get(row, "DoI")),
                "stock_status": _clean(get(row, "Stock Status")),
                "po_balance": _num(get(row, "PO Balance")),
                "source_allocation": _num(get(row, "Allocation")),
                "source_new_doi": _ceil_doi(get(row, "DoI after PO Bal")),
                "stock_status_after_po": _clean(get(row, "Stock Status after PO Bal")),
                "reorder": _num(get(row, "Re-order")),
            })
        return out, title

    def management_dashboard(
        self,
        brand: str = "All Brands",
        class_key: str = "All Classes",
        status: str = "All Statuses",
        model: str = "All Models",
        allocations: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        """Management order-planning view sourced from the workbook's Management/Re-order sheet.

        v2.11 locks Class, Brand, Model, DoI and Stock Status as workbook
        reference fields. Management may edit Unit Cost, Current Inventory,
        PO Balance, Order Quantity and Remarks while New DoI and Total Amount
        remain formula-driven. Saved edits stay outside the imported workbook.
        """
        with self._lock:
            records = [dict(r) for r in self.management_records]
            title = self.management_title
            sheet_name = self.management_sheet_name
        allocations = allocations or {}

        def effective_record(r: Dict[str, Any]) -> Dict[str, Any]:
            row = dict(r)
            saved = allocations.get(r["key"], None)
            remarks = ""
            allocation = _num(r.get("source_allocation"))
            if isinstance(saved, dict):
                # Product classification and stock-condition fields are source-of-truth
                # references. Only management planning inputs may override workbook values.
                for field in ("unit_cost", "inventory", "po_balance"):
                    if field in saved:
                        row[field] = max(0.0, _num(saved.get(field)))
                if "quantity" in saved or "allocation" in saved:
                    allocation = max(0.0, _num(saved.get("quantity", saved.get("allocation", allocation))))
                remarks = _clean(saved.get("remarks", ""))
            elif saved is not None:
                allocation = max(0.0, _num(saved))
            row["allocation"] = round(allocation, 6)
            row["remarks"] = remarks
            inv_after_po = row["inventory"] + row["po_balance"] + allocation
            row["inventory_after_po"] = round(inv_after_po, 6)
            row["new_doi"] = _ceil_doi((inv_after_po / row["avg_daily_sale"]) if row["avg_daily_sale"] > 0 else 0.0)
            row["total_amount"] = round(row["unit_cost"] * allocation, 4)
            return row

        effective = [effective_record(r) for r in records]
        all_brands = sorted({r["brand"] for r in effective if r["brand"]})
        all_models = sorted({r["model"] for r in effective if r["model"]})
        all_statuses = sorted({r["stock_status"] for r in effective if r["stock_status"]})

        def matches(r: Dict[str, Any]) -> bool:
            if brand and brand != "All Brands" and r["brand"] != brand:
                return False
            if class_key and class_key != "All Classes" and r["class"] != class_key:
                return False
            if status and status != "All Statuses" and r["stock_status"].lower() != status.lower():
                return False
            if model and model != "All Models" and r["model"] != model:
                return False
            return True

        rows = [r for r in effective if matches(r)]
        class_order = {"A": 0, "B": 1, "C": 2, None: 3}
        rows.sort(key=lambda r: (r["brand"], class_order.get(r["class"], 3), r["rank"], r["model"]))
        summary = {
            "models": len(rows),
            "current_inventory": round(sum(r["inventory"] for r in rows), 4),
            "po_balance": round(sum(r["po_balance"] for r in rows), 4),
            "allocation_order": round(sum(r["allocation"] for r in rows), 4),
            "grand_total": round(sum(r["total_amount"] for r in rows), 4),
        }
        return {
            "source_sheet": sheet_name,
            "title": title,
            "selected": {"brand": brand, "class": class_key, "status": status, "model": model},
            "filters": {
                "brands": ["All Brands"] + all_brands,
                "classes": ["All Classes", "A", "B", "C"],
                "statuses": ["All Statuses"] + all_statuses,
                "models": ["All Models"] + all_models,
            },
            "summary": summary,
            "rows": rows,
        }

    def reorder_model_card(self) -> Dict[str, Any]:
        """Executive ABC model snapshot sourced directly from the imported Re-order sheet.

        Rank is preserved from the workbook so the dashboard never maintains a
        second ranking source. Rows are sorted A -> B -> C, then imported Rank.
        """
        with self._lock:
            records = [dict(r) for r in self.management_records]
            title = self.management_title
            sheet_name = self.management_sheet_name

        class_order = {"A": 0, "B": 1, "C": 2}
        rows = [r for r in records if r.get("class") in class_order]
        rows.sort(key=lambda r: (class_order.get(r.get("class"), 9), int(r.get("rank", 999999)), r.get("model", "")))

        class_counts = {cls: sum(1 for r in rows if r.get("class") == cls) for cls in ("A", "B", "C")}
        class_avg_doi = {}
        for cls in ("A", "B", "C"):
            values = [float(r.get("doi", 0) or 0) for r in rows if r.get("class") == cls]
            class_avg_doi[cls] = _ceil_doi(sum(values) / len(values)) if values else 0

        return {
            "source_sheet": sheet_name,
            "title": title,
            "summary": {"counts": class_counts, "avg_doi": class_avg_doi, "models": len(rows)},
            "rows": [
                {
                    "model": r.get("model", ""),
                    "brand": r.get("brand", ""),
                    "class": r.get("class", ""),
                    "rank": int(r.get("rank", 999999)),
                    "doi": _ceil_doi(r.get("doi", 0)),
                    "stock_status": r.get("stock_status", ""),
                }
                for r in rows
            ],
        }

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
        display_values = [_pct(v) for v in values] if meta["unit"] == "percent" else [_ceil_doi(v) for v in values]
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
                    "doi": _ceil_doi(r["doi"]),
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
            "doi": _ceil_doi(r["doi"]),
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
        """Area-first Brand + Model stock-status intelligence from Raw/Distribution records."""
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
            if status and status != "All Statuses" and (r["stock_status"] or "").lower() != status.lower():
                return False
            return True

        filtered = [r for r in records if matches(r)]

        def summarize(rows: List[Dict[str, Any]]) -> Dict[str, Any]:
            valid = [r for r in rows if r["stock_status"]]
            statuses: Dict[str, int] = defaultdict(int)
            for r in valid:
                statuses[r["stock_status"]] += 1
            so = sum(1 for r in valid if _is_stockout(r["stock_status"]))
            return {
                "records": len(rows),
                "branches": len({r["branch"] for r in rows if r["branch"]}),
                "models": len({(r["brand"] or "Unspecified", r["model"]) for r in rows if r["model"]}),
                "brands": len({r["brand"] or "Unspecified" for r in rows}),
                "areas": len({r["area"] for r in rows if r["area"]}),
                "inventory": round(sum(r["inventory"] for r in rows), 4),
                "suggested_transfer": round(sum(r["suggested_transfer"] for r in rows), 4),
                "avg_doi": _ceil_doi(sum(r["doi"] for r in rows) / len(rows)) if rows else 0,
                "stockout_count": so,
                "stockout_rate": _pct(so / len(valid)) if valid else 0.0,
                "statuses": dict(sorted(statuses.items())),
            }

        summary = summarize(filtered)
        status_breakdown = summary["statuses"]

        brand_groups: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
        model_groups: Dict[Tuple[str, str], List[Dict[str, Any]]] = defaultdict(list)
        area_status_groups: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
        area_brand_class_groups: Dict[Tuple[str, str, str], List[Dict[str, Any]]] = defaultdict(list)
        area_model_groups: Dict[Tuple[str, str, str, str], List[Dict[str, Any]]] = defaultdict(list)

        for r in filtered:
            b = r["brand"] or "Unspecified"
            a = r["area"] or "Unassigned Area"
            c = r["class"] or "—"
            brand_groups[b].append(r)
            model_groups[(b, r["model"])].append(r)
            area_status_groups[a].append(r)
            area_brand_class_groups[(a, b, c)].append(r)
            area_model_groups[(a, b, r["model"], c)].append(r)

        brand_rows = []
        for b, rows in brand_groups.items():
            m = summarize(rows)
            m.update({"brand": b})
            brand_rows.append(m)
        brand_rows.sort(key=lambda x: (-x["stockout_rate"], -x["stockout_count"], x["brand"]))

        model_rows = []
        for (b, mod), rows in model_groups.items():
            m = summarize(rows)
            classes = sorted({r["class"] for r in rows if r["class"]})
            m.update({"brand": b, "model": mod, "class": "/".join(classes) if classes else "—"})
            model_rows.append(m)
        model_rows.sort(key=lambda x: (-x["stockout_rate"], -x["stockout_count"], x["brand"], x["model"]))

        area_status_rows = []
        for a, rows in area_status_groups.items():
            m = summarize(rows)
            m.update({"area": a})
            area_status_rows.append(m)
        area_status_rows.sort(key=lambda x: x["area"])

        area_brand_rows = []
        for (a, b, c), rows in area_brand_class_groups.items():
            m = summarize(rows)
            m.update({"area": a, "brand": b, "class": c})
            area_brand_rows.append(m)
        area_brand_rows.sort(key=lambda x: (x["area"], x["class"] != "A", x["class"], -x["stockout_rate"], -x["stockout_count"], x["brand"]))

        area_model_rows = []
        for (a, b, mod, c), rows in area_model_groups.items():
            m = summarize(rows)
            m.update({"area": a, "brand": b, "model": mod, "class": c})
            area_model_rows.append(m)
        area_model_rows.sort(key=lambda x: (x["area"], x["class"] != "A", x["class"], -x["stockout_rate"], -x["stockout_count"], x["brand"], x["model"]))

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
            "status_breakdown": status_breakdown,
            "brands": brand_rows,
            "models": model_rows,
            "area_status": area_status_rows,
            "area_brands": area_brand_rows,
            "area_models": area_model_rows,
        }



    def bootstrap(self, role: str) -> Dict[str, Any]:
        branch = self.branches[0] if self.branches else ""
        return {
            "role": role,
            "generated_at": self.generated_at.isoformat(),
            "has_data": bool(self.raw_records or self.management_records or any((v.get("ytd", {}).get("values") or v.get("weekly", {}).get("values")) for v in self.kpis.values())),
            "areas": ["Overall"] + self.areas,
            "branches": self.branches,
            "kpis": self.kpis,
            "area": self.area_dashboard("Overall"),
            "branch": self.branch_dashboard(branch),
            "status_summary": self.status_summary(),
            "reorder_card": self.reorder_model_card(),
        }
