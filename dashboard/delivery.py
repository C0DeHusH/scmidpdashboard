from __future__ import annotations

from collections import defaultdict
from pathlib import Path
from threading import RLock
from typing import Any, Dict, Iterable, List, Optional, Tuple
import csv
import json
import math
import shutil

from .xlsx_reader import XlsxReader

DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]


def _clean(v: Any) -> str:
    return "" if v is None else str(v).strip()


def _num(v: Any, default: float = 0.0) -> float:
    try:
        x = float(v)
        return x if math.isfinite(x) else default
    except (TypeError, ValueError):
        return default


def _norm_header(v: Any) -> str:
    return " ".join(_clean(v).lower().replace("_", " ").replace("-", " ").split())


def _class_rank(cls: Optional[str]) -> int:
    return {"A": 0, "B": 1, "C": 2}.get((cls or "").upper(), 3)


def _is_risk_status(value: Any) -> bool:
    s = _clean(value).lower().replace("-", " ").replace("_", " ")
    return any(k in s for k in ("stockout", "stock out", "re order", "reorder", "critical"))


def _status_rank(value: Any) -> int:
    """Lower is more urgent. Class priority is applied before this status rank."""
    s = _clean(value).lower().replace("-", " ").replace("_", " ")
    if "stockout" in s or "stock out" in s:
        return 0
    if "re order" in s or "reorder" in s:
        return 1
    if "critical" in s or "low" in s:
        return 2
    if "over" in s or "ok" in s:
        return 4
    return 3


def _priority_label(cls: Any, status: Any) -> str:
    c = _clean(cls).upper() or "—"
    sr = _status_rank(status)
    if c == "A":
        if sr == 0:
            return "P1 • CLASS A • STOCKOUT"
        if sr == 1:
            return "P1 • CLASS A • RE-ORDER"
        return "P1 • CLASS A"
    if c == "B":
        return "P2 • CLASS B" + (" • RISK" if sr <= 2 else "")
    if c == "C":
        return "P3 • CLASS C" + (" • RISK" if sr <= 2 else "")
    return "P4 • OTHER"


class DeliveryStore:
    """Editable delivery-planning state + truck capacity analysis.

    The reference master is shipped as JSON. User edits/imports are persisted to a writable
    state directory so Render can point SCM_DATA_DIR at a persistent disk when desired.
    """

    def __init__(self, default_master_path: str | Path, state_dir: str | Path):
        self._lock = RLock()
        self.default_master_path = Path(default_master_path)
        self.state_dir = Path(state_dir)
        self.state_dir.mkdir(parents=True, exist_ok=True)
        self.master_path = self.state_dir / "delivery_master.json"
        self.allocations_path = self.state_dir / "delivery_allocations.json"
        self.schedule_path = self.state_dir / "delivery_schedule.json"
        if not self.master_path.exists():
            shutil.copyfile(self.default_master_path, self.master_path)
        self.master: Dict[str, Any] = {}
        self.allocations: List[Dict[str, Any]] = []
        self.schedule: List[Dict[str, str]] = []
        self.load()

    def load(self) -> None:
        with self._lock:
            self.master = json.loads(self.master_path.read_text(encoding="utf-8"))
            self.allocations = self._read_json_list(self.allocations_path)
            self.schedule = self._read_json_list(self.schedule_path)
            self._normalize_master()
            self._migrate_legacy_schedule()

    @staticmethod
    def _read_json_list(path: Path) -> List[Dict[str, Any]]:
        if not path.exists():
            return []
        try:
            value = json.loads(path.read_text(encoding="utf-8"))
            return value if isinstance(value, list) else []
        except Exception:
            return []

    def _normalize_master(self) -> None:
        self.master.setdefault("trucks", [])
        self.master.setdefault("models", [])
        self.master.setdefault("branches", [])
        self.master.setdefault("settings", {})
        s = self.master["settings"]
        s.setdefault("underutilized_pct", 75)
        s.setdefault("full_pct", 90)
        s.setdefault("max_branches_per_truck", 2)

    def _migrate_legacy_schedule(self) -> None:
        """Upgrade v2.0 Day+Branch rows to v2.1 Day+Truck+Branch rows.

        Legacy schedules had no fixed truck. To avoid losing a user's existing weekly plan,
        distribute legacy rows deterministically across active trucks, respecting the configured
        per-truck branch limit. The user can then refine the truck assignments in the new UI.
        """
        if not self.schedule or all(_clean(x.get("plate")) for x in self.schedule):
            return
        trucks = [_clean(t.get("plate")) for t in self.master.get("trucks", []) if t.get("active", True) and _clean(t.get("plate"))]
        if not trucks:
            return
        max_branches = int(max(1, min(2, _num(self.master.get("settings", {}).get("max_branches_per_truck"), 2))))
        slot_counts: Dict[Tuple[str, str], int] = defaultdict(int)
        migrated: List[Dict[str, str]] = []
        seen_branches = set()
        for row in sorted(self.schedule, key=lambda x: (DAYS.index(_clean(x.get("day"))) if _clean(x.get("day")) in DAYS else 99, _clean(x.get("branch")))):
            day = _clean(row.get("day"))
            branch = _clean(row.get("branch"))
            if day not in DAYS or not branch or branch.upper() in seen_branches:
                continue
            plate = _clean(row.get("plate"))
            if not plate:
                plate = next((p for p in trucks if slot_counts[(day, p.upper())] < max_branches), trucks[0])
            slot_counts[(day, plate.upper())] += 1
            seen_branches.add(branch.upper())
            migrated.append({"day": day, "plate": plate, "branch": branch})
        self.schedule = migrated
        self._save_schedule()

    def _save_master(self) -> None:
        self.master_path.write_text(json.dumps(self.master, indent=2, ensure_ascii=False), encoding="utf-8")

    def _save_allocations(self) -> None:
        self.allocations_path.write_text(json.dumps(self.allocations, indent=2, ensure_ascii=False), encoding="utf-8")

    def _save_schedule(self) -> None:
        self.schedule_path.write_text(json.dumps(self.schedule, indent=2, ensure_ascii=False), encoding="utf-8")

    def sync_dashboard_branches(self, dashboard_records: Iterable[Dict[str, Any]]) -> None:
        """Add branches missing from the delivery master without overwriting user edits."""
        with self._lock:
            existing = {str(x.get("branch", "")).upper() for x in self.master.get("branches", [])}
            changed = False
            first_by_branch: Dict[str, str] = {}
            for r in dashboard_records:
                b = _clean(r.get("branch"))
                if b and b not in first_by_branch:
                    first_by_branch[b] = _clean(r.get("area"))
            for b, a in sorted(first_by_branch.items()):
                if b.upper() not in existing:
                    self.master["branches"].append({"branch": b, "area": a, "active": True})
                    existing.add(b.upper())
                    changed = True
            if changed:
                self.master["branches"].sort(key=lambda x: (_clean(x.get("area")), _clean(x.get("branch"))))
                self._save_master()

    def bootstrap(self, day: str, dashboard_records: Iterable[Dict[str, Any]]) -> Dict[str, Any]:
        analysis = self.analyze(day, dashboard_records)
        with self._lock:
            master = json.loads(json.dumps(self.master))
            allocations = list(self.allocations)
            schedule = list(self.schedule)
        return {
            "days": DAYS,
            "selected_day": day if day in DAYS else DAYS[0],
            "master": master,
            "allocations": allocations,
            "allocation_mapping": self.allocation_mapping(),
            "schedule": schedule,
            "analysis": analysis,
        }

    def update_master(self, section: str, rows: Any) -> Dict[str, Any]:
        with self._lock:
            if section not in {"trucks", "models", "branches", "settings"}:
                raise ValueError("Unsupported masterlist section.")
            if section == "settings":
                if not isinstance(rows, dict):
                    raise ValueError("Settings must be an object.")
                under = min(100.0, max(0.0, _num(rows.get("underutilized_pct"), 75)))
                full = min(100.0, max(under, _num(rows.get("full_pct"), 90)))
                max_br = int(max(1, min(2, _num(rows.get("max_branches_per_truck"), 2))))
                self.master["settings"] = {"underutilized_pct": under, "full_pct": full, "max_branches_per_truck": max_br}
            elif section == "trucks":
                clean_rows = []
                for r in rows or []:
                    plate = _clean(r.get("plate"))
                    cap = _num(r.get("capacity"))
                    if not plate or cap <= 0:
                        continue
                    clean_rows.append({
                        "plate": plate,
                        "capacity": cap,
                        "description": _clean(r.get("description")) or plate,
                        "active": bool(r.get("active", True)),
                    })
                self.master["trucks"] = clean_rows
            elif section == "models":
                by_model: Dict[str, Dict[str, Any]] = {}
                for r in rows or []:
                    model = _clean(r.get("model"))
                    idx = _num(r.get("index_size"))
                    if not model or idx <= 0:
                        continue
                    by_model[model.upper()] = {"model": model, "index_size": idx}
                self.master["models"] = sorted(by_model.values(), key=lambda x: x["model"])
            elif section == "branches":
                by_branch: Dict[str, Dict[str, Any]] = {}
                for r in rows or []:
                    branch = _clean(r.get("branch"))
                    if not branch:
                        continue
                    by_branch[branch.upper()] = {"branch": branch, "area": _clean(r.get("area")), "active": bool(r.get("active", True))}
                self.master["branches"] = sorted(by_branch.values(), key=lambda x: (x["area"], x["branch"]))
            self._save_master()
            return json.loads(json.dumps(self.master))

    def update_schedule(self, rows: Any) -> List[Dict[str, str]]:
        """Save the fixed weekly trip plan: Day + Truck + Branch.

        A branch has one weekly schedule slot so an imported allocation can map to a
        deterministic day/truck. Each truck/day accepts up to the configured branch limit.
        """
        clean_rows: List[Dict[str, str]] = []
        seen_slots = set()
        seen_branches = set()
        active_branches = {_clean(b.get("branch")).upper() for b in self.master.get("branches", []) if b.get("active", True)}
        active_trucks = {_clean(t.get("plate")).upper() for t in self.master.get("trucks", []) if t.get("active", True)}
        max_branches = int(max(1, min(2, _num(self.master.get("settings", {}).get("max_branches_per_truck"), 2))))
        slot_counts: Dict[Tuple[str, str], int] = defaultdict(int)
        for r in rows or []:
            day = _clean(r.get("day"))
            plate = _clean(r.get("plate"))
            branch = _clean(r.get("branch"))
            if day not in DAYS or not plate or not branch:
                continue
            if plate.upper() not in active_trucks or branch.upper() not in active_branches:
                continue
            # One allocation destination per branch avoids ambiguity after import.
            if branch.upper() in seen_branches:
                continue
            key = (day, plate.upper(), branch.upper())
            slot_key = (day, plate.upper())
            if key in seen_slots or slot_counts[slot_key] >= max_branches:
                continue
            seen_slots.add(key)
            seen_branches.add(branch.upper())
            slot_counts[slot_key] += 1
            clean_rows.append({"day": day, "plate": plate, "branch": branch})
        clean_rows.sort(key=lambda x: (DAYS.index(x["day"]), x["plate"], x["branch"]))
        with self._lock:
            self.schedule = clean_rows
            self._save_schedule()
        return list(clean_rows)

    def allocation_mapping(self) -> List[Dict[str, Any]]:
        """Map every imported allocation to the branch's saved weekly Day + Truck."""
        with self._lock:
            allocations = list(self.allocations)
            schedule = list(self.schedule)
        schedule_by_branch = {_clean(x.get("branch")).upper(): x for x in schedule if _clean(x.get("branch"))}
        out: List[Dict[str, Any]] = []
        for i, a in enumerate(allocations):
            branch = _clean(a.get("branch"))
            slot = schedule_by_branch.get(branch.upper())
            out.append({
                **a,
                "allocation_index": i,
                "scheduled_day": _clean(slot.get("day")) if slot else "",
                "plate": _clean(slot.get("plate")) if slot else "",
                "schedule_status": "SCHEDULED" if slot else "UNSCHEDULED",
            })
        return out

    def replace_allocations(self, rows: Any) -> List[Dict[str, Any]]:
        clean_rows: List[Dict[str, Any]] = []
        for r in rows or []:
            model = _clean(r.get("model"))
            branch = _clean(r.get("branch"))
            qty = _num(r.get("quantity"))
            if not model or not branch or qty <= 0:
                continue
            clean_rows.append({"model": model, "branch": branch, "quantity": qty, "remarks": _clean(r.get("remarks"))})
        with self._lock:
            self.allocations = clean_rows
            self._save_allocations()
        return list(clean_rows)

    def import_allocations(self, path: str | Path) -> Tuple[List[Dict[str, Any]], List[str]]:
        path = Path(path)
        candidate_sheets: List[Tuple[str, List[List[Any]]]] = []
        if path.suffix.lower() == ".csv":
            with path.open("r", encoding="utf-8-sig", newline="") as f:
                candidate_sheets = [(path.name, list(csv.reader(f)))]
        else:
            reader = XlsxReader(path)
            if not reader.sheet_names:
                raise ValueError("Workbook has no worksheets.")
            candidate_sheets = [(name, reader.read_sheet(name).rows) for name in reader.sheet_names]

        header_idx = None
        mapping: Dict[str, int] = {}
        rows: List[List[Any]] = []
        source_name = ""
        aliases = {
            "model": {"model", "unit", "item", "item description", "unit / item description", "standard description"},
            "branch": {"branch", "retail branch", "branch name"},
            "quantity": {"quantity", "qty", "allocation", "allocated qty", "allocated quantity"},
            "remarks": {"remarks", "remark", "notes", "note"},
        }
        for sheet_name, sheet_rows in candidate_sheets:
            for i, row in enumerate(sheet_rows[:20]):
                found: Dict[str, int] = {}
                for c, val in enumerate(row):
                    h = _norm_header(val)
                    for key, names in aliases.items():
                        if h in names and key not in found:
                            found[key] = c
                if all(k in found for k in ("model", "branch", "quantity")):
                    header_idx = i
                    mapping = found
                    rows = sheet_rows
                    source_name = sheet_name
                    break
            if header_idx is not None:
                break
        if header_idx is None:
            raise ValueError("Import requires columns: Model, Branch, Quantity, Remarks on any worksheet.")

        parsed = []
        warnings: List[str] = []
        for rno, row in enumerate(rows[header_idx + 1:], header_idx + 2):
            def val(key: str) -> Any:
                c = mapping.get(key)
                return row[c] if c is not None and c < len(row) else None
            model = _clean(val("model"))
            branch = _clean(val("branch"))
            qty = _num(val("quantity"))
            remarks = _clean(val("remarks"))
            if not model and not branch and qty == 0:
                continue
            if not model or not branch or qty <= 0:
                warnings.append(f"Row {rno} skipped: Model, Branch and positive Quantity are required.")
                continue
            parsed.append({"model": model, "branch": branch, "quantity": qty, "remarks": remarks})
        if not parsed:
            raise ValueError("No valid allocation rows were found.")
        self.replace_allocations(parsed)
        return parsed, warnings

    def _master_maps(self) -> Tuple[Dict[str, float], Dict[str, str]]:
        model_index = {_clean(x.get("model")).upper(): _num(x.get("index_size"), 1.0) for x in self.master.get("models", []) if _clean(x.get("model"))}
        branch_area = {_clean(x.get("branch")).upper(): _clean(x.get("area")) for x in self.master.get("branches", []) if _clean(x.get("branch"))}
        return model_index, branch_area

    @staticmethod
    def _dashboard_lookup(dashboard_records: Iterable[Dict[str, Any]]) -> Dict[Tuple[str, str], Dict[str, Any]]:
        out: Dict[Tuple[str, str], Dict[str, Any]] = {}
        for r in dashboard_records:
            key = (_clean(r.get("branch")).upper(), _clean(r.get("model")).upper())
            if not all(key):
                continue
            old = out.get(key)
            if old is None or _class_rank(r.get("class")) < _class_rank(old.get("class")):
                out[key] = r
        return out

    def analyze(self, day: str, dashboard_records: Iterable[Dict[str, Any]]) -> Dict[str, Any]:
        """Analyze the saved truck schedule for one day against the imported allocation.

        The schedule is authoritative. Imported rows do not cause branches to be reassigned to
        another truck. If a scheduled branch has no allocation, the trip remains valid and is
        shown as "Scheduled / No Allocation". If a scheduled truck is overloaded, the loading
        recommendation prioritizes Class A first, then stock-risk status within each class.
        """
        if day not in DAYS:
            day = DAYS[0]
        with self._lock:
            allocations = list(self.allocations)
            schedule = list(self.schedule)
            master = json.loads(json.dumps(self.master))

        model_index, branch_area = self._master_maps()
        dashboard_lookup = self._dashboard_lookup(dashboard_records)
        active_trucks = [t for t in master.get("trucks", []) if t.get("active", True) and _num(t.get("capacity")) > 0]
        truck_by_plate = {_clean(t.get("plate")).upper(): t for t in active_trucks}
        day_schedule = [x for x in schedule if x.get("day") == day and _clean(x.get("plate"))]
        schedule_by_branch = {_clean(x.get("branch")).upper(): x for x in schedule if _clean(x.get("branch"))}

        allocations_by_branch: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
        unscheduled_rows: List[Dict[str, Any]] = []
        for i, a in enumerate(allocations):
            b = _clean(a.get("branch"))
            slot = schedule_by_branch.get(b.upper())
            mapped = {**a, "allocation_index": i, "scheduled_day": _clean(slot.get("day")) if slot else "", "plate": _clean(slot.get("plate")) if slot else ""}
            if not slot:
                unscheduled_rows.append(mapped)
                continue
            allocations_by_branch[b.upper()].append(mapped)

        # Build the selected day's fixed truck slots.
        branches_by_truck: Dict[str, List[str]] = defaultdict(list)
        for slot in day_schedule:
            plate_key = _clean(slot.get("plate")).upper()
            branch = _clean(slot.get("branch"))
            if plate_key in truck_by_plate and branch:
                branches_by_truck[plate_key].append(branch)

        under_pct = _num(master.get("settings", {}).get("underutilized_pct"), 75)
        full_pct = _num(master.get("settings", {}).get("full_pct"), 90)
        max_branches = int(master.get("settings", {}).get("max_branches_per_truck", 2) or 2)
        assignments: List[Dict[str, Any]] = []
        branch_priorities: List[Dict[str, Any]] = []

        for truck in active_trucks:
            plate = _clean(truck.get("plate"))
            plate_key = plate.upper()
            cap = _num(truck.get("capacity"))
            scheduled_branches = branches_by_truck.get(plate_key, [])[:max_branches]
            branch_details: List[Dict[str, Any]] = []
            all_items: List[Dict[str, Any]] = []
            unknown_models = set()

            for branch in scheduled_branches:
                bkey = branch.upper()
                raw_rows = allocations_by_branch.get(bkey, [])
                b_qty = 0.0
                b_load = 0.0
                b_class_a = 0.0
                b_risk = 0.0
                b_items: List[Dict[str, Any]] = []
                for a in raw_rows:
                    model = _clean(a.get("model"))
                    qty = _num(a.get("quantity"))
                    dash = dashboard_lookup.get((bkey, model.upper()), {})
                    idx_known = model.upper() in model_index
                    idx = model_index.get(model.upper(), 1.0)
                    cls = _clean(dash.get("class")) or "—"
                    stock_status = _clean(dash.get("stock_status")) or "—"
                    load_index = qty * idx
                    if not idx_known:
                        unknown_models.add(model)
                    item = {
                        "branch": branch,
                        "area": branch_area.get(bkey) or _clean(dash.get("area")),
                        "model": model,
                        "quantity": qty,
                        "remarks": _clean(a.get("remarks")),
                        "class": cls,
                        "stock_status": stock_status,
                        "index_size": idx,
                        "load_index": load_index,
                        "index_known": idx_known,
                        "priority": _priority_label(cls, stock_status),
                    }
                    b_items.append(item)
                    all_items.append(item)
                    b_qty += qty
                    b_load += load_index
                    if cls.upper() == "A":
                        b_class_a += qty
                    if _is_risk_status(stock_status):
                        b_risk += qty
                b_items.sort(key=lambda x: (_class_rank(x.get("class")), _status_rank(x.get("stock_status")), -x["load_index"], x["model"]))
                branch_details.append({
                    "branch": branch,
                    "area": branch_area.get(bkey) or (b_items[0].get("area") if b_items else ""),
                    "has_allocation": bool(raw_rows),
                    "quantity": b_qty,
                    "load_index": b_load,
                    "class_a_qty": b_class_a,
                    "risk_qty": b_risk,
                    "items": b_items,
                })
                branch_priorities.append({
                    "branch": branch,
                    "plate": plate,
                    "area": branch_details[-1]["area"],
                    "has_allocation": bool(raw_rows),
                    "quantity": b_qty,
                    "load_index": b_load,
                    "class_a_qty": b_class_a,
                    "risk_qty": b_risk,
                    "items": b_items,
                })

            all_items.sort(key=lambda x: (_class_rank(x.get("class")), _status_rank(x.get("stock_status")), -x["load_index"], x["branch"], x["model"]))
            total_load = sum(x["load_index"] for x in all_items)
            total_qty = sum(x["quantity"] for x in all_items)
            util = (total_load / cap * 100.0) if cap else 0.0

            if not scheduled_branches:
                status = "IDLE"
            elif total_load <= 0:
                status = "SCHEDULED / NO ALLOCATION"
            elif util > 100.0001:
                status = "OVERLOAD"
            elif util >= full_pct:
                status = "FULL / HIGH UTILIZATION"
            elif util >= under_pct:
                status = "OPTIMIZED"
            else:
                status = "UNDERUTILIZED"

            # Capacity-aware loading recommendation. Class A always wins before B/C.
            remaining_cap = cap
            deferred_units = 0.0
            planned_load = 0.0
            priority_items: List[Dict[str, Any]] = []
            for item in all_items:
                idx = max(0.0001, _num(item.get("index_size"), 1.0))
                requested = _num(item.get("quantity"))
                fit_units = max(0, math.floor((remaining_cap + 1e-9) / idx))
                planned = min(requested, fit_units)
                deferred = max(0.0, requested - planned)
                if planned > 0:
                    used = planned * idx
                    remaining_cap -= used
                    planned_load += used
                deferred_units += deferred
                priority_items.append({**item, "planned_units": planned, "deferred_units": deferred})

            assignments.append({
                "plate": plate,
                "description": _clean(truck.get("description")),
                "capacity": cap,
                "branches": scheduled_branches,
                "branch_details": branch_details,
                "areas": sorted({x.get("area", "") for x in branch_details if x.get("area")}),
                "load_index": total_load,
                "planned_load_index": planned_load,
                "quantity": total_qty,
                "utilization": util,
                "status": status,
                "class_a_qty": sum(x["quantity"] for x in all_items if _clean(x.get("class")).upper() == "A"),
                "deferred_units": deferred_units,
                "priority_items": priority_items,
                "unknown_models": sorted(m for m in unknown_models if m),
                "empty_scheduled_branches": [x["branch"] for x in branch_details if not x["has_allocation"]],
            })

        status_counts: Dict[str, int] = defaultdict(int)
        for a in assignments:
            status_counts[a["status"]] += 1

        scheduled_assignments = [a for a in assignments if a["branches"]]
        scheduled_capacity = sum(a["capacity"] for a in scheduled_assignments)
        total_load = sum(a["load_index"] for a in scheduled_assignments)
        day_alloc_rows = sum(len(allocations_by_branch.get(b.upper(), [])) for b in [x.get("branch", "") for x in day_schedule])
        branches_with_allocation = sum(1 for p in branch_priorities if p["has_allocation"])
        branches_without_allocation = sum(1 for p in branch_priorities if not p["has_allocation"])
        branch_priorities.sort(key=lambda x: (0 if x["has_allocation"] else 1, -x["class_a_qty"], -x["risk_qty"], -x["load_index"], x["branch"]))

        # Imported branches not found in the schedule are not silently assigned elsewhere.
        unscheduled_by_branch: Dict[str, Dict[str, Any]] = {}
        for row in unscheduled_rows:
            b = _clean(row.get("branch"))
            g = unscheduled_by_branch.setdefault(b.upper(), {"branch": b, "quantity": 0.0, "rows": 0})
            g["quantity"] += _num(row.get("quantity"))
            g["rows"] += 1

        return {
            "day": day,
            "scheduled_branches": [x.get("branch", "") for x in day_schedule],
            "scheduled_count": len(day_schedule),
            "scheduled_trucks": len(scheduled_assignments),
            "allocation_rows": len(allocations),
            "scoped_allocation_rows": day_alloc_rows,
            "unscheduled_allocation_rows": len(unscheduled_rows),
            "summary": {
                "total_load_index": total_load,
                "total_capacity": scheduled_capacity,
                "fleet_utilization": (total_load / scheduled_capacity * 100.0) if scheduled_capacity else 0.0,
                "scheduled_branches": len(day_schedule),
                "branches_with_allocation": branches_with_allocation,
                "branches_without_allocation": branches_without_allocation,
                "class_a_units": sum(a["class_a_qty"] for a in scheduled_assignments),
                "overloaded_trucks": status_counts.get("OVERLOAD", 0),
                "underutilized_trucks": status_counts.get("UNDERUTILIZED", 0),
                "idle_trucks": status_counts.get("IDLE", 0),
            },
            "assignments": sorted(assignments, key=lambda x: (0 if x["branches"] else 1, x["plate"])),
            "branch_priorities": branch_priorities,
            "unassigned": sorted(unscheduled_by_branch.values(), key=lambda x: (-x["quantity"], x["branch"])),
            "thresholds": {"underutilized_pct": under_pct, "full_pct": full_pct, "max_branches_per_truck": max_branches},
        }

