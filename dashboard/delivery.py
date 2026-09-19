from __future__ import annotations

from collections import defaultdict
from copy import deepcopy
from pathlib import Path
from threading import RLock
from typing import Any, Dict, Iterable, List, Optional, Tuple
import csv
import json
import math
import shutil

from .xlsx_reader import XlsxReader

DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]


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


def _allocation_class(v: Any) -> str:
    """Normalize allocation Class values from imports/manual entry to A/B/C."""
    s = _clean(v).upper().replace("_", " ").replace("-", " ")
    s = " ".join(s.split())
    if s in {"A", "CLASS A", "A CLASS"}:
        return "A"
    if s in {"B", "CLASS B", "B CLASS"}:
        return "B"
    if s in {"C", "CLASS C", "C CLASS"}:
        return "C"
    return ""


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
            local_master = json.loads(self.master_path.read_text(encoding="utf-8"))
            local_allocations = self._read_json_list(self.allocations_path)
            local_schedule = self._read_json_list(self.schedule_path)
            self.master = local_master
            self.allocations = local_allocations
            self.schedule = local_schedule
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
        s.setdefault("max_branches_per_truck", 6)
        # v2.26: the uploaded operating schedule contains route days with more than two branches.
        # Upgrade legacy settings so valid multi-stop routes are not rejected.
        changed = False
        if int(max(1, _num(s.get("max_branches_per_truck"), 6))) < 6:
            s["max_branches_per_truck"] = 6
            changed = True

        # v2.30 master-data migration: the current allocation template contains BURGMAN15.
        # Apply once to both new installs and existing persistent Render master files, then
        # allow future user edits without re-overwriting the value on every restart.
        migrations = self.master.setdefault("_system_migrations", [])
        migration_key = "v2.30_burgman15_index_1_5"
        if migration_key not in migrations:
            models = self.master.get("models", [])
            match = next((x for x in models if _clean(x.get("model")).upper() == "BURGMAN15"), None)
            if match is None:
                models.append({"model": "BURGMAN15", "index_size": 1.5})
            else:
                match["model"] = "BURGMAN15"
                match["index_size"] = 1.5
            models.sort(key=lambda x: _clean(x.get("model")).upper())
            migrations.append(migration_key)
            changed = True

        # Delivery-plan audit fix: the approved weekly schedule includes MUTI SURIGAO
        # under AREA VI. Add it only when absent so valid imported trips are not silently
        # rejected, while preserving any later user edits/active status.
        branch_migration = "v2.30_muti_surigao_area_vi"
        if branch_migration not in migrations:
            branches = self.master.get("branches", [])
            existing_branch = next((x for x in branches if _clean(x.get("branch")).upper() == "MUTI SURIGAO"), None)
            if existing_branch is None:
                branches.append({"branch": "MUTI SURIGAO", "area": "AREA VI", "active": True})
                branches.sort(key=lambda x: (_clean(x.get("area")).upper(), _clean(x.get("branch")).upper()))
                changed = True
            migrations.append(branch_migration)
            changed = True
        if changed:
            self._save_master()

    def _migrate_legacy_schedule(self) -> None:
        """Upgrade legacy schedules without removing repeated weekly branch visits.

        v2.26 treats every Day + Truck + Branch row as a real schedule slot. A branch may
        therefore appear more than once during the week (or even on two truck routes on the
        same day). Older rows without a truck are distributed to active trucks while preserving
        source order.
        """
        if not self.schedule:
            return
        trucks = [_clean(t.get("plate")) for t in self.master.get("trucks", []) if t.get("active", True) and _clean(t.get("plate"))]
        if not trucks:
            return
        max_branches = int(max(1, min(12, _num(self.master.get("settings", {}).get("max_branches_per_truck"), 6))))
        slot_counts: Dict[Tuple[str, str], int] = defaultdict(int)
        migrated: List[Dict[str, Any]] = []
        seen_exact = set()
        for order, row in enumerate(self.schedule):
            day = _clean(row.get("day"))
            branch = _clean(row.get("branch"))
            if day not in DAYS or not branch:
                continue
            plate = _clean(row.get("plate"))
            if not plate:
                plate = next((p for p in trucks if slot_counts[(day, p.upper())] < max_branches), trucks[0])
            key = (day, plate.upper(), branch.upper())
            if key in seen_exact:
                continue
            seen_exact.add(key)
            slot_counts[(day, plate.upper())] += 1
            migrated.append({"day": day, "plate": plate, "branch": branch, "sequence": int(_num(row.get("sequence"), order + 1))})
        migrated.sort(key=lambda x: (DAYS.index(x["day"]), int(x.get("sequence", 0)), x["plate"], x["branch"]))
        if migrated != self.schedule:
            self.schedule = migrated
            self._save_schedule()

    @staticmethod
    def _atomic_write_json(path: Path, payload: Any) -> None:
        """Persist JSON with replace semantics so interrupted writes cannot corrupt state."""
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_name(f".{path.name}.tmp")
        tmp.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
        tmp.replace(path)

    def _save_master(self) -> None:
        self._atomic_write_json(self.master_path, self.master)

    def _save_allocations(self) -> None:
        self._atomic_write_json(self.allocations_path, self.allocations)

    def _save_schedule(self) -> None:
        self._atomic_write_json(self.schedule_path, self.schedule)

    def reset_to_defaults(self) -> None:
        """Restore shipped Delivery master data and clear all saved weekly planning rows."""
        with self._lock:
            default_master = json.loads(self.default_master_path.read_text(encoding="utf-8"))
            self.master = default_master
            self.allocations = []
            self.schedule = []
            self._normalize_master()
            self._save_master()
            self._save_allocations()
            self._save_schedule()

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
        selected_day = day if day in DAYS or day == "Whole Week" else "Whole Week"
        plan = self._build_week_plan(dashboard_records)
        if selected_day == "Whole Week":
            analysis = dict(plan["weekly"])
        else:
            analysis = dict(plan["daily"].get(selected_day) or {})
            # Keep the complete Monday-Saturday ribbon visible even while drilling into one day.
            analysis["daily_summaries"] = list(plan["weekly"].get("daily_summaries") or [])
        with self._lock:
            master = deepcopy(self.master)
            allocations = list(self.allocations)
            schedule = list(self.schedule)
        return {
            "days": DAYS,
            "analysis_days": ["Whole Week", *DAYS],
            "selected_day": selected_day,
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
                max_br = int(max(1, min(12, _num(rows.get("max_branches_per_truck"), 6))))
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
            return deepcopy(self.master)

    def import_schedule(self, path: str | Path) -> Tuple[List[Dict[str, str]], List[str]]:
        """Import and replace the weekly truck schedule from XLSX/XLSM/CSV.

        Repeated branches are intentionally allowed. Each Day + Truck + Branch row is one
        schedule slot, enabling branches to receive one, two or more deliveries per week.
        """
        path = Path(path)
        warnings: List[str] = []
        source_rows: List[Dict[str, Any]] = []
        aliases = {
            "day": {"day", "delivery day", "schedule day"},
            "plate": {"truck", "plate", "plate no", "plate number", "truck plate"},
            "area": {"area", "branch area"},
            "branch": {"branch", "branch name", "destination"},
        }

        def map_headers(headers: List[Any]) -> Dict[str, int]:
            out: Dict[str, int] = {}
            for i, h in enumerate(headers):
                n = _norm_header(h)
                for key, opts in aliases.items():
                    if n in opts and key not in out:
                        out[key] = i
            return out

        if path.suffix.lower() == ".csv":
            with path.open("r", encoding="utf-8-sig", newline="") as f:
                all_rows = list(csv.reader(f))
            header_i = None
            col_map: Dict[str, int] = {}
            for i, row in enumerate(all_rows[:20]):
                m = map_headers(row)
                if {"day", "plate", "branch"}.issubset(m):
                    header_i, col_map = i, m
                    break
            if header_i is None:
                raise ValueError("Weekly Schedule import requires Day, Truck and Branch columns.")
            data_rows = all_rows[header_i + 1:]
        else:
            reader = XlsxReader(path)
            sheet_name = "Weekly Schedule" if "Weekly Schedule" in reader.sheet_names else (reader.sheet_names[0] if reader.sheet_names else "")
            if not sheet_name:
                raise ValueError("The schedule workbook contains no worksheets.")
            all_rows = reader.read_sheet(sheet_name).rows
            header_i = None
            col_map = {}
            for i, row in enumerate(all_rows[:25]):
                m = map_headers(row)
                if {"day", "plate", "branch"}.issubset(m):
                    header_i, col_map = i, m
                    break
            if header_i is None:
                raise ValueError("Weekly Schedule import requires Day, Truck and Branch columns.")
            data_rows = all_rows[header_i + 1:]

        def val(row: List[Any], key: str) -> str:
            idx = col_map.get(key)
            return _clean(row[idx]) if idx is not None and idx < len(row) else ""

        trucks = {_clean(t.get("plate")).upper(): t for t in self.master.get("trucks", []) if t.get("active", True) and _clean(t.get("plate"))}
        branches = {_clean(b.get("branch")).upper(): b for b in self.master.get("branches", []) if b.get("active", True) and _clean(b.get("branch"))}
        seen_exact = set()
        slot_counts: Dict[Tuple[str, str], int] = defaultdict(int)
        max_branches = int(max(1, min(12, _num(self.master.get("settings", {}).get("max_branches_per_truck"), 6))))

        for n, row in enumerate(data_rows, start=(header_i or 0) + 2):
            day_raw = val(row, "day")
            plate_raw = val(row, "plate")
            branch_raw = val(row, "branch")
            area_raw = val(row, "area")
            if not any((day_raw, plate_raw, branch_raw, area_raw)):
                continue
            day = next((d for d in DAYS if d.lower() == day_raw.lower()), "")
            truck = trucks.get(plate_raw.upper())
            branch = branches.get(branch_raw.upper())
            if not day:
                warnings.append(f"Row {n}: invalid Day '{day_raw}' skipped. Delivery week is Monday to Saturday.")
                continue
            if not truck:
                warnings.append(f"Row {n}: Truck '{plate_raw}' is not active/in Truck Master; skipped.")
                continue
            if not branch:
                warnings.append(f"Row {n}: Branch '{branch_raw}' is not active/in Branch Master; skipped.")
                continue
            canonical_branch = _clean(branch.get("branch"))
            canonical_plate = _clean(truck.get("plate"))
            canonical_area = _clean(branch.get("area"))
            if area_raw and canonical_area and area_raw.strip().upper() != canonical_area.upper():
                warnings.append(f"Row {n}: Area '{area_raw}' corrected to '{canonical_area}' for {canonical_branch}.")
            key = (day, canonical_plate.upper(), canonical_branch.upper())
            if key in seen_exact:
                warnings.append(f"Row {n}: exact duplicate {day} / {canonical_plate} / {canonical_branch} skipped.")
                continue
            slot_key = (day, canonical_plate.upper())
            if slot_counts[slot_key] >= max_branches:
                warnings.append(f"Row {n}: {canonical_plate} on {day} already has the configured maximum {max_branches} route stop(s); skipped.")
                continue
            seen_exact.add(key)
            slot_counts[slot_key] += 1
            source_rows.append({
                "day": day,
                "plate": canonical_plate,
                "branch": canonical_branch,
                "area": canonical_area,
                "sequence": len(source_rows) + 1,
            })

        if not source_rows:
            raise ValueError("No valid Weekly Truck Schedule rows were found. The existing schedule was not changed.")
        saved = self.update_schedule(source_rows)
        if len(saved) != len(source_rows):
            warnings.append("Some schedule rows were removed by final schedule validation.")
        return saved, warnings

    def update_schedule(self, rows: Any) -> List[Dict[str, str]]:
        """Save the fixed weekly trip plan while allowing repeated branch delivery slots."""
        clean_rows: List[Dict[str, Any]] = []
        seen_exact = set()
        active_branches = {_clean(b.get("branch")).upper() for b in self.master.get("branches", []) if b.get("active", True)}
        active_trucks = {_clean(t.get("plate")).upper() for t in self.master.get("trucks", []) if t.get("active", True)}
        max_branches = int(max(1, min(12, _num(self.master.get("settings", {}).get("max_branches_per_truck"), 6))))
        slot_counts: Dict[Tuple[str, str], int] = defaultdict(int)
        for order, r in enumerate(rows or []):
            day = _clean(r.get("day"))
            plate = _clean(r.get("plate"))
            branch = _clean(r.get("branch"))
            if day not in DAYS or not plate or not branch:
                continue
            if plate.upper() not in active_trucks or branch.upper() not in active_branches:
                continue
            key = (day, plate.upper(), branch.upper())
            slot_key = (day, plate.upper())
            if key in seen_exact or slot_counts[slot_key] >= max_branches:
                continue
            seen_exact.add(key)
            slot_counts[slot_key] += 1
            clean_rows.append({
                "day": day,
                "plate": plate,
                "branch": branch,
                "sequence": int(_num(r.get("sequence"), order + 1)),
            })
        clean_rows.sort(key=lambda x: (DAYS.index(x["day"]), int(x.get("sequence", 0)), x["plate"], x["branch"]))
        # Re-number after sorting so the order stays deterministic after later edits/imports.
        for i, row in enumerate(clean_rows, 1):
            row["sequence"] = i
        with self._lock:
            self.schedule = clean_rows
            self._save_schedule()
        return [dict(x) for x in clean_rows]

    def allocation_mapping(self) -> List[Dict[str, Any]]:
        """Map each allocation to every saved delivery slot for that branch."""
        with self._lock:
            allocations = list(self.allocations)
            schedule = list(self.schedule)
        ordered = sorted(
            [x for x in schedule if _clean(x.get("branch")) and _clean(x.get("plate")) and _clean(x.get("day")) in DAYS],
            key=lambda x: (DAYS.index(_clean(x.get("day"))), int(_num(x.get("sequence"), 0)), _clean(x.get("plate"))),
        )
        slots_by_branch: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
        for slot in ordered:
            slots_by_branch[_clean(slot.get("branch")).upper()].append(slot)
        out: List[Dict[str, Any]] = []
        for i, a in enumerate(allocations):
            branch = _clean(a.get("branch"))
            slots = slots_by_branch.get(branch.upper(), [])
            preferred_day = _clean(a.get("preferred_day"))
            preferred_plate = _clean(a.get("preferred_plate"))
            first = slots[0] if slots else None
            trip_labels = [f"{_clean(x.get('day'))} • {_clean(x.get('plate'))}" for x in slots]
            if preferred_day and preferred_plate:
                trip_labels.insert(0, f"MANUAL • {preferred_day} • {preferred_plate}")
            out.append({
                **a,
                "allocation_index": i,
                "scheduled_day": preferred_day or (_clean(first.get("day")) if first else ""),
                "plate": preferred_plate or (_clean(first.get("plate")) if first else ""),
                "scheduled_trips": trip_labels,
                "delivery_frequency": len(slots),
                "schedule_status": (f"MANUAL REBALANCE • {preferred_day} • {preferred_plate}" if preferred_day and preferred_plate else (f"SCHEDULED • {len(slots)} TRIP{'S' if len(slots) != 1 else ''}/WEEK" if slots else "UNSCHEDULED")),
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
            preferred_day = _clean(r.get("preferred_day"))
            preferred_plate = _clean(r.get("preferred_plate"))
            if not preferred_day or not preferred_plate:
                preferred_day = ""
                preferred_plate = ""
            clean_rows.append({
                "model": model,
                "branch": branch,
                "quantity": qty,
                "class": _allocation_class(r.get("class")),
                "remarks": _clean(r.get("remarks")),
                "preferred_day": preferred_day,
                "preferred_plate": preferred_plate,
            })
        with self._lock:
            self.allocations = clean_rows
            self._save_allocations()
        return list(clean_rows)

    def update_allocation(self, index: int, changes: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Update one persisted allocation line and return the full saved allocation list."""
        with self._lock:
            if index < 0 or index >= len(self.allocations):
                raise IndexError("Allocation line was not found. Refresh the Delivery Plan and try again.")
            current = dict(self.allocations[index])
            for key in ("model", "branch", "quantity", "class", "remarks", "preferred_day", "preferred_plate"):
                if key in (changes or {}):
                    current[key] = changes.get(key)
            model = _clean(current.get("model"))
            branch = _clean(current.get("branch"))
            qty = _num(current.get("quantity"))
            if not model or not branch:
                raise ValueError("Model and Branch are required for an allocation line.")
            if qty <= 0:
                raise ValueError("Quantity must be greater than zero. Use Delete to remove the allocation line.")
            preferred_day = _clean(current.get("preferred_day"))
            preferred_plate = _clean(current.get("preferred_plate"))
            if not preferred_day or not preferred_plate:
                preferred_day = ""
                preferred_plate = ""
            if preferred_day and preferred_plate:
                trip_rows = [x for x in self.schedule if _clean(x.get("day")) == preferred_day and _clean(x.get("plate")).upper() == preferred_plate.upper()]
                if not trip_rows:
                    raise ValueError("The selected target truck trip is not in the saved Weekly Truck Schedule.")
                target_branches = {_clean(x.get("branch")).upper() for x in trip_rows if _clean(x.get("branch"))}
                max_stops = max(1, int(_num((self.master.get("settings") or {}).get("max_branches_per_truck"), 6)))
                if branch.upper() not in target_branches and len(target_branches) >= max_stops:
                    raise ValueError(f"Target truck already has the maximum {max_stops} route stops. Choose another truck or adjust the Weekly Truck Schedule.")
            self.allocations[index] = {
                "model": model,
                "branch": branch,
                "quantity": qty,
                "class": _allocation_class(current.get("class")),
                "remarks": _clean(current.get("remarks")),
                "preferred_day": preferred_day,
                "preferred_plate": preferred_plate,
            }
            self._save_allocations()
            return [dict(x) for x in self.allocations]

    def transfer_allocation(self, index: int, target_day: str, target_plate: str, quantity: Any = None) -> List[Dict[str, Any]]:
        """Move all or part of one allocation to a selected saved truck trip for load rebalancing."""
        day = _clean(target_day)
        plate = _clean(target_plate)
        if day not in DAYS or not plate:
            raise ValueError("Select a valid target Day and Truck.")
        with self._lock:
            if index < 0 or index >= len(self.allocations):
                raise IndexError("Allocation line was not found. Refresh the Delivery Plan and try again.")
            trip_rows = [x for x in self.schedule if _clean(x.get("day")) == day and _clean(x.get("plate")).upper() == plate.upper()]
            if not trip_rows:
                raise ValueError("The selected target truck trip is not in the saved Weekly Truck Schedule.")
            current = dict(self.allocations[index])
            branch = _clean(current.get("branch"))
            target_branches = {_clean(x.get("branch")).upper() for x in trip_rows if _clean(x.get("branch"))}
            max_stops = max(1, int(_num((self.master.get("settings") or {}).get("max_branches_per_truck"), 6)))
            if branch and branch.upper() not in target_branches and len(target_branches) >= max_stops:
                raise ValueError(f"Target truck already has the maximum {max_stops} route stops. Choose another truck or adjust the Weekly Truck Schedule.")
            current_qty = _num(current.get("quantity"))
            move_qty = current_qty if quantity in (None, "") else _num(quantity)
            if move_qty <= 0:
                raise ValueError("Transfer quantity must be greater than zero.")
            if move_qty > current_qty + 1e-9:
                raise ValueError("Transfer quantity cannot exceed the allocation quantity.")
            if move_qty >= current_qty - 1e-9:
                current["preferred_day"] = day
                current["preferred_plate"] = plate
                self.allocations[index] = current
            else:
                self.allocations[index]["quantity"] = current_qty - move_qty
                moved = dict(current)
                moved["quantity"] = move_qty
                moved["preferred_day"] = day
                moved["preferred_plate"] = plate
                self.allocations.insert(index + 1, moved)
            self._save_allocations()
            return [dict(x) for x in self.allocations]

    def delete_allocation(self, index: int) -> List[Dict[str, Any]]:
        """Delete one persisted allocation line and return the remaining allocations."""
        with self._lock:
            if index < 0 or index >= len(self.allocations):
                raise IndexError("Allocation line was not found. Refresh the Delivery Plan and try again.")
            self.allocations.pop(index)
            self._save_allocations()
            return [dict(x) for x in self.allocations]

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
            "class": {"class", "classification", "abc class", "class abc", "item class", "model class"},
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
            raise ValueError("Import requires columns: Model, Branch and Quantity. The approved template also includes CLASS and Remarks.")

        parsed = []
        warnings: List[str] = []
        if "class" not in mapping:
            warnings.append("CLASS column was not found. Class will fall back to Dashboard data where available; use the new Unit Allocation Template for explicit A/B/C allocation priority.")
        for rno, row in enumerate(rows[header_idx + 1:], header_idx + 2):
            def val(key: str) -> Any:
                c = mapping.get(key)
                return row[c] if c is not None and c < len(row) else None
            model = _clean(val("model"))
            branch = _clean(val("branch"))
            qty = _num(val("quantity"))
            raw_class = _clean(val("class"))
            alloc_class = _allocation_class(raw_class)
            remarks = _clean(val("remarks"))
            if not model and not branch and qty == 0 and not raw_class:
                continue
            if not model or not branch or qty <= 0:
                warnings.append(f"Row {rno} skipped: Model, Branch and positive Quantity are required.")
                continue
            if raw_class and not alloc_class:
                warnings.append(f"Row {rno}: CLASS '{raw_class}' is not A, B or C; dashboard class will be used when available.")
            parsed.append({"model": model, "branch": branch, "quantity": qty, "class": alloc_class, "remarks": remarks, "preferred_day": "", "preferred_plate": ""})
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

    def _build_week_plan(self, dashboard_records: Iterable[Dict[str, Any]]) -> Dict[str, Any]:
        """Build one capacity-aware Monday-Saturday dispatch plan.

        Allocation is consumed only once across the week. If a branch has another saved
        schedule slot later in the week, capacity overflow automatically rolls forward to
        that next slot. Residual quantity after the branch's last slot is marked for route
        adjustment / following-week backorder. Unscheduled branches remain unassigned.
        """
        with self._lock:
            allocations = [dict(x) for x in self.allocations]
            schedule = [dict(x) for x in self.schedule]
            master = deepcopy(self.master)

        model_index, branch_area = self._master_maps()
        dashboard_lookup = self._dashboard_lookup(dashboard_records)
        truck_by_plate = {
            _clean(t.get("plate")).upper(): t
            for t in master.get("trucks", [])
            if _clean(t.get("plate")) and _num(t.get("capacity")) > 0
        }
        active_trucks = {
            k: v for k, v in truck_by_plate.items() if v.get("active", True)
        }
        under_pct = _num(master.get("settings", {}).get("underutilized_pct"), 75)
        full_pct = _num(master.get("settings", {}).get("full_pct"), 90)
        max_branches = int(max(1, min(12, _num(master.get("settings", {}).get("max_branches_per_truck"), 6))))

        # Normalize slots and preserve imported/manual route order. Repeated branches are valid.
        slots: List[Dict[str, Any]] = []
        seen_exact = set()
        for order, row in enumerate(schedule):
            day = _clean(row.get("day"))
            plate = _clean(row.get("plate"))
            branch = _clean(row.get("branch"))
            if day not in DAYS or not plate or not branch or plate.upper() not in active_trucks:
                continue
            key = (day, plate.upper(), branch.upper())
            if key in seen_exact:
                continue
            seen_exact.add(key)
            slots.append({
                "day": day,
                "plate": plate,
                "branch": branch,
                "area": branch_area.get(branch.upper(), ""),
                "sequence": int(_num(row.get("sequence"), order + 1)),
            })
        slots.sort(key=lambda x: (DAYS.index(x["day"]), x["sequence"], x["plate"], x["branch"]))

        # A trip is one Day + Truck route. Multiple branch rows under the same route share capacity.
        trips: List[Dict[str, Any]] = []
        trip_by_key: Dict[Tuple[str, str], Dict[str, Any]] = {}
        for slot in slots:
            key = (slot["day"], slot["plate"].upper())
            trip = trip_by_key.get(key)
            if trip is None:
                truck = active_trucks.get(slot["plate"].upper(), {})
                trip = {
                    "day": slot["day"],
                    "plate": slot["plate"],
                    "description": _clean(truck.get("description")) or slot["plate"],
                    "capacity": _num(truck.get("capacity")),
                    "sequence": slot["sequence"],
                    "branches": [],
                    "slots": [],
                }
                trip_by_key[key] = trip
                trips.append(trip)
            if slot["branch"].upper() not in {b.upper() for b in trip["branches"]}:
                trip["branches"].append(slot["branch"])
                trip["slots"].append(slot)
        trips.sort(key=lambda x: (DAYS.index(x["day"]), x["sequence"], x["plate"]))
        for i, trip in enumerate(trips):
            trip["trip_index"] = i
            trip["day_trip_no"] = 1 + sum(1 for x in trips[:i] if x["day"] == trip["day"])

        branch_trip_indices: Dict[str, List[int]] = defaultdict(list)
        trip_index_by_key: Dict[Tuple[str, str], int] = {}
        for i, trip in enumerate(trips):
            trip_index_by_key[(trip["day"], trip["plate"].upper())] = i
            for branch in trip["branches"]:
                branch_trip_indices[branch.upper()].append(i)

        # Enrich allocation lines once and maintain remaining quantity through the week.
        states: List[Dict[str, Any]] = []
        allocations_by_branch: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
        for i, a in enumerate(allocations):
            branch = _clean(a.get("branch"))
            model = _clean(a.get("model"))
            qty = _num(a.get("quantity"))
            if not branch or not model or qty <= 0:
                continue
            dash = dashboard_lookup.get((branch.upper(), model.upper()), {})
            idx_known = model.upper() in model_index
            idx = model_index.get(model.upper(), 1.0)
            cls = _allocation_class(a.get("class")) or _clean(dash.get("class")) or "—"
            stock_status = _clean(dash.get("stock_status")) or "—"
            preferred_day = _clean(a.get("preferred_day"))
            preferred_plate = _clean(a.get("preferred_plate"))
            preferred_trip_index = trip_index_by_key.get((preferred_day, preferred_plate.upper())) if preferred_day and preferred_plate else None
            normal_trip_indices = list(branch_trip_indices.get(branch.upper(), []))
            if preferred_trip_index is not None:
                # Manual rebalance: force this line to the selected truck first, then retain later
                # branch trips as automatic rollover options if the target truck still cannot fit it.
                scheduled_trip_indices = [preferred_trip_index] + [x for x in normal_trip_indices if x > preferred_trip_index]
                scheduled_trip_indices = list(dict.fromkeys(scheduled_trip_indices))
            else:
                scheduled_trip_indices = normal_trip_indices
            state = {
                "allocation_index": i,
                "branch": branch,
                "area": branch_area.get(branch.upper()) or _clean(dash.get("area")),
                "model": model,
                "class": cls,
                "stock_status": stock_status,
                "index_size": idx,
                "index_known": idx_known,
                "original_quantity": qty,
                "remaining_quantity": qty,
                "user_remarks": _clean(a.get("remarks")),
                "priority": _priority_label(cls, stock_status),
                "preferred_day": preferred_day,
                "preferred_plate": preferred_plate,
                "preferred_trip_index": preferred_trip_index,
                "manual_transfer": preferred_trip_index is not None,
                "scheduled_trip_indices": scheduled_trip_indices,
                "planned_segments": [],
            }
            states.append(state)
            allocations_by_branch[branch.upper()].append(state)

        daily_assignments: Dict[str, List[Dict[str, Any]]] = {d: [] for d in DAYS}
        daily_priorities: Dict[str, List[Dict[str, Any]]] = {d: [] for d in DAYS}

        for ti, trip in enumerate(trips):
            # An allocation is eligible only on one of its resolved trip indices. This supports
            # a direct manual truck transfer without double-loading the same allocation on its
            # original route before the selected rebalance trip.
            eligible = [s for s in states if ti in s.get("scheduled_trip_indices", []) and s["remaining_quantity"] > 1e-9]
            eligible.sort(key=lambda x: (
                _class_rank(x.get("class")),
                _status_rank(x.get("stock_status")),
                -(_num(x.get("remaining_quantity")) * _num(x.get("index_size"), 1)),
                x.get("allocation_index", 0),
            ))

            requested_load = sum(_num(x.get("remaining_quantity")) * _num(x.get("index_size"), 1) for x in eligible)
            requested_units = sum(_num(x.get("remaining_quantity")) for x in eligible)
            remaining_cap = _num(trip.get("capacity"))
            planned_load = 0.0
            planned_units_total = 0.0
            priority_items: List[Dict[str, Any]] = []
            unknown_models = set()

            for state in eligible:
                before = _num(state.get("remaining_quantity"))
                idx = max(0.0001, _num(state.get("index_size"), 1.0))
                fit_units = max(0, math.floor((remaining_cap + 1e-9) / idx))
                planned = min(before, fit_units)
                if planned > 0:
                    used = planned * idx
                    remaining_cap = max(0.0, remaining_cap - used)
                    planned_load += used
                    planned_units_total += planned
                    state["remaining_quantity"] = max(0.0, before - planned)
                remaining_after = _num(state.get("remaining_quantity"))
                later = [x for x in state.get("scheduled_trip_indices", []) if x > ti]
                next_trip = trips[later[0]] if later else None
                carryover = remaining_after if remaining_after > 1e-9 and next_trip else 0.0
                backorder = remaining_after if remaining_after > 1e-9 and not next_trip else 0.0
                if carryover > 0:
                    system_remarks = f"CARRYOVER TO {next_trip['day'].upper()} • {next_trip['plate']}"
                    dispatch_status = "CARRYOVER"
                elif backorder > 0:
                    system_remarks = "BACKORDER • FOLLOWING WEEK / ROUTE ADJUSTMENT REQUIRED"
                    dispatch_status = "BACKORDER"
                else:
                    system_remarks = "PLANNED THIS TRIP"
                    dispatch_status = "PLANNED"
                if state.get("manual_transfer") and state.get("preferred_trip_index") == ti and state.get("preferred_day") and state.get("preferred_plate"):
                    system_remarks = f"MANUAL REBALANCE TO {state.get('preferred_day').upper()} • {state.get('preferred_plate')} • {system_remarks}"
                if not state.get("index_known"):
                    unknown_models.add(state.get("model"))
                combined_remarks = " • ".join(x for x in [state.get("user_remarks"), system_remarks] if x)
                item = {
                    "allocation_index": state.get("allocation_index"),
                    "branch": state.get("branch"),
                    "area": state.get("area"),
                    "model": state.get("model"),
                    "class": state.get("class"),
                    "stock_status": state.get("stock_status"),
                    "index_size": idx,
                    "index_known": state.get("index_known"),
                    "priority": state.get("priority"),
                    "quantity": before,
                    "original_quantity": state.get("original_quantity"),
                    "planned_units": planned,
                    "deferred_units": max(0.0, before - planned),
                    "carryover_units": carryover,
                    "backorder_units": backorder,
                    "load_index": before * idx,
                    "planned_load_index": planned * idx,
                    "dispatch_status": dispatch_status,
                    "next_trip_day": next_trip.get("day", "") if next_trip else "",
                    "next_trip_plate": next_trip.get("plate", "") if next_trip else "",
                    "system_remarks": system_remarks,
                    "user_remarks": state.get("user_remarks", ""),
                    "remarks": combined_remarks,
                    "preferred_day": state.get("preferred_day", ""),
                    "preferred_plate": state.get("preferred_plate", ""),
                    "manual_transfer": bool(state.get("manual_transfer")),
                }
                priority_items.append(item)
                if planned > 0:
                    state["planned_segments"].append({
                        "day": trip["day"], "plate": trip["plate"], "quantity": planned,
                        "load_index": planned * idx,
                    })

            branch_details: List[Dict[str, Any]] = []
            effective_branches = list(trip["branches"])
            for moved_state in eligible:
                moved_branch = _clean(moved_state.get("branch"))
                if moved_branch and moved_branch.upper() not in {b.upper() for b in effective_branches}:
                    effective_branches.append(moved_branch)
            for branch in effective_branches:
                bkey = branch.upper()
                original_states = allocations_by_branch.get(bkey, [])
                items = [x for x in priority_items if _clean(x.get("branch")).upper() == bkey]
                had_allocation = bool(original_states)
                completed_earlier = bool(had_allocation and not items and all(_num(x.get("remaining_quantity")) <= 1e-9 for x in original_states))
                quantity = sum(_num(x.get("quantity")) for x in items)
                planned_qty = sum(_num(x.get("planned_units")) for x in items)
                carryover_qty = sum(_num(x.get("carryover_units")) for x in items)
                backorder_qty = sum(_num(x.get("backorder_units")) for x in items)
                load_index = sum(_num(x.get("load_index")) for x in items)
                class_a_qty = sum(_num(x.get("quantity")) for x in items if _clean(x.get("class")).upper() == "A")
                planned_class_a_qty = sum(_num(x.get("planned_units")) for x in items if _clean(x.get("class")).upper() == "A")
                risk_qty = sum(_num(x.get("quantity")) for x in items if _is_risk_status(x.get("stock_status")))
                next_slots = sorted({
                    x for state_row in original_states for x in state_row.get("scheduled_trip_indices", []) if x > ti
                })
                next_trip = trips[next_slots[0]] if next_slots else None
                frequency = len(sorted({x for state_row in original_states for x in state_row.get("scheduled_trip_indices", [])})) if original_states else len(branch_trip_indices.get(bkey, []))
                detail = {
                    "day": trip["day"],
                    "trip_no": trip["day_trip_no"],
                    "plate": trip["plate"],
                    "branch": branch,
                    "area": branch_area.get(bkey) or (items[0].get("area") if items else ""),
                    "frequency": frequency,
                    "has_allocation": had_allocation,
                    "has_trip_demand": bool(items),
                    "completed_earlier": completed_earlier,
                    "quantity": quantity,
                    "planned_qty": planned_qty,
                    "load_index": load_index,
                    "class_a_qty": class_a_qty,
                    "planned_class_a_qty": planned_class_a_qty,
                    "risk_qty": risk_qty,
                    "carryover_qty": carryover_qty,
                    "backorder_qty": backorder_qty,
                    "next_trip_day": next_trip.get("day", "") if next_trip else "",
                    "next_trip_plate": next_trip.get("plate", "") if next_trip else "",
                    "items": items,
                }
                branch_details.append(detail)
                daily_priorities[trip["day"]].append(dict(detail))

            planned_util = (planned_load / _num(trip.get("capacity")) * 100.0) if _num(trip.get("capacity")) else 0.0
            requested_util = (requested_load / _num(trip.get("capacity")) * 100.0) if _num(trip.get("capacity")) else 0.0
            carryover_units = sum(_num(x.get("carryover_units")) for x in priority_items)
            backorder_units = sum(_num(x.get("backorder_units")) for x in priority_items)
            has_any_allocation = any(x.get("has_allocation") for x in branch_details)
            has_trip_demand = bool(priority_items)
            if not has_any_allocation:
                status = "SCHEDULED / NO ALLOCATION"
            elif not has_trip_demand:
                status = "SCHEDULED / COMPLETED EARLIER"
            elif requested_load > _num(trip.get("capacity")) + 1e-9:
                status = "OVERLOAD / CARRYOVER" if carryover_units > 0 else "OVERLOAD / BACKORDER"
            elif planned_util >= full_pct:
                status = "FULL / HIGH UTILIZATION"
            elif planned_util >= under_pct:
                status = "OPTIMIZED"
            else:
                status = "UNDERUTILIZED"

            assignment = {
                "day": trip["day"],
                "trip_no": trip["day_trip_no"],
                "plate": trip["plate"],
                "description": trip["description"],
                "capacity": trip["capacity"],
                "branches": list(effective_branches),
                "branch_details": branch_details,
                "areas": sorted({x.get("area", "") for x in branch_details if x.get("area")}),
                "load_index": requested_load,
                "planned_load_index": planned_load,
                "quantity": requested_units,
                "planned_units": planned_units_total,
                "utilization": requested_util,
                "planned_utilization": planned_util,
                "status": status,
                "class_a_qty": sum(_num(x.get("quantity")) for x in priority_items if _clean(x.get("class")).upper() == "A"),
                "planned_class_a_qty": sum(_num(x.get("planned_units")) for x in priority_items if _clean(x.get("class")).upper() == "A"),
                "deferred_units": carryover_units + backorder_units,
                "carryover_units": carryover_units,
                "backorder_units": backorder_units,
                "priority_items": priority_items,
                "unknown_models": sorted(m for m in unknown_models if m),
                "empty_scheduled_branches": [x["branch"] for x in branch_details if not x["has_allocation"] and x["branch"].upper() in {b.upper() for b in trip["branches"]}],
                "completed_earlier_branches": [x["branch"] for x in branch_details if x.get("completed_earlier")],
            }
            daily_assignments[trip["day"]].append(assignment)

        # Any remaining scheduled quantity has exhausted its final weekly slot: backorder.
        backorder_rows: List[Dict[str, Any]] = []
        unscheduled_rows: List[Dict[str, Any]] = []
        for state in states:
            remaining = _num(state.get("remaining_quantity"))
            if not state.get("scheduled_trip_indices"):
                unscheduled_rows.append({
                    "branch": state.get("branch"), "model": state.get("model"), "class": state.get("class"),
                    "quantity": state.get("original_quantity"), "remarks": state.get("user_remarks"),
                    "required_action": "ADD TO WEEKLY SCHEDULE / ROUTE ADJUSTMENT",
                })
            elif remaining > 1e-9:
                backorder_rows.append({
                    "branch": state.get("branch"), "area": state.get("area"), "model": state.get("model"),
                    "class": state.get("class"), "quantity": remaining,
                    "remarks": "BACKORDER • FOLLOWING WEEK / ROUTE ADJUSTMENT REQUIRED",
                })

        unscheduled_by_branch: Dict[str, Dict[str, Any]] = {}
        for row in unscheduled_rows:
            key = _clean(row.get("branch")).upper()
            g = unscheduled_by_branch.setdefault(key, {"branch": row.get("branch"), "quantity": 0.0, "rows": 0})
            g["quantity"] += _num(row.get("quantity"))
            g["rows"] += 1
        backorder_by_branch: Dict[str, Dict[str, Any]] = {}
        for row in backorder_rows:
            key = _clean(row.get("branch")).upper()
            g = backorder_by_branch.setdefault(key, {"branch": row.get("branch"), "area": row.get("area"), "quantity": 0.0, "rows": 0})
            g["quantity"] += _num(row.get("quantity"))
            g["rows"] += 1

        total_allocation_units = sum(_num(x.get("original_quantity")) for x in states)
        planned_week_units = sum(sum(_num(i.get("planned_units")) for i in a.get("priority_items", [])) for day in DAYS for a in daily_assignments[day])
        planned_week_load = sum(sum(_num(i.get("planned_load_index")) for i in a.get("priority_items", [])) for day in DAYS for a in daily_assignments[day])
        total_week_capacity = sum(_num(a.get("capacity")) for day in DAYS for a in daily_assignments[day])
        backorder_units_total = sum(_num(x.get("quantity")) for x in backorder_rows)
        unscheduled_units_total = sum(_num(x.get("quantity")) for x in unscheduled_rows)

        daily_results: Dict[str, Dict[str, Any]] = {}
        daily_summaries: List[Dict[str, Any]] = []
        for day in DAYS:
            assigns = daily_assignments[day]
            priorities = daily_priorities[day]
            status_counts: Dict[str, int] = defaultdict(int)
            for a in assigns:
                status_counts[a.get("status", "")] += 1
            day_slots = [x for x in slots if x["day"] == day]
            day_capacity = sum(_num(a.get("capacity")) for a in assigns)
            day_requested_load = sum(_num(a.get("load_index")) for a in assigns)
            day_planned_load = sum(_num(a.get("planned_load_index")) for a in assigns)
            day_planned_units = sum(_num(a.get("planned_units")) for a in assigns)
            day_carryover = sum(_num(a.get("carryover_units")) for a in assigns)
            day_backorder = sum(_num(a.get("backorder_units")) for a in assigns)
            priorities.sort(key=lambda x: (
                0 if x.get("has_trip_demand") else 1,
                -_num(x.get("class_a_qty")), -_num(x.get("risk_qty")), -_num(x.get("load_index")),
                _clean(x.get("branch")),
            ))
            summary = {
                "total_load_index": day_requested_load,
                "planned_load_index": day_planned_load,
                "total_capacity": day_capacity,
                "fleet_utilization": (day_requested_load / day_capacity * 100.0) if day_capacity else 0.0,
                "planned_utilization": (day_planned_load / day_capacity * 100.0) if day_capacity else 0.0,
                "scheduled_branches": len(day_slots),
                "branches_with_allocation": sum(1 for x in priorities if x.get("has_allocation")),
                "branches_without_allocation": sum(1 for x in priorities if not x.get("has_allocation")),
                "class_a_units": sum(_num(a.get("class_a_qty")) for a in assigns),
                "allocation_units": sum(_num(a.get("quantity")) for a in assigns),
                "requested_units": sum(_num(a.get("quantity")) for a in assigns),
                "planned_units": day_planned_units,
                "carryover_units": day_carryover,
                "backorder_units": day_backorder,
                "overloaded_trucks": sum(1 for a in assigns if "OVERLOAD" in _clean(a.get("status")).upper()),
                "underutilized_trucks": status_counts.get("UNDERUTILIZED", 0),
                "idle_trucks": 0,
            }
            result = {
                "day": day,
                "scheduled_branches": [x.get("branch", "") for x in day_slots],
                "scheduled_count": len(day_slots),
                "scheduled_trucks": len(assigns),
                "allocation_rows": len(allocations),
                "scoped_allocation_rows": sum(len(allocations_by_branch.get(x.get("branch", "").upper(), [])) for x in day_slots),
                "unscheduled_allocation_rows": len(unscheduled_rows),
                "summary": summary,
                "assignments": assigns,
                "branch_priorities": priorities,
                "unassigned": sorted(unscheduled_by_branch.values(), key=lambda x: (-x["quantity"], x["branch"])),
                "backorders": sorted(backorder_by_branch.values(), key=lambda x: (-x["quantity"], x["branch"])),
                "thresholds": {"underutilized_pct": under_pct, "full_pct": full_pct, "max_branches_per_truck": max_branches},
            }
            daily_results[day] = result
            daily_summaries.append({
                "day": day,
                "trips": len(assigns),
                "branch_slots": len(day_slots),
                "planned_units": day_planned_units,
                "planned_load_index": day_planned_load,
                "capacity": day_capacity,
                "utilization": summary["planned_utilization"],
                "carryover_units": day_carryover,
                "backorder_units": day_backorder,
                "overloaded_trucks": summary["overloaded_trucks"],
            })

        weekly_assignments = [dict(a) for d in DAYS for a in daily_assignments[d]]
        weekly_priorities = [dict(p) for d in DAYS for p in daily_priorities[d]]
        weekly_summary = {
            "total_load_index": sum(_num(a.get("load_index")) for a in weekly_assignments),
            "planned_load_index": planned_week_load,
            "total_capacity": total_week_capacity,
            "fleet_utilization": (planned_week_load / total_week_capacity * 100.0) if total_week_capacity else 0.0,
            "planned_utilization": (planned_week_load / total_week_capacity * 100.0) if total_week_capacity else 0.0,
            "scheduled_branches": len(slots),
            "branches_with_allocation": sum(1 for p in weekly_priorities if p.get("has_allocation")),
            "branches_without_allocation": sum(1 for p in weekly_priorities if not p.get("has_allocation")),
            "class_a_units": sum(_num(a.get("planned_class_a_qty")) for a in weekly_assignments),
            "allocation_units": total_allocation_units,
            "planned_units": planned_week_units,
            "carryover_units": sum(_num(a.get("carryover_units")) for a in weekly_assignments),
            "backorder_units": backorder_units_total,
            "unscheduled_units": unscheduled_units_total,
            "overloaded_trucks": sum(1 for a in weekly_assignments if "OVERLOAD" in _clean(a.get("status")).upper()),
            "underutilized_trucks": sum(1 for a in weekly_assignments if a.get("status") == "UNDERUTILIZED"),
            "idle_trucks": 0,
        }
        weekly = {
            "day": "Whole Week",
            "scheduled_branches": [x.get("branch", "") for x in slots],
            "scheduled_count": len(slots),
            "scheduled_trucks": len(weekly_assignments),
            "allocation_rows": len(allocations),
            "scoped_allocation_rows": sum(1 for s in states if s.get("scheduled_trip_indices")),
            "unscheduled_allocation_rows": len(unscheduled_rows),
            "summary": weekly_summary,
            "assignments": weekly_assignments,
            "branch_priorities": weekly_priorities,
            "unassigned": sorted(unscheduled_by_branch.values(), key=lambda x: (-x["quantity"], x["branch"])),
            "backorders": sorted(backorder_by_branch.values(), key=lambda x: (-x["quantity"], x["branch"])),
            "daily_summaries": daily_summaries,
            "thresholds": {"underutilized_pct": under_pct, "full_pct": full_pct, "max_branches_per_truck": max_branches},
        }
        return {"daily": daily_results, "weekly": weekly}

    def analysis_bundle(self, dashboard_records: Iterable[Dict[str, Any]]) -> Dict[str, Any]:
        """Build the Monday-Saturday plan once for exports that need every day."""
        return self._build_week_plan(dashboard_records)


