from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional
import re
import zipfile
import xml.etree.ElementTree as ET

_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def _col_index(cell_ref: str) -> int:
    m = re.match(r"([A-Z]+)", cell_ref or "")
    if not m:
        return 0
    n = 0
    for ch in m.group(1):
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def excel_serial_to_date(value: Any) -> Optional[datetime]:
    """Normalize Excel serials and common date text into a datetime.

    KPI period headers can arrive from Excel either as numeric serials or as
    text/formula results (for example ``09/21/2026``).  Treating only numeric
    serials as dates made a newly-added YTD/Weekly column silently disappear.
    Keep this dependency-free and deliberately conservative: numeric values use
    Excel's 1900 date system, while text is accepted only when it matches a
    known date shape.
    """
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)

    # Native numeric cell / numeric string -> Excel 1900 serial.
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            serial = float(value)
            if serial <= 0:
                return None
            return datetime(1899, 12, 30) + timedelta(days=serial)
        except (TypeError, ValueError, OverflowError):
            return None

    text = str(value).strip()
    if not text:
        return None

    # Numeric strings are also valid cached Excel serials.
    try:
        serial = float(text)
        if serial > 0:
            return datetime(1899, 12, 30) + timedelta(days=serial)
    except (TypeError, ValueError, OverflowError):
        pass

    # Common Excel/display formats used by the control-tower YTD/Weekly sheets.
    formats = (
        "%m/%d/%Y", "%m/%d/%y", "%d/%m/%Y", "%d/%m/%y",
        "%Y-%m-%d", "%Y/%m/%d",
        "%d-%b-%Y", "%d-%b-%y",
        "%b %d, %Y", "%B %d, %Y", "%b %d %Y", "%B %d %Y",
        "%b-%Y", "%b %Y",
    )
    for fmt in formats:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue

    # ISO datetime text may include a time component or UTC offset.
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError:
        return None


@dataclass
class SheetData:
    name: str
    rows: List[List[Any]]


class XlsxReader:
    """Small, dependency-free XLSX reader tailored to dashboard imports.

    It reads cached formula values, shared strings, inline strings, booleans,
    and numeric values. This keeps deployment lightweight and avoids requiring
    Excel on the dashboard server.
    """

    def __init__(self, path: str | Path):
        self.path = Path(path)
        self._shared_strings: List[str] = []
        self._sheet_paths: Dict[str, str] = {}
        self._sheet_cache: Dict[str, SheetData] = {}
        self._load_metadata()

    @property
    def sheet_names(self) -> List[str]:
        return list(self._sheet_paths)

    def _load_metadata(self) -> None:
        if not self.path.exists():
            raise FileNotFoundError(self.path)
        with zipfile.ZipFile(self.path) as z:
            if "xl/sharedStrings.xml" in z.namelist():
                root = ET.fromstring(z.read("xl/sharedStrings.xml"))
                for si in root.findall(f"{{{_NS_MAIN}}}si"):
                    parts = []
                    for t in si.iter(f"{{{_NS_MAIN}}}t"):
                        parts.append(t.text or "")
                    self._shared_strings.append("".join(parts))

            wb = ET.fromstring(z.read("xl/workbook.xml"))
            rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            rel_map = {
                rel.attrib["Id"]: rel.attrib["Target"]
                for rel in rels.findall(f"{{{_NS_PKG_REL}}}Relationship")
            }
            sheets = wb.find(f"{{{_NS_MAIN}}}sheets")
            if sheets is None:
                return
            for sh in sheets.findall(f"{{{_NS_MAIN}}}sheet"):
                name = sh.attrib.get("name", "")
                rid = sh.attrib.get(f"{{{_NS_REL}}}id")
                target = rel_map.get(rid or "", "")
                if target.startswith("/"):
                    target = target.lstrip("/")
                elif not target.startswith("xl/"):
                    target = "xl/" + target
                self._sheet_paths[name] = target

    def read_sheet(self, name: str) -> SheetData:
        if name not in self._sheet_paths:
            raise KeyError(f"Missing worksheet: {name}")
        cached = self._sheet_cache.get(name)
        if cached is not None:
            return cached

        with zipfile.ZipFile(self.path) as z:
            root = ET.fromstring(z.read(self._sheet_paths[name]))

        sheet_data = root.find(f"{{{_NS_MAIN}}}sheetData")
        rows: List[List[Any]] = []
        if sheet_data is not None:
            for row_el in sheet_data.findall(f"{{{_NS_MAIN}}}row"):
                values: List[Any] = []
                for c in row_el.findall(f"{{{_NS_MAIN}}}c"):
                    ref = c.attrib.get("r", "A1")
                    idx = _col_index(ref)
                    while len(values) <= idx:
                        values.append(None)
                    values[idx] = self._cell_value(c)
                rows.append(values)

        result = SheetData(name, rows)
        self._sheet_cache[name] = result
        return result

    def _cell_value(self, cell: ET.Element) -> Any:
        cell_type = cell.attrib.get("t")
        if cell_type == "inlineStr":
            is_el = cell.find(f"{{{_NS_MAIN}}}is")
            if is_el is None:
                return ""
            return "".join((t.text or "") for t in is_el.iter(f"{{{_NS_MAIN}}}t"))

        v = cell.find(f"{{{_NS_MAIN}}}v")
        if v is None or v.text is None:
            # Formula cells may have no cached value.
            return None
        raw = v.text
        if cell_type == "s":
            try:
                return self._shared_strings[int(raw)]
            except (ValueError, IndexError):
                return raw
        if cell_type == "str":
            return raw
        if cell_type == "b":
            return raw == "1"
        try:
            num = float(raw)
            if num.is_integer():
                return int(num)
            return num
        except ValueError:
            return raw
