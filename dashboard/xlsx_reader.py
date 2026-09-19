from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta
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
    try:
        serial = float(value)
    except (TypeError, ValueError):
        return None
    # Excel's 1900 date system (including the historic leap-year quirk).
    return datetime(1899, 12, 30) + timedelta(days=serial)


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
