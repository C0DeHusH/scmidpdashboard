from __future__ import annotations

from datetime import datetime, timezone
from io import BytesIO
from typing import Any, Dict, Iterable, List
from xml.sax.saxutils import escape
from decimal import Decimal, ROUND_HALF_UP
import zipfile
import re


def _whole(value: Any) -> Any:
    if not isinstance(value, (int, float)):
        return value
    return int(Decimal(str(value)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def _cell(ref: str, value: Any, style: int = 0, num: bool = False) -> str:
    if value is None:
        return f'<c r="{ref}" s="{style}"/>'
    if num and isinstance(value, (int, float)):
        return f'<c r="{ref}" s="{style}"><v>{value}</v></c>'
    text = escape(str(value))
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'


def _row(r: int, cells: Iterable[str], height: int | None = None) -> str:
    ht = f' ht="{height}" customHeight="1"' if height else ""
    return f'<row r="{r}"{ht}>' + "".join(cells) + "</row>"


def build_branch_request_xlsx(branch: str, area: str, items: List[Dict[str, Any]]) -> bytes:
    now = datetime.now().astimezone()
    rows = []
    rows.append(_row(1, [_cell("A1", "BRANCH REQUEST STATUS REPORT", 1)], 24))
    rows.append(_row(2, [_cell("A2", "MUTI MC SCM Executive Control Tower • Branch Request Report", 2)], 18))
    rows.append(_row(3, [] , 8))
    rows.append(_row(4, [_cell("A4", "REQUESTING BRANCH", 3), _cell("C4", branch, 4), _cell("F4", "REPORT GENERATED", 3), _cell("H4", now.strftime("%d %b %Y • %I:%M %p"), 4)], 18))
    rows.append(_row(5, [_cell("A5", "AREA", 3), _cell("C5", area, 4)], 18))
    rows.append(_row(6, [], 8))

    headers = ["NO.", "MODEL", "CLASS", "INVENTORY", "REQUEST QTY", "STOCK STATUS", "CURRENT DoI", "NEW DoI", "REMARKS"]
    rows.append(_row(7, [_cell(f"{chr(65+i)}7", h, 6) for i, h in enumerate(headers)], 22))
    start = 8
    for idx, item in enumerate(items, 1):
        r = start + idx - 1
        vals = [
            idx, item.get("model", ""), item.get("class", ""), _whole(item.get("inventory", 0)),
            _whole(item.get("requested_qty", 0)), item.get("stock_status", ""), _whole(item.get("doi", 0)),
            _whole(item.get("new_doi", "N/A")), item.get("remarks", ""),
        ]
        cells = []
        for c, value in enumerate(vals):
            ref = f"{chr(65+c)}{r}"
            numeric = c in {0, 3, 4, 6, 7} and isinstance(value, (int, float))
            cells.append(_cell(ref, value, 7 if r % 2 else 8, numeric))
        rows.append(_row(r, cells, 24))

    sig = start + max(1, len(items)) + 3
    rows.append(_row(sig, [_cell(f"A{sig}", "__________________________", 9), _cell(f"D{sig}", "__________________________", 9), _cell(f"G{sig}", "____________", 9)], 18))
    rows.append(_row(sig + 1, [_cell(f"A{sig+1}", "Prepared / Reviewed By", 10), _cell(f"D{sig+1}", "Approved By", 10), _cell(f"G{sig+1}", "Date", 10)], 16))

    merges = ["A1:I1", "A2:I2", "A4:B4", "C4:E4", "F4:G4", "H4:I4", "A5:B5", "C5:E5", f"A{sig}:C{sig}", f"D{sig}:F{sig}", f"G{sig}:I{sig}", f"A{sig+1}:C{sig+1}", f"D{sig+1}:F{sig+1}", f"G{sig+1}:I{sig+1}"]

    sheet_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
<sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>
<cols>
<col min="1" max="1" width="5.5" customWidth="1"/><col min="2" max="2" width="20" customWidth="1"/>
<col min="3" max="3" width="8" customWidth="1"/><col min="4" max="4" width="9" customWidth="1"/>
<col min="5" max="5" width="10" customWidth="1"/><col min="6" max="6" width="13" customWidth="1"/>
<col min="7" max="8" width="10" customWidth="1"/><col min="9" max="9" width="22" customWidth="1"/>
</cols>
<sheetData>{''.join(rows)}</sheetData>
<mergeCells count="{len(merges)}">{''.join(f'<mergeCell ref="{m}"/>' for m in merges)}</mergeCells>
<printOptions horizontalCentered="1" verticalCentered="0"/>
<pageMargins left="0.25" right="0.25" top="0.35" bottom="0.35" header="0.15" footer="0.15"/>
<pageSetup orientation="portrait" paperSize="1" fitToWidth="1" fitToHeight="0" horizontalDpi="300" verticalDpi="300"/>
</worksheet>'''

    styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="8">
<font><sz val="8.5"/><name val="Aptos"/></font>
<font><b/><sz val="14"/><color rgb="FFFFFFFF"/><name val="Aptos Display"/></font>
<font><b/><sz val="8"/><color rgb="FFFBBF24"/><name val="Aptos"/></font>
<font><b/><sz val="8.5"/><color rgb="FF0F172A"/><name val="Aptos"/></font>
<font><sz val="8"/><color rgb="FF475569"/><name val="Aptos"/></font>
<font><b/><sz val="11"/><color rgb="FF0F172A"/><name val="Aptos"/></font>
<font><b/><sz val="8"/><color rgb="FFFFFFFF"/><name val="Aptos"/></font>
<font><b/><sz val="8.5"/><color rgb="FFFFFFFF"/><name val="Aptos"/></font>
</fonts>
<fills count="7">
<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF0F172A"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1E293B"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFBBF24"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF8FAFC"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF1F5F9"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2"><border/><border><left style="thin"><color rgb="FFE2E8F0"/></left><right style="thin"><color rgb="FFE2E8F0"/></right><top style="thin"><color rgb="FFE2E8F0"/></top><bottom style="thin"><color rgb="FFE2E8F0"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="11">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0"><alignment horizontal="left" vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="7" fillId="3" borderId="1" xfId="0"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="5" fillId="4" borderId="1" xfId="0"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="6" fillId="2" borderId="1" xfId="0"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="3" fillId="5" borderId="1" xfId="0"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="3" fillId="6" borderId="1" xfId="0"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0"><alignment horizontal="center"/></xf>
<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0"><alignment horizontal="center"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''

    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''
    wb_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Branch Request" sheetId="1" r:id="rId1"/></sheets>
<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">'Branch Request'!$A$1:$I${sig+1}</definedName><definedName name="_xlnm.Print_Titles" localSheetId="0">'Branch Request'!$7:$7</definedName></definedNames>
</workbook>'''
    wb_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''
    stamp = now.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    core = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>SCM Branch Request</dc:title><dc:creator>MUTI MC SCM Executive Control Tower</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">{stamp}</dcterms:created></cp:coreProperties>'''
    app = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>SCM IDP Dashboard</Application></Properties>'''

    out = BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("xl/workbook.xml", wb_xml)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/styles.xml", styles_xml)
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        z.writestr("docProps/core.xml", core)
        z.writestr("docProps/app.xml", app)
    return out.getvalue()


def _col_letter(n: int) -> str:
    out = ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def _status_style(status: str, base: int = 8) -> int:
    s = str(status or "").upper()
    if s == "OVERLOAD":
        return 10
    if "FULL" in s or "HIGH" in s:
        return 11
    if s == "OPTIMIZED":
        return 12
    if s == "UNDERUTILIZED":
        return 13
    if "NO ALLOCATION" in s:
        return 14
    return base


def _build_delivery_plan_xlsx_legacy(
    day: str,
    analysis: Dict[str, Any],
    weekly_analysis: Dict[str, Dict[str, Any]],
    schedule: List[Dict[str, Any]],
) -> bytes:
    """Build a highly formatted, Letter landscape, print-ready Delivery Plan workbook.

    Sheet 1 is the selected day's operational dispatch plan. Sheet 2 is the complete
    weekly truck schedule enriched with allocation and utilization results.
    """
    now = datetime.now().astimezone()
    s = analysis.get("summary", {}) or {}
    assignments = analysis.get("assignments", []) or []
    branch_priorities = analysis.get("branch_priorities", []) or []
    unassigned = analysis.get("unassigned", []) or []

    # ---------------- Sheet 1: Daily Delivery Plan ----------------
    rows1: List[str] = []
    rows1.append(_row(1, [_cell("A1", "SCM DAILY DELIVERY PLAN", 1)], 28))
    rows1.append(_row(2, [_cell("A2", "Inventory & Distribution Planning • Truck Capacity & Priority Loading Report", 2)], 18))
    rows1.append(_row(3, [], 7))
    rows1.append(_row(4, [
        _cell("A4", "DELIVERY DAY", 3), _cell("C4", day, 4),
        _cell("F4", "REPORT GENERATED", 3), _cell("H4", now.strftime("%d %b %Y • %I:%M %p"), 4),
    ], 20))
    rows1.append(_row(5, [], 7))

    kpis = [
        ("Scheduled Branches", s.get("scheduled_branches", 0)),
        ("Branches With Allocation", s.get("branches_with_allocation", 0)),
        ("Total Capacity Index", s.get("total_capacity", 0)),
        ("Required Load Index", s.get("total_load_index", 0)),
        ("Fleet Utilization", f"{_whole(s.get('fleet_utilization', 0))}%"),
        ("Overloaded Trucks", s.get("overloaded_trucks", 0)),
    ]
    label_cells, value_cells = [], []
    for i, (label, value) in enumerate(kpis):
        c1 = 1 + i * 2
        c2 = c1 + 1
        label_cells.append(_cell(f"{_col_letter(c1)}6", label, 5))
        value_cells.append(_cell(f"{_col_letter(c1)}7", _whole(value) if isinstance(value, (int, float)) else value, 6, isinstance(value, (int, float))))
    rows1.append(_row(6, label_cells, 18))
    rows1.append(_row(7, value_cells, 25))
    rows1.append(_row(8, [], 8))

    rows1.append(_row(9, [_cell("A9", "TRUCK DISPATCH & CAPACITY SUMMARY", 15)], 22))
    truck_headers = ["NO.", "TRUCK", "SCHEDULED BRANCHES", "AREA(S)", "CAPACITY", "REQ. LOAD", "UTILIZATION", "STATUS", "UNITS", "CLASS A", "DEFERRED", "PLANNED LOAD", "NOTES"]
    rows1.append(_row(10, [_cell(f"{_col_letter(i+1)}10", h, 7) for i, h in enumerate(truck_headers)], 25))
    r = 11
    visible_assignments = [a for a in assignments if a.get("branches") or str(a.get("status", "")).upper() != "IDLE"]
    for idx, a in enumerate(visible_assignments, 1):
        base = 8 if idx % 2 else 9
        notes = []
        if a.get("empty_scheduled_branches"):
            notes.append("No Allocation: " + ", ".join(a.get("empty_scheduled_branches") or []))
        if a.get("unknown_models"):
            notes.append("Index missing: " + ", ".join(a.get("unknown_models") or []))
        values = [
            idx,
            a.get("plate", ""),
            " / ".join(a.get("branches") or []) or "—",
            " / ".join(a.get("areas") or []) or "—",
            _whole(a.get("capacity", 0)),
            _whole(a.get("load_index", 0)),
            f"{_whole(a.get('utilization', 0))}%",
            a.get("status", ""),
            _whole(a.get("quantity", 0)),
            _whole(a.get("class_a_qty", 0)),
            _whole(a.get("deferred_units", 0)),
            _whole(a.get("planned_load_index", 0)),
            " • ".join(notes),
        ]
        cells = []
        for c, value in enumerate(values, 1):
            st = base
            if c == 8:
                st = _status_style(str(value), base)
            elif c == 10 and isinstance(value, (int, float)) and value > 0:
                st = 16
            elif c == 11 and isinstance(value, (int, float)) and value > 0:
                st = 17
            numeric = c in {1, 5, 6, 9, 10, 11, 12} and isinstance(value, (int, float))
            cells.append(_cell(f"{_col_letter(c)}{r}", value, st, numeric))
        rows1.append(_row(r, cells, 26))
        r += 1
    if not visible_assignments:
        rows1.append(_row(r, [_cell(f"A{r}", "No truck schedule exists for the selected day.", 8)], 24))
        r += 1

    r += 1
    rows1.append(_row(r, [_cell(f"A{r}", "MODEL PRIORITY LOADING RECOMMENDATION", 15)], 22)); priority_title_row = r
    r += 1
    priority_headers = ["TRUCK", "BRANCH", "AREA", "MODEL", "CLASS", "STOCK STATUS", "PRIORITY", "REQUESTED", "PLANNED", "DEFERRED", "INDEX", "LOAD INDEX", "REMARKS"]
    rows1.append(_row(r, [_cell(f"{_col_letter(i+1)}{r}", h, 7) for i, h in enumerate(priority_headers)], 25)); priority_header_row = r
    r += 1
    pidx = 0
    for a in assignments:
        if not a.get("branches"):
            continue
        for item in a.get("priority_items", []) or []:
            pidx += 1
            base = 8 if pidx % 2 else 9
            values = [
                a.get("plate", ""), item.get("branch", ""), item.get("area", ""), item.get("model", ""), item.get("class", ""),
                item.get("stock_status", ""), item.get("priority", ""), _whole(item.get("quantity", 0)), _whole(item.get("planned_units", 0)),
                _whole(item.get("deferred_units", 0)), item.get("index_size", 1), _whole(item.get("load_index", 0)), item.get("remarks", ""),
            ]
            cells = []
            for c, value in enumerate(values, 1):
                st = base
                if c == 5 and str(value).upper() == "A": st = 16
                if c == 6 and str(value).lower() in {"stockout", "stock out", "re-order", "reorder", "critical"}: st = 17
                if c == 10 and isinstance(value, (int, float)) and value > 0: st = 17
                numeric = c in {8, 9, 10, 11, 12} and isinstance(value, (int, float))
                cells.append(_cell(f"{_col_letter(c)}{r}", value, st, numeric))
            rows1.append(_row(r, cells, 25))
            r += 1
    if pidx == 0:
        rows1.append(_row(r, [_cell(f"A{r}", "No allocation models mapped to this day's scheduled trips.", 8)], 24)); r += 1

    r += 1
    branch_title_row = r
    rows1.append(_row(r, [_cell(f"A{r}", "BRANCH LOADING PRIORITY", 15)], 22)); r += 1
    branch_headers = ["PRIORITY", "TRUCK", "BRANCH", "AREA", "ALLOCATION STATUS", "TOTAL UNITS", "CLASS A", "RISK UNITS", "REQ. INDEX", "ACTION"]
    rows1.append(_row(r, [_cell(f"{_col_letter(i+1)}{r}", h, 7) for i, h in enumerate(branch_headers)], 25)); r += 1
    for idx, b in enumerate(branch_priorities, 1):
        base = 8 if idx % 2 else 9
        action = "Load Class A first" if _num_local(b.get("class_a_qty")) > 0 else ("Scheduled / No Allocation" if not b.get("has_allocation") else "Load as planned")
        vals = [idx, b.get("plate", ""), b.get("branch", ""), b.get("area", ""), "ALLOCATED" if b.get("has_allocation") else "NO ALLOCATION", _whole(b.get("quantity", 0)), _whole(b.get("class_a_qty", 0)), _whole(b.get("risk_qty", 0)), _whole(b.get("load_index", 0)), action]
        cells=[]
        for c,v in enumerate(vals,1):
            st=base
            if c==5 and v=="NO ALLOCATION": st=14
            if c==7 and isinstance(v,(int,float)) and v>0: st=16
            if c==8 and isinstance(v,(int,float)) and v>0: st=17
            numeric=c in {1,6,7,8,9} and isinstance(v,(int,float))
            cells.append(_cell(f"{_col_letter(c)}{r}",v,st,numeric))
        rows1.append(_row(r,cells,24)); r+=1
    if not branch_priorities:
        rows1.append(_row(r, [_cell(f"A{r}", "No scheduled branches for this day.", 8)], 24)); r+=1

    unscheduled_title_row = None
    if unassigned:
        r += 1
        unscheduled_title_row = r
        rows1.append(_row(r, [_cell(f"A{r}", "UNSCHEDULED IMPORTED ALLOCATIONS", 15)], 22)); r += 1
        uns_headers=["BRANCH","QUANTITY","ROWS","REQUIRED ACTION"]
        rows1.append(_row(r,[_cell(f"{_col_letter(i+1)}{r}",h,7) for i,h in enumerate(uns_headers)],25)); r+=1
        for idx,u in enumerate(unassigned,1):
            base=8 if idx%2 else 9
            vals=[u.get("branch",""),_whole(u.get("quantity",0)),_whole(u.get("rows",0)),"Add branch to Weekly Truck Schedule"]
            rows1.append(_row(r,[_cell(f"{_col_letter(c+1)}{r}",v,17 if c==3 else base,c in {1,2} and isinstance(v,(int,float))) for c,v in enumerate(vals)],24));r+=1

    end1=r
    merges1 = ["A1:M1", "A2:M2", "A4:B4", "C4:E4", "F4:G4", "H4:M4", "A9:M9", f"A{priority_title_row}:M{priority_title_row}", f"A{branch_title_row}:M{branch_title_row}"]
    if unscheduled_title_row:
        merges1.append(f"A{unscheduled_title_row}:M{unscheduled_title_row}")
    # KPI labels/values each span two columns.
    for i in range(6):
        c1=1+i*2; c2=c1+1
        merges1 += [f"{_col_letter(c1)}6:{_col_letter(c2)}6", f"{_col_letter(c1)}7:{_col_letter(c2)}7"]

    cols1 = [6,14,26,13,10,11,11,24,9,9,9,11,28]
    cols_xml1 = "".join(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i,w in enumerate(cols1,1))
    sheet1_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
<sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>
<cols>{cols_xml1}</cols><sheetData>{''.join(rows1)}</sheetData>
<mergeCells count="{len(merges1)}">{''.join(f'<mergeCell ref="{m}"/>' for m in merges1)}</mergeCells>
<printOptions horizontalCentered="1" verticalCentered="0"/>
<pageMargins left="0.2" right="0.2" top="0.35" bottom="0.35" header="0.15" footer="0.15"/>
<pageSetup orientation="landscape" paperSize="1" fitToWidth="1" fitToHeight="0" horizontalDpi="300" verticalDpi="300"/>
</worksheet>'''

    # ---------------- Sheet 2: Weekly Schedule ----------------
    rows2: List[str] = []
    rows2.append(_row(1, [_cell("A1", "WEEKLY TRUCK DELIVERY SCHEDULE", 1)], 28))
    rows2.append(_row(2, [_cell("A2", "Saved Day + Truck + Branch Schedule • Allocation Mapping & Utilization", 2)], 18))
    rows2.append(_row(3, [], 8))
    rows2.append(_row(4, [_cell("A4", "GENERATED", 3), _cell("C4", now.strftime("%d %b %Y • %I:%M %p"), 4), _cell("F4", "SCHEDULE ROWS", 3), _cell("H4", len(schedule), 4, True)], 20))
    rows2.append(_row(5, [], 8))
    wh = ["DAY", "TRUCK", "BRANCH", "AREA", "ALLOCATION", "UNITS", "CLASS A", "RISK UNITS", "REQ. INDEX", "TRUCK UTIL.", "TRUCK STATUS", "ACTION"]
    rows2.append(_row(6, [_cell(f"{_col_letter(i+1)}6", h, 7) for i,h in enumerate(wh)], 25))
    rr=7
    for dname in ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]:
        da = weekly_analysis.get(dname, {}) or {}
        pr_map={(str(x.get("plate","")).upper(),str(x.get("branch","")).upper()):x for x in da.get("branch_priorities",[]) or []}
        tr_map={str(x.get("plate","")).upper():x for x in da.get("assignments",[]) or []}
        day_rows=[x for x in schedule if x.get("day")==dname]
        for slot in day_rows:
            plate=str(slot.get("plate", "")); branch=str(slot.get("branch", "")); p=pr_map.get((plate.upper(),branch.upper()),{}); t=tr_map.get(plate.upper(),{})
            has=bool(p.get("has_allocation")); status=t.get("status", "")
            action="Priority Class A loading" if _num_local(p.get("class_a_qty"))>0 else ("Scheduled / No Allocation" if not has else "Load as planned")
            vals=[dname,plate,branch,p.get("area", ""),"ALLOCATED" if has else "NO ALLOCATION",_whole(p.get("quantity",0)),_whole(p.get("class_a_qty",0)),_whole(p.get("risk_qty",0)),_whole(p.get("load_index",0)),f"{_whole(t.get('utilization',0))}%",status,action]
            base=8 if rr%2 else 9
            cells=[]
            for c,v in enumerate(vals,1):
                st=base
                if c==5 and v=="NO ALLOCATION": st=14
                if c==7 and isinstance(v,(int,float)) and v>0: st=16
                if c==8 and isinstance(v,(int,float)) and v>0: st=17
                if c==11: st=_status_style(str(v),base)
                numeric=c in {6,7,8,9} and isinstance(v,(int,float))
                cells.append(_cell(f"{_col_letter(c)}{rr}",v,st,numeric))
            rows2.append(_row(rr,cells,24)); rr+=1
    if rr==7:
        rows2.append(_row(rr,[_cell(f"A{rr}","No weekly schedule saved.",8)],24));rr+=1
    end2=rr-1
    merges2=["A1:L1","A2:L2","A4:B4","C4:E4","F4:G4","H4:I4"]
    cols2=[11,14,28,14,15,9,9,10,11,11,24,24]
    cols_xml2="".join(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i,w in enumerate(cols2,1))
    sheet2_xml=f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
<sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>
<cols>{cols_xml2}</cols><sheetData>{''.join(rows2)}</sheetData>
<mergeCells count="{len(merges2)}">{''.join(f'<mergeCell ref="{m}"/>' for m in merges2)}</mergeCells>
<printOptions horizontalCentered="1" verticalCentered="0"/><pageMargins left="0.2" right="0.2" top="0.35" bottom="0.35" header="0.15" footer="0.15"/>
<pageSetup orientation="landscape" paperSize="1" fitToWidth="1" fitToHeight="0" horizontalDpi="300" verticalDpi="300"/>
</worksheet>'''

    styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="10">
<font><sz val="8.5"/><name val="Aptos"/></font>
<font><b/><sz val="16"/><color rgb="FFFFFFFF"/><name val="Aptos Display"/></font>
<font><b/><sz val="8.5"/><color rgb="FFFBBF24"/><name val="Aptos"/></font>
<font><b/><sz val="8.5"/><color rgb="FF0F172A"/><name val="Aptos"/></font>
<font><sz val="8"/><color rgb="FF475569"/><name val="Aptos"/></font>
<font><b/><sz val="12"/><color rgb="FF0F172A"/><name val="Aptos"/></font>
<font><b/><sz val="8"/><color rgb="FFFFFFFF"/><name val="Aptos"/></font>
<font><b/><sz val="9"/><color rgb="FF0F172A"/><name val="Aptos"/></font>
<font><b/><sz val="8.5"/><color rgb="FFB91C1C"/><name val="Aptos"/></font>
<font><b/><sz val="8.5"/><color rgb="FF92400E"/><name val="Aptos"/></font>
</fonts>
<fills count="12">
<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF0F172A"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1E293B"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFBBF24"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF8FAFC"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF1F5F9"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFEE2E2"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFEF3C7"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFDCFCE7"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFEDE9FE"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFFBEB"/></patternFill></fill>
</fills>
<borders count="2"><border/><border><left style="thin"><color rgb="FFE2E8F0"/></left><right style="thin"><color rgb="FFE2E8F0"/></right><top style="thin"><color rgb="FFE2E8F0"/></top><bottom style="thin"><color rgb="FFE2E8F0"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="19">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0"><alignment horizontal="left" vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="6" fillId="3" borderId="1" xfId="0"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="6" fillId="3" borderId="1" xfId="0"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="5" fillId="4" borderId="1" xfId="0"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="6" fillId="2" borderId="1" xfId="0"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="3" fillId="5" borderId="1" xfId="0"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="3" fillId="6" borderId="1" xfId="0"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="8" fillId="7" borderId="1" xfId="0"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="9" fillId="8" borderId="1" xfId="0"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="7" fillId="9" borderId="1" xfId="0"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="9" fillId="11" borderId="1" xfId="0"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="7" fillId="10" borderId="1" xfId="0"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="6" fillId="3" borderId="1" xfId="0"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="9" fillId="4" borderId="1" xfId="0"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="8" fillId="7" borderId="1" xfId="0"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0"><alignment horizontal="center" vertical="center"/></xf>
</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''

    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''
    wb_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Daily Delivery Plan" sheetId="1" r:id="rId1"/><sheet name="Weekly Schedule" sheetId="2" r:id="rId2"/></sheets>
<definedNames>
<definedName name="_xlnm.Print_Area" localSheetId="0">'Daily Delivery Plan'!$A$1:$M${end1}</definedName>
<definedName name="_xlnm.Print_Titles" localSheetId="0">'Daily Delivery Plan'!$1:$2</definedName>
<definedName name="_xlnm.Print_Area" localSheetId="1">'Weekly Schedule'!$A$1:$L${max(6,end2)}</definedName>
<definedName name="_xlnm.Print_Titles" localSheetId="1">'Weekly Schedule'!$1:$6</definedName>
</definedNames></workbook>'''
    wb_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''
    stamp = now.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    core = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>SCM Delivery Plan</dc:title><dc:creator>MUTI MC SCM Executive Control Tower</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">{stamp}</dcterms:created></cp:coreProperties>'''
    app_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>SCM IDP Dashboard</Application></Properties>'''

    out = BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("xl/workbook.xml", wb_xml)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/styles.xml", styles_xml)
        z.writestr("xl/worksheets/sheet1.xml", sheet1_xml)
        z.writestr("xl/worksheets/sheet2.xml", sheet2_xml)
        z.writestr("docProps/core.xml", core)
        z.writestr("docProps/app.xml", app_xml)
    return out.getvalue()


def build_delivery_plan_xlsx(
    day: str,
    analysis: Dict[str, Any],
    weekly_analysis: Dict[str, Dict[str, Any]],
    schedule: List[Dict[str, Any]],
) -> bytes:
    """Build a whole-week Delivery Plan workbook with one day per worksheet.

    Each Monday-Sunday worksheet is the complete operational daily plan for that day,
    followed by a Weekly Schedule worksheet. All sheets are Letter landscape, fit one
    page wide, and intentionally have no frozen panes.
    """
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    daily_xml: Dict[str, str] = {}
    styles_xml = core_xml = app_xml = weekly_xml = ""

    for idx, day_name in enumerate(days):
        day_analysis = weekly_analysis.get(day_name) or analysis
        legacy = _build_delivery_plan_xlsx_legacy(day_name, day_analysis, weekly_analysis, schedule)
        with zipfile.ZipFile(BytesIO(legacy), "r") as z:
            daily_xml[day_name] = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
            if idx == 0:
                weekly_xml = z.read("xl/worksheets/sheet2.xml").decode("utf-8")
                styles_xml = z.read("xl/styles.xml").decode("utf-8")
                core_xml = z.read("docProps/core.xml").decode("utf-8")
                app_xml = z.read("docProps/app.xml").decode("utf-8")

    pane_re = re.compile(r'<pane[^>]*/>')
    daily_xml = {k: pane_re.sub('', v) for k, v in daily_xml.items()}
    weekly_xml = pane_re.sub('', weekly_xml)

    sheet_names = days + ["Weekly Schedule"]
    sheet_xmls = [daily_xml[d] for d in days] + [weekly_xml]
    active_tab = days.index(day) if day in days else 0

    def _last_row(xml: str) -> int:
        rows = [int(x) for x in re.findall(r'<row r="(\d+)"', xml)]
        return max(rows) if rows else 1

    content_types = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                     '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
                     '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                     '<Default Extension="xml" ContentType="application/xml"/>\n'
                     '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>']
    for i in range(1, len(sheet_names) + 1):
        content_types.append(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    content_types.append('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                         '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
                         '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
                         '</Types>')
    content_types_xml = ''.join(content_types)

    root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                 '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
                 '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
                 '</Relationships>')

    sheet_nodes = ''.join(
        f'<sheet name="{escape(name)}" sheetId="{i}" r:id="rId{i}"/>'
        for i, name in enumerate(sheet_names, 1)
    )
    defined = []
    for i, (name, xml) in enumerate(zip(sheet_names, sheet_xmls)):
        last = _last_row(xml)
        last_col = 'L' if name == 'Weekly Schedule' else 'M'
        title_rows = '$1:$6' if name == 'Weekly Schedule' else '$1:$2'
        defined.append(f'<definedName name="_xlnm.Print_Area" localSheetId="{i}">\'{name}\'!$A$1:${last_col}${last}</definedName>')
        defined.append(f'<definedName name="_xlnm.Print_Titles" localSheetId="{i}">\'{name}\'!{title_rows}</definedName>')
    workbook_xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    f'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                    f'<bookViews><workbookView activeTab="{active_tab}"/></bookViews>'
                    f'<sheets>{sheet_nodes}</sheets><definedNames>{"".join(defined)}</definedNames></workbook>')

    rel_nodes = ''.join(
        f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
        for i in range(1, len(sheet_names) + 1)
    )
    style_rid = len(sheet_names) + 1
    workbook_rels = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                     f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                     f'{rel_nodes}<Relationship Id="rId{style_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                     f'</Relationships>')

    out = BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types_xml)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook_xml)
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        z.writestr("xl/styles.xml", styles_xml)
        for i, xml in enumerate(sheet_xmls, 1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", xml)
        z.writestr("docProps/core.xml", core_xml)
        z.writestr("docProps/app.xml", app_xml)
    return out.getvalue()

def _num_local(v: Any) -> float:
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


def build_management_order_xlsx(data: Dict[str, Any]) -> bytes:
    """Build a print-ready Management Order Plan workbook with one sheet per Brand.

    v2.15: ordered items are grouped into separate Brand worksheets. Printable columns
    are Line No., Brand, Model, Unit Cost, Inv., DoI, Stock Status, PO Bal., Order Qty,
    New DoI, Total Amount and Remarks. Column widths and key alignments follow the
    approved management print specification.
    """
    now = datetime.now().astimezone()
    items = list(data.get("rows") or [])
    selected = data.get("selected") or {}
    title = str(data.get("title") or "Management Order Plan")

    def safe_sheet_name(value: Any, used: set[str]) -> str:
        raw = str(value or "Unspecified").strip() or "Unspecified"
        cleaned = "".join("_" if ch in '[]:*?/\\' else ch for ch in raw).strip(" '") or "Brand"
        base = cleaned[:31]
        name = base
        n = 2
        while name.lower() in used:
            suffix = f"_{n}"
            name = base[: max(1, 31 - len(suffix))] + suffix
            n += 1
        used.add(name.lower())
        return name

    grouped: Dict[str, List[Dict[str, Any]]] = {}
    for item in items:
        brand = str(item.get("brand") or "Unspecified").strip() or "Unspecified"
        grouped.setdefault(brand, []).append(item)
    if not grouped:
        grouped = {"Management Order Plan": []}

    used_names: set[str] = set()
    sheets: List[Dict[str, Any]] = []
    widths = [10.0, 13.0, 18.0, 15.0, 9.0, 8.0, 13.0, 18.0, 12.0, 10.0, 19.0, 30.0]
    headers = [
        "LINE NO.", "BRAND", "MODEL", "UNIT COST", "INV.", "DoI",
        "STOCK STATUS", "PO BAL.", "ORDER QTY", "NEW DoI", "TOTAL AMOUNT", "REMARKS",
    ]

    for brand in sorted(grouped, key=lambda x: x.lower()):
        brand_items = grouped[brand]
        sheet_name = safe_sheet_name(brand, used_names)
        rows_xml: List[str] = []
        rows_xml.append(_row(1, [_cell("A1", "MANAGEMENT ORDER PLAN", 1)], 23))
        rows_xml.append(_row(2, [_cell("A2", title, 2)], 17))
        rows_xml.append(_row(3, [], 5))
        filter_bits = [f"Brand: {brand}"]
        if selected.get("class") and selected.get("class") != "All Classes":
            filter_bits.append(f"Class: {selected.get('class')}")
        if selected.get("status") and selected.get("status") != "All Statuses":
            filter_bits.append(f"Status: {selected.get('status')}")
        if selected.get("model") and selected.get("model") != "All Models":
            filter_bits.append(f"Model: {selected.get('model')}")
        rows_xml.append(_row(4, [_cell("A4", " | ".join(filter_bits), 3)], 17))
        rows_xml.append(_row(5, [], 5))
        rows_xml.append(_row(6, [_cell(f"{_col_letter(i)}6", h, 4) for i, h in enumerate(headers, 1)], 25))

        start = 7
        order_qty_total = 0.0
        amount_total = 0.0
        for idx, item in enumerate(brand_items, 1):
            r = start + idx - 1
            odd = idx % 2 == 1
            order_qty = float(item.get("allocation", 0) or 0)
            total_amount = float(item.get("total_amount", 0) or 0)
            order_qty_total += order_qty
            amount_total += total_amount
            vals = [
                idx,
                item.get("brand", ""),
                item.get("model", ""),
                float(item.get("unit_cost", 0) or 0),
                _whole(item.get("inventory", 0)),
                _whole(item.get("doi", 0)),
                item.get("stock_status", ""),
                _whole(item.get("po_balance", 0)),
                _whole(order_qty),
                _whole(item.get("new_doi", 0)),
                total_amount,
                item.get("remarks", ""),
            ]
            cells: List[str] = []
            for c, value in enumerate(vals, 1):
                ref = f"{_col_letter(c)}{r}"
                if c in {1, 5, 6, 8, 9, 10}:
                    cells.append(_cell(ref, value, 7 if odd else 8, True))
                elif c in {4, 11}:
                    cells.append(_cell(ref, value, 9 if odd else 10, True))
                elif c == 7:
                    cells.append(_cell(ref, value, 18 if odd else 19))
                elif c == 12:
                    cells.append(_cell(ref, value, 11 if odd else 12))
                elif c == 3:
                    cells.append(_cell(ref, value, 13 if odd else 14))
                else:
                    cells.append(_cell(ref, value, 5 if odd else 6))
            rows_xml.append(_row(r, cells, 24))

        total_row = start + max(1, len(brand_items))
        if not brand_items:
            rows_xml.append(_row(start, [_cell("A7", "No ordered models for this Brand.", 20)], 26))
        rows_xml.append(_row(total_row, [
            _cell(f"A{total_row}", "GRAND TOTAL", 15),
            _cell(f"I{total_row}", _whole(order_qty_total), 16, True),
            _cell(f"K{total_row}", amount_total, 17, True),
        ], 22))

        merges = ["A1:L1", "A2:L2", "A4:L4", f"A{total_row}:H{total_row}"]
        if not brand_items:
            merges.append("A7:L7")
        cols_xml = "".join(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i, w in enumerate(widths, 1))
        sheet_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
<sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>
<cols>{cols_xml}</cols>
<sheetData>{''.join(rows_xml)}</sheetData>
<mergeCells count="{len(merges)}">{''.join(f'<mergeCell ref="{m}"/>' for m in merges)}</mergeCells>
<printOptions horizontalCentered="1" verticalCentered="0"/>
<pageMargins left="0.08" right="0.08" top="0.16" bottom="0.16" header="0.05" footer="0.05"/>
<pageSetup orientation="portrait" paperSize="1" fitToWidth="1" fitToHeight="0" horizontalDpi="300" verticalDpi="300"/>
</worksheet>'''
        sheets.append({"name": sheet_name, "xml": sheet_xml, "end_row": total_row})

    styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2"><numFmt numFmtId="164" formatCode="#,##0"/><numFmt numFmtId="165" formatCode="&quot;₱&quot;#,##0"/></numFmts>
<fonts count="8">
<font><sz val="7.5"/><name val="Aptos"/></font>
<font><b/><sz val="14"/><color rgb="FFFFFFFF"/><name val="Aptos Display"/></font>
<font><b/><sz val="7.5"/><color rgb="FFFBBF24"/><name val="Aptos"/></font>
<font><sz val="7"/><color rgb="FF475569"/><name val="Aptos"/></font>
<font><b/><sz val="7"/><color rgb="FFFFFFFF"/><name val="Aptos"/></font>
<font><b/><sz val="7.5"/><color rgb="FF0F172A"/><name val="Aptos"/></font>
<font><sz val="7"/><color rgb="FF334155"/><name val="Aptos"/></font>
<font><b/><sz val="7.5"/><color rgb="FFFBBF24"/><name val="Aptos"/></font>
</fonts>
<fills count="8">
<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF0F172A"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1E293B"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF8FAFC"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF1F5F9"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFFBEB"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFBBF24"/></patternFill></fill>
</fills>
<borders count="2"><border/><border><left style="thin"><color rgb="FFE2E8F0"/></left><right style="thin"><color rgb="FFE2E8F0"/></right><top style="thin"><color rgb="FFE2E8F0"/></top><bottom style="thin"><color rgb="FFE2E8F0"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="21">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0"><alignment horizontal="left" vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0"><alignment horizontal="left" vertical="center"/></xf>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0"><alignment vertical="center" shrinkToFit="1"/></xf>
<xf numFmtId="0" fontId="4" fillId="2" borderId="1" xfId="0"><alignment horizontal="center" vertical="center" shrinkToFit="1"/></xf>
<xf numFmtId="0" fontId="5" fillId="4" borderId="1" xfId="0"><alignment horizontal="left" vertical="center" shrinkToFit="1"/></xf>
<xf numFmtId="0" fontId="5" fillId="5" borderId="1" xfId="0"><alignment horizontal="left" vertical="center" shrinkToFit="1"/></xf>
<xf numFmtId="164" fontId="5" fillId="4" borderId="1" xfId="0"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="164" fontId="5" fillId="5" borderId="1" xfId="0"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="165" fontId="5" fillId="4" borderId="1" xfId="0"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="165" fontId="5" fillId="5" borderId="1" xfId="0"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="0" fontId="6" fillId="4" borderId="1" xfId="0"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="6" fillId="5" borderId="1" xfId="0"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="5" fillId="4" borderId="1" xfId="0"><alignment horizontal="left" vertical="center" shrinkToFit="1"/></xf>
<xf numFmtId="0" fontId="5" fillId="5" borderId="1" xfId="0"><alignment horizontal="left" vertical="center" shrinkToFit="1"/></xf>
<xf numFmtId="0" fontId="4" fillId="3" borderId="1" xfId="0"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="164" fontId="4" fillId="3" borderId="1" xfId="0"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="165" fontId="4" fillId="3" borderId="1" xfId="0"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="0" fontId="5" fillId="6" borderId="1" xfId="0"><alignment horizontal="center" vertical="center" shrinkToFit="1"/></xf>
<xf numFmtId="0" fontId="5" fillId="6" borderId="1" xfId="0"><alignment horizontal="center" vertical="center" shrinkToFit="1"/></xf>
<xf numFmtId="0" fontId="6" fillId="4" borderId="1" xfId="0"><alignment horizontal="left" vertical="center" shrinkToFit="1"/></xf>
</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''

    content_overrides = ''.join(
        f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for i in range(1, len(sheets) + 1)
    )
    content_types = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
{content_overrides}
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

    sheet_entries = []
    defined_names = []
    wb_rel_entries = []
    for i, sh in enumerate(sheets, 1):
        name_attr = escape(sh["name"], {'"': '&quot;'})
        sheet_entries.append(f'<sheet name="{name_attr}" sheetId="{i}" r:id="rId{i}"/>')
        formula_name = sh["name"].replace("'", "''")
        defined_names.append(f'<definedName name="_xlnm.Print_Area" localSheetId="{i-1}">\'{escape(formula_name)}\'!$A$1:$L${sh["end_row"]}</definedName>')
        defined_names.append(f'<definedName name="_xlnm.Print_Titles" localSheetId="{i-1}">\'{escape(formula_name)}\'!$6:$6</definedName>')
        wb_rel_entries.append(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>')
    styles_rid = len(sheets) + 1
    wb_rel_entries.append(f'<Relationship Id="rId{styles_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
    wb_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>{''.join(sheet_entries)}</sheets>
<definedNames>{''.join(defined_names)}</definedNames>
</workbook>'''
    wb_rels = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{''.join(wb_rel_entries)}</Relationships>'''
    stamp = now.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    core = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>SCM Management Order Plan</dc:title><dc:creator>MUTI MC SCM Executive Control Tower</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">{stamp}</dcterms:created></cp:coreProperties>'''
    app_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>SCM IDP Dashboard</Application></Properties>'''

    out = BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("xl/workbook.xml", wb_xml)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/styles.xml", styles_xml)
        for i, sh in enumerate(sheets, 1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", sh["xml"])
        z.writestr("docProps/core.xml", core)
        z.writestr("docProps/app.xml", app_xml)
    return out.getvalue()



def build_weekly_schedule_template_xlsx(
    schedule: List[Dict[str, Any]],
    master: Dict[str, Any],
) -> bytes:
    """Build an import-compatible Weekly Truck Schedule workbook.

    Users can export the current schedule, edit Day / Truck / Area / Branch,
    then import the same file back. Area is informational and is reconciled
    against the Branch Master during import.
    """
    now = datetime.now().astimezone()
    branch_area = {
        str(b.get("branch", "")).strip().upper(): str(b.get("area", "")).strip()
        for b in (master.get("branches", []) or [])
        if str(b.get("branch", "")).strip()
    }
    valid_days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    day_rank = {d: i for i, d in enumerate(valid_days)}
    ordered = sorted(
        [dict(x) for x in (schedule or [])],
        key=lambda x: (
            day_rank.get(str(x.get("day", "")).strip(), 99),
            str(x.get("plate", "")),
            str(x.get("branch", "")),
        ),
    )

    rows: List[str] = []
    rows.append(_row(1, [_cell("A1", "WEEKLY TRUCK SCHEDULE TEMPLATE", 1)], 28))
    rows.append(_row(2, [_cell("A2", "SCM Inventory & Distribution Planning • Bulk Schedule Import / Export", 2)], 18))
    rows.append(_row(3, [_cell("A3", "Edit Day, Truck, Area and Branch. Importing this file replaces the saved weekly schedule after validation.", 3)], 22))
    rows.append(_row(4, [], 7))
    headers = ["DAY", "TRUCK", "AREA", "BRANCH"]
    rows.append(_row(5, [_cell(f"{chr(65+i)}5", h, 4) for i, h in enumerate(headers)], 22))
    start = 6
    for idx, item in enumerate(ordered):
        r = start + idx
        branch = str(item.get("branch", "") or "").strip()
        vals = [
            str(item.get("day", "") or "").strip(),
            str(item.get("plate", "") or "").strip(),
            branch_area.get(branch.upper(), str(item.get("area", "") or "").strip()),
            branch,
        ]
        style = 5 if idx % 2 == 0 else 6
        rows.append(_row(r, [_cell(f"{chr(65+i)}{r}", v, style) for i, v in enumerate(vals)], 21))
    end_row = max(6, start + len(ordered) - 1)

    sheet_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
<sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>
<cols>
<col min="1" max="1" width="15" customWidth="1"/>
<col min="2" max="2" width="18" customWidth="1"/>
<col min="3" max="3" width="18" customWidth="1"/>
<col min="4" max="4" width="34" customWidth="1"/>
</cols>
<sheetData>{''.join(rows)}</sheetData>
<mergeCells count="3"><mergeCell ref="A1:D1"/><mergeCell ref="A2:D2"/><mergeCell ref="A3:D3"/></mergeCells>
<autoFilter ref="A5:D{end_row}"/>
<printOptions horizontalCentered="1" verticalCentered="0"/>
<pageMargins left="0.25" right="0.25" top="0.35" bottom="0.35" header="0.15" footer="0.15"/>
<pageSetup orientation="landscape" paperSize="1" fitToWidth="1" fitToHeight="0" horizontalDpi="300" verticalDpi="300"/>
</worksheet>'''

    styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="6">
<font><sz val="9"/><name val="Aptos"/></font>
<font><b/><sz val="15"/><color rgb="FFFFFFFF"/><name val="Aptos Display"/></font>
<font><b/><sz val="8.5"/><color rgb="FFFBBF24"/><name val="Aptos"/></font>
<font><sz val="8.5"/><color rgb="FF475569"/><name val="Aptos"/></font>
<font><b/><sz val="9"/><color rgb="FFFFFFFF"/><name val="Aptos"/></font>
<font><sz val="9"/><color rgb="FF0F172A"/><name val="Aptos"/></font>
</fonts>
<fills count="6">
<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF0F172A"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1E293B"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF8FAFC"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF1F5F9"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2"><border/><border><left style="thin"><color rgb="FFE2E8F0"/></left><right style="thin"><color rgb="FFE2E8F0"/></right><top style="thin"><color rgb="FFE2E8F0"/></top><bottom style="thin"><color rgb="FFE2E8F0"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="7">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="4" fillId="3" borderId="1" xfId="0"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="5" fillId="4" borderId="1" xfId="0"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="5" fillId="5" borderId="1" xfId="0"><alignment vertical="center"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''

    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''
    wb_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Weekly Schedule" sheetId="1" r:id="rId1"/></sheets>
<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">'Weekly Schedule'!$A$1:$D${end_row}</definedName><definedName name="_xlnm.Print_Titles" localSheetId="0">'Weekly Schedule'!$5:$5</definedName></definedNames>
</workbook>'''
    wb_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''
    stamp = now.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    core = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>Weekly Truck Schedule Template</dc:title><dc:creator>SCM IDP Dashboard</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">{stamp}</dcterms:created></cp:coreProperties>'''
    app_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>SCM IDP Dashboard</Application></Properties>'''

    out = BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("xl/workbook.xml", wb_xml)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/styles.xml", styles_xml)
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        z.writestr("docProps/core.xml", core)
        z.writestr("docProps/app.xml", app_xml)
    return out.getvalue()
