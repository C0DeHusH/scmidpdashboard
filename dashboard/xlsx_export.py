from __future__ import annotations

from datetime import datetime, timezone
from io import BytesIO
from typing import Any, Dict, Iterable, List
from xml.sax.saxutils import escape
from decimal import Decimal, ROUND_HALF_UP
import zipfile


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
    total_qty = _whole(sum(float(i.get("requested_qty", 0) or 0) for i in items))
    risk = sum(1 for i in items if str(i.get("stock_status", "")).lower() in {"stockout", "critical", "re-order", "reorder"})

    rows = []
    rows.append(_row(1, [_cell("A1", "BRANCH REQUEST STATUS REPORT", 1)], 24))
    rows.append(_row(2, [_cell("A2", "MUTI MC SCM Executive Control Tower • Branch Request Report", 2)], 18))
    rows.append(_row(3, [] , 8))
    rows.append(_row(4, [_cell("A4", "REQUESTING BRANCH", 3), _cell("C4", branch, 4), _cell("F4", "REPORT GENERATED", 3), _cell("H4", now.strftime("%d %b %Y • %I:%M %p"), 4)], 18))
    rows.append(_row(5, [_cell("A5", "AREA", 3), _cell("C5", area, 4)], 18))
    rows.append(_row(6, [], 8))
    rows.append(_row(7, [_cell("A7", "REQUESTED ITEMS", 3), _cell("C7", "TOTAL REQUEST QTY", 3), _cell("F7", "RISK ITEMS", 3)], 18))
    rows.append(_row(8, [_cell("A8", len(items), 5, True), _cell("C8", total_qty, 5, True), _cell("F8", risk, 5, True)], 22))
    rows.append(_row(9, [], 8))

    headers = ["NO.", "MODEL", "CLASS", "INVENTORY", "REQUEST QTY", "STOCK STATUS", "CURRENT DoI", "NEW DoI", "REMARKS"]
    rows.append(_row(10, [_cell(f"{chr(65+i)}10", h, 6) for i, h in enumerate(headers)], 22))
    start = 11
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

    merges = ["A1:I1", "A2:I2", "A4:B4", "C4:E4", "F4:G4", "H4:I4", "A5:B5", "C5:E5", "A7:B7", "C7:E7", "F7:G7", "A8:B8", "C8:E8", "F8:G8", f"A{sig}:C{sig}", f"D{sig}:F{sig}", f"G{sig}:I{sig}", f"A{sig+1}:C{sig+1}", f"D{sig+1}:F{sig+1}", f"G{sig+1}:I{sig+1}"]

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
<pageSetup orientation="portrait" paperSize="9" fitToWidth="1" fitToHeight="0" horizontalDpi="300" verticalDpi="300"/>
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
<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">'Branch Request'!$A$1:$I${sig+1}</definedName><definedName name="_xlnm.Print_Titles" localSheetId="0">'Branch Request'!$10:$10</definedName></definedNames>
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


def build_delivery_plan_xlsx(
    day: str,
    analysis: Dict[str, Any],
    weekly_analysis: Dict[str, Dict[str, Any]],
    schedule: List[Dict[str, Any]],
) -> bytes:
    """Build a highly formatted, A4 landscape, print-ready Delivery Plan workbook.

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
<sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="10" topLeftCell="A11" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
<cols>{cols_xml1}</cols><sheetData>{''.join(rows1)}</sheetData>
<mergeCells count="{len(merges1)}">{''.join(f'<mergeCell ref="{m}"/>' for m in merges1)}</mergeCells>
<printOptions horizontalCentered="1" verticalCentered="0"/>
<pageMargins left="0.2" right="0.2" top="0.35" bottom="0.35" header="0.15" footer="0.15"/>
<pageSetup orientation="landscape" paperSize="9" fitToWidth="1" fitToHeight="0" horizontalDpi="300" verticalDpi="300"/>
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
<sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="6" topLeftCell="A7" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
<cols>{cols_xml2}</cols><sheetData>{''.join(rows2)}</sheetData>
<mergeCells count="{len(merges2)}">{''.join(f'<mergeCell ref="{m}"/>' for m in merges2)}</mergeCells>
<printOptions horizontalCentered="1" verticalCentered="0"/><pageMargins left="0.2" right="0.2" top="0.35" bottom="0.35" header="0.15" footer="0.15"/>
<pageSetup orientation="landscape" paperSize="9" fitToWidth="1" fitToHeight="0" horizontalDpi="300" verticalDpi="300"/>
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
<xf numFmtId="0" fontId="3" fillId="5" borderId="1" xfId="0"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="3" fillId="6" borderId="1" xfId="0"><alignment vertical="center" wrapText="1"/></xf>
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


def _num_local(v: Any) -> float:
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0
