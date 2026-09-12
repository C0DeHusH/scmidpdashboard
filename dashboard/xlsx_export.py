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
