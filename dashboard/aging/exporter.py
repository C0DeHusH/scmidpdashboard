from __future__ import annotations

from datetime import datetime, timezone
from io import BytesIO
from typing import Any, Iterable
from xml.sax.saxutils import escape
import zipfile


def _col_letter(index: int) -> str:
    """1-based Excel column index to letter."""
    out = ""
    n = int(index)
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def _cell(ref: str, value: Any, style: int = 0, *, numeric: bool = False) -> str:
    if value is None or value == "":
        return f'<c r="{ref}" s="{style}"/>'
    if numeric and isinstance(value, (int, float)):
        return f'<c r="{ref}" s="{style}"><v>{value}</v></c>'
    text = escape(str(value))
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'


def _row(row_num: int, cells: Iterable[str], *, height: float | None = None) -> str:
    ht = f' ht="{height}" customHeight="1"' if height else ""
    return f'<row r="{row_num}"{ht}>' + "".join(cells) + "</row>"


def _merge_xml(merges: list[str]) -> str:
    if not merges:
        return ""
    return f'<mergeCells count="{len(merges)}">' + "".join(f'<mergeCell ref="{m}"/>' for m in merges) + "</mergeCells>"


def _pane_xml(y_split: int = 0, x_split: int = 0, top_left: str | None = None) -> str:
    if not y_split and not x_split:
        return '<sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>'
    attrs = []
    if x_split:
        attrs.append(f'xSplit="{x_split}"')
    if y_split:
        attrs.append(f'ySplit="{y_split}"')
    attrs.append(f'topLeftCell="{top_left or "A1"}"')
    active = "bottomRight" if x_split and y_split else ("bottomLeft" if y_split else "topRight")
    attrs.append(f'activePane="{active}"')
    attrs.append('state="frozen"')
    return '<sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ' + " ".join(attrs) + '/></sheetView></sheetViews>'


def _sheet_xml(
    rows: list[str],
    cols: list[tuple[int, float]],
    *,
    merges: list[str] | None = None,
    autofilter: str | None = None,
    freeze_rows: int = 0,
    freeze_cols: int = 0,
    top_left: str | None = None,
    orientation: str = "landscape",
    print_area: str | None = None,
) -> str:
    col_xml = "".join(
        f'<col min="{i}" max="{i}" width="{width}" customWidth="1"/>' for i, width in cols
    )
    auto = f'<autoFilter ref="{autofilter}"/>' if autofilter else ""
    # print_area is expressed through workbook defined names; retained in args for readability.
    _ = print_area
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
{_pane_xml(freeze_rows, freeze_cols, top_left)}
<cols>{col_xml}</cols>
<sheetData>{''.join(rows)}</sheetData>
{_merge_xml(merges or [])}
{auto}
<printOptions horizontalCentered="0" verticalCentered="0"/>
<pageMargins left="0.28" right="0.28" top="0.45" bottom="0.45" header="0.2" footer="0.2"/>
<pageSetup orientation="{orientation}" paperSize="9" fitToWidth="1" fitToHeight="0" horizontalDpi="300" verticalDpi="300"/>
</worksheet>'''




def _money(value: Any) -> float:
    try:
        return round(float(value or 0), 2)
    except (TypeError, ValueError):
        return 0.0

def _age_style(days: int, alt: bool = False) -> int:
    if days > 365:
        return 16
    if days > 180:
        return 15
    if days > 90:
        return 14
    return 13


def _risk_style(level: str) -> int:
    return {"High": 18, "Watch": 19, "Controlled": 20}.get(str(level), 10)


def build_aging_report_xlsx(
    *,
    summary: dict[str, Any],
    unit_rows: list[dict[str, Any]],
    as_of: str,
    basis: str,
    filters: dict[str, str],
    unit_filters: dict[str, str],
    source_filename: str = "",
    filter_warnings: list[str] | None = None,
    base_row_count: int | None = None,
) -> bytes:
    """Build a management-ready multi-sheet Motorcycle Aging workbook.

    The workbook intentionally contains values only (no macros/external links),
    which keeps downloads portable across Excel desktop, web and mobile.
    """
    generated = datetime.now().astimezone()
    generated_utc = generated.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    basis_label = "Branch / Created On" if basis == "branch" else "Company / Incoming Date"
    filter_warnings = list(filter_warnings or [])
    if base_row_count is None:
        base_row_count = len(unit_rows)

    # ------------------------------------------------------------------
    # Sheet 1: Executive Summary
    # ------------------------------------------------------------------
    exec_rows: list[str] = []
    exec_merges = ["A1:H1", "A2:H2", "A4:H4", "A10:H10", "A20:H20", "A25:H25"]
    exec_rows.append(_row(1, [_cell("A1", "MOTORCYCLE AGING INTELLIGENCE", 1)], height=28))
    exec_rows.append(_row(2, [_cell("A2", "SCM IDP Control Tower · Executive Aging Export", 2)], height=20))
    exec_rows.append(_row(3, [
        _cell("A3", "As-of Date", 4), _cell("B3", as_of, 5),
        _cell("C3", "Age Basis", 4), _cell("D3", basis_label, 5),
        _cell("E3", "Generated", 4), _cell("F3", generated.strftime("%d %b %Y %I:%M %p"), 5),
        _cell("G3", "Source", 4), _cell("H3", source_filename or "Consolidated SCM Workbook", 5),
    ], height=20))
    exec_rows.append(_row(4, [_cell("A4", "EXECUTIVE SNAPSHOT", 3)], height=21))

    kpis = [
        ("Total Units", summary.get("total_qty", 0), 6),
        ("Inventory Value", _money(summary.get("total_value", 0)), 7),
        ("Average Age", summary.get("avg_age", 0), 6),
        ("91+ Day Units", summary.get("aged_90", 0), 6),
        ("91+ Exposure", float(summary.get("aged_90_pct", 0)) / 100, 8),
        ("91+ Aged Value", _money(summary.get("aged_90_value", 0)), 7),
        ("180+ Day Units", summary.get("aged_180", 0), 6),
        ("366+ Day Units", summary.get("aged_365", 0), 6),
        ("Healthy ≤90", summary.get("healthy_qty", 0), 6),
        ("Oldest Unit (Days)", summary.get("oldest", 0), 6),
        ("Branches", summary.get("branch_count", 0), 6),
        ("Areas", summary.get("area_count", 0), 6),
    ]
    for pair_index in range(0, len(kpis), 4):
        row_num = 5 + pair_index // 4
        cells = []
        for offset, (label, value, style) in enumerate(kpis[pair_index:pair_index + 4]):
            c = 1 + offset * 2
            cells.append(_cell(f"{_col_letter(c)}{row_num}", label, 4))
            cells.append(_cell(f"{_col_letter(c+1)}{row_num}", value, style, numeric=isinstance(value, (int, float))))
        exec_rows.append(_row(row_num, cells, height=22))

    risk_row = 8
    exec_rows.append(_row(risk_row, [
        _cell(f"A{risk_row}", "Management Exposure", 4),
        _cell(f"B{risk_row}", summary.get("risk_level", "—"), _risk_style(str(summary.get("risk_level", "")))),
        _cell(f"C{risk_row}", summary.get("risk_message", ""), 21),
    ], height=34))
    exec_merges.append(f"C{risk_row}:H{risk_row}")

    exec_rows.append(_row(10, [_cell("A10", "AGE PROFILE", 3)], height=21))
    exec_rows.append(_row(11, [
        _cell("A11", "Age Band", 9), _cell("B11", "Units", 9), _cell("C11", "Share", 9),
        _cell("D11", "Management Interpretation", 9),
    ], height=22))
    exec_merges.extend(["D11:H11"])
    total = float(summary.get("total_qty", 0) or 0)
    band_labels = summary.get("detailed_bucket_labels", [])
    band_values = summary.get("detailed_bucket_values", [])
    band_notes = {
        "0-30 DAYS": "Fresh inventory",
        "31-60 DAYS": "Normal rotation",
        "61-90 DAYS": "Monitor before action threshold",
        "91-180 DAYS": "Action required: sell / transfer / focus",
        "181-365 DAYS": "High aging exposure",
        "366+ DAYS": "Critical liquidation / recovery review",
    }
    for i, label in enumerate(band_labels):
        r = 12 + i
        value = float(band_values[i] if i < len(band_values) else 0)
        pct = value / total if total else 0
        age_style = 13 if i <= 2 else (14 if i == 3 else 15 if i == 4 else 16)
        exec_rows.append(_row(r, [
            _cell(f"A{r}", label, 10),
            _cell(f"B{r}", value, age_style, numeric=True),
            _cell(f"C{r}", pct, 8, numeric=True),
            _cell(f"D{r}", band_notes.get(label, ""), 10),
        ], height=20))
        exec_merges.append(f"D{r}:H{r}")

    exec_rows.append(_row(20, [_cell("A20", "TOP MANAGEMENT PRIORITIES", 3)], height=21))
    exec_rows.append(_row(21, [
        _cell("A21", "Priority", 9), _cell("B21", "Name", 9), _cell("C21", "91+ Units", 9),
        _cell("D21", "91+ %", 9), _cell("E21", "Aged Value", 9), _cell("F21", "Oldest", 9),
        _cell("G21", "Coverage", 9), _cell("H21", "Action", 9),
    ], height=22))
    priorities = [
        ("Area", summary.get("top_area")),
        ("Branch", summary.get("top_branch")),
        ("Model", summary.get("top_model")),
    ]
    for i, (kind, item) in enumerate(priorities, start=22):
        item = item or {}
        name = item.get("name") or item.get("standard_description") or "—"
        coverage = ""
        if kind == "Area":
            coverage = f"{item.get('branch_count', 0)} branches"
        elif kind == "Branch":
            coverage = str(item.get("area") or "—")
        else:
            coverage = f"{item.get('branch_count', 0)} branches / {item.get('area_count', 0)} areas"
        pct = float(item.get("aged_90_pct", 0) or 0) / 100
        action = "Prioritize action" if item else "No current exposure"
        exec_rows.append(_row(i, [
            _cell(f"A{i}", kind, 4), _cell(f"B{i}", name, 10),
            _cell(f"C{i}", item.get("aged_90", 0), 6, numeric=True),
            _cell(f"D{i}", pct, 8, numeric=True),
            _cell(f"E{i}", _money(item.get("aged_value", 0)), 7, numeric=True),
            _cell(f"F{i}", item.get("oldest", 0), 6, numeric=True),
            _cell(f"G{i}", coverage, 10), _cell(f"H{i}", action, 10),
        ], height=22))

    exec_rows.append(_row(25, [_cell("A25", "APPLIED FILTERS", 3)], height=21))
    filter_pairs = [
        ("Area", filters.get("area") or "All Areas"),
        ("Branch", filters.get("branch") or "All Branches"),
        ("Brand", filters.get("brand") or "All Brands"),
        ("Model Search", filters.get("std") or "All Models"),
        ("Unit Search", filters.get("q") or "None"),
        ("Trace Age Band", unit_filters.get("age") or "all"),
        ("Trace Search", unit_filters.get("q") or "None"),
        ("Trace Sort", unit_filters.get("sort") or "oldest"),
        ("Dashboard Rows", f"{int(base_row_count):,}"),
        ("Scope Repair", " | ".join(filter_warnings) if filter_warnings else "None"),
    ]
    # Build two filter pairs per row so the exported context stays compact.
    for pair_idx in range(0, len(filter_pairs), 2):
        r = 26 + pair_idx // 2
        cells = []
        for j, (label, value) in enumerate(filter_pairs[pair_idx:pair_idx + 2]):
            c = 1 + j * 4
            cells.extend([_cell(f"{_col_letter(c)}{r}", label, 4), _cell(f"{_col_letter(c+1)}{r}", value, 5)])
            # span value over three columns for readability
            exec_merges.append(f"{_col_letter(c+1)}{r}:{_col_letter(c+3)}{r}")
        exec_rows.append(_row(r, cells, height=20))

    sheet1 = _sheet_xml(
        exec_rows,
        [(1, 18), (2, 22), (3, 15), (4, 22), (5, 16), (6, 18), (7, 18), (8, 28)],
        merges=exec_merges,
        freeze_rows=3,
        top_left="A4",
        orientation="landscape",
    )

    # ------------------------------------------------------------------
    # Sheet 2: Area Intelligence
    # ------------------------------------------------------------------
    area_rows_data = summary.get("area_ranking_all") or summary.get("area_ranking") or []
    area_rows: list[str] = [
        _row(1, [_cell("A1", "PER-AREA AGING INTELLIGENCE", 1)], height=28),
        _row(2, [_cell("A2", f"As of {as_of} · {basis_label} · Areas ranked by 91+ aging exposure", 2)], height=20),
        _row(3, [_cell("A3", "Use this sheet to compare aging exposure, capital at risk, severity, and branch coverage by Area.", 22)], height=20),
        _row(5, [_cell(f"{_col_letter(i)}5", h, 9) for i, h in enumerate([
            "Rank", "Area", "Exposure", "Branches", "Units", "Inventory Value", "Avg Age",
            "91+ Units", "91+ %", "Aged Value", "180+ Units", "366+ Units", "Oldest Days"
        ], start=1)], height=24),
    ]
    area_merges = ["A1:M1", "A2:M2", "A3:M3"]
    if not area_rows_data:
        area_rows.append(_row(6, [
            _cell("A6", "No records matched the current dashboard filters. Review Applied Filters in Executive Summary.", 17),
        ], height=28))
        area_merges.append("A6:M6")
    for idx, a in enumerate(area_rows_data, start=1):
        r = 5 + idx
        pct = float(a.get("aged_90_pct", 0) or 0) / 100
        style = 10 if idx % 2 else 11
        risk = str(a.get("risk_level") or ("High" if pct >= .40 else "Watch" if pct >= .20 else "Controlled"))
        area_rows.append(_row(r, [
            _cell(f"A{r}", idx, style, numeric=True),
            _cell(f"B{r}", a.get("name"), style),
            _cell(f"C{r}", risk, _risk_style(risk)),
            _cell(f"D{r}", a.get("branch_count", 0), style, numeric=True),
            _cell(f"E{r}", a.get("qty", 0), style, numeric=True),
            _cell(f"F{r}", _money(a.get("value", 0)), 7, numeric=True),
            _cell(f"G{r}", a.get("avg_age", 0), style, numeric=True),
            _cell(f"H{r}", a.get("aged_90", 0), style, numeric=True),
            _cell(f"I{r}", pct, 8, numeric=True),
            _cell(f"J{r}", _money(a.get("aged_value", 0)), 7, numeric=True),
            _cell(f"K{r}", a.get("aged_180", 0), style, numeric=True),
            _cell(f"L{r}", a.get("aged_365", 0), style, numeric=True),
            _cell(f"M{r}", a.get("oldest", 0), _age_style(int(a.get("oldest", 0) or 0)), numeric=True),
        ], height=21))
    area_end = max(6, 5 + len(area_rows_data))
    sheet2 = _sheet_xml(
        area_rows,
        [(1, 7), (2, 22), (3, 13), (4, 10), (5, 10), (6, 17), (7, 10),
         (8, 11), (9, 10), (10, 17), (11, 11), (12, 11), (13, 12)],
        merges=area_merges,
        autofilter=f"A5:M{area_end}",
        freeze_rows=5,
        freeze_cols=2,
        top_left="C6",
        orientation="landscape",
    )

    # ------------------------------------------------------------------
    # Sheet 3: Model Intelligence
    # ------------------------------------------------------------------
    model_rows_data = summary.get("model_summary_all") or summary.get("model_summary") or []
    model_rows: list[str] = [
        _row(1, [_cell("A1", "MODEL-LEVEL AGING INTELLIGENCE", 1)], height=28),
        _row(2, [_cell("A2", f"As of {as_of} · {basis_label} · Sorted by 91+ Units Highest → Lowest", 2)], height=20),
        _row(4, [_cell(f"{_col_letter(i)}4", h, 9) for i, h in enumerate([
            "Rank", "Standard Description", "Exposure", "Units", "Avg Age", "91+ Units", "91+ %", "Aged Value", "Oldest", "Branches", "Areas"
        ], start=1)], height=24),
    ]
    model_merges = ["A1:K1", "A2:K2"]
    if not model_rows_data:
        model_rows.append(_row(5, [
            _cell("A5", "No model records matched the current dashboard filters. Review Applied Filters in Executive Summary.", 17),
        ], height=28))
        model_merges.append("A5:K5")
    for idx, m in enumerate(model_rows_data, start=1):
        r = 4 + idx
        pct = float(m.get("aged_90_pct", 0) or 0) / 100
        style = 10 if idx % 2 else 11
        model_rows.append(_row(r, [
            _cell(f"A{r}", idx, style, numeric=True),
            _cell(f"B{r}", m.get("standard_description"), style),
            _cell(f"C{r}", m.get("risk_level"), _risk_style(str(m.get("risk_level", "")))),
            _cell(f"D{r}", m.get("qty", 0), style, numeric=True),
            _cell(f"E{r}", m.get("avg_age", 0), style, numeric=True),
            _cell(f"F{r}", m.get("aged_90", 0), style, numeric=True),
            _cell(f"G{r}", pct, 8, numeric=True),
            _cell(f"H{r}", _money(m.get("aged_value", 0)), 7, numeric=True),
            _cell(f"I{r}", m.get("oldest", 0), _age_style(int(m.get("oldest", 0) or 0)), numeric=True),
            _cell(f"J{r}", m.get("branch_count", 0), style, numeric=True),
            _cell(f"K{r}", m.get("area_count", 0), style, numeric=True),
        ], height=21))
    model_end = max(5, 4 + len(model_rows_data))
    sheet3 = _sheet_xml(
        model_rows,
        [(1, 7), (2, 42), (3, 13), (4, 10), (5, 10), (6, 11), (7, 10), (8, 16), (9, 10), (10, 10), (11, 9)],
        merges=model_merges,
        autofilter=f"A4:K{model_end}",
        freeze_rows=4,
        freeze_cols=2,
        top_left="C5",
        orientation="landscape",
    )

    # ------------------------------------------------------------------
    # Sheet 4: Unit Detail
    # ------------------------------------------------------------------
    unit_headers = [
        "Branch", "Area", "Standard Description", "Brand", "Engine No.", "Chassis", "Created On",
        "Incoming Date", "Age Days", "Qty", "Unit Cost", "Inventory Value", "Location", "Barcode"
    ]
    unit_sheet_rows: list[str] = [
        _row(1, [_cell("A1", "UNIT-LEVEL AGING TRACEABILITY", 1)], height=28),
        _row(2, [_cell("A2", f"{len(unit_rows):,} matching rows · {as_of} · {basis_label}", 2)], height=20),
        _row(3, [_cell("A3", f"Unit filter: {unit_filters.get('age','all')} · Sort: {unit_filters.get('sort','oldest')} · Search: {unit_filters.get('q') or 'None'}", 22)], height=18),
        _row(5, [_cell(f"{_col_letter(i)}5", h, 9) for i, h in enumerate(unit_headers, start=1)], height=24),
    ]
    unit_merges = ["A1:N1", "A2:N2", "A3:N3"]
    if not unit_rows:
        reason = "No units matched the Unit-Level Traceability filters." if base_row_count else "No records matched the current dashboard filters."
        unit_sheet_rows.append(_row(6, [
            _cell("A6", reason + " Review Applied Filters in Executive Summary.", 17),
        ], height=28))
        unit_merges.append("A6:N6")
    for idx, u in enumerate(unit_rows, start=1):
        r = 5 + idx
        style = 10 if idx % 2 else 11
        age = int(u.get("age_days", 0) or 0)
        unit_sheet_rows.append(_row(r, [
            _cell(f"A{r}", u.get("branch_name"), style),
            _cell(f"B{r}", u.get("area"), style),
            _cell(f"C{r}", u.get("standard_description"), style),
            _cell(f"D{r}", u.get("brand"), style),
            _cell(f"E{r}", u.get("engine_no"), style),
            _cell(f"F{r}", u.get("chassis"), style),
            _cell(f"G{r}", u.get("created_on"), style),
            _cell(f"H{r}", u.get("incoming_date"), style),
            _cell(f"I{r}", age, _age_style(age), numeric=True),
            _cell(f"J{r}", u.get("qty", 0), 6, numeric=True),
            _cell(f"K{r}", _money(u.get("amount", 0)), 7, numeric=True),
            _cell(f"L{r}", _money(u.get("inventory_value", 0)), 7, numeric=True),
            _cell(f"M{r}", u.get("location"), style),
            _cell(f"N{r}", u.get("barcode"), style),
        ], height=20))
    unit_end = max(6, 5 + len(unit_rows))
    sheet4 = _sheet_xml(
        unit_sheet_rows,
        [(1, 24), (2, 14), (3, 42), (4, 14), (5, 22), (6, 23), (7, 13), (8, 13), (9, 10), (10, 8), (11, 14), (12, 16), (13, 22), (14, 19)],
        merges=unit_merges,
        autofilter=f"A5:N{unit_end}",
        freeze_rows=5,
        freeze_cols=2,
        top_left="C6",
        orientation="landscape",
    )

    # ------------------------------------------------------------------
    # Sheet 5: Data Dictionary
    # ------------------------------------------------------------------
    dictionary = [
        ("Branch", "Branch currently holding or assigned to the motorcycle unit."),
        ("Area", "Operational / geographic cluster mapped to the branch."),
        ("Standard Description", "Standardized motorcycle model description used for grouping."),
        ("Brand", "Motorcycle manufacturer or supplier brand."),
        ("Engine No.", "Unique engine identifier used for unit-level traceability."),
        ("Chassis", "Unique frame / chassis identifier."),
        ("Created On", "Branch aging basis date."),
        ("Incoming Date", "Company aging basis date."),
        ("Age Days", "Exact elapsed days from the selected aging basis date to the report as-of date."),
        ("Qty", "Quantity represented by the row."),
        ("Unit Cost", "Imported unit amount / cost basis."),
        ("Inventory Value", "Qty multiplied by Unit Cost."),
        ("Location", "Original location reference from the consolidated import file."),
        ("91+ Units", "Units whose Age Days are greater than 90."),
        ("Aged Value", "Inventory value tied to units over 90 days."),
        ("Exposure", "High ≥40% aged; Watch ≥20%; Controlled <20%, using 91+ share of model inventory."),
    ]
    dict_rows: list[str] = [
        _row(1, [_cell("A1", "AGING EXPORT GUIDE", 1)], height=28),
        _row(2, [_cell("A2", "Definitions, thresholds and workbook usage notes", 2)], height=20),
        _row(4, [_cell("A4", "Field / Measure", 9), _cell("B4", "Definition", 9)], height=22),
    ]
    for idx, (field, meaning) in enumerate(dictionary, start=5):
        style = 10 if idx % 2 else 11
        dict_rows.append(_row(idx, [_cell(f"A{idx}", field, 4), _cell(f"B{idx}", meaning, style)], height=28))
    notes_start = 6 + len(dictionary)
    dict_rows.append(_row(notes_start, [_cell(f"A{notes_start}", "AGE COLOR GUIDE", 3)], height=21))
    dict_rows.append(_row(notes_start + 1, [_cell(f"A{notes_start+1}", "≤90 Days", 13), _cell(f"B{notes_start+1}", "Healthy / normal inventory rotation", 10)]))
    dict_rows.append(_row(notes_start + 2, [_cell(f"A{notes_start+2}", "91–180 Days", 14), _cell(f"B{notes_start+2}", "Action required", 10)]))
    dict_rows.append(_row(notes_start + 3, [_cell(f"A{notes_start+3}", "181–365 Days", 15), _cell(f"B{notes_start+3}", "High aging exposure", 10)]))
    dict_rows.append(_row(notes_start + 4, [_cell(f"A{notes_start+4}", "366+ Days", 16), _cell(f"B{notes_start+4}", "Critical management review", 10)]))
    dict_merges = ["A1:B1", "A2:B2", f"A{notes_start}:B{notes_start}"]
    sheet5 = _sheet_xml(
        dict_rows,
        [(1, 25), (2, 92)],
        merges=dict_merges,
        freeze_rows=4,
        top_left="A5",
        orientation="portrait",
    )

    styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="3"><numFmt numFmtId="164" formatCode="#,##0"/><numFmt numFmtId="165" formatCode="₱#,##0.00"/><numFmt numFmtId="166" formatCode="0.0%"/></numFmts>
<fonts count="10">
<font><sz val="10"/><color rgb="FF20364F"/><name val="Aptos"/></font>
<font><b/><sz val="17"/><color rgb="FFFFFFFF"/><name val="Aptos Display"/></font>
<font><b/><sz val="10"/><color rgb="FF93C5FD"/><name val="Aptos"/></font>
<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Aptos"/></font>
<font><b/><sz val="9"/><color rgb="FF315D88"/><name val="Aptos"/></font>
<font><sz val="9.5"/><color rgb="FF20364F"/><name val="Aptos"/></font>
<font><b/><sz val="9.5"/><color rgb="FF0F172A"/><name val="Aptos"/></font>
<font><b/><sz val="9"/><color rgb="FF9F1239"/><name val="Aptos"/></font>
<font><b/><sz val="9"/><color rgb="FF92400E"/><name val="Aptos"/></font>
<font><b/><sz val="9"/><color rgb="FF047857"/><name val="Aptos"/></font>
</fonts>
<fills count="14">
<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF0B1F3A"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF173A5E"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFEAF4FF"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF7FAFD"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFECFDF5"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFFBEB"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFF7ED"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFF1F2"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFEE2E2"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFEF3C7"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFD1FAE5"/></patternFill></fill>
</fills>
<borders count="2"><border/><border><left style="thin"><color rgb="FFD9E5F1"/></left><right style="thin"><color rgb="FFD9E5F1"/></right><top style="thin"><color rgb="FFD9E5F1"/></top><bottom style="thin"><color rgb="FFD9E5F1"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="23">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="3" fillId="3" borderId="0" xfId="0"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="4" fillId="4" borderId="1" xfId="0"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="5" fillId="5" borderId="1" xfId="0"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="164" fontId="6" fillId="5" borderId="1" xfId="0"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="165" fontId="6" fillId="5" borderId="1" xfId="0"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="166" fontId="6" fillId="5" borderId="1" xfId="0"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="5" fillId="5" borderId="1" xfId="0"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="5" fillId="6" borderId="1" xfId="0"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="164" fontId="6" fillId="6" borderId="1" xfId="0"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="164" fontId="9" fillId="7" borderId="1" xfId="0"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="164" fontId="8" fillId="8" borderId="1" xfId="0"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="164" fontId="8" fillId="9" borderId="1" xfId="0"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="164" fontId="7" fillId="10" borderId="1" xfId="0"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="0" fontId="5" fillId="4" borderId="1" xfId="0"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="7" fillId="11" borderId="1" xfId="0"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="8" fillId="12" borderId="1" xfId="0"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="9" fillId="13" borderId="1" xfId="0"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="5" fillId="5" borderId="1" xfId="0"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''

    sheets = [
        ("Executive Summary", sheet1),
        ("Area Intelligence", sheet2),
        ("Model Intelligence", sheet3),
        ("Unit Detail", sheet4),
        ("Data Dictionary", sheet5),
    ]
    content_types = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">', '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>', '<Default Extension="xml" ContentType="application/xml"/>', '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>', '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>', '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>', '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>']
    for i in range(1, len(sheets) + 1):
        content_types.append(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    content_types.append('</Types>')

    root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''
    workbook_xml = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>']
    workbook_rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    for i, (name, _) in enumerate(sheets, start=1):
        workbook_xml.append(f'<sheet name="{escape(name)}" sheetId="{i}" r:id="rId{i}"/>')
        workbook_rels.append(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>')
    workbook_xml.append('</sheets></workbook>')
    workbook_rels.append(f'<Relationship Id="rId{len(sheets)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
    workbook_rels.append('</Relationships>')

    core = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>Motorcycle Aging Intelligence</dc:title><dc:creator>SCM IDP Control Tower</dc:creator><dc:subject>Aging management export</dc:subject><dcterms:created xsi:type="dcterms:W3CDTF">{generated_utc}</dcterms:created></cp:coreProperties>'''
    app = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>SCM IDP Dashboard</Application><AppVersion>2.47.5</AppVersion></Properties>'''

    out = BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", "".join(content_types))
        archive.writestr("_rels/.rels", root_rels)
        archive.writestr("xl/workbook.xml", "".join(workbook_xml))
        archive.writestr("xl/_rels/workbook.xml.rels", "".join(workbook_rels))
        archive.writestr("xl/styles.xml", styles_xml)
        archive.writestr("docProps/core.xml", core)
        archive.writestr("docProps/app.xml", app)
        for i, (_, xml) in enumerate(sheets, start=1):
            archive.writestr(f"xl/worksheets/sheet{i}.xml", xml)
    return out.getvalue()
