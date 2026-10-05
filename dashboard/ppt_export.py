from __future__ import annotations

from io import BytesIO
from typing import Any, Dict, List, Sequence
from datetime import datetime
import math

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# v2.48.8: intentionally simple presentation export.
# No decorative background images, image-heavy templates, or visual effects.
WHITE = RGBColor(255, 255, 255)
NAVY = RGBColor(24, 45, 72)
BLUE = RGBColor(37, 99, 235)
TEXT = RGBColor(31, 41, 55)
MUTED = RGBColor(100, 116, 139)
BORDER = RGBColor(203, 213, 225)
LIGHT = RGBColor(241, 245, 249)
RED = RGBColor(185, 28, 28)
GREEN = RGBColor(21, 128, 61)


def _num(value: Any, decimals: int = 0) -> str:
    try:
        n = float(value)
    except Exception:
        return "—" if value in (None, "") else str(value)
    if not math.isfinite(n):
        return "—"
    if decimals <= 0:
        return f"{int(round(n)):,}"
    return f"{n:,.{decimals}f}"


def _pct(value: Any) -> str:
    return f"{_num(value, 0)}%"


def _date_text(data: Dict[str, Any]) -> str:
    raw = (data or {}).get("export_generated_at") or (data or {}).get("generated_at")
    try:
        dt = datetime.fromisoformat(str(raw)) if raw else datetime.now()
    except Exception:
        dt = datetime.now()
    return dt.strftime("%B %d, %Y").replace(" 0", " ")


def _set_white(slide) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = WHITE


def _text(slide, x, y, w, h, value, size=14, bold=False, color=TEXT, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.text = str(value if value is not None else "")
    p.alignment = align
    run = p.runs[0]
    run.font.name = "Aptos"
    run.font.size = Pt(size)
    run.font.bold = bool(bold)
    run.font.color.rgb = color
    return box


def _line(slide, x, y, w, color=BORDER, height=0.015):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(height))
    sh.fill.solid(); sh.fill.fore_color.rgb = color
    sh.line.fill.background()
    return sh


def _header(slide, title: str, subtitle: str = "") -> None:
    _text(slide, 0.55, 0.28, 12.1, 0.42, title, 21, True, NAVY)
    if subtitle:
        _text(slide, 0.55, 0.74, 12.1, 0.25, subtitle, 9.5, False, MUTED)
    _line(slide, 0.55, 1.08, 12.1, BLUE, 0.025)


def _footer(slide, page_no: int | None = None) -> None:
    _line(slide, 0.55, 7.08, 12.1, BORDER, 0.01)
    _text(slide, 0.55, 7.12, 5.4, 0.18, "SCM • Inventory & Distribution Planning", 7.2, False, MUTED)
    if page_no is not None:
        _text(slide, 11.8, 7.12, 0.85, 0.18, str(page_no), 7.2, False, MUTED, PP_ALIGN.RIGHT)


def _add_cover(prs, blank, data: Dict[str, Any]):
    slide = prs.slides.add_slide(blank); _set_white(slide)
    _text(slide, 0.75, 1.35, 11.8, 0.44, "SCM IDP REPORT", 15, True, BLUE)
    _text(slide, 0.75, 2.05, 11.8, 0.75, "Inventory and Distribution Planning", 32, True, NAVY)
    _text(slide, 0.75, 2.86, 11.8, 0.42, "Management Performance Report", 18, False, TEXT)
    _line(slide, 0.75, 3.55, 3.0, BLUE, 0.04)
    _text(slide, 0.75, 3.88, 6.5, 0.32, _date_text(data), 13, False, MUTED)
    _text(slide, 0.75, 6.55, 11.8, 0.30, "Simple export • Data, tables and management information only", 9.5, False, MUTED)
    return slide


def _table(slide, x, y, w, h, headers: Sequence[str], rows: Sequence[Sequence[Any]], widths: Sequence[float] | None = None, font_size: float = 8.2):
    rows = list(rows)
    table = slide.shapes.add_table(len(rows) + 1, len(headers), Inches(x), Inches(y), Inches(w), Inches(h)).table
    if widths:
        total = float(sum(widths)) or 1.0
        for i, part in enumerate(widths):
            table.columns[i].width = Inches(w * float(part) / total)
    for c, head in enumerate(headers):
        cell = table.cell(0, c)
        cell.text = str(head)
        cell.fill.solid(); cell.fill.fore_color.rgb = NAVY
        cell.margin_left = cell.margin_right = Inches(0.05)
        cell.margin_top = cell.margin_bottom = Inches(0.03)
        for p in cell.text_frame.paragraphs:
            p.alignment = PP_ALIGN.CENTER
            for r in p.runs:
                r.font.name = "Aptos"; r.font.size = Pt(font_size); r.font.bold = True; r.font.color.rgb = WHITE
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, value in enumerate(row):
            cell = table.cell(r_idx, c_idx)
            cell.text = str(value if value not in (None, "") else "—")
            cell.fill.solid(); cell.fill.fore_color.rgb = WHITE if r_idx % 2 else LIGHT
            cell.margin_left = cell.margin_right = Inches(0.05)
            cell.margin_top = cell.margin_bottom = Inches(0.025)
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.LEFT if c_idx in (0, 1, 2, 3) else PP_ALIGN.CENTER
                for rr in p.runs:
                    rr.font.name = "Aptos"; rr.font.size = Pt(font_size); rr.font.color.rgb = TEXT
    return table


def _kpi_rows(data: Dict[str, Any]):
    out = []
    for _, kpi in (data.get("kpis") or {}).items():
        meta = kpi.get("meta") or {}
        label = meta.get("label") or "KPI"
        unit = meta.get("unit") or ""
        ytd = kpi.get("ytd") or {}
        wk = kpi.get("weekly") or {}
        suffix = "%" if unit == "percent" else (" days" if unit == "days" else "")
        def val(v):
            return f"{_num(v)}{suffix}" if v is not None else "—"
        out.append([label, val(ytd.get("latest")), _num(ytd.get("delta")), val(wk.get("latest")), _num(wk.get("delta"))])
    return out


def _add_kpi_summary(prs, blank, data: Dict[str, Any]):
    slide = prs.slides.add_slide(blank); _set_white(slide)
    _header(slide, "Executive KPI Summary", "Latest YTD and Weekly KPI results")
    rows = _kpi_rows(data)
    _table(slide, 0.55, 1.35, 12.1, 4.95,
           ["KPI", "YTD Latest", "YTD Δ", "Weekly Latest", "Weekly Δ"], rows,
           [5.2, 1.6, 1.1, 1.7, 1.1], 9)
    _text(slide, 0.55, 6.45, 12.0, 0.35, "Use this page as the management checkpoint before reviewing Area, Branch and Model detail.", 9.5, False, MUTED)
    return slide


def _period_rows(period: Dict[str, Any], unit: str):
    labels = list(period.get("labels") or [])
    values = list(period.get("values") or [])
    trend = list(period.get("trend") or [])
    suffix = "%" if unit == "percent" else (" d" if unit == "days" else "")
    rows = []
    for i, label in enumerate(labels[-9:]):
        source_i = len(labels) - min(9, len(labels)) + i
        v = values[source_i] if source_i < len(values) else None
        t = trend[source_i] if source_i < len(trend) else None
        rows.append([label, f"{_num(v)}{suffix}" if v is not None else "—", f"{_num(t,1)}{suffix}" if t is not None else "—"])
    return rows


def _add_kpi_detail(prs, blank, name: str, kpi: Dict[str, Any]):
    slide = prs.slides.add_slide(blank); _set_white(slide)
    label = (kpi.get("meta") or {}).get("label") or name
    unit = (kpi.get("meta") or {}).get("unit") or ""
    _header(slide, label, "YTD and Weekly data points")
    ytd = kpi.get("ytd") or {}; weekly = kpi.get("weekly") or {}
    suffix = "%" if unit == "percent" else (" days" if unit == "days" else "")
    _text(slide, 0.65, 1.35, 5.7, 0.32, f"YTD Latest: {_num(ytd.get('latest'))}{suffix}   |   Change: {_num(ytd.get('delta'))}", 11, True, NAVY)
    _table(slide, 0.65, 1.78, 5.75, 4.75, ["Period", "Actual", "Trend"], _period_rows(ytd, unit), [2.5, 1.2, 1.2], 8.5)
    _text(slide, 6.85, 1.35, 5.7, 0.32, f"Weekly Latest: {_num(weekly.get('latest'))}{suffix}   |   Change: {_num(weekly.get('delta'))}", 11, True, NAVY)
    _table(slide, 6.85, 1.78, 5.75, 4.75, ["Period", "Actual", "Trend"], _period_rows(weekly, unit), [2.5, 1.2, 1.2], 8.5)
    return slide


def _add_area_pages(prs, blank, area_data: Dict[str, Any]):
    ranking = list(area_data.get("ranking") or [])
    if not ranking:
        return
    per_page = 15
    for start in range(0, len(ranking), per_page):
        chunk = ranking[start:start + per_page]
        slide = prs.slides.add_slide(blank); _set_white(slide)
        suffix = f" ({start + 1}-{start + len(chunk)} of {len(ranking)})" if len(ranking) > per_page else ""
        _header(slide, "Stock-Out Rate by Area" + suffix, "Class A / B / C and overall average")
        rows = [[r.get("area", "—"), _pct(r.get("class_a")), _pct(r.get("class_b")), _pct(r.get("class_c")), _pct(r.get("overall_avg"))] for r in chunk]
        _table(slide, 0.65, 1.40, 12.0, 5.55, ["Area", "Class A", "Class B", "Class C", "Overall Avg"], rows, [3.7, 1.3, 1.3, 1.3, 1.5], 9)


def _branch_summary_rows(all_branches: List[Dict[str, Any]]):
    rows = []
    for b in all_branches or []:
        s = b.get("summary") or {}; cr = s.get("class_rates") or {}
        rows.append([b.get("branch", "—"), b.get("area", "—"), _pct(cr.get("A")), _pct(cr.get("B")), _pct(cr.get("C")), _pct(s.get("overall_avg"))])
    return rows


def _add_branch_summary_pages(prs, blank, all_branches: List[Dict[str, Any]]):
    rows = _branch_summary_rows(all_branches)
    per_page = 18
    for start in range(0, len(rows), per_page):
        chunk = rows[start:start + per_page]
        slide = prs.slides.add_slide(blank); _set_white(slide)
        suffix = f" ({start + 1}-{start + len(chunk)} of {len(rows)})" if len(rows) > per_page else ""
        _header(slide, "Stock-Out Rate by Branch" + suffix, "Branch Class A / B / C and overall average")
        _table(slide, 0.45, 1.35, 12.45, 5.7, ["Branch", "Area", "Class A", "Class B", "Class C", "Overall"], chunk, [3.4, 2.0, 1.15, 1.15, 1.15, 1.2], 8.3)


def _flatten_branch_models(branch: Dict[str, Any]):
    rows = []
    for cls in ("A", "B", "C"):
        for r in (branch.get("classes") or {}).get(cls, []) or []:
            rows.append([
                f"Class {cls}", r.get("rank", "—"), r.get("brand", "—"), r.get("model", "—"),
                r.get("stock_status", "—"), _num(r.get("inventory")), _num(r.get("suggested_transfer")), f"{_num(r.get('doi'))} d"
            ])
    return rows


def _add_branch_model_pages(prs, blank, all_branches: List[Dict[str, Any]]):
    per_page = 17
    for branch in all_branches or []:
        rows = _flatten_branch_models(branch)
        if not rows:
            continue
        summary = branch.get("summary") or {}; cr = summary.get("class_rates") or {}
        for start in range(0, len(rows), per_page):
            chunk = rows[start:start + per_page]
            slide = prs.slides.add_slide(blank); _set_white(slide)
            page = start // per_page + 1; pages = math.ceil(len(rows) / per_page)
            title = f"{branch.get('branch','Branch')} — Model Detail"
            if pages > 1: title += f" ({page}/{pages})"
            subtitle = f"{branch.get('area','')} • Stock-Out A {_pct(cr.get('A'))} • B {_pct(cr.get('B'))} • C {_pct(cr.get('C'))} • Overall {_pct(summary.get('overall_avg'))}"
            _header(slide, title, subtitle)
            _table(slide, 0.30, 1.35, 12.75, 5.75,
                   ["Class", "Rank", "Brand", "Model", "Stock Status", "Inventory", "Suggested", "DoI"],
                   chunk, [1.0, 0.65, 1.45, 3.25, 1.45, 0.9, 0.9, 0.75], 7.7)


def _add_reorder_pages(prs, blank, reorder: Dict[str, Any], brand_filter: str):
    rows = list(reorder.get("rows") or [])
    if brand_filter and brand_filter != "All Brands":
        rows = [r for r in rows if str(r.get("brand") or "") == brand_filter]
    rows.sort(key=lambda r: (str(r.get("class") or ""), int(r.get("rank") or 999999), str(r.get("model") or "")))
    if not rows:
        return
    per_page = 18
    for start in range(0, len(rows), per_page):
        chunk = rows[start:start + per_page]
        slide = prs.slides.add_slide(blank); _set_white(slide)
        scope = f" • Brand: {brand_filter}" if brand_filter and brand_filter != "All Brands" else ""
        suffix = f" ({start + 1}-{start + len(chunk)} of {len(rows)})" if len(rows) > per_page else ""
        _header(slide, "Reorder Model Position" + suffix, f"{reorder.get('title','Reorder review')}{scope}")
        table_rows = [[f"Class {r.get('class','—')}", r.get("rank", "—"), r.get("brand", "—"), r.get("model", "—"), f"{_num(r.get('doi'))} d", r.get("stock_status", "—")] for r in chunk]
        _table(slide, 0.55, 1.38, 12.1, 5.65, ["Class", "Rank", "Brand", "Model", "DoI", "Stock Status"], table_rows, [1.1, 0.7, 1.6, 4.0, 1.0, 1.6], 8.1)


def _add_aging_pages(prs, blank, aging: Dict[str, Any] | None):
    if not aging:
        return
    slide = prs.slides.add_slide(blank); _set_white(slide)
    _header(slide, "Motorcycle Aging Summary", "Simple aging exposure and priority view")
    metrics = [
        ["Total Units", _num(aging.get("total_qty") or aging.get("units"))],
        ["91+ Day Units", _num(aging.get("aged_90") or aging.get("aged_90_qty"))],
        ["91+ Exposure", _pct(aging.get("aged_90_pct"))],
        ["Healthy Inventory", _pct(aging.get("healthy_pct"))],
        ["Branches", _num(aging.get("branch_count"))],
        ["Areas", _num(aging.get("area_count"))],
    ]
    _table(slide, 0.70, 1.55, 5.6, 3.4, ["Metric", "Value"], metrics, [3.6, 1.4], 10)
    priorities = []
    for label, obj, key in [
        ("Top Risk Area", aging.get("top_area") or {}, "name"),
        ("Top Risk Branch", aging.get("top_branch") or {}, "name"),
        ("Top Aged Model", aging.get("top_model") or {}, "standard_description"),
    ]:
        priorities.append([label, obj.get(key, "—"), _num(obj.get("aged_90")), _pct(obj.get("aged_90_pct")), f"{_num(obj.get('oldest'))} d"])
    _table(slide, 6.65, 1.55, 5.95, 3.4, ["Priority", "Name", "91+", "91+ %", "Oldest"], priorities, [1.3, 2.6, 0.7, 0.8, 0.8], 8.5)
    msg = aging.get("risk_message") or "Review 91+ day exposure and prioritize transfer, sell-through or disposition actions."
    _text(slide, 0.70, 5.45, 11.9, 0.75, msg, 11, False, TEXT)


def build_presentation(
    data: Dict[str, Any],
    area_data: Dict[str, Any],
    branch_data: Dict[str, Any],
    all_branches: List[Dict[str, Any]] | None = None,
    reorder_brand: str = "All Brands",
) -> bytes:
    """Build a simple management presentation without decorative images/design.

    The export intentionally prioritizes readability and data over styling:
    white background, plain headings, simple tables, no external pictures.
    """
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    _add_cover(prs, blank, data)
    _add_kpi_summary(prs, blank, data)
    for name, kpi in (data.get("kpis") or {}).items():
        _add_kpi_detail(prs, blank, name, kpi)
    _add_area_pages(prs, blank, area_data or {})
    branches = all_branches or ([branch_data] if branch_data else [])
    _add_branch_summary_pages(prs, blank, branches)
    _add_branch_model_pages(prs, blank, branches)
    _add_reorder_pages(prs, blank, data.get("reorder_card") or {}, reorder_brand)
    _add_aging_pages(prs, blank, data.get("aging_summary"))

    # Apply a small footer after content is complete.
    for idx, slide in enumerate(prs.slides, start=1):
        if idx > 1:
            _footer(slide, idx)

    out = BytesIO()
    prs.save(out)
    return out.getvalue()
