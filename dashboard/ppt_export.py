from __future__ import annotations

from io import BytesIO
from typing import Any, Dict, List
import math
import tempfile
from pathlib import Path
from decimal import Decimal, ROUND_HALF_UP

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

DARK = RGBColor(15, 23, 42)
DARK2 = RGBColor(30, 41, 59)
GOLD = RGBColor(251, 191, 36)
WHITE = RGBColor(248, 250, 252)
MUTED = RGBColor(148, 163, 184)
RED = RGBColor(248, 113, 113)
GREEN = RGBColor(74, 222, 128)
TEAL = RGBColor(20, 184, 166)
LOGO_PATH = Path(__file__).resolve().parents[1] / "static" / "brilliant4_logo.png"


def _whole(value):
    if value is None or value == "—":
        return value
    try:
        return int(Decimal(str(value)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    except Exception:
        return value


def _doi_whole(value):
    if value is None or value == "—":
        return value
    try:
        return int(math.ceil(max(0.0, float(value))))
    except Exception:
        return value


def _delta_favorable(delta, good):
    if delta is None:
        return None
    try:
        d = float(delta)
    except Exception:
        return None
    if abs(d) < 1e-12 or good == "balanced":
        return None
    return d > 0 if good == "high" else d < 0


def _set_bg(slide, color=DARK):
    fill = slide.background.fill
    fill.solid(); fill.fore_color.rgb = color


def _textbox(slide, x, y, w, h, text, size=18, bold=False, color=WHITE, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame; tf.clear()
    p = tf.paragraphs[0]; p.alignment = align
    run = p.add_run(); run.text = str(text); run.font.size = Pt(size); run.font.bold = bold; run.font.color.rgb = color
    return box


def _add_logo(slide, x=9.70, y=0.32, w=2.85):
    if LOGO_PATH.exists():
        try:
            slide.shapes.add_picture(str(LOGO_PATH), Inches(x), Inches(y), width=Inches(w))
        except Exception:
            pass


def _title(slide, title, subtitle=None):
    _add_logo(slide)
    _textbox(slide, 0.65, 0.36, 8.7, 0.46, title, 22, True, WHITE)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.65), Inches(0.92), Inches(1.55), Inches(0.055))
    bar.fill.solid(); bar.fill.fore_color.rgb = TEAL; bar.line.fill.background()
    bar2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2.22), Inches(0.92), Inches(0.70), Inches(0.055))
    bar2.fill.solid(); bar2.fill.fore_color.rgb = GOLD; bar2.line.fill.background()
    if subtitle:
        _textbox(slide, 0.65, 1.03, 10.2, 0.35, subtitle, 10, False, MUTED)


def _card(slide, x, y, w, h, label, value, suffix="", delta=None, good="low"):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = DARK2; sh.line.color.rgb = RGBColor(51,65,85)
    accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x+0.10), Inches(y+0.10), Inches(0.055), Inches(max(0.25, h-0.20)))
    accent.fill.solid(); accent.fill.fore_color.rgb = GOLD; accent.line.fill.background()
    _textbox(slide, x+0.22, y+0.12, w-0.42, 0.25, label, 8.5, True, MUTED)
    _textbox(slide, x+0.22, y+0.45, w-0.42, 0.42, f"{value}{suffix}", 22, True, WHITE)
    if delta is not None:
        favorable = _delta_favorable(delta, good)
        col = GREEN if favorable is True else (RED if favorable is False else MUTED)
        arrow = "▼" if delta < 0 else ("▲" if delta > 0 else "•")
        meaning = "" if favorable is None else (" • Positive" if favorable else " • Negative")
        _textbox(slide, x+0.18, y+h-0.38, w-0.36, 0.22, f"{arrow} {_whole(abs(delta))} vs previous{meaning}", 8, True, col)


def _chart_png(labels: List[str], values: List[float], trend: List[float], title: str, percent=False, period="ytd", good="low") -> str:
    # Actual and Trend share one chart. The regression trend remains a straight
    # directional line and is shifted upward by one constant display offset only.
    fig, ax = plt.subplots(figsize=(9.5, 3.15), dpi=160)
    fig.patch.set_facecolor("#0f172a")
    ax.set_facecolor("#0f172a")
    for spine in ax.spines.values():
        spine.set_visible(False)

    rounded_values = [_whole(v) for v in values]
    true_trend = [float(v) for v in trend]
    x = list(range(len(rounded_values)))

    ax.plot(x, rounded_values, marker="o", markersize=3.8, linewidth=2.6,
            label="Actual", color="#fbbf24", zorder=3)
    ax.fill_between(x, rounded_values, alpha=0.08, color="#fbbf24", zorder=1)

    # Show data labels on every Actual point in the exported presentation.
    # Labels use the approved half-up whole-number rounding.
    if rounded_values:
        numeric_actual = [v for v in rounded_values if isinstance(v, (int, float))]
        label_gap = max(0.8, ((max(numeric_actual) - min(numeric_actual)) if numeric_actual else 1) * 0.045)
        for xi, yi in zip(x, rounded_values):
            if isinstance(yi, (int, float)):
                ax.text(xi, yi + label_gap, f"{yi}", color="#f8fafc", fontsize=7,
                        fontweight="bold", ha="center", va="bottom", zorder=5)

    # Presentation-only separation: apply ONE constant vertical offset to the
    # entire regression line. This preserves a straight direction line instead
    # of making the visual trend follow the Actual curve.
    numeric = [v for v in rounded_values + true_trend if isinstance(v, (int, float))]
    spread = max(numeric) - min(numeric) if numeric else 1
    gap = max(1, spread * 0.10)
    required_offset = max(
        [((a + gap) - t) for a, t in zip(rounded_values, true_trend)
         if isinstance(a, (int, float)) and isinstance(t, (int, float))] or [0]
    )
    visual_offset = max(0, required_offset)
    display_trend = [(t + visual_offset) if isinstance(t, (int, float)) else t for t in true_trend]
    first_t = next((v for v in true_trend if isinstance(v, (int, float))), None)
    last_t = next((v for v in reversed(true_trend) if isinstance(v, (int, float))), None)
    if first_t is None or last_t is None or abs(last_t - first_t) < 1e-9:
        trend_arrow = "→"
    else:
        trend_arrow = "↑" if last_t > first_t else "↓"
    latest_note = ""
    if len(rounded_values) > 1 and all(isinstance(v, (int, float)) for v in rounded_values[-2:]):
        latest_delta = rounded_values[-1] - rounded_values[-2]
        latest_arrow = "→" if abs(latest_delta) < 1e-9 else ("↑" if latest_delta > 0 else "↓")
        favorable = _delta_favorable(latest_delta, good)
        latest_meaning = "" if favorable is None else (" Positive" if favorable else " Negative")
        latest_note = f" • Latest {latest_arrow}{latest_meaning}"
    ax.plot(x, display_trend, linestyle=(0, (6, 4)), linewidth=2.0, color="#cbd5e1",
            label=f"Trend Direction {trend_arrow}{latest_note}", zorder=4, solid_capstyle="butt")

    # Keep data labels and elevated trend direction line inside the plot area.
    all_plot_values = [v for v in (rounded_values + display_trend) if isinstance(v, (int, float))]
    if all_plot_values:
        ymin, ymax = min(all_plot_values), max(all_plot_values)
        pad = max(2, (ymax - ymin) * 0.16)
        ax.set_ylim(ymin - pad * 0.45, ymax + pad)

    ax.grid(True, axis="y", alpha=0.16, color="#cbd5e1")
    ax.tick_params(colors="#cbd5e1", labelsize=8)

    if labels:
        if period == "ytd":
            shown = [str(v)[:3] for v in labels]
            idxs = list(range(len(labels)))
        else:
            shown = [str(v).replace(", 2026", "") for v in labels]
            step = max(1, len(labels)//6)
            idxs = list(range(0, len(labels), step))
            if idxs[-1] != len(labels)-1:
                idxs.append(len(labels)-1)
        ax.set_xticks(idxs)
        ax.set_xticklabels([shown[i] for i in idxs], rotation=0)

    ax.set_title(title, color="#f8fafc", fontsize=11, fontweight="bold", loc="left", pad=5)
    ax.legend(loc="upper right", frameon=False, labelcolor="#cbd5e1", fontsize=7, ncol=2)
    if percent:
        ax.set_ylabel("%", color="#94a3b8", fontsize=8)
    else:
        ax.set_ylabel("Days", color="#94a3b8", fontsize=8)

    path = tempfile.NamedTemporaryFile(suffix=".png", delete=False).name
    fig.subplots_adjust(left=0.07, right=0.99, top=0.90, bottom=0.13)
    fig.savefig(path, transparent=False, facecolor="#0f172a", bbox_inches="tight")
    plt.close(fig)
    return path



def _table_header(slide, x, y, columns, widths, h=0.30):
    x0 = x
    for title, w in zip(columns, widths):
        sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x0), Inches(y), Inches(w), Inches(h))
        sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor(10, 35, 54); sh.line.color.rgb = RGBColor(43, 76, 96)
        _textbox(slide, x0+0.04, y+0.07, w-0.08, h-0.08, title, 6.8, True, WHITE, PP_ALIGN.CENTER)
        x0 += w


def _table_row(slide, x, y, values, widths, h=0.28, danger=False):
    x0 = x
    for i, (val, w) in enumerate(zip(values, widths)):
        sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x0), Inches(y), Inches(w), Inches(h))
        sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor(18, 29, 48) if not danger else RGBColor(58, 28, 36)
        sh.line.color.rgb = RGBColor(51,65,85)
        col = RED if danger and i in (3,4,5,6) else WHITE
        align = PP_ALIGN.CENTER if i not in (1,2) else PP_ALIGN.LEFT
        _textbox(slide, x0+0.04, y+0.06, w-0.08, h-0.08, val, 6.5, i in (0,3,4,5,6), col, align)
        x0 += w


def _safe_pct(value):
    v = _whole(value)
    return "—" if v is None or v == "—" else f"{v}%"


def _branch_rows(all_branches: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    rows = []
    for b in all_branches or []:
        s = b.get("summary", {})
        rates = s.get("class_rates", {})
        rows.append({
            "branch": b.get("branch", ""),
            "area": b.get("area", ""),
            "class_a": rates.get("A", 0),
            "class_b": rates.get("B", 0),
            "class_c": rates.get("C", 0),
            "overall": s.get("overall_avg", 0),
        })
    rows.sort(key=lambda r: (-float(r.get("overall") or 0), -float(r.get("class_a") or 0), r.get("area", ""), r.get("branch", "")))
    return rows


def _add_kpi_summary_slide(prs, blank, data: Dict[str, Any]):
    slide = prs.slides.add_slide(blank); _set_bg(slide)
    _title(slide, "EXECUTIVE KPI SUMMARY", "Latest YTD and Weekly results across the six core Supply Chain KPIs.")
    cols = ["KPI", "YTD", "YTD Δ", "Weekly", "Weekly Δ", "Direction"]
    widths = [4.45, 1.25, 1.25, 1.25, 1.25, 1.50]
    _table_header(slide, 0.72, 1.35, cols, widths, 0.33)
    y = 1.76
    for name, kpi in data.get("kpis", {}).items():
        meta = kpi.get("meta", {})
        unit = "%" if meta.get("unit") == "percent" else " d"
        ytd = kpi.get("ytd", {})
        weekly = kpi.get("weekly", {})
        ytd_latest = ytd.get("latest")
        weekly_latest = weekly.get("latest")
        ytd_delta = ytd.get("delta")
        weekly_delta = weekly.get("delta")
        fmt_value = _whole if meta.get("unit") == "percent" else _doi_whole
        arrow = "→" if ytd_delta is None or abs(float(ytd_delta)) < 1e-12 else ("↑" if ytd_delta > 0 else "↓")
        favorable = _delta_favorable(ytd_delta, meta.get("good", "low"))
        direction = f"{arrow} " + ("Positive" if favorable is True else "Negative" if favorable is False else "Neutral")
        vals = [
            meta.get("label", name),
            "—" if ytd_latest is None else f"{fmt_value(ytd_latest)}{unit}",
            "—" if ytd_delta is None else f"{_whole(ytd_delta)}{unit}",
            "—" if weekly_latest is None else f"{fmt_value(weekly_latest)}{unit}",
            "—" if weekly_delta is None else f"{_whole(weekly_delta)}{unit}",
            direction,
        ]
        _table_row(slide, 0.72, y, vals, widths, 0.36, danger=(meta.get("unit") == "percent" and (ytd_latest or 0) > 20))
        y += 0.40
    _textbox(slide, 0.72, 6.62, 11.9, 0.26, "Use this page as the quick management checkpoint before reviewing YTD, Weekly, Area and all-Branch drilldowns.", 9.5, False, MUTED)


def _add_area_slide(prs, blank, area_data: Dict[str, Any]):
    slide = prs.slides.add_slide(blank); _set_bg(slide)
    _title(slide, "AREA PERFORMANCE · ALL AREAS", "Identify Areas and Branches with the highest stock-out exposure and Class A service risk.")
    s = area_data.get("summary", {})
    cards = [("Class A Stock-Out", _whole(s.get("class_rates", {}).get("A", 0)), "%"), ("Class B Stock-Out", _whole(s.get("class_rates", {}).get("B", 0)), "%"), ("Class C Stock-Out", _whole(s.get("class_rates", {}).get("C", 0)), "%"), ("Overall Average", _whole(s.get("overall_avg", 0)), "%")]
    for i,(lab,val,suf) in enumerate(cards):
        _card(slide, 0.72+i*3.08, 1.42, 2.72, 1.08, lab, val, suf)
    ranking = area_data.get("ranking", [])
    _textbox(slide, 0.72, 2.86, 11.8, 0.25, "AREA STOCK-OUT RANKING", 11, True, GOLD)
    cols=["#","Area","A","B","C","Overall"]
    widths=[0.45,2.3,0.9,0.9,0.9,1.1]
    _table_header(slide, 0.72, 3.20, cols, widths, 0.28)
    y=3.52
    for i,r in enumerate(ranking[:10], start=1):
        vals=[i, r.get('area',''), _safe_pct(r.get('class_a')), _safe_pct(r.get('class_b')), _safe_pct(r.get('class_c')), _safe_pct(r.get('overall_avg'))]
        _table_row(slide, 0.72, y, vals, widths, 0.27, danger=i<=3)
        y += 0.29
    _textbox(slide, 7.55, 2.86, 4.75, 0.25, "TOP BRANCH CLASS A STOCK-OUT", 11, True, GOLD)
    cols2=["#","Branch","Area","A %"]
    widths2=[0.45,2.45,1.05,0.80]
    _table_header(slide, 7.55, 3.20, cols2, widths2, 0.28)
    y=3.52
    for i,r in enumerate(area_data.get("branch_class_a", [])[:10], start=1):
        _table_row(slide, 7.55, y, [i, r.get('branch',''), r.get('area',''), _safe_pct(r.get('class_a'))], widths2, 0.27, danger=i<=5)
        y += 0.29


def _add_branch_class_a_ranking_slide(prs, blank, area_data: Dict[str, Any]):
    rows = area_data.get("branch_class_a", [])
    per_slide = 22
    for page, start in enumerate(range(0, len(rows), per_slide), start=1):
        slide = prs.slides.add_slide(blank); _set_bg(slide)
        total_pages = max(1, math.ceil(len(rows)/per_slide))
        _title(slide, f"BRANCH CLASS A STOCK-OUT RANKING · {page}/{total_pages}", "All branches are captured. Highest Class A exposure appears first.")
        cols=["#","Branch","Area","Class A %","SO / Status Count"]
        widths=[0.45,4.15,1.20,1.20,1.65]
        _table_header(slide, 0.72, 1.35, cols, widths, 0.30)
        y=1.70
        for i,r in enumerate(rows[start:start+per_slide], start=start+1):
            counts = r.get("class_a_counts", {}) or {}
            count_txt = f"{_whole(counts.get('stockout', 0))} / {_whole(counts.get('status_count', 0))}"
            vals=[i, r.get('branch',''), r.get('area',''), _safe_pct(r.get('class_a')), count_txt]
            _table_row(slide, 0.72, y, vals, widths, 0.23, danger=i<=10)
            y += 0.245


def _add_branch_all_slides(prs, blank, all_branches: List[Dict[str, Any]]):
    rows = _branch_rows(all_branches)
    per_slide = 17
    for page, start in enumerate(range(0, len(rows), per_slide), start=1):
        slide = prs.slides.add_slide(blank); _set_bg(slide)
        total_pages = max(1, math.ceil(len(rows)/per_slide))
        _title(slide, f"BRANCH PERFORMANCE · ALL BRANCHES · {page}/{total_pages}", "All branches are ranked by Overall Average Stock-Out, then Class A risk.")
        cols=["#","Branch","Area","Class A","Class B","Class C","Overall"]
        widths=[0.45,3.9,1.2,1.0,1.0,1.0,1.1]
        _table_header(slide, 0.72, 1.35, cols, widths, 0.30)
        y=1.72
        for i,r in enumerate(rows[start:start+per_slide], start=start+1):
            vals=[i, r.get('branch',''), r.get('area',''), _safe_pct(r.get('class_a')), _safe_pct(r.get('class_b')), _safe_pct(r.get('class_c')), _safe_pct(r.get('overall'))]
            _table_row(slide, 0.72, y, vals, widths, 0.27, danger=i<=10)
            y += 0.30


def _add_priority_model_slide(prs, blank, all_branches: List[Dict[str, Any]]):
    # Capture all-branch model intelligence in one prioritized management list.
    model_rows=[]
    for b in all_branches or []:
        for cls in ("A","B","C"):
            for m in (b.get("classes", {}) or {}).get(cls, [])[:3]:
                status = m.get("stock_status", "") or ""
                risk = 2 if "stock" in status.lower() else (1 if "re" in status.lower() or "critical" in status.lower() else 0)
                model_rows.append({
                    "branch": b.get("branch", ""), "area": b.get("area", ""), "class": cls,
                    "rank": m.get("rank", 999999), "model": m.get("model", ""), "brand": m.get("brand", ""),
                    "status": status, "inventory": m.get("inventory", 0), "doi": m.get("doi", 0), "risk": risk,
                })
    model_rows.sort(key=lambda r: (r["class"] != "A", -r["risk"], r["rank"], r["area"], r["branch"]))
    slide = prs.slides.add_slide(blank); _set_bg(slide)
    _title(slide, "ALL-BRANCH MODEL PRIORITY SNAPSHOT", "Top model-level exceptions from branch Model Intelligence, prioritizing Class A and stock-risk conditions.")
    cols=["#","Area","Branch","Cls","Model","Status","Inv","DoI"]
    widths=[0.42,0.92,2.45,0.55,3.15,1.42,0.65,0.65]
    _table_header(slide, 0.72, 1.35, cols, widths, 0.30)
    y=1.70
    for i,r in enumerate(model_rows[:20], start=1):
        vals=[i, r['area'], r['branch'], r['class'], r['model'], r['status'], _whole(r['inventory']), _doi_whole(r['doi'])]
        _table_row(slide, 0.72, y, vals, widths, 0.245, danger=(r['risk']>0 and r['class']=='A'))
        y += 0.265


def _model_detail_rows(all_branches: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    rows = []
    for b in all_branches or []:
        for cls in ("A", "B", "C"):
            for m in (b.get("classes", {}) or {}).get(cls, []) or []:
                rows.append({
                    "area": b.get("area", ""),
                    "branch": b.get("branch", ""),
                    "class": cls,
                    "brand": m.get("brand", ""),
                    "model": m.get("model", ""),
                    "status": m.get("stock_status", ""),
                    "inventory": m.get("inventory", 0),
                    "suggested_transfer": m.get("suggested_transfer", 0),
                    "doi": m.get("doi", 0),
                    "rank": m.get("rank", 999999),
                })
    rows.sort(key=lambda r: (r.get("area", ""), r.get("branch", ""), r.get("class", ""), r.get("rank", 999999), r.get("model", "")))
    return rows


def _add_branch_model_detail_slides(prs, blank, all_branches: List[Dict[str, Any]]):
    """Separated per-Branch Class A/B/C model visibility for management review.

    v2.19: Class A, Class B and Class C are no longer mixed in one table.
    Each Branch/Class gets its own slide page(s) so management can review the
    exact class segment independently: Stock Status, Inventory, Suggested
    Transfer and DoI.
    """
    branches = sorted(all_branches or [], key=lambda b: (b.get("area", ""), b.get("branch", "")))
    if not branches:
        slide = prs.slides.add_slide(blank); _set_bg(slide)
        _title(slide, "BRANCH MODEL INTELLIGENCE", "No Class A/B/C model rows available from the current Raw Distribution data.")
        return

    per_slide = 16
    for b in branches:
        branch = b.get("branch", "") or "—"
        area = b.get("area", "") or "—"
        classes = b.get("classes", {}) or {}
        for cls in ("A", "B", "C"):
            rows = list(classes.get(cls, []) or [])
            # Still create a small blank-class slide when there is no data, so the
            # branch A/B/C structure is consistent across the presentation.
            chunks = [rows[i:i+per_slide] for i in range(0, len(rows), per_slide)] or [[]]
            total_pages = len(chunks)
            for page, chunk in enumerate(chunks, start=1):
                slide = prs.slides.add_slide(blank); _set_bg(slide)
                page_suffix = f" · {page}/{total_pages}" if total_pages > 1 else ""
                _title(slide, f"{branch} · CLASS {cls} MODELS{page_suffix}", f"{area} • Stock Status, Inventory, Suggested Transfer and DoI are separated by Class for branch-level action.")
                # Class scorecards
                risk_count = sum(1 for m in rows if any(x in str(m.get('stock_status','')).lower() for x in ['stock', 're', 'critical']))
                total_inv = sum(float(m.get('inventory') or 0) for m in rows)
                total_sug = sum(float(m.get('suggested_transfer') or 0) for m in rows)
                avg_doi = (sum(float(m.get('doi') or 0) for m in rows) / len(rows)) if rows else 0
                cards = [("Models", len(rows), ""), ("Risk Items", risk_count, ""), ("Inventory", _whole(total_inv), ""), ("Suggested Transfer", _whole(total_sug), ""), ("Avg DoI", _doi_whole(avg_doi), " d")]
                for i, (lab, val, suf) in enumerate(cards):
                    _card(slide, 0.52 + i*2.48, 1.25, 2.18, 0.84, lab, val, suf)
                cols = ["#", "Brand", "Model", "Stock Status", "Inv", "Sug. Trf", "DoI"]
                widths = [0.45, 1.25, 4.85, 1.80, 0.78, 0.95, 0.78]
                _table_header(slide, 0.72, 2.35, cols, widths, 0.32)
                y = 2.72
                if not chunk:
                    _textbox(slide, 0.72, y+0.25, 11.5, 0.35, f"No Class {cls} model rows available for this branch.", 11, False, MUTED)
                    continue
                for i, m in enumerate(chunk, start=1 + (page-1)*per_slide):
                    status = str(m.get("stock_status", "") or "")
                    danger = (cls == "A" and any(x in status.lower() for x in ["stock", "re", "critical"]))
                    vals = [
                        i,
                        m.get("brand", ""),
                        m.get("model", ""),
                        status,
                        _whole(m.get("inventory", 0)),
                        _whole(m.get("suggested_transfer", 0)),
                        _doi_whole(m.get("doi", 0)),
                    ]
                    _table_row(slide, 0.72, y, vals, widths, 0.27, danger=danger)
                    y += 0.30



def _add_reorder_model_position_slides(prs, blank, reorder_card: Dict[str, Any], brand_filter: str = "All Brands"):
    rows = list((reorder_card or {}).get("rows", []) or [])
    if brand_filter and brand_filter != "All Brands":
        rows = [r for r in rows if (r.get("brand") or "Unspecified") == brand_filter]
    class_order = {"A": 0, "B": 1, "C": 2}
    rows.sort(key=lambda r: (class_order.get(r.get("class"), 9), int(r.get("rank", 999999)), r.get("brand", ""), r.get("model", "")))
    scope = brand_filter if brand_filter and brand_filter != "All Brands" else "All Brands"
    counts = {c: sum(1 for r in rows if r.get("class") == c) for c in ("A", "B", "C")}
    avgs = {}
    for c in ("A", "B", "C"):
        vals = [_doi_whole(r.get("doi", 0)) for r in rows if r.get("class") == c]
        vals = [v for v in vals if isinstance(v, (int, float))]
        avgs[c] = _doi_whole(sum(vals)/len(vals)) if vals else 0

    slide = prs.slides.add_slide(blank); _set_bg(slide)
    _title(slide, "CLASS A / B / C MODEL POSITION", f"Imported Reorder priority • Brand filter: {scope} • DoI values follow the approved round-up rule.")
    cards=[("Class A",counts["A"],f"Avg DoI {avgs['A']} d"),("Class B",counts["B"],f"Avg DoI {avgs['B']} d"),("Class C",counts["C"],f"Avg DoI {avgs['C']} d"),("ABC Models",len(rows),scope)]
    for i,(lab,val,sub) in enumerate(cards):
        x=0.72+i*3.08
        sh=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.55), Inches(2.72), Inches(1.15))
        sh.fill.solid(); sh.fill.fore_color.rgb=DARK2; sh.line.color.rgb=RGBColor(51,65,85)
        _textbox(slide,x+0.18,1.73,2.3,0.22,lab,8.5,True,MUTED)
        _textbox(slide,x+0.18,2.02,2.3,0.34,val,20,True,WHITE)
        _textbox(slide,x+0.18,2.42,2.3,0.18,sub,7.5,True,GOLD)
    _textbox(slide,0.72,3.10,11.8,0.30,"Decision rule: prioritize Class A first; within each class, review lower DoI positions and imported Class Rank for replenishment action.",10.5,False,MUTED)
    if not rows:
        _textbox(slide,0.72,3.75,11.8,0.35,"No Reorder models match the selected Brand filter.",15,True,RED)
        return

    per_slide=20
    total_pages=max(1,math.ceil(len(rows)/per_slide))
    for page,start in enumerate(range(0,len(rows),per_slide),start=1):
        slide=prs.slides.add_slide(blank); _set_bg(slide)
        _title(slide,f"ABC MODEL POSITION · {page}/{total_pages}",f"Brand: {scope} • Sorted Class A → B → C, then imported Class Rank • DoI rounded up.")
        cols=["#","Class","Rank","Brand","Model","DoI","Stock Status"]
        widths=[0.45,0.70,0.65,1.45,4.35,0.70,2.55]
        _table_header(slide,0.72,1.35,cols,widths,0.30)
        y=1.70
        for i,r in enumerate(rows[start:start+per_slide],start=start+1):
            danger=(r.get("class")=="A" and ("stock" in str(r.get("stock_status","")).lower() or _doi_whole(r.get("doi",0)) <= 3))
            vals=[i,r.get("class","—"),r.get("rank","—"),r.get("brand","—"),r.get("model","—"),f"{_doi_whole(r.get('doi',0))} d",r.get("stock_status","—")]
            _table_row(slide,0.72,y,vals,widths,0.245,danger=danger)
            y+=0.265


def build_presentation(data: Dict[str, Any], area_data: Dict[str, Any], branch_data: Dict[str, Any], all_branches: List[Dict[str, Any]] | None = None, reorder_brand: str = "All Brands") -> bytes:
    """Build the management deck.

    v2.18: Export captures branded YTD and Weekly KPI trends, full Area Performance rates, Branch rankings, and per-Branch Class A/B/C model details.
    """
    prs = Presentation(); prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # Slide 1 Cover — branded executive opening
    slide = prs.slides.add_slide(blank); _set_bg(slide)
    if LOGO_PATH.exists():
        slide.shapes.add_picture(str(LOGO_PATH), Inches(0.68), Inches(0.45), width=Inches(4.25))
    ribbon = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(6.95), Inches(13.333), Inches(0.10))
    ribbon.fill.solid(); ribbon.fill.fore_color.rgb = TEAL; ribbon.line.fill.background()
    ribbon2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(7.05), Inches(5.1), Inches(0.06))
    ribbon2.fill.solid(); ribbon2.fill.fore_color.rgb = GOLD; ribbon2.line.fill.background()
    _textbox(slide, 0.72, 1.55, 11.9, 0.58, "SCM INVENTORY & DISTRIBUTION PLANNING", 30, True, WHITE)
    _textbox(slide, 0.72, 2.17, 11.9, 0.35, "Executive KPI Review • YTD • Weekly • ABC Model Position • Area • Branch Details", 14, True, GOLD)
    _textbox(slide, 0.72, 2.70, 11.4, 0.52, "A branded management deck capturing KPI movement, imported ABC Model Position, Area stock-out rates, branch risk and model-level action details.", 14, False, MUTED)
    latest_cards = []
    for name, kpi in data.get("kpis", {}).items():
        latest = kpi["ytd"].get("latest")
        delta = kpi["ytd"].get("delta")
        unit = kpi["meta"].get("unit")
        latest_value = (_whole(latest) if unit == "percent" else _doi_whole(latest)) if latest is not None else "—"
        latest_cards.append((kpi["meta"].get("label", name), latest_value, "%" if unit == "percent" else " d", delta, kpi["meta"].get("good", "low")))
    for i, c in enumerate(latest_cards):
        x = 0.72 + (i % 3) * 4.12; y = 3.0 + (i // 3) * 1.65
        _card(slide, x, y, 3.72, 1.32, *c)
    branch_count = len(all_branches or [])
    area_count = len(data.get("areas", [])) - (1 if data.get("areas") and data.get("areas", [])[0] == "Overall" else 0)
    _textbox(slide, 0.72, 6.84, 11.8, 0.25, f"Coverage: {area_count} Areas • {branch_count} Branches • Source: KPI_YTD_Input, KPI_WEEKLY_Input and Raw Distribution data", 8.5, False, MUTED)

    _add_kpi_summary_slide(prs, blank, data)
    _add_reorder_model_position_slides(prs, blank, data.get("reorder_card", {}), reorder_brand)

    # YTD / Weekly trend slides — all six KPIs captured.
    for period_key, period_title in [("ytd", "YTD KPI TRENDS"), ("weekly", "WEEKLY KPI TRENDS")]:
        kpi_items = list(data.get("kpis", {}).items())
        for page in range(math.ceil(len(kpi_items)/3) or 1):
            slide = prs.slides.add_slide(blank); _set_bg(slide)
            subset = kpi_items[page*3:(page+1)*3]
            total_pages = max(1, math.ceil(len(kpi_items)/3))
            _title(slide, f"{period_title} · {page+1}/{total_pages}", "Average, High, Low and Trend Direction highlight momentum, volatility and emerging exceptions.")
            y = 1.45
            for name, kpi in subset:
                d = kpi[period_key]
                path = _chart_png(d.get("labels", []), d.get("values", []), d.get("trend", []), kpi["meta"].get("label", name), kpi["meta"].get("unit") == "percent", period_key, kpi["meta"].get("good", "low"))
                slide.shapes.add_picture(path, Inches(0.72), Inches(y), width=Inches(11.9), height=Inches(1.6))
                y += 1.78

    _add_area_slide(prs, blank, area_data)
    _add_branch_class_a_ranking_slide(prs, blank, area_data)
    _add_branch_all_slides(prs, blank, all_branches or ([branch_data] if branch_data else []))
    _add_branch_model_detail_slides(prs, blank, all_branches or ([branch_data] if branch_data else []))
    _add_priority_model_slide(prs, blank, all_branches or ([branch_data] if branch_data else []))

    # Closing slide
    slide = prs.slides.add_slide(blank); _set_bg(slide)
    _textbox(slide, 0.72, 1.0, 11.9, 0.55, "MANAGEMENT ACTION FRAME", 28, True, WHITE)
    _textbox(slide, 0.72, 1.65, 11.6, 0.4, "Use the deck to identify Area exposure, branch priorities and model-level actions from a single KPI export.", 15, False, MUTED)
    actions = [
        "1  Review YTD and Weekly KPI movement before approving corrective actions.",
        "2  Review Class A/B/C Model Position, starting with Class A Rank and low DoI exceptions.",
        "3  Prioritize high stock-out Areas and Class A branch exposure.",
        "4  Use all-Branch rankings to assign follow-up owners and replenishment priority.",
        "5  Re-import the latest workbook and regenerate this deck for every management review.",
    ]
    for i,a in enumerate(actions):
        sh=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.72), Inches(2.45+i*0.72), Inches(11.8), Inches(0.55))
        sh.fill.solid(); sh.fill.fore_color.rgb=DARK2; sh.line.color.rgb=RGBColor(51,65,85)
        _textbox(slide, 1.0, 2.60+i*0.72, 11.2, 0.24, a, 13, i==0, GOLD if i==0 else WHITE)

    out = BytesIO(); prs.save(out); return out.getvalue()
