from __future__ import annotations

from io import BytesIO
from typing import Any, Dict, List
import math
import tempfile
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


def _whole(value):
    if value is None or value == "—":
        return value
    try:
        return int(Decimal(str(value)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    except Exception:
        return value


def _set_bg(slide, color=DARK):
    fill = slide.background.fill
    fill.solid(); fill.fore_color.rgb = color


def _textbox(slide, x, y, w, h, text, size=18, bold=False, color=WHITE, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame; tf.clear()
    p = tf.paragraphs[0]; p.alignment = align
    run = p.add_run(); run.text = str(text); run.font.size = Pt(size); run.font.bold = bold; run.font.color.rgb = color
    return box


def _title(slide, title, subtitle=None):
    _textbox(slide, 0.65, 0.35, 11.9, 0.48, title, 24, True, WHITE)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.65), Inches(0.92), Inches(1.4), Inches(0.05))
    bar.fill.solid(); bar.fill.fore_color.rgb = GOLD; bar.line.fill.background()
    if subtitle:
        _textbox(slide, 0.65, 1.02, 11.7, 0.35, subtitle, 10, False, MUTED)


def _card(slide, x, y, w, h, label, value, suffix="", delta=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = DARK2; sh.line.color.rgb = RGBColor(51,65,85)
    _textbox(slide, x+0.18, y+0.12, w-0.36, 0.25, label, 9, True, MUTED)
    _textbox(slide, x+0.18, y+0.45, w-0.36, 0.42, f"{value}{suffix}", 22, True, WHITE)
    if delta is not None:
        col = GREEN if delta <= 0 else RED
        arrow = "▼" if delta < 0 else ("▲" if delta > 0 else "•")
        _textbox(slide, x+0.18, y+h-0.38, w-0.36, 0.22, f"{arrow} {_whole(abs(delta))} vs previous", 8, True, col)


def _chart_png(labels: List[str], values: List[float], trend: List[float], title: str, percent=False, period="ytd") -> str:
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
    ax.plot(x, display_trend, linestyle=(0, (6, 4)), linewidth=2.0, color="#cbd5e1",
            label=f"Trend Direction {trend_arrow}", zorder=4, solid_capstyle="butt")

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


def build_presentation(data: Dict[str, Any], area_data: Dict[str, Any], branch_data: Dict[str, Any]) -> bytes:
    prs = Presentation(); prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # Slide 1
    slide = prs.slides.add_slide(blank); _set_bg(slide)
    _textbox(slide, 0.72, 0.65, 11.9, 0.55, "SCM INVENTORY & DISTRIBUTION PLANNING", 28, True, WHITE)
    _textbox(slide, 0.72, 1.25, 11.9, 0.35, "Executive Control Tower • Management Presentation", 14, False, GOLD)
    _textbox(slide, 0.72, 1.85, 11.4, 0.5, "Current performance, stock-out exposure, inventory days and branch request readiness.", 14, False, MUTED)
    latest_cards = []
    for name, kpi in data["kpis"].items():
        latest = kpi["ytd"].get("latest")
        delta = kpi["ytd"].get("delta")
        unit = kpi["meta"]["unit"]
        latest_cards.append((kpi["meta"]["label"], _whole(latest) if latest is not None else "—", "%" if unit == "percent" else " d", delta))
    for i, c in enumerate(latest_cards):
        x = 0.72 + (i % 3) * 4.12; y = 3.0 + (i // 3) * 1.65
        _card(slide, x, y, 3.72, 1.32, *c)
    _textbox(slide, 0.72, 6.88, 11.8, 0.25, "Source: KPI_YTD_Input, KPI_WEEKLY_Input and Raw Distribution data", 8, False, MUTED)

    # Slides 2-3 YTD / Weekly
    for period_key, period_title in [("ytd", "YTD KPI TREND"), ("weekly", "WEEKLY KPI TREND")]:
        for page in range(2):
            slide = prs.slides.add_slide(blank); _set_bg(slide)
            subset = list(data["kpis"].items())[page*3:(page+1)*3]
            _title(slide, f"{period_title} · {page+1}/2", "Actual and Trend share one chart; Trend is a straight directional regression line kept above Actual; YTD axis displays month only.")
            y = 1.45
            for name, kpi in subset:
                d = kpi[period_key]
                path = _chart_png(d["labels"], d["values"], d["trend"], kpi["meta"]["label"], kpi["meta"]["unit"] == "percent", period_key)
                slide.shapes.add_picture(path, Inches(0.72), Inches(y), width=Inches(11.9), height=Inches(1.6))
                y += 1.78

    # Area slide
    slide = prs.slides.add_slide(blank); _set_bg(slide)
    _title(slide, "AREA PERFORMANCE", f"Selected view: {area_data['selected']} • Stock-out rate = Stockout count ÷ stock-status count within the same class")
    s = area_data["summary"]
    cards = [("Class A Stock-Out", _whole(s["class_rates"]["A"]), "%"), ("Class B Stock-Out", _whole(s["class_rates"]["B"]), "%"), ("Class C Stock-Out", _whole(s["class_rates"]["C"]), "%"), ("Overall Average", _whole(s["overall_avg"]), "%")]
    for i,(lab,val,suf) in enumerate(cards): _card(slide, 0.72+i*3.08, 1.55, 2.72, 1.2, lab, val, suf)
    ranking = area_data["ranking"][:10]
    _textbox(slide, 0.72, 3.12, 5.7, 0.3, "AREA STOCK-OUT RANKING · HIGHEST TO LOWEST", 11, True, GOLD)
    for i, r in enumerate(ranking):
        y=3.55+i*0.3
        _textbox(slide, 0.78, y, 2.1, 0.22, f"{i+1}. {r['area']}", 9, True, WHITE)
        _textbox(slide, 3.0, y, 1.0, 0.22, f"{_whole(r['overall_avg'])}%", 9, True, RED if i<3 else WHITE)
        _textbox(slide, 4.1, y, 1.0, 0.22, f"A {_whole(r['class_a'])}%", 8, False, MUTED)
    _textbox(slide, 7.0, 3.12, 5.2, 0.3, "BRANCH CLASS A STOCK-OUT RATE", 11, True, GOLD)
    for i, r in enumerate(area_data.get("branch_class_a", [])[:10]):
        y=3.55+i*0.3
        _textbox(slide, 7.05, y, 2.75, 0.22, f"{i+1}. {r['branch']}", 8, True, WHITE)
        _textbox(slide, 10.0, y, 1.0, 0.22, f"{_whole(r['class_a'])}%", 9, True, RED if i<5 else WHITE)
        _textbox(slide, 11.1, y, 1.2, 0.22, r.get('area',''), 7, False, MUTED)

    # Branch slide
    slide = prs.slides.add_slide(blank); _set_bg(slide)
    _title(slide, "BRANCH PERFORMANCE", f"{branch_data['branch']} • {branch_data['area']}")
    s = branch_data["summary"]
    cards = [("Class A Stock-Out", _whole(s["class_rates"]["A"]), "%"), ("Class B Stock-Out", _whole(s["class_rates"]["B"]), "%"), ("Class C Stock-Out", _whole(s["class_rates"]["C"]), "%"), ("Overall Average", _whole(s["overall_avg"]), "%")]
    for i,(lab,val,suf) in enumerate(cards): _card(slide, 0.72+i*3.08, 1.45, 2.72, 1.18, lab, val, suf)
    # Unified model matrix replaces the old three-column card presentation.
    _textbox(slide, 0.72, 2.88, 11.8, 0.3, "MODEL INTELLIGENCE · TOP PRIORITY MODELS BY CLASS", 11, True, GOLD)
    headers = [("CLASS",0.72,0.62),("RANK",1.38,0.62),("MODEL",2.08,3.25),("STATUS",5.40,1.65),("INV",7.12,0.8),("SUGG.",7.98,0.92),("DoI",8.98,0.72)]
    y_head=3.24
    for text,x,w in headers:
        _textbox(slide, x, y_head, w, 0.22, text, 7, True, MUTED)
    row_y=3.55
    for cls in ("A","B","C"):
        for m in branch_data["classes"][cls][:4]:
            col = GOLD if cls == "A" else (RGBColor(56,189,248) if cls == "B" else MUTED)
            _textbox(slide, 0.72, row_y, 0.55, 0.22, f"{cls}", 8, True, col)
            _textbox(slide, 1.38, row_y, 0.6, 0.22, f"#{_whole(m['rank'])}", 8, True, WHITE)
            _textbox(slide, 2.08, row_y, 3.2, 0.22, m['model'], 8, True, WHITE)
            status=m['stock_status'] or "—"
            _textbox(slide, 5.40, row_y, 1.6, 0.22, status, 7, True, RED if ('stock' in status.lower() or 'critical' in status.lower()) else MUTED)
            _textbox(slide, 7.12, row_y, 0.78, 0.22, _whole(m['inventory']), 8, True, WHITE)
            _textbox(slide, 7.98, row_y, 0.9, 0.22, _whole(m['suggested_transfer']), 8, True, WHITE)
            _textbox(slide, 8.98, row_y, 0.7, 0.22, _whole(m['doi']), 8, True, WHITE)
            row_y += 0.25
        row_y += 0.08

    # Closing slide
    slide = prs.slides.add_slide(blank); _set_bg(slide)
    _textbox(slide, 0.72, 1.0, 11.9, 0.55, "MANAGEMENT ACTION FRAME", 28, True, WHITE)
    _textbox(slide, 0.72, 1.65, 11.6, 0.4, "Focus replenishment decisions on Class A exposure first, then use branch DoI and Suggested Transfer to balance inventory.", 15, False, MUTED)
    actions = [
        "1  Prioritize the highest Class A stock-out areas and branches.",
        "2  Review models with Stockout / Re-order status against Suggested Transfer.",
        "3  Use the Request Simulator to test projected DoI before approving transfers.",
        "4  Re-import the latest workbook and regenerate this deck for every operating review.",
    ]
    for i,a in enumerate(actions):
        sh=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.72), Inches(2.55+i*0.86), Inches(11.8), Inches(0.65))
        sh.fill.solid(); sh.fill.fore_color.rgb=DARK2; sh.line.color.rgb=RGBColor(51,65,85)
        _textbox(slide, 1.0, 2.73+i*0.86, 11.2, 0.28, a, 13, i==0, GOLD if i==0 else WHITE)

    out = BytesIO(); prs.save(out); return out.getvalue()
