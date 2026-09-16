from __future__ import annotations

from io import BytesIO
from typing import Any, Dict, List
import math
import tempfile
from pathlib import Path
from decimal import Decimal, ROUND_HALF_UP
from datetime import datetime

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
COVER_TEMPLATE_PATH = Path(__file__).resolve().parents[1] / "static" / "ppt_cover_template_clean.png"
KPI_TEMPLATE_PATH = Path(__file__).resolve().parents[1] / "static" / "ppt_kpi_template.png"
NAVY = RGBColor(2, 20, 38)
NAVY_PANEL = RGBColor(3, 29, 49)
BLUE_LINE = RGBColor(18, 130, 193)
CYAN = RGBColor(68, 200, 255)
AQUA = RGBColor(32, 240, 205)


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

def _add_full_bleed_picture(slide, path: Path):
    if path.exists():
        try:
            slide.shapes.add_picture(str(path), 0, 0, width=Inches(13.333), height=Inches(7.5))
            return True
        except Exception:
            return False
    return False


def _source_date(data: Dict[str, Any]) -> str:
    dates = []
    for kpi in (data or {}).get("kpis", {}).values():
        labels = (kpi.get("ytd") or {}).get("labels", []) or []
        for raw in labels[-2:]:
            try:
                dates.append(datetime.strptime(str(raw), "%b %d, %Y"))
            except Exception:
                pass
    dt = max(dates) if dates else datetime.now()
    return dt.strftime("%B %d, %Y").replace(" 0", " ")


def _kpi_heading(label: str):
    u = (label or "").upper()
    if "BEFORE PO" in u:
        return "STOCK-OUT", "• BEFORE PO BALANCE", "VISIBILITY  •  CONTROL  •  SUPPLY RISK  •  A STRONGER TOMORROW"
    if "AFTER PO" in u:
        return "STOCK-OUT", "• AFTER PO BALANCE", "FULFILMENT  •  CONTROL  •  SUPPLY RECOVERY  •  A STRONGER TOMORROW"
    if "PER BRANCH" in u:
        return "STOCK-OUT", "• PER BRANCH", "BRANCH VISIBILITY  •  PRIORITY  •  SERVICE LEVEL  •  GROWTH"
    if "CLASS A" in u and "DOI" in u:
        return "CLASS A DOI", "• AVAILABILITY", "PRIORITY MODELS  •  COVERAGE  •  REPLENISHMENT  •  GROWTH"
    if "CLASS A" in u and "STOCK" in u:
        return "CLASS A STOCK-OUT", "• NETWORK PRIORITY", "CLASS A VISIBILITY  •  CONTROL  •  REPLENISHMENT  •  GROWTH"
    if "DOI" in u:
        return "DAYS OF INVENTORY", "• NETWORK COVERAGE", "AVAILABILITY  •  BALANCE  •  INVENTORY HEALTH  •  GROWTH"
    return label.upper(), "", "VISIBILITY  •  CONTROL  •  SUPPLY  •  GROWTH"


def _kpi_actions(label: str):
    u = (label or "").upper()
    if "STOCK" in u:
        return [
            ("01", "REDUCE STOCK-OUTS", "Protect demand before PO arrival"),
            ("02", "STRENGTHEN PLANNING", "Escalate supply and back-order actions"),
            ("03", "DRIVE CONVERSION", "Use Class B & C alternatives when needed"),
            ("04", "ENABLE GROWTH", "Support sales with supply visibility"),
        ]
    return [
        ("01", "PROTECT AVAILABILITY", "Sustain healthy days of inventory"),
        ("02", "REBALANCE STOCK", "Move units to priority branches and models"),
        ("03", "FOCUS CLASS A", "Prioritize high-rank, low-DoI models"),
        ("04", "ENABLE GROWTH", "Support sales with inventory visibility"),
    ]


def _kpi_chart_png(labels: List[str], values: List[float], trend: List[float], percent=False, period="ytd", good="low") -> str:
    fig, ax = plt.subplots(figsize=(10.4, 3.25), dpi=175)
    fig.patch.set_alpha(0)
    ax.set_facecolor((0,0,0,0))
    for spine in ax.spines.values():
        spine.set_color("#9fb8cc")
        spine.set_alpha(0.55)
    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)

    vals = [_whole(v) for v in values]
    tr = [float(v) for v in trend]
    x = list(range(len(vals)))
    ax.plot(x, vals, marker="o", markersize=6.5, linewidth=3.2, color="#fbc52b", label="Actual", zorder=4)
    if vals:
        ax.fill_between(x, vals, color="#fbc52b", alpha=0.16, zorder=1)

    numeric = [v for v in vals + tr if isinstance(v, (int,float))]
    spread = max(numeric)-min(numeric) if numeric else 1
    gap = max(1, spread*0.10)
    required = max([((a+gap)-t) for a,t in zip(vals,tr) if isinstance(a,(int,float)) and isinstance(t,(int,float))] or [0])
    display_trend=[t+max(0,required) for t in tr]
    ax.plot(x, display_trend, linestyle=(0,(6,4)), linewidth=2.5, color="#edf4fb", label="Trend Direction", zorder=3)

    data_gap=max(0.7, spread*0.045)
    for xi,yi in zip(x,vals):
        if isinstance(yi,(int,float)):
            txt=f"{yi}%" if percent else f"{yi}"
            ax.text(xi, yi+data_gap, txt, color="#ffd22e", fontsize=8.8, fontweight="bold", ha="center", va="bottom", zorder=5)

    if labels:
        if period=="ytd":
            shown=[]
            for raw in labels:
                try:
                    dt=datetime.strptime(str(raw), "%b %d, %Y")
                    shown.append(dt.strftime("%b") if dt.day not in (14,15) else dt.strftime("%b %d").replace(" 0"," "))
                except Exception:
                    shown.append(str(raw))
            idxs=list(range(len(labels)))
        else:
            shown=[]
            for raw in labels:
                try:
                    dt=datetime.strptime(str(raw), "%b %d, %Y")
                    shown.append(dt.strftime("%b %d").replace(" 0"," "))
                except Exception:
                    shown.append(str(raw))
            idxs=list(range(len(labels)))
        ax.set_xticks(idxs); ax.set_xticklabels([shown[i] for i in idxs], fontsize=8.5, color="#eef4fa")

    allv=[v for v in vals+display_trend if isinstance(v,(int,float))]
    if allv:
        lo,hi=min(allv),max(allv)
        if percent:
            lower=min(0, math.floor((lo-max(5,spread*.14))/10)*10)
            upper=max(10, math.ceil((hi+max(5,spread*.14))/10)*10)
        else:
            lower=max(0, math.floor(lo-max(3,spread*.12)))
            upper=math.ceil(hi+max(3,spread*.15))
        if upper<=lower: upper=lower+10
        ax.set_ylim(lower,upper)

    ax.grid(True, axis="both", linestyle="--", linewidth=0.65, alpha=0.18, color="#aac4d8")
    ax.tick_params(colors="#e5eef6", labelsize=8.5)
    from matplotlib.ticker import FuncFormatter
    if percent:
        ax.yaxis.set_major_formatter(FuncFormatter(lambda y,_: f"{int(y)}%"))
    else:
        ax.yaxis.set_major_formatter(FuncFormatter(lambda y,_: f"{int(y)}"))
    ax.legend(loc="upper right", bbox_to_anchor=(1.0,1.12), frameon=False, labelcolor="#f2f7fb", fontsize=8.5, ncol=2, handlelength=3)
    fig.subplots_adjust(left=0.07,right=0.99,top=0.88,bottom=0.16)
    path=tempfile.NamedTemporaryFile(suffix=".png",delete=False).name
    fig.savefig(path, transparent=True, bbox_inches="tight", pad_inches=0.03)
    plt.close(fig)
    return path


def _add_kpi_template_slide(prs, blank, name: str, kpi: Dict[str, Any], period_key: str, period_title: str):
    slide=prs.slides.add_slide(blank)
    if not _add_full_bleed_picture(slide, KPI_TEMPLATE_PATH):
        _set_bg(slide, NAVY)

    # Cover the reference sample content while retaining the globe / ambient background.
    for x,y,w,h in [(0.0,0.52,10.25,1.12),(0.17,1.60,13.05,4.18),(0.0,5.84,13.333,1.42)]:
        sh=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        sh.fill.solid(); sh.fill.fore_color.rgb=NAVY; sh.line.fill.background()

    meta=kpi.get("meta",{})
    label=meta.get("label",name)
    unit=meta.get("unit")
    good=meta.get("good","low")
    d=kpi.get(period_key,{}) or {}
    values=list(d.get("values",[]) or [])
    labels=list(d.get("labels",[]) or [])
    trend=list(d.get("trend",[]) or [])
    latest=d.get("latest")
    delta=d.get("delta")
    primary,secondary,strap=_kpi_heading(label)

    # Header
    secondary_x = min(6.05, 0.48 + 0.22 * len(primary))
    _textbox(slide,0.46,0.66,max(2.3,secondary_x-0.55),0.50,primary,25,True,GOLD)
    _textbox(slide,secondary_x,0.66,10.05-secondary_x,0.50,secondary,25,True,WHITE)
    _textbox(slide,0.47,1.23,9.25,0.25,strap,8.8,False,WHITE)

    # Left KPI panel
    panel=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.22), Inches(1.68), Inches(3.55), Inches(4.02))
    panel.fill.solid(); panel.fill.fore_color.rgb=NAVY_PANEL; panel.line.color.rgb=BLUE_LINE; panel.line.width=Pt(1.1)
    _textbox(slide,0.46,1.84,3.05,0.30,label.upper(),10.5,True,WHITE)
    card=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.35), Inches(2.18), Inches(3.28), Inches(1.52))
    card.fill.solid(); card.fill.fore_color.rgb=RGBColor(5,36,58); card.line.color.rgb=GOLD; card.line.width=Pt(1.2)
    circle=slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.55), Inches(2.43), Inches(0.86), Inches(0.86))
    circle.fill.solid(); circle.fill.fore_color.rgb=RGBColor(8,45,67); circle.line.color.rgb=GOLD; circle.line.width=Pt(1.3)
    _textbox(slide,0.67,2.67,0.62,0.24,"KPI",10,True,GOLD,PP_ALIGN.CENTER)
    latest_fmt=("—" if latest is None else (_whole(latest) if unit=="percent" else _doi_whole(latest)))
    suffix="%" if unit=="percent" else " d"
    _textbox(slide,1.58,2.37,1.80,0.62,f"{latest_fmt}{suffix}",35,True,WHITE,PP_ALIGN.CENTER)
    if delta is not None:
        favorable=_delta_favorable(delta,good)
        col=AQUA if favorable is True else (RED if favorable is False else MUTED)
        arrow="▼" if float(delta)<0 else ("▲" if float(delta)>0 else "•")
        meaning="Positive" if favorable is True else ("Negative" if favorable is False else "Neutral")
        delta_fmt=_whole(abs(delta))
        _textbox(slide,1.32,3.27,2.15,0.26,f"{arrow} {delta_fmt} vs previous • {meaning}",8.6,True,col,PP_ALIGN.CENTER)

    # Summary panel
    summ=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.22), Inches(3.87), Inches(3.55), Inches(1.83))
    summ.fill.solid(); summ.fill.fore_color.rgb=NAVY_PANEL; summ.line.color.rgb=BLUE_LINE; summ.line.width=Pt(1.0)
    _textbox(slide,0.46,4.05,2.9,0.27,f"{period_title} SUMMARY",10.5,True,WHITE)
    numeric=[v for v in values if isinstance(v,(int,float))]
    if numeric:
        if unit=="percent":
            avg=_whole(sum(numeric)/len(numeric)); hi=_whole(max(numeric)); lo=_whole(min(numeric))
        else:
            avg=_doi_whole(sum(numeric)/len(numeric)); hi=_doi_whole(max(numeric)); lo=_doi_whole(min(numeric))
    else:
        avg=hi=lo="—"
    for i,(lab,val) in enumerate([("AVERAGE",avg),("HIGH",hi),("LOW",lo)]):
        x=0.40+i*1.06
        sh=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(4.43), Inches(0.94), Inches(0.92))
        sh.fill.solid(); sh.fill.fore_color.rgb=RGBColor(10,49,70); sh.line.color.rgb=BLUE_LINE
        _textbox(slide,x+0.07,4.60,0.80,0.18,lab,7.5,True,GOLD,PP_ALIGN.CENTER)
        _textbox(slide,x+0.03,4.91,0.88,0.30,f"{val}{suffix}",17,True,WHITE,PP_ALIGN.CENTER)

    # Right chart panel
    right=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.89), Inches(1.68), Inches(9.25), Inches(4.02))
    right.fill.solid(); right.fill.fore_color.rgb=NAVY_PANEL; right.line.color.rgb=BLUE_LINE; right.line.width=Pt(1.1)
    _textbox(slide,4.15,1.86,4.3,0.30,f"{period_title} • ACTUAL + TREND",12.5,True,GOLD)
    view="Month view • fits screen" if period_key=="ytd" else "Week view • fits screen"
    _textbox(slide,10.60,1.87,2.22,0.24,view,8,False,WHITE,PP_ALIGN.RIGHT)
    path=_kpi_chart_png(labels,values,trend,percent=(unit=="percent"),period=period_key,good=good)
    slide.shapes.add_picture(path, Inches(4.18), Inches(2.26), width=Inches(8.62), height=Inches(2.85))
    latest_dir="→"
    latest_meaning="Neutral"
    if delta is not None and abs(float(delta))>1e-12:
        latest_dir="↑" if float(delta)>0 else "↓"
        fav=_delta_favorable(delta,good)
        latest_meaning="Positive" if fav is True else ("Negative" if fav is False else "Neutral")
    _textbox(slide,4.17,5.24,8.60,0.22,f"Latest movement: {latest_dir} {latest_meaning}  •  Performance rule: {'higher DoI is better' if good=='high' else 'lower stock-out is better' if good=='low' else 'manage within healthy coverage'}",7.7,False,MUTED)

    # Bottom management action strip
    actions=_kpi_actions(label)
    for i,(num,head,body) in enumerate(actions):
        x=0.40+i*3.18
        if i>0:
            ln=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x-0.13), Inches(6.05), Inches(0.012), Inches(0.70))
            ln.fill.solid(); ln.fill.fore_color.rgb=RGBColor(88,139,173); ln.line.fill.background()
        circ=slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(6.02), Inches(0.58), Inches(0.58))
        circ.fill.solid(); circ.fill.fore_color.rgb=NAVY; circ.line.color.rgb=GOLD; circ.line.width=Pt(1.4)
        _textbox(slide,x+0.08,6.19,0.42,0.18,num,8.5,True,GOLD,PP_ALIGN.CENTER)
        _textbox(slide,x+0.72,6.03,2.15,0.22,head,8.6,True,GOLD)
        _textbox(slide,x+0.72,6.34,2.14,0.40,body,8.2,False,WHITE)
    _textbox(slide,0.40,7.13,7.2,0.17,"RIGHT PRODUCTS.   RIGHT CUSTOMERS.   A STRONGER TOMORROW.",6.7,False,WHITE)
    _textbox(slide,9.45,7.13,3.45,0.17,"PEOPLE  •  PROCESS  •  PERFORMANCE",6.7,False,WHITE,PP_ALIGN.RIGHT)


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

    v2.37: Export uses the approved Operations Excellence cover and full-slide KPI template while retaining Area, Branch and ABC drilldowns.
    """
    prs = Presentation(); prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # Slide 1 Cover — approved Operations Excellence design template.
    slide = prs.slides.add_slide(blank)
    if not _add_full_bleed_picture(slide, COVER_TEMPLATE_PATH):
        _set_bg(slide, NAVY)
        _textbox(slide, 0.68, 1.55, 11.9, 0.58, "INVENTORY AND", 30, True, WHITE)
        _textbox(slide, 0.68, 2.18, 11.9, 0.58, "DISTRIBUTION PLANNING", 30, True, GOLD)
        _textbox(slide, 0.68, 2.81, 11.9, 0.58, "DEPARTMENT - SCM", 30, True, WHITE)
    # The supplied reference artwork is cleaned once and the source date is rendered dynamically.
    _textbox(slide, 0.62, 4.55, 4.30, 0.33, _source_date(data), 15.5, False, WHITE)

    _add_kpi_summary_slide(prs, blank, data)
    _add_reorder_model_position_slides(prs, blank, data.get("reorder_card", {}), reorder_brand)

    # Approved KPI template: one full executive slide per KPI for both YTD and Weekly views.
    for period_key, period_title in [("ytd", "YTD"), ("weekly", "WEEKLY")]:
        for name, kpi in data.get("kpis", {}).items():
            _add_kpi_template_slide(prs, blank, name, kpi, period_key, period_title)

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
