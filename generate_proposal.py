"""
Amazon Listing Automation — Official Cost Proposal PDF Generator
"""
import io
import math
from datetime import date

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.ticker as mticker
from matplotlib.patches import FancyBboxPatch
import numpy as np

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm, cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, Image as RLImage, KeepTogether
)
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.pdfgen import canvas as rl_canvas

# ── Constants ─────────────────────────────────────────────────────────────────
USD_TO_INR        = 94.71
GROQ_IN_PER_1M    = 0.11          # $/1M input tokens  — Scout 17B
GROQ_OUT_PER_1M   = 0.34          # $/1M output tokens — Scout 17B
FAL_PER_IMAGE_USD = 0.003         # fal-ai/flux/schnell per image
SYSTEM_CHARGE_INR = 0.20          # fixed processing charge per SKU
GROQ_IN_TOKENS    = 3_500         # avg input tokens per SKU
GROQ_OUT_TOKENS   = 2_600         # avg output tokens per SKU
MONTHLY_SKUS      = 200_000

# ── Colours ───────────────────────────────────────────────────────────────────
NAVY   = "#0D1B2A"
GOLD   = "#C9A84C"
STEEL  = "#2E4057"
SLATE  = "#4A6FA5"
CREAM  = "#FDF8F0"
WHITE  = "#FFFFFF"
LIGHT  = "#F0F4F8"
MID    = "#8FA3BF"
RED    = "#C0392B"
GREEN  = "#27AE60"

# ── Derived costs ─────────────────────────────────────────────────────────────
groq_usd = (GROQ_IN_TOKENS * GROQ_IN_PER_1M / 1_000_000 +
            GROQ_OUT_TOKENS * GROQ_OUT_PER_1M / 1_000_000)
groq_inr  = groq_usd * USD_TO_INR

def fal_inr(n_images: int) -> float:
    return n_images * FAL_PER_IMAGE_USD * USD_TO_INR

def total_inr(n_images: int) -> float:
    return groq_inr + fal_inr(n_images) + SYSTEM_CHARGE_INR

def monthly_inr(n_images: int) -> float:
    return total_inr(n_images) * MONTHLY_SKUS


# ══════════════════════════════════════════════════════════════════════════════
# Chart builder
# ══════════════════════════════════════════════════════════════════════════════
def build_chart() -> bytes:
    fig, axes = plt.subplots(1, 2, figsize=(13.5, 3.9),
                             facecolor=WHITE,
                             gridspec_kw={"width_ratios": [1.15, 1], "wspace": 0.42})

    # ── LEFT: Stacked bar — Monthly cost breakdown ────────────────────────────
    ax1 = axes[0]
    ax1.set_facecolor(WHITE)

    imgs      = [1, 2, 3]
    g_vals    = [groq_inr * MONTHLY_SKUS / 1_000   for _ in imgs]   # in ₹ thousands
    f_vals    = [fal_inr(n) * MONTHLY_SKUS / 1_000 for n in imgs]
    s_vals    = [SYSTEM_CHARGE_INR * MONTHLY_SKUS / 1_000 for _ in imgs]
    totals_k  = [(g + f + s) for g, f, s in zip(g_vals, f_vals, s_vals)]

    bar_w   = 0.42
    x       = np.arange(len(imgs))
    c_groq  = SLATE
    c_fal   = GOLD
    c_sys   = "#A8C5DA"

    b1 = ax1.bar(x, g_vals, bar_w, color=c_groq,  label="Groq (AI Analysis + Copy)", zorder=3, linewidth=0)
    b2 = ax1.bar(x, f_vals, bar_w, bottom=g_vals,  color=c_fal,   label="FAL (Image Generation)", zorder=3, linewidth=0)
    b3 = ax1.bar(x, s_vals, bar_w,
                 bottom=[g + f for g, f in zip(g_vals, f_vals)],
                 color=c_sys,  label="System Processing", zorder=3, linewidth=0)

    # Total labels above each bar
    for xi, total in zip(x, totals_k):
        ax1.text(xi, total + 0.8, f"₹{total:,.0f}K",
                 ha="center", va="bottom", fontsize=9.5, fontweight="bold",
                 color=NAVY)

    ax1.set_xticks(x)
    ax1.set_xticklabels(["1 Image\nper SKU", "2 Images\nper SKU", "3 Images\nper SKU"],
                        fontsize=9, color=NAVY)
    ax1.set_ylabel("Monthly Cost (₹ Thousands)", fontsize=9, color=NAVY, labelpad=8)
    ax1.set_title("Monthly Cost at 2,00,000 SKUs — Image Variation",
                  fontsize=10.5, fontweight="bold", color=NAVY, pad=12)
    ax1.set_ylim(0, max(totals_k) * 1.18)
    ax1.tick_params(axis="y", colors=NAVY, labelsize=8.5)
    ax1.spines["top"].set_visible(False)
    ax1.spines["right"].set_visible(False)
    ax1.spines["left"].set_color("#D0D8E4")
    ax1.spines["bottom"].set_color("#D0D8E4")
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"₹{v:,.0f}K"))
    ax1.grid(axis="y", color="#E8EEF4", linewidth=0.6, zorder=0)
    ax1.legend(fontsize=7.8, frameon=False, loc="upper left",
               labelcolor=NAVY, handlelength=1.2, handletextpad=0.5)

    # ── RIGHT: Per-SKU cost breakdown donut (at 3 images — worst case) ────────
    ax2 = axes[1]
    ax2.set_facecolor(WHITE)

    # Show stacked horizontal bar per image scenario
    scenarios = ["1 Image", "2 Images", "3 Images"]
    g_sku = [groq_inr] * 3
    f_sku = [fal_inr(n) for n in imgs]
    s_sku = [SYSTEM_CHARGE_INR] * 3

    y_pos = np.arange(len(scenarios))
    bh_w  = 0.38

    ax2.barh(y_pos, g_sku, bh_w, color=c_groq, label="Groq",   zorder=3, linewidth=0)
    ax2.barh(y_pos, f_sku, bh_w, left=g_sku,   color=c_fal,    label="FAL",    zorder=3, linewidth=0)
    ax2.barh(y_pos, s_sku, bh_w,
             left=[g + f for g, f in zip(g_sku, f_sku)],
             color=c_sys, label="System", zorder=3, linewidth=0)

    totals_sku = [g + f + s for g, f, s in zip(g_sku, f_sku, s_sku)]
    for yi, tot in zip(y_pos, totals_sku):
        ax2.text(tot + 0.004, yi, f"₹{tot:.3f}",
                 va="center", fontsize=9, fontweight="bold", color=NAVY)

    ax2.set_yticks(y_pos)
    ax2.set_yticklabels(scenarios, fontsize=9, color=NAVY)
    ax2.set_xlabel("Cost per SKU (₹)", fontsize=9, color=NAVY, labelpad=8)
    ax2.set_title("Per-SKU Cost Breakdown", fontsize=10.5, fontweight="bold",
                  color=NAVY, pad=12)
    ax2.tick_params(axis="x", colors=NAVY, labelsize=8.5)
    ax2.spines["top"].set_visible(False)
    ax2.spines["right"].set_visible(False)
    ax2.spines["left"].set_color("#D0D8E4")
    ax2.spines["bottom"].set_color("#D0D8E4")
    ax2.xaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"₹{v:.2f}"))
    ax2.grid(axis="x", color="#E8EEF4", linewidth=0.6, zorder=0)
    ax2.set_xlim(0, max(totals_sku) * 1.22)

    fig.tight_layout(pad=1.5)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=160, bbox_inches="tight",
                facecolor=WHITE)
    plt.close(fig)
    buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════════════════════════
# PDF builder
# ══════════════════════════════════════════════════════════════════════════════
OUTPUT_PATH = "/home/venom/amazon_project/Amazon_Listing_Automation_Proposal.pdf"

def build_pdf():
    doc = SimpleDocTemplate(
        OUTPUT_PATH,
        pagesize=A4,
        leftMargin=16*mm, rightMargin=16*mm,
        topMargin=4*mm,   bottomMargin=5*mm,
    )

    W, H = A4
    content_w = W - 32*mm

    styles = getSampleStyleSheet()

    def sty(name, **kw):
        return ParagraphStyle(name, **kw)

    # ── Style palette ──────────────────────────────────────────────────────────
    s_meta    = sty("meta",    fontName="Helvetica",       fontSize=7.5, textColor=colors.HexColor(MID),
                               leading=10, alignment=TA_RIGHT)
    s_tagline = sty("tag",     fontName="Helvetica-Oblique", fontSize=8.5, textColor=colors.HexColor(GOLD),
                               leading=11, alignment=TA_LEFT, spaceAfter=2)
    s_section = sty("sec",     fontName="Helvetica-Bold",  fontSize=9,   textColor=colors.HexColor(NAVY),
                               leading=12, spaceBefore=5, spaceAfter=3,
                               borderPad=0)
    s_body    = sty("body",    fontName="Helvetica",       fontSize=7.2, textColor=colors.HexColor(STEEL),
                               leading=10, alignment=TA_JUSTIFY)
    s_bullet  = sty("bul",     fontName="Helvetica",       fontSize=7.0, textColor=colors.HexColor(STEEL),
                               leading=9.3, leftIndent=10, firstLineIndent=-7, spaceAfter=0.5)
    s_note    = sty("note",    fontName="Helvetica-Oblique", fontSize=6.8, textColor=colors.HexColor(MID),
                               leading=9, alignment=TA_CENTER)
    s_footer  = sty("footer",  fontName="Helvetica",       fontSize=6.5, textColor=colors.HexColor(MID),
                               leading=9, alignment=TA_CENTER)

    elems = []

    # ══ HEADER BAND ═══════════════════════════════════════════════════════════
    header_data = [[
        Paragraph(
            '<font name="Helvetica-Bold" size="15" color="#0D1B2A">Amazon Listing Automation</font><br/>'
            '<font name="Helvetica" size="8.5" color="#4A6FA5">AI-Powered Catalog Enrichment Platform</font>',
            sty("hl", fontName="Helvetica", fontSize=15, leading=20)
        ),
        Paragraph(
            f'<font name="Helvetica" size="7.5" color="#8FA3BF">Prepared by: Technology Division<br/>'
            f'Date: {date.today().strftime("%d %B %Y")}<br/>'
            f'Ref: ALA-PROP-2025-001<br/>'
            f'Exchange Rate: 1 USD = ₹{USD_TO_INR:.0f}</font>',
            s_meta
        ),
    ]]
    ht = Table(header_data, colWidths=[content_w * 0.62, content_w * 0.38])
    ht.setStyle(TableStyle([
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))

    # Gold rule under header
    gold_rule = Table([[""]], colWidths=[content_w])
    gold_rule.setStyle(TableStyle([
        ("LINEABOVE",     (0, 0), (-1, -1), 2.5, colors.HexColor(GOLD)),
        ("LINEBELOW",     (0, 0), (-1, -1), 0.4, colors.HexColor("#D0D8E4")),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))

    elems += [ht, gold_rule, Spacer(1, 1.5*mm)]

    # ══ TWO-COLUMN BODY ═══════════════════════════════════════════════════════
    col_gap = 5*mm
    col_w   = (content_w - col_gap) / 2

    # ── Left column ───────────────────────────────────────────────────────────
    def bullet(text):
        return Paragraph(f'<font color="{GOLD}">▸</font>  {text}', s_bullet)

    overview_paras = [
        Paragraph("PROJECT OVERVIEW", s_section),
        Paragraph(
            "This platform automates the end-to-end creation of Amazon-compliant product listings "
            "at catalog scale. Given a product image and SKU, the system produces a structured "
            "listing — title, five bullet points, description, HSN code, keywords — and generates "
            "up to three professional marketing images per product, all without human intervention.",
            s_body),
        Spacer(1, 1.5*mm),

        Paragraph("WHY GROQ API", s_section),
        bullet("Only publicly-priced <b>vision-capable</b> model on Groq's infrastructure"),
        bullet("Llama 4 Scout 17B delivers <b>594 tokens/second</b> on LPU silicon — "
               "~10× faster than equivalent GPU-hosted models"),
        bullet("<b>One-shot architecture:</b> single API call returns both visual product analysis "
               "and full listing copy, halving network round-trips"),
        bullet("Free tier covers up to 100 vision requests/day — <b>zero cost for prototyping</b>"),
        bullet("128k context window handles large constrained-fields guides from "
               "Amazon category templates"),
        bullet("Structured JSON output reliability at 17B scale — critical for "
               "deterministic field mapping"),
        Spacer(1, 1.5*mm),

        Paragraph("WHY FAL (fal.ai) API", s_section),
        bullet("FLUX.1-schnell delivers production-quality images at <b>$0.003/image</b> "
               "— the lowest price point for this quality tier"),
        bullet("<b>Serverless, zero-ops:</b> no GPU provisioning, auto-scales "
               "from 1 to 20,000 SKUs without configuration"),
        bullet("4-step inference on schnell model — <b>2–5 seconds per image</b>, "
               "compatible with real-time batch pipelines"),
        bullet("Theme-aware prompts (wellness, gaming, food, baby, etc.) produce "
               "on-brand lifestyle scenes without manual art direction"),
        bullet("PIL fallback always available — <b>zero dependency failure risk</b> "
               "if FAL is unreachable"),
    ]

    # ── Right column ──────────────────────────────────────────────────────────
    def inr(v, decimals=3):
        return f"₹{v:.{decimals}f}"

    r = USD_TO_INR  # shorthand

    cost_header = [
        Paragraph("MODELS &amp; PUBLISHED RATES", s_section),
    ]

    # Model reference table
    model_data = [
        ["API", "Model ID", "Input", "Output / Unit"],
        ["Groq\n(Vision +\nOne-Shot)",
         "meta-llama/\nllama-4-scout-\n17b-16e-instruct",
         "$0.11\n/1M tokens\n(₹10.42/1M)",
         "$0.34\n/1M tokens\n(₹32.20/1M)"],
        ["Groq\n(Text\nFallback)",
         "llama-3.1-\n8b-instant",
         "$0.05\n/1M tokens\n(₹4.74/1M)",
         "$0.08\n/1M tokens\n(₹7.58/1M)"],
        ["FAL\n(Image\nGen)",
         "fal-ai/\nflux/schnell",
         "—",
         "$0.003/image\n(₹0.284/image)\n4-step inference"],
    ]
    model_cw = [col_w * 0.18, col_w * 0.34, col_w * 0.24, col_w * 0.24]
    model_tbl = Table(model_data, colWidths=model_cw, repeatRows=1)
    model_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0), colors.HexColor(NAVY)),
        ("TEXTCOLOR",     (0, 0), (-1, 0), colors.HexColor(GOLD)),
        ("FONTNAME",      (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, -1), 6.5),
        ("FONTNAME",      (0, 1), (-1, -1), "Helvetica"),
        ("TEXTCOLOR",     (0, 1), (-1, -1), colors.HexColor(STEEL)),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING",   (0, 0), (-1, -1), 3),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 3),
        ("ALIGN",         (2, 0), (-1, -1), "CENTER"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND",    (0, 1), (-1, 1), colors.HexColor("#EDF1F7")),
        ("BACKGROUND",    (0, 2), (-1, 2), colors.HexColor("#F7F9FC")),
        ("BACKGROUND",    (0, 3), (-1, 3), colors.HexColor("#EDF1F7")),
        # Gold left border on model ID column
        ("LINEAFTER",     (0, 0), (0, -1), 1.2, colors.HexColor(GOLD)),
        ("GRID",          (0, 0), (-1, -1), 0.3, colors.HexColor("#D0D8E4")),
        ("LINEBELOW",     (0, 0), (-1, 0), 1.2, colors.HexColor(GOLD)),
        # Bold model IDs
        ("FONTNAME",      (1, 1), (1, -1), "Helvetica-Bold"),
        ("TEXTCOLOR",     (1, 1), (1, -1), colors.HexColor(SLATE)),
        # Bold rates
        ("FONTNAME",      (2, 1), (-1, -1), "Helvetica-Bold"),
        ("TEXTCOLOR",     (2, 1), (-1, -1), colors.HexColor(NAVY)),
    ]))

    cost_subheader = [
        Spacer(1, 1.5*mm),
        Paragraph("API COST BREAKDOWN — PER SKU", s_section),
    ]

    # Per-SKU micro-table
    micro_data = [
        ["Component", "Calculation", "Per SKU (₹)"],
        ["Groq — Input",
         f"3,500 tok × $0.11/1M × ₹{r:.2f}",
         inr(GROQ_IN_TOKENS * GROQ_IN_PER_1M / 1_000_000 * r)],
        ["Groq — Output",
         f"2,600 tok × $0.34/1M × ₹{r:.2f}",
         inr(GROQ_OUT_TOKENS * GROQ_OUT_PER_1M / 1_000_000 * r)],
        ["Groq Subtotal", "Vision analysis + listing copy (1 API call)", inr(groq_inr)],
        ["FAL — 1 Image", f"1 × $0.003 × ₹{r:.2f}",  inr(fal_inr(1))],
        ["FAL — 2 Images", f"2 × $0.003 × ₹{r:.2f}", inr(fal_inr(2))],
        ["FAL — 3 Images", f"3 × $0.003 × ₹{r:.2f}", inr(fal_inr(3))],
        ["System Charge", "Fixed processing charge/SKU", inr(SYSTEM_CHARGE_INR)],
    ]
    total_rows = [
        ["TOTAL (1 img)", "", inr(total_inr(1))],
        ["TOTAL (2 imgs)", "", inr(total_inr(2))],
        ["TOTAL (3 imgs)", "", inr(total_inr(3))],
    ]

    all_micro = micro_data + total_rows
    micro_cw  = [col_w * 0.40, col_w * 0.38, col_w * 0.22]
    micro_tbl = Table(all_micro, colWidths=micro_cw, repeatRows=1)

    ts = TableStyle([
        # Header
        ("BACKGROUND",    (0, 0), (-1, 0), colors.HexColor(NAVY)),
        ("TEXTCOLOR",     (0, 0), (-1, 0), colors.HexColor(GOLD)),
        ("FONTNAME",      (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, 0), 7),
        ("TOPPADDING",    (0, 0), (-1, 0), 4),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 4),
        ("ALIGN",         (2, 0), (2, -1), "RIGHT"),
        # Body
        ("FONTNAME",      (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE",      (0, 1), (-1, -1), 7),
        ("TOPPADDING",    (0, 1), (-1, -1), 2.5),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 2.5),
        ("TEXTCOLOR",     (0, 1), (-1, -1), colors.HexColor(STEEL)),
        # Alternating rows
        ("BACKGROUND",    (0, 1), (-1, 1), colors.HexColor("#F7F9FC")),
        ("BACKGROUND",    (0, 3), (-1, 3), colors.HexColor("#EDF1F7")),
        ("BACKGROUND",    (0, 5), (-1, 5), colors.HexColor("#F7F9FC")),
        ("BACKGROUND",    (0, 7), (-1, 7), colors.HexColor("#EDF1F7")),
        # Groq subtotal highlight
        ("BACKGROUND",    (0, 3), (-1, 3), colors.HexColor("#E8F0FE")),
        ("FONTNAME",      (0, 3), (-1, 3), "Helvetica-Bold"),
        ("TEXTCOLOR",     (0, 3), (-1, 3), colors.HexColor(SLATE)),
        # Total rows highlight
        ("BACKGROUND",    (0, 8), (-1, 8), colors.HexColor("#FFF8E8")),
        ("BACKGROUND",    (0, 9), (-1, 9), colors.HexColor("#FDF3D0")),
        ("BACKGROUND",    (0, 10),(-1, 10),colors.HexColor("#F9E8A0")),
        ("FONTNAME",      (0, 8), (-1, 10), "Helvetica-Bold"),
        ("TEXTCOLOR",     (0, 8), (-1, 10), colors.HexColor(NAVY)),
        # Grid
        ("GRID",          (0, 0), (-1, -1), 0.3, colors.HexColor("#D0D8E4")),
        ("LINEBELOW",     (0, 0), (-1, 0), 1.2, colors.HexColor(GOLD)),
        ("LINEBELOW",     (0, 2), (-1, 2), 0.6, colors.HexColor(GOLD)),
        ("LINEABOVE",     (0, 8), (-1, 8), 1.0, colors.HexColor(NAVY)),
    ])
    micro_tbl.setStyle(ts)

    # Indian digit-grouping formatter (e.g. 1118000 -> 11,18,000)
    def inr_lakh(n):
        n = int(round(n))
        s = str(n)
        if len(s) <= 3:
            return f"₹{s}"
        last3, rest = s[-3:], s[:-3]
        parts = []
        while len(rest) > 2:
            parts.insert(0, rest[-2:])
            rest = rest[:-2]
        if rest:
            parts.insert(0, rest)
        return "₹" + ",".join(parts) + "," + last3

    # Monthly summary table
    monthly_data = [
        ["Cost Component", "1 Image", "2 Images", "3 Images"],
        ["Groq  (AI Analysis + Copy)",
         inr_lakh(groq_inr*MONTHLY_SKUS),
         inr_lakh(groq_inr*MONTHLY_SKUS),
         inr_lakh(groq_inr*MONTHLY_SKUS)],
        ["FAL  (Image Generation)",
         inr_lakh(fal_inr(1)*MONTHLY_SKUS),
         inr_lakh(fal_inr(2)*MONTHLY_SKUS),
         inr_lakh(fal_inr(3)*MONTHLY_SKUS)],
        ["System Processing",
         inr_lakh(SYSTEM_CHARGE_INR*MONTHLY_SKUS),
         inr_lakh(SYSTEM_CHARGE_INR*MONTHLY_SKUS),
         inr_lakh(SYSTEM_CHARGE_INR*MONTHLY_SKUS)],
        ["GRAND TOTAL / MONTH",
         inr_lakh(monthly_inr(1)),
         inr_lakh(monthly_inr(2)),
         inr_lakh(monthly_inr(3))],
        ["Effective Cost / SKU",
         inr(total_inr(1)), inr(total_inr(2)), inr(total_inr(3))],
    ]
    mon_cw = [col_w * 0.40, col_w * 0.20, col_w * 0.20, col_w * 0.20]
    mon_tbl = Table(monthly_data, colWidths=mon_cw, repeatRows=1)
    mon_tbl.setStyle(TableStyle([
        # Header row — navy/gold to match brand
        ("BACKGROUND",    (0, 0), (-1, 0), colors.HexColor(NAVY)),
        ("TEXTCOLOR",     (0, 0), (-1, 0), colors.HexColor(GOLD)),
        ("FONTNAME",      (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, 0), 7.3),
        ("ALIGN",         (0, 0), (0, -1), "LEFT"),
        ("ALIGN",         (1, 0), (-1, -1), "RIGHT"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING",   (0, 0), (-1, -1), 5),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 5),
        # Body rows
        ("FONTNAME",      (0, 1), (-1, 3), "Helvetica"),
        ("FONTSIZE",      (0, 1), (-1, 3), 7.3),
        ("TEXTCOLOR",     (0, 1), (-1, 3), colors.HexColor(STEEL)),
        ("BACKGROUND",    (0, 1), (-1, 1), colors.white),
        ("BACKGROUND",    (0, 2), (-1, 2), colors.HexColor("#F7F9FC")),
        ("BACKGROUND",    (0, 3), (-1, 3), colors.white),
        # Grand total row — gold band, bold navy
        ("BACKGROUND",    (0, 4), (-1, 4), colors.HexColor(GOLD)),
        ("TEXTCOLOR",     (0, 4), (-1, 4), colors.HexColor(NAVY)),
        ("FONTNAME",      (0, 4), (-1, 4), "Helvetica-Bold"),
        ("FONTSIZE",      (0, 4), (-1, 4), 8.2),
        ("TOPPADDING",    (0, 4), (-1, 4), 5),
        ("BOTTOMPADDING", (0, 4), (-1, 4), 5),
        # Per-SKU reference row — light footer style
        ("BACKGROUND",    (0, 5), (-1, 5), colors.HexColor("#FDF8F0")),
        ("FONTNAME",      (0, 5), (-1, 5), "Helvetica-Oblique"),
        ("FONTSIZE",      (0, 5), (-1, 5), 6.8),
        ("TEXTCOLOR",     (0, 5), (-1, 5), colors.HexColor(MID)),
        # Grid + accent rules
        ("GRID",          (0, 0), (-1, -1), 0.4, colors.HexColor("#D8DEE8")),
        ("LINEBELOW",     (0, 0), (-1, 0), 1.4, colors.HexColor(GOLD)),
        ("LINEABOVE",     (0, 4), (-1, 4), 1.2, colors.HexColor(NAVY)),
        ("LINEBELOW",     (0, 4), (-1, 4), 1.2, colors.HexColor(NAVY)),
        ("ROUNDEDCORNERS", [3, 3, 3, 3]),
    ]))

    # Prompt-cache saving estimate (static 1,000 tokens × 50% × 200k SKUs)
    cache_saving = int(1_000 * 0.50 * GROQ_IN_PER_1M / 1_000_000 * r * MONTHLY_SKUS)

    right_col = cost_header + [
        model_tbl,
    ] + cost_subheader + [
        micro_tbl,
        Spacer(1, 1.5*mm),
        Paragraph("MONTHLY COST SUMMARY  —  2,00,000 SKUs / Month", s_section),
        mon_tbl,
        Spacer(1, 1*mm),
        Paragraph(
            f"★  Groq prompt caching (50% off static input tokens) saves ~₹{cache_saving:,}/month at 2,00,000 SKUs.",
            sty("tip", fontName="Helvetica-Oblique", fontSize=6.8,
                textColor=colors.HexColor(GOLD), leading=9)
        ),
    ]

    body_data = [[overview_paras, right_col]]
    body_tbl  = Table(body_data, colWidths=[col_w, col_w],
                      hAlign="LEFT")
    body_tbl.setStyle(TableStyle([
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (0, -1),  col_gap / 2),
        ("LEFTPADDING",   (1, 0), (1, -1),  col_gap / 2),
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("LINEAFTER",     (0, 0), (0, -1),  0.4, colors.HexColor("#D0D8E4")),
    ]))
    elems.append(body_tbl)
    elems.append(Spacer(1, 1*mm))

    # ══ CHART ═════════════════════════════════════════════════════════════════
    chart_png = build_chart()
    chart_img = RLImage(io.BytesIO(chart_png), width=content_w, height=content_w * (3.9/13.5))
    elems.append(chart_img)
    elems.append(Spacer(1, 1*mm))

    # ══ FOOTER ════════════════════════════════════════════════════════════════
    footer_rule = Table([[""]], colWidths=[content_w])
    footer_rule.setStyle(TableStyle([
        ("LINEABOVE",     (0, 0), (-1, -1), 0.4, colors.HexColor("#D0D8E4")),
        ("TOPPADDING",    (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    elems.append(footer_rule)
    elems.append(Paragraph(
        f"Amazon Listing Automation — Confidential Proposal  ·  "
        f"Groq: Llama 4 Scout 17B-16E @ $0.11/$0.34 per 1M tokens  ·  "
        f"FAL: flux/schnell @ $0.003/image  ·  "
        f"1 USD = ₹{USD_TO_INR:.2f}  ·  Volume: 2,00,000 SKUs/month  ·  "
        f"Generated {date.today().strftime('%d %b %Y')}",
        s_footer
    ))

    doc.build(elems)
    print(f"PDF saved → {OUTPUT_PATH}")


if __name__ == "__main__":
    build_pdf()
