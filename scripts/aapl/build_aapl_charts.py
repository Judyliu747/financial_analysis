"""
Apple AAPL Q1 FY2026 (Calendar Q4 2025) Earnings Update
Chart Generation — 10 charts for equity research report
"""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import matplotlib.patches as mpatches
import numpy as np

OUT_DIR = "/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis/output/AAPL"

# ── Brand palette ─────────────────────────────────────────────────────────────
APPLE_GREY   = "#636366"
APPLE_BLUE   = "#0071E3"
APPLE_GREEN  = "#34C759"
APPLE_RED    = "#FF3B30"
APPLE_ORANGE = "#FF9500"
APPLE_PURPLE = "#AF52DE"
CHART_BG     = "#FAFAFA"
GRID_COLOR   = "#E5E5EA"
TEXT_COLOR   = "#1C1C1E"

FONT_FAMILY = "DejaVu Sans"

plt.rcParams.update({
    "font.family":       FONT_FAMILY,
    "axes.spines.top":   False,
    "axes.spines.right": False,
    "axes.edgecolor":    GRID_COLOR,
    "axes.labelcolor":   TEXT_COLOR,
    "axes.titlecolor":   TEXT_COLOR,
    "axes.facecolor":    CHART_BG,
    "figure.facecolor":  CHART_BG,
    "xtick.color":       TEXT_COLOR,
    "ytick.color":       TEXT_COLOR,
    "grid.color":        GRID_COLOR,
    "grid.linewidth":    0.8,
})

def save(fig, name):
    path = os.path.join(OUT_DIR, name)
    fig.savefig(path, dpi=150, bbox_inches="tight", facecolor=CHART_BG)
    plt.close(fig)
    print(f"  Saved: {name}")

# ── Chart 1: Quarterly Revenue Progression (8 quarters) ──────────────────────
def chart1_revenue():
    quarters = ["Q2\nFY24", "Q3\nFY24", "Q4\nFY24", "Q1\nFY25",
                "Q2\nFY25", "Q3\nFY25", "Q4\nFY25", "Q1\nFY26"]
    revenue  = [90.75, 85.78, 94.93, 124.30, 95.36, 85.78, 102.50, 143.77]

    fig, ax = plt.subplots(figsize=(9, 4.5))
    colors = [APPLE_GREY] * 7 + [APPLE_BLUE]
    bars = ax.bar(quarters, revenue, color=colors, width=0.6, zorder=3)

    # Label bars
    for bar, val in zip(bars, revenue):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 1.5,
                f"${val:.1f}B", ha="center", va="bottom", fontsize=9,
                color=TEXT_COLOR, fontweight="bold" if val == 143.77 else "normal")

    # Highlight Q1 FY26
    bars[-1].set_edgecolor(APPLE_BLUE)
    bars[-1].set_linewidth(2)

    ax.set_ylim(0, 165)
    ax.set_ylabel("Revenue ($B)", fontsize=10)
    ax.set_title("Quarterly Revenue Progression", fontsize=12, fontweight="bold", pad=10)
    ax.yaxis.grid(True, zorder=0)
    ax.set_axisbelow(True)

    # Beat annotation
    ax.annotate("Beat by $5.4B\n(+3.9% vs. consensus)", xy=(7, 143.77),
                xytext=(5.8, 153), fontsize=8.5, color=APPLE_GREEN, fontweight="bold",
                arrowprops=dict(arrowstyle="->", color=APPLE_GREEN, lw=1.5))

    ax.text(0.01, -0.12,
            "Source: Apple Earnings Releases, FQ4 2023 – FQ1 2026 (apple.com/newsroom). Q1 FY2026 reported January 29, 2026.",
            transform=ax.transAxes, fontsize=7, color=APPLE_GREY)

    fig.tight_layout()
    save(fig, "chart1_revenue.png")

# ── Chart 2: EPS Progression ──────────────────────────────────────────────────
def chart2_eps():
    quarters = ["Q2\nFY24", "Q3\nFY24", "Q4\nFY24", "Q1\nFY25",
                "Q2\nFY25", "Q3\nFY25", "Q4\nFY25", "Q1\nFY26"]
    eps      = [1.53, 1.40, 1.64, 2.40, 1.65, 1.40, 1.85, 2.84]
    consensus_q1_26 = 2.675

    fig, ax = plt.subplots(figsize=(9, 4.5))

    ax.plot(quarters, eps, color=APPLE_BLUE, lw=2.5, marker="o",
            markersize=7, zorder=3, label="Diluted EPS (GAAP)")
    ax.scatter([quarters[-1]], [consensus_q1_26], color=APPLE_ORANGE, s=90,
               marker="D", zorder=4, label=f"Consensus Q1 FY26 (${consensus_q1_26:.2f})")

    for i, (q, e) in enumerate(zip(quarters, eps)):
        offset = 0.07
        ax.text(i, e + offset, f"${e:.2f}", ha="center", va="bottom", fontsize=9,
                color=APPLE_BLUE if i == len(eps)-1 else TEXT_COLOR,
                fontweight="bold" if i == len(eps)-1 else "normal")

    ax.set_ylim(0.9, 3.4)
    ax.set_ylabel("Diluted EPS ($)", fontsize=10)
    ax.set_title("Diluted EPS Progression", fontsize=12, fontweight="bold", pad=10)
    ax.yaxis.grid(True, zorder=0)
    ax.set_axisbelow(True)
    ax.legend(fontsize=9, loc="upper left")

    ax.annotate(f"+18.3% YoY\nBeat +$0.17", xy=(7, 2.84),
                xytext=(6.0, 3.1), fontsize=8.5, color=APPLE_GREEN, fontweight="bold",
                arrowprops=dict(arrowstyle="->", color=APPLE_GREEN, lw=1.5))

    ax.text(0.01, -0.12,
            "Source: Apple Earnings Releases FQ4 2023 – FQ1 2026; Consensus from Bloomberg.",
            transform=ax.transAxes, fontsize=7, color=APPLE_GREY)

    fig.tight_layout()
    save(fig, "chart2_eps.png")

# ── Chart 3: Revenue by Segment — Q1 FY26 vs Q1 FY25 ────────────────────────
def chart3_segments():
    segments = ["iPhone", "Services", "Mac", "iPad", "Wearables &\nHome/Acc."]
    q1_fy25  = [69.14, 26.34, 9.01, 8.09, 11.76]
    q1_fy26  = [85.27, 30.01, 8.40, 8.60, 11.50]
    yoy      = [(a-b)/b*100 for a,b in zip(q1_fy26, q1_fy25)]

    x = np.arange(len(segments))
    w = 0.35

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))

    # Left: grouped bars
    b1 = ax1.bar(x - w/2, q1_fy25, w, label="Q1 FY2025", color=APPLE_GREY, alpha=0.8)
    b2 = ax1.bar(x + w/2, q1_fy26, w, label="Q1 FY2026", color=APPLE_BLUE)

    for bar, val in zip(b2, q1_fy26):
        ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
                 f"${val:.1f}B", ha="center", va="bottom", fontsize=8.5, color=APPLE_BLUE)

    ax1.set_xticks(x)
    ax1.set_xticklabels(segments, fontsize=9)
    ax1.set_ylabel("Revenue ($B)", fontsize=10)
    ax1.set_title("Revenue by Segment: Q1 FY26 vs. Q1 FY25", fontsize=11, fontweight="bold")
    ax1.yaxis.grid(True, zorder=0); ax1.set_axisbelow(True)
    ax1.legend(fontsize=9)
    ax1.set_ylim(0, 105)

    # Right: YoY growth bars
    growth_colors = [APPLE_GREEN if g > 0 else APPLE_RED for g in yoy]
    bars = ax2.barh(segments[::-1], yoy[::-1], color=growth_colors[::-1], height=0.5)
    ax2.axvline(0, color=TEXT_COLOR, lw=0.8)

    for bar, val in zip(bars, yoy[::-1]):
        offset = 0.5 if val >= 0 else -0.5
        ha = "left" if val >= 0 else "right"
        ax2.text(val + offset, bar.get_y() + bar.get_height()/2,
                 f"{val:+.1f}%", va="center", ha=ha, fontsize=9, fontweight="bold",
                 color=APPLE_GREEN if val > 0 else APPLE_RED)

    ax2.set_xlabel("YoY Growth (%)", fontsize=10)
    ax2.set_title("YoY Growth by Segment (Q1 FY26 vs. Q1 FY25)", fontsize=11, fontweight="bold")
    ax2.yaxis.grid(False); ax2.xaxis.grid(True, zorder=0)

    fig.text(0.01, -0.04,
             "Source: Apple Q1 FY2026 Earnings Release, January 29, 2026 (apple.com/newsroom).",
             fontsize=7, color=APPLE_GREY)

    fig.tight_layout()
    save(fig, "chart3_segments.png")

# ── Chart 4: Gross Margin Trend ───────────────────────────────────────────────
def chart4_margins():
    quarters = ["Q2\nFY24", "Q3\nFY24", "Q4\nFY24", "Q1\nFY25",
                "Q2\nFY25", "Q3\nFY25", "Q4\nFY25", "Q1\nFY26"]
    total_gm   = [46.6, 46.3, 46.2, 46.9, 47.1, 46.3, 46.5, 48.2]
    svc_gm     = [73.8, 74.6, 74.0, 75.0, 74.9, 74.8, 74.9, 76.5]
    prod_gm    = [35.5, 35.3, 35.5, 35.5, 36.0, 35.0, 35.2, 40.7]

    fig, ax = plt.subplots(figsize=(10, 5))
    x = range(len(quarters))

    ax.plot(x, total_gm, color=APPLE_BLUE,  lw=2.5, marker="o", markersize=7, label="Total Gross Margin")
    ax.plot(x, svc_gm,   color=APPLE_GREEN, lw=2,   marker="s", markersize=6, label="Services Gross Margin", linestyle="--")
    ax.plot(x, prod_gm,  color=APPLE_ORANGE,lw=2,   marker="^", markersize=6, label="Products Gross Margin", linestyle=":")

    # Label key Q1 FY26 points
    for series, col in [(total_gm, APPLE_BLUE), (svc_gm, APPLE_GREEN), (prod_gm, APPLE_ORANGE)]:
        ax.text(len(quarters)-1 + 0.15, series[-1], f"{series[-1]:.1f}%",
                va="center", fontsize=9, color=col, fontweight="bold")

    ax.set_xticks(range(len(quarters)))
    ax.set_xticklabels(quarters, fontsize=9)
    ax.set_ylim(28, 82)
    ax.set_ylabel("Gross Margin (%)", fontsize=10)
    ax.set_title("Gross Margin Trends — Total, Products & Services", fontsize=12, fontweight="bold", pad=10)
    ax.yaxis.grid(True, zorder=0); ax.set_axisbelow(True)
    ax.legend(fontsize=9, loc="lower left")

    # Highlight Q1 FY26 products margin step-up
    ax.annotate("+540 bps QoQ\n(Products)", xy=(7, 40.7), xytext=(5.5, 44),
                fontsize=8.5, color=APPLE_ORANGE, fontweight="bold",
                arrowprops=dict(arrowstyle="->", color=APPLE_ORANGE, lw=1.5))

    ax.text(0.01, -0.12,
            "Source: Apple Earnings Releases FQ4 2023 – FQ1 2026; Q1 FY2026 10-Q filed February 2026.",
            transform=ax.transAxes, fontsize=7, color=APPLE_GREY)

    fig.tight_layout()
    save(fig, "chart4_margins.png")

# ── Chart 5: Services Revenue Growth Trajectory ───────────────────────────────
def chart5_services():
    quarters = ["Q1\nFY24", "Q2\nFY24", "Q3\nFY24", "Q4\nFY24",
                "Q1\nFY25", "Q2\nFY25", "Q3\nFY25", "Q4\nFY25", "Q1\nFY26"]
    svc_rev  = [23.12, 23.87, 24.21, 24.97, 26.34, 26.65, 27.42, 28.75, 30.01]
    yoy_gw   = [None, None, None, None, 13.9, 11.7, 13.3, 15.1, 14.0]

    fig, ax1 = plt.subplots(figsize=(10, 5))
    ax2 = ax1.twinx()

    x = range(len(quarters))
    ax1.bar(x, svc_rev, color=APPLE_GREEN, alpha=0.85, width=0.5, zorder=2, label="Services Revenue ($B)")

    for i, v in enumerate(svc_rev):
        ax1.text(i, v + 0.25, f"${v:.2f}B", ha="center", va="bottom", fontsize=8.5,
                 color=APPLE_GREEN, fontweight="bold" if i == len(svc_rev)-1 else "normal")

    yoy_plot = [v for v in yoy_gw if v is not None]
    x_yoy    = [i for i, v in enumerate(yoy_gw) if v is not None]
    ax2.plot(x_yoy, yoy_plot, color=APPLE_BLUE, lw=2, marker="o", markersize=6,
             label="YoY Growth (%)", linestyle="--")

    ax1.set_xticks(x)
    ax1.set_xticklabels(quarters, fontsize=9)
    ax1.set_ylim(0, 36)
    ax1.set_ylabel("Services Revenue ($B)", fontsize=10, color=APPLE_GREEN)
    ax2.set_ylim(8, 20)
    ax2.set_ylabel("YoY Growth (%)", fontsize=10, color=APPLE_BLUE)
    ax1.set_title("Services Revenue: Growth Trajectory (Q1 FY24 – Q1 FY26)", fontsize=12, fontweight="bold", pad=10)
    ax1.yaxis.grid(True, zorder=0); ax1.set_axisbelow(True)

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, fontsize=9, loc="upper left")

    # $30B milestone annotation
    ax1.axhline(30, color=APPLE_ORANGE, lw=1.2, linestyle=":", alpha=0.7)
    ax1.text(8.5, 30.3, "$30B milestone", fontsize=8, color=APPLE_ORANGE, ha="right")

    ax1.text(0.01, -0.12,
             "Source: Apple Earnings Releases Q1 FY2024 – Q1 FY2026; Apple Annual Reports.",
             transform=ax1.transAxes, fontsize=7, color=APPLE_GREY)

    fig.tight_layout()
    save(fig, "chart5_services.png")

# ── Chart 6: Geographic Revenue Breakdown ─────────────────────────────────────
def chart6_geography():
    regions  = ["Americas", "Europe", "Greater\nChina", "Japan", "Rest of\nAsia Pac."]
    q1_fy25  = [50.65, 33.89, 18.51, 7.10, 14.15]
    q1_fy26  = [58.50, 38.10, 25.50, 7.46, 14.21]
    yoy      = [(a-b)/b*100 for a, b in zip(q1_fy26, q1_fy25)]

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))

    # Stacked / side-by-side comparison
    x = np.arange(len(regions))
    w = 0.38
    ax1.bar(x - w/2, q1_fy25, w, label="Q1 FY2025", color=APPLE_GREY, alpha=0.8)
    ax1.bar(x + w/2, q1_fy26, w, label="Q1 FY2026", color=APPLE_BLUE)

    for bar, val in zip(ax1.patches[len(regions):], q1_fy26):
        ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
                 f"${val:.1f}B", ha="center", va="bottom", fontsize=8.5, color=APPLE_BLUE)

    ax1.set_xticks(x)
    ax1.set_xticklabels(regions, fontsize=9)
    ax1.set_ylabel("Revenue ($B)", fontsize=10)
    ax1.set_title("Geographic Revenue: Q1 FY26 vs. Q1 FY25", fontsize=11, fontweight="bold")
    ax1.yaxis.grid(True, zorder=0); ax1.set_axisbelow(True)
    ax1.legend(fontsize=9)
    ax1.set_ylim(0, 70)

    # YoY growth
    growth_colors = [APPLE_GREEN if g > 0 else APPLE_RED for g in yoy]
    bars = ax2.barh(regions[::-1], yoy[::-1], color=growth_colors[::-1], height=0.45)
    ax2.axvline(0, color=TEXT_COLOR, lw=0.8)

    for bar, val in zip(bars, yoy[::-1]):
        offset = 0.4 if val >= 0 else -0.4
        ha = "left" if val >= 0 else "right"
        ax2.text(val + offset, bar.get_y() + bar.get_height()/2,
                 f"{val:+.1f}%", va="center", ha=ha, fontsize=9.5, fontweight="bold",
                 color=APPLE_GREEN if val > 0 else APPLE_RED)

    ax2.set_xlabel("YoY Growth (%)", fontsize=10)
    ax2.set_title("YoY Growth by Region (Q1 FY26 vs. Q1 FY25)", fontsize=11, fontweight="bold")
    ax2.xaxis.grid(True, zorder=0)

    # Greater China highlight
    ax2.get_yticklabels()[4].set_color(APPLE_GREEN)
    ax2.get_yticklabels()[4].set_fontweight("bold")

    fig.text(0.01, -0.04,
             "Source: Apple Q1 FY2026 Earnings Release, January 29, 2026. Japan/Rest of Asia Pacific estimated from total.",
             fontsize=7, color=APPLE_GREY)

    fig.tight_layout()
    save(fig, "chart6_geography.png")

# ── Chart 7: Beat/Miss Summary ────────────────────────────────────────────────
def chart7_beat_miss():
    metrics    = ["Revenue\n($B)", "Diluted EPS\n($)", "Gross\nMargin (%)"]
    actual     = [143.77, 2.84, 48.2]
    consensus  = [138.40, 2.68, 47.0]
    beat_pct   = [(a - c) / c * 100 for a, c in zip(actual, consensus)]
    beat_abs   = ["+$5.37B", "+$0.16", "+120 bps"]

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, 4.5))

    # Left: actual vs consensus
    x = np.arange(len(metrics))
    w = 0.35
    b_cons = ax1.bar(x - w/2, consensus, w, label="Consensus Estimate", color=APPLE_GREY, alpha=0.8)
    b_act  = ax1.bar(x + w/2, actual,    w, label="Actual Result",       color=APPLE_BLUE)

    ax1.set_xticks(x); ax1.set_xticklabels(metrics, fontsize=10)
    ax1.set_title("Actual vs. Consensus Estimate", fontsize=11, fontweight="bold")
    ax1.yaxis.grid(True, zorder=0); ax1.set_axisbelow(True)
    ax1.legend(fontsize=9)

    for bar, ba in zip(b_act, beat_abs):
        ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + bar.get_height()*0.01,
                 ba, ha="center", va="bottom", fontsize=9, color=APPLE_GREEN, fontweight="bold")

    # Right: beat% bars
    colors = [APPLE_GREEN if b > 0 else APPLE_RED for b in beat_pct]
    beat_bars = ax2.bar(metrics, beat_pct, color=colors, width=0.4)
    ax2.axhline(0, color=TEXT_COLOR, lw=0.8)
    ax2.set_ylabel("Beat / Miss (%)", fontsize=10)
    ax2.set_title("Q1 FY2026 Beat vs. Street Consensus", fontsize=11, fontweight="bold")
    ax2.yaxis.grid(True, zorder=0); ax2.set_axisbelow(True)

    for bar, val in zip(beat_bars, beat_pct):
        ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.05,
                 f"{val:+.1f}%", ha="center", va="bottom", fontsize=10, fontweight="bold",
                 color=APPLE_GREEN)

    ax2.set_ylim(-1, 6)

    fig.text(0.01, -0.04,
             "Source: Apple Q1 FY2026 Earnings Release (Jan 29, 2026); Bloomberg consensus as of Jan 28, 2026.",
             fontsize=7, color=APPLE_GREY)

    fig.tight_layout()
    save(fig, "chart7_beat_miss.png")

# ── Chart 8: Operating Cash Flow & Free Cash Flow ────────────────────────────
def chart8_cashflow():
    quarters = ["Q2\nFY24", "Q3\nFY24", "Q4\nFY24", "Q1\nFY25",
                "Q2\nFY25", "Q3\nFY25", "Q4\nFY25", "Q1\nFY26"]
    ocf      = [22.7, 29.0, 26.8, 39.9, 21.5, 28.9, 26.0, 53.9]
    capex    = [2.8,  2.8,  2.9,  2.6,  2.2,  2.8,  2.7,  5.1]
    fcf      = [o - c for o, c in zip(ocf, capex)]

    fig, ax = plt.subplots(figsize=(10, 4.5))
    x = range(len(quarters))
    w = 0.35

    ax.bar([i - w/2 for i in x], ocf, w, label="Operating Cash Flow", color=APPLE_BLUE, alpha=0.9)
    ax.bar([i + w/2 for i in x], fcf, w, label="Free Cash Flow (est.)", color=APPLE_GREEN, alpha=0.9)

    ax.set_xticks(list(x))
    ax.set_xticklabels(quarters, fontsize=9)
    ax.set_ylabel("Cash Flow ($B)", fontsize=10)
    ax.set_title("Operating Cash Flow & Free Cash Flow", fontsize=12, fontweight="bold", pad=10)
    ax.yaxis.grid(True, zorder=0); ax.set_axisbelow(True)
    ax.legend(fontsize=9)

    # Annotate Q1 FY26 OCF record
    ax.text(7 - w/2, 53.9 + 1, "$53.9B\n(All-time record)", ha="center", va="bottom",
            fontsize=8.5, color=APPLE_BLUE, fontweight="bold")

    ax.text(0.01, -0.12,
            "Source: Apple Earnings Releases; FCF = Operating Cash Flow – CapEx (estimated). Q1 FY26 OCF per Apple press release.",
            transform=ax.transAxes, fontsize=7, color=APPLE_GREY)

    ax.set_ylim(0, 65)
    fig.tight_layout()
    save(fig, "chart8_cashflow.png")

# ── Chart 9: Capital Return Program ───────────────────────────────────────────
def chart9_buybacks():
    fy_labels = ["FY2021", "FY2022", "FY2023", "FY2024", "FY2025", "Q1 FY26\n(Qtrly)"]
    buybacks  = [85.5,  89.4,  77.6,  94.9,  90.7,  25.0]
    dividends = [14.5,  14.8,  15.1,  15.2,  15.4,   3.9]

    x = np.arange(len(fy_labels))
    w = 0.55

    fig, ax = plt.subplots(figsize=(10, 4.5))
    b1 = ax.bar(x, buybacks,  w, label="Share Repurchases", color=APPLE_BLUE, alpha=0.9)
    b2 = ax.bar(x, dividends, w, bottom=buybacks, label="Dividends", color=APPLE_GREEN, alpha=0.7)

    totals = [b + d for b, d in zip(buybacks, dividends)]
    for i, t in enumerate(totals):
        ax.text(i, t + 1, f"${t:.1f}B", ha="center", va="bottom", fontsize=9,
                fontweight="bold" if i >= len(fy_labels)-2 else "normal", color=TEXT_COLOR)

    ax.set_xticks(x); ax.set_xticklabels(fy_labels, fontsize=9)
    ax.set_ylabel("Capital Returned ($B)", fontsize=10)
    ax.set_title("Capital Return Program: Buybacks & Dividends", fontsize=12, fontweight="bold", pad=10)
    ax.yaxis.grid(True, zorder=0); ax.set_axisbelow(True)
    ax.legend(fontsize=9, loc="upper left")
    ax.set_ylim(0, 130)

    ax.text(0.01, -0.12,
            "Source: Apple Annual Reports FY2021–FY2025; Q1 FY2026 per Earnings Release (Jan 29, 2026). "
            "Q1 FY26 shown as quarterly figure (not annualized).",
            transform=ax.transAxes, fontsize=7, color=APPLE_GREY)

    fig.tight_layout()
    save(fig, "chart9_buybacks.png")

# ── Chart 10: Valuation — NTM P/E vs. 5-Year Range ───────────────────────────
def chart10_valuation():
    quarters_val = ["Q2\nFY23", "Q3\nFY23", "Q4\nFY23", "Q1\nFY24",
                    "Q2\nFY24", "Q3\nFY24", "Q4\nFY24", "Q1\nFY25",
                    "Q2\nFY25", "Q3\nFY25", "Q4\nFY25", "Q1\nFY26",
                    "Current\nMar 2026"]
    ntm_pe = [27.5, 30.0, 28.9, 30.2, 26.8, 31.0, 30.5, 33.1,
              29.5, 30.2, 31.5, 30.5, 32.1]

    x = range(len(quarters_val))
    fig, ax = plt.subplots(figsize=(11, 5))

    ax.plot(x, ntm_pe, color=APPLE_BLUE, lw=2.5, marker="o", markersize=6, zorder=3)
    ax.fill_between(x, ntm_pe, alpha=0.1, color=APPLE_BLUE)

    # Reference bands
    five_yr_avg = 29.5
    ax.axhline(five_yr_avg, color=APPLE_ORANGE, lw=1.5, linestyle="--",
               label=f"5-Year Avg. NTM P/E: {five_yr_avg:.1f}x")
    ax.fill_between(x, 25, 35, alpha=0.05, color=APPLE_GREY, label="Historical range band (25–35x)")

    ax.text(len(x)-1+0.2, ntm_pe[-1], f" {ntm_pe[-1]:.1f}x", va="center",
            fontsize=10, color=APPLE_BLUE, fontweight="bold")

    ax.set_xticks(list(x)); ax.set_xticklabels(quarters_val, fontsize=8.5)
    ax.set_ylim(18, 40)
    ax.set_ylabel("NTM P/E Multiple", fontsize=10)
    ax.set_title("AAPL NTM P/E Valuation vs. Historical Trend", fontsize=12, fontweight="bold", pad=10)
    ax.yaxis.grid(True, zorder=0); ax.set_axisbelow(True)
    ax.legend(fontsize=9, loc="upper left")

    # Current vs. 52-wk high annotation
    ax.annotate("13% below\n52-wk high ($288.62)", xy=(12, 32.1), xytext=(10.2, 37.5),
                fontsize=8.5, color=APPLE_RED, fontweight="bold",
                arrowprops=dict(arrowstyle="->", color=APPLE_RED, lw=1.5))

    ax.text(0.01, -0.12,
            "Source: Bloomberg NTM consensus estimates; AAPL price data. Current price $250.12 as of March 13, 2026.",
            transform=ax.transAxes, fontsize=7, color=APPLE_GREY)

    fig.tight_layout()
    save(fig, "chart10_valuation.png")


# ── Run all ───────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("Generating AAPL Q1 FY2026 charts...")
    chart1_revenue()
    chart2_eps()
    chart3_segments()
    chart4_margins()
    chart5_services()
    chart6_geography()
    chart7_beat_miss()
    chart8_cashflow()
    chart9_buybacks()
    chart10_valuation()
    print("Done — 10 charts saved.")
