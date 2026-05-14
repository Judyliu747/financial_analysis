"""
Coherent Corp. (COHR) Q3 FY2026 Earnings Update
Chart Generation — 10 charts for equity research report
"""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np

OUT_DIR = "/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis/output/COHR"

# ── Brand palette ─────────────────────────────────────────────────────────────
COHR_BLUE    = "#003DA5"
COHR_LIGHT   = "#4A90D9"
COHR_DARK    = "#1A1A2E"
COHR_GREEN   = "#2D9D78"
COHR_RED     = "#CC3333"
COHR_ORANGE  = "#E68619"
COHR_PURPLE  = "#7B61FF"
COHR_TEAL    = "#0099CC"
CHART_BG     = "#FAFAFA"
GRID_COLOR   = "#E5E5EA"
TEXT_COLOR    = "#1C1C1E"

plt.rcParams.update({
    "font.family":       "DejaVu Sans",
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

# ═══════════════════════════════════════════════════════════════════════════════
# DATA
# ═══════════════════════════════════════════════════════════════════════════════
# Quarterly revenue ($B) — Q4 FY24 through Q3 FY26
quarters_8   = ["Q4\nFY24","Q1\nFY25","Q2\nFY25","Q3\nFY25",
                "Q4\nFY25","Q1\nFY26","Q2\nFY26","Q3\nFY26"]
revenue_8    = [1.32, 1.35, 1.41, 1.50, 1.53, 1.58, 1.69, 1.81]

# Non-GAAP EPS
eps_8        = [0.63, 0.71, 0.82, 0.91, 0.95, 1.16, 1.29, 1.41]

# GAAP EPS
gaap_eps_8   = [0.32, 0.28, 0.50, 0.38, 0.58, 1.19, 0.76, 0.97]

# Consensus for Q3 FY26
rev_consensus = 1.78
eps_consensus = 1.39

# Segment revenue ($B) — Q3 FY26
seg_dc   = 1.36   # Datacenter & Communications
seg_ind  = 0.444  # Industrial
seg_dc_yoy  = 36.0
seg_ind_yoy = -16.0

# Margins (%) — last 6 quarters
margin_qtrs      = ["Q2\nFY25","Q3\nFY25","Q4\nFY25","Q1\nFY26","Q2\nFY26","Q3\nFY26"]
gaap_gross_mg    = [35.8, 36.5, 36.3, 36.6, 36.9, 37.7]
nongaap_gross_mg = [37.5, 38.1, 38.3, 38.7, 39.0, 39.6]
nongaap_op_mg    = [17.0, 17.8, 18.5, 19.1, 19.6, 20.3]

# Q4 FY26 Guidance
q4_guide_rev_low  = 1.91
q4_guide_rev_high = 2.05
q4_guide_eps_low  = 1.52
q4_guide_eps_high = 1.72

# YoY comparison — Q3 FY25 vs Q3 FY26
q3_fy25_rev = 1.50
q3_fy26_rev = 1.81
q3_fy25_eps = 0.91
q3_fy26_eps = 1.41

# Balance sheet
cash_q2 = 1.5
cash_q3 = 2.23   # excl NVIDIA, total ~3.0B with NVIDIA
leverage_q2 = 1.7
leverage_q3 = 0.5

# CapEx ($M)
capex_qtrs = ["Q1\nFY26","Q2\nFY26","Q3\nFY26"]
capex_vals = [130, 154, 290]

# ═══════════════════════════════════════════════════════════════════════════════
# CHART 1: Quarterly Revenue Progression (8 quarters)
# ═══════════════════════════════════════════════════════════════════════════════
def chart1_revenue():
    fig, ax = plt.subplots(figsize=(9, 4.5))
    colors = [COHR_DARK] * 7 + [COHR_BLUE]
    bars = ax.bar(quarters_8, revenue_8, color=colors, width=0.6, zorder=3)

    for bar, val in zip(bars, revenue_8):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.02,
                f"${val:.2f}B", ha="center", va="bottom", fontsize=9,
                color=TEXT_COLOR, fontweight="bold" if val == 1.81 else "normal")

    bars[-1].set_edgecolor(COHR_BLUE)
    bars[-1].set_linewidth(2)

    ax.set_ylim(0, 2.2)
    ax.set_ylabel("Revenue ($B)", fontsize=10)
    ax.set_title("Quarterly Revenue Progression", fontsize=12, fontweight="bold", pad=10)
    ax.yaxis.grid(True, zorder=0)
    ax.set_axisbelow(True)

    ax.annotate("Record: $1.81B\nBeat by ~$30M (+1.7%)", xy=(7, 1.81),
                xytext=(5.2, 2.05), fontsize=8.5, color=COHR_GREEN, fontweight="bold",
                arrowprops=dict(arrowstyle="->", color=COHR_GREEN, lw=1.5))

    ax.text(0.01, -0.12,
            "Source: Coherent Corp Earnings Releases, Q4 FY2024 – Q3 FY2026. Q3 FY2026 reported May 6, 2026.",
            transform=ax.transAxes, fontsize=7, color=COHR_DARK)

    fig.tight_layout()
    save(fig, "cohr_chart1_revenue.png")


# ═══════════════════════════════════════════════════════════════════════════════
# CHART 2: Non-GAAP EPS Progression
# ═══════════════════════════════════════════════════════════════════════════════
def chart2_eps():
    fig, ax = plt.subplots(figsize=(9, 4.5))

    ax.plot(quarters_8, eps_8, color=COHR_BLUE, lw=2.5, marker="o",
            markersize=7, zorder=3, label="Non-GAAP EPS")
    ax.scatter([quarters_8[-1]], [eps_consensus], color=COHR_ORANGE, s=90,
               marker="D", zorder=4, label=f"Consensus (${eps_consensus:.2f})")

    for i, (q, e) in enumerate(zip(quarters_8, eps_8)):
        ax.text(i, e + 0.03, f"${e:.2f}", ha="center", va="bottom", fontsize=9,
                color=COHR_BLUE if i == len(eps_8)-1 else TEXT_COLOR,
                fontweight="bold" if i == len(eps_8)-1 else "normal")

    ax.set_ylim(0.3, 1.8)
    ax.set_ylabel("Non-GAAP Diluted EPS ($)", fontsize=10)
    ax.set_title("Non-GAAP EPS Progression", fontsize=12, fontweight="bold", pad=10)
    ax.yaxis.grid(True, zorder=0)
    ax.set_axisbelow(True)
    ax.legend(fontsize=9, loc="upper left")

    ax.annotate(f"+55% YoY\nBeat +$0.02", xy=(7, 1.41),
                xytext=(5.5, 1.65), fontsize=8.5, color=COHR_GREEN, fontweight="bold",
                arrowprops=dict(arrowstyle="->", color=COHR_GREEN, lw=1.5))

    ax.text(0.01, -0.12,
            "Source: Coherent Corp Earnings Releases; Consensus from Bloomberg.",
            transform=ax.transAxes, fontsize=7, color=COHR_DARK)

    fig.tight_layout()
    save(fig, "cohr_chart2_eps.png")


# ═══════════════════════════════════════════════════════════════════════════════
# CHART 3: Revenue by Segment — Q3 FY26
# ═══════════════════════════════════════════════════════════════════════════════
def chart3_segments():
    labels = ["Datacenter &\nCommunications", "Industrial"]
    values = [seg_dc, seg_ind]
    colors_bar = [COHR_BLUE, COHR_TEAL]
    yoy = [seg_dc_yoy, seg_ind_yoy]

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))

    # Pie chart
    explode = (0.03, 0.03)
    wedges, texts, autotexts = ax1.pie(
        values, labels=labels, colors=colors_bar, autopct="%1.0f%%",
        startangle=90, explode=explode, textprops={"fontsize": 10})
    for at in autotexts:
        at.set_fontweight("bold")
        at.set_color("white")
    ax1.set_title("Revenue Mix — Q3 FY2026", fontsize=12, fontweight="bold", pad=10)

    # YoY growth bar chart
    bar_colors = [COHR_GREEN if y > 0 else COHR_RED for y in yoy]
    bars = ax2.barh(labels, yoy, color=bar_colors, height=0.5, zorder=3)
    for bar, val in zip(bars, yoy):
        xpos = bar.get_width() + 1 if val > 0 else bar.get_width() - 5
        ax2.text(xpos, bar.get_y() + bar.get_height()/2,
                 f"{val:+.0f}%", va="center", fontsize=11, fontweight="bold",
                 color=COHR_GREEN if val > 0 else COHR_RED)
    ax2.set_title("YoY Revenue Growth by Segment", fontsize=12, fontweight="bold", pad=10)
    ax2.axvline(0, color=TEXT_COLOR, lw=0.8)
    ax2.xaxis.grid(True, zorder=0)
    ax2.set_xlabel("YoY Growth (%)")

    fig.text(0.01, 0.01,
             "Source: Coherent Corp Q3 FY2026 Earnings Release (May 6, 2026).",
             fontsize=7, color=COHR_DARK)

    fig.tight_layout()
    save(fig, "cohr_chart3_segments.png")


# ═══════════════════════════════════════════════════════════════════════════════
# CHART 4: Margin Trends
# ═══════════════════════════════════════════════════════════════════════════════
def chart4_margins():
    fig, ax = plt.subplots(figsize=(9, 4.5))

    ax.plot(margin_qtrs, nongaap_gross_mg, color=COHR_BLUE, lw=2.5, marker="o",
            markersize=7, zorder=3, label="Non-GAAP Gross Margin")
    ax.plot(margin_qtrs, nongaap_op_mg, color=COHR_GREEN, lw=2.5, marker="s",
            markersize=7, zorder=3, label="Non-GAAP Operating Margin")
    ax.plot(margin_qtrs, gaap_gross_mg, color=COHR_LIGHT, lw=2, marker="^",
            markersize=6, zorder=3, label="GAAP Gross Margin", linestyle="--")

    for i, (gm, om) in enumerate(zip(nongaap_gross_mg, nongaap_op_mg)):
        ax.text(i, gm + 0.4, f"{gm:.1f}%", ha="center", va="bottom", fontsize=8,
                color=COHR_BLUE)
        ax.text(i, om - 0.8, f"{om:.1f}%", ha="center", va="top", fontsize=8,
                color=COHR_GREEN)

    ax.set_ylim(14, 44)
    ax.set_ylabel("Margin (%)", fontsize=10)
    ax.set_title("Margin Trends — Sustained Expansion", fontsize=12, fontweight="bold", pad=10)
    ax.yaxis.grid(True, zorder=0)
    ax.set_axisbelow(True)
    ax.legend(fontsize=9, loc="upper left")

    ax.text(0.01, -0.12,
            "Source: Coherent Corp Earnings Releases, Q2 FY2025 – Q3 FY2026.",
            transform=ax.transAxes, fontsize=7, color=COHR_DARK)

    fig.tight_layout()
    save(fig, "cohr_chart4_margins.png")


# ═══════════════════════════════════════════════════════════════════════════════
# CHART 5: Beat/Miss Summary — Q3 FY26
# ═══════════════════════════════════════════════════════════════════════════════
def chart5_beat_miss():
    fig, ax = plt.subplots(figsize=(9, 4.5))

    metrics = ["Revenue\n($B)", "Non-GAAP EPS\n($)", "Non-GAAP\nGross Margin (%)"]
    actuals = [1.81, 1.41, 39.6]
    estimates = [1.78, 1.39, 39.2]
    beats = [(a - e) / e * 100 for a, e in zip(actuals, estimates)]

    x = np.arange(len(metrics))
    w = 0.3
    bars1 = ax.bar(x - w/2, actuals, w, color=COHR_BLUE, label="Actual", zorder=3)
    bars2 = ax.bar(x + w/2, estimates, w, color=COHR_LIGHT, label="Consensus", zorder=3, alpha=0.7)

    for bar, val in zip(bars1, actuals):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.3,
                f"{val:.2f}" if val < 10 else f"{val:.1f}%", ha="center",
                fontsize=9, fontweight="bold", color=COHR_BLUE)

    for i, b in enumerate(beats):
        ax.text(i, max(actuals[i], estimates[i]) + 1.2,
                f"Beat +{b:.1f}%", ha="center", fontsize=9,
                fontweight="bold", color=COHR_GREEN)

    ax.set_xticks(x)
    ax.set_xticklabels(metrics, fontsize=10)
    ax.set_title("Q3 FY2026 Results vs. Consensus", fontsize=12, fontweight="bold", pad=10)
    ax.yaxis.grid(True, zorder=0)
    ax.set_axisbelow(True)
    ax.legend(fontsize=9)
    ax.set_ylim(0, 48)

    ax.text(0.01, -0.12,
            "Source: Coherent Corp Q3 FY2026 Earnings Release; Consensus from Bloomberg as of May 5, 2026.",
            transform=ax.transAxes, fontsize=7, color=COHR_DARK)

    fig.tight_layout()
    save(fig, "cohr_chart5_beat_miss.png")


# ═══════════════════════════════════════════════════════════════════════════════
# CHART 6: Q4 FY26 Guidance Bridge
# ═══════════════════════════════════════════════════════════════════════════════
def chart6_guidance():
    fig, ax = plt.subplots(figsize=(9, 4.5))

    quarters = ["Q1\nFY26\nActual","Q2\nFY26\nActual","Q3\nFY26\nActual","Q4\nFY26\nGuidance"]
    rev_act = [1.58, 1.69, 1.81, None]
    guide_mid = (q4_guide_rev_low + q4_guide_rev_high) / 2

    bar_colors = [COHR_BLUE, COHR_BLUE, COHR_BLUE, COHR_ORANGE]
    bar_vals = [1.58, 1.69, 1.81, guide_mid]

    bars = ax.bar(quarters, bar_vals, color=bar_colors, width=0.55, zorder=3)

    # Q4 guidance range
    ax.errorbar(3, guide_mid, yerr=[[guide_mid - q4_guide_rev_low],
                [q4_guide_rev_high - guide_mid]], fmt="none",
                ecolor=COHR_DARK, capsize=8, lw=2, zorder=4)

    for bar, val in zip(bars, bar_vals):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.02,
                f"${val:.2f}B", ha="center", va="bottom", fontsize=9,
                fontweight="bold")

    ax.text(3, q4_guide_rev_high + 0.06,
            f"Range: ${q4_guide_rev_low:.2f} – ${q4_guide_rev_high:.2f}B",
            ha="center", fontsize=8, color=COHR_ORANGE, fontweight="bold")

    implied_fy = 1.58 + 1.69 + 1.81 + guide_mid
    ax.text(0.98, 0.95, f"Implied FY2026: ~${implied_fy:.2f}B\n(+21% YoY)",
            transform=ax.transAxes, ha="right", va="top", fontsize=9,
            color=COHR_GREEN, fontweight="bold",
            bbox=dict(boxstyle="round,pad=0.3", facecolor="white", edgecolor=COHR_GREEN, alpha=0.8))

    ax.set_ylim(0, 2.4)
    ax.set_ylabel("Revenue ($B)", fontsize=10)
    ax.set_title("FY2026 Revenue Trajectory & Q4 Guidance", fontsize=12, fontweight="bold", pad=10)
    ax.yaxis.grid(True, zorder=0)
    ax.set_axisbelow(True)

    ax.text(0.01, -0.12,
            "Source: Coherent Corp Earnings Releases & Guidance, Q3 FY2026.",
            transform=ax.transAxes, fontsize=7, color=COHR_DARK)

    fig.tight_layout()
    save(fig, "cohr_chart6_guidance.png")


# ═══════════════════════════════════════════════════════════════════════════════
# CHART 7: CapEx Ramp
# ═══════════════════════════════════════════════════════════════════════════════
def chart7_capex():
    fig, ax = plt.subplots(figsize=(9, 4.5))

    bars = ax.bar(capex_qtrs, capex_vals, color=[COHR_LIGHT, COHR_LIGHT, COHR_BLUE],
                  width=0.5, zorder=3)

    for bar, val in zip(bars, capex_vals):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 5,
                f"${val}M", ha="center", fontsize=10, fontweight="bold")

    ax.annotate("Nearly 2x QoQ\nCapacity build-out", xy=(2, 290),
                xytext=(0.5, 320), fontsize=9, color=COHR_RED, fontweight="bold",
                arrowprops=dict(arrowstyle="->", color=COHR_RED, lw=1.5))

    ax.set_ylim(0, 380)
    ax.set_ylabel("CapEx ($M)", fontsize=10)
    ax.set_title("Capital Expenditure — Accelerating Capacity Investment", fontsize=12,
                 fontweight="bold", pad=10)
    ax.yaxis.grid(True, zorder=0)
    ax.set_axisbelow(True)

    ax.text(0.01, -0.12,
            "Source: Coherent Corp Earnings Releases, FY2026.",
            transform=ax.transAxes, fontsize=7, color=COHR_DARK)

    fig.tight_layout()
    save(fig, "cohr_chart7_capex.png")


# ═══════════════════════════════════════════════════════════════════════════════
# CHART 8: Leverage Ratio Improvement
# ═══════════════════════════════════════════════════════════════════════════════
def chart8_leverage():
    fig, ax = plt.subplots(figsize=(9, 4.5))

    lev_qtrs = ["Q3\nFY25","Q4\nFY25","Q1\nFY26","Q2\nFY26","Q3\nFY26"]
    lev_vals = [2.1, 1.9, 1.8, 1.7, 0.5]

    ax.plot(lev_qtrs, lev_vals, color=COHR_BLUE, lw=2.5, marker="o",
            markersize=8, zorder=3)
    ax.fill_between(range(len(lev_qtrs)), lev_vals, alpha=0.15, color=COHR_BLUE)

    for i, val in enumerate(lev_vals):
        ax.text(i, val + 0.08, f"{val:.1f}x", ha="center", fontsize=10,
                fontweight="bold", color=COHR_BLUE)

    ax.annotate("NVIDIA $2B equity\ninvestment impact", xy=(4, 0.5),
                xytext=(2.5, 0.8), fontsize=9, color=COHR_GREEN, fontweight="bold",
                arrowprops=dict(arrowstyle="->", color=COHR_GREEN, lw=1.5))

    ax.set_ylim(0, 2.8)
    ax.set_ylabel("Net Leverage Ratio (x)", fontsize=10)
    ax.set_title("Leverage Ratio — Rapid Deleveraging", fontsize=12, fontweight="bold", pad=10)
    ax.yaxis.grid(True, zorder=0)
    ax.set_axisbelow(True)

    ax.text(0.01, -0.12,
            "Source: Coherent Corp Earnings Releases, Q3 FY2025 – Q3 FY2026.",
            transform=ax.transAxes, fontsize=7, color=COHR_DARK)

    fig.tight_layout()
    save(fig, "cohr_chart8_leverage.png")


# ═══════════════════════════════════════════════════════════════════════════════
# CHART 9: YoY Comparison — Q3 FY25 vs Q3 FY26
# ═══════════════════════════════════════════════════════════════════════════════
def chart9_yoy():
    fig, axes = plt.subplots(1, 3, figsize=(13, 4.5))

    metrics = [("Revenue ($B)", q3_fy25_rev, q3_fy26_rev),
               ("Non-GAAP EPS ($)", q3_fy25_eps, q3_fy26_eps),
               ("Non-GAAP GM (%)", 38.1, 39.6)]
    bar_labels = ["Q3 FY25", "Q3 FY26"]

    for ax, (title, old, new) in zip(axes, metrics):
        x = [0, 1]
        bars = ax.bar(x, [old, new], color=[COHR_LIGHT, COHR_BLUE], width=0.5, zorder=3)
        ax.set_xticks(x)
        ax.set_xticklabels(bar_labels, fontsize=10)
        ax.set_title(title, fontsize=11, fontweight="bold")
        ax.yaxis.grid(True, zorder=0)
        ax.set_axisbelow(True)

        for bar, val in zip(bars, [old, new]):
            fmt = f"${val:.2f}" if "%" not in title else f"{val:.1f}%"
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.01 * max(old, new),
                    fmt, ha="center", fontsize=10, fontweight="bold")

        pct = (new - old) / old * 100
        ax.text(0.5, 0.92, f"+{pct:.0f}% YoY" if pct > 0 else f"{pct:.0f}% YoY",
                transform=ax.transAxes, ha="center", fontsize=9,
                color=COHR_GREEN if pct > 0 else COHR_RED, fontweight="bold")

    fig.text(0.01, 0.01,
             "Source: Coherent Corp Q3 FY2025 & Q3 FY2026 Earnings Releases.",
             fontsize=7, color=COHR_DARK)

    fig.tight_layout()
    save(fig, "cohr_chart9_yoy.png")


# ═══════════════════════════════════════════════════════════════════════════════
# CHART 10: Datacenter Revenue Ramp
# ═══════════════════════════════════════════════════════════════════════════════
def chart10_datacenter():
    fig, ax = plt.subplots(figsize=(9, 4.5))

    dc_qtrs = ["Q1\nFY25","Q2\nFY25","Q3\nFY25","Q4\nFY25","Q1\nFY26","Q2\nFY26","Q3\nFY26"]
    dc_rev  = [0.82, 0.90, 1.00, 1.05, 1.12, 1.20, 1.36]

    bars = ax.bar(dc_qtrs, dc_rev, color=[COHR_LIGHT]*6 + [COHR_BLUE], width=0.55, zorder=3)

    for bar, val in zip(bars, dc_rev):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.02,
                f"${val:.2f}B", ha="center", fontsize=9,
                fontweight="bold" if val == 1.36 else "normal")

    ax.annotate("75% of total revenue\n+36% YoY", xy=(6, 1.36),
                xytext=(4, 1.55), fontsize=9, color=COHR_GREEN, fontweight="bold",
                arrowprops=dict(arrowstyle="->", color=COHR_GREEN, lw=1.5))

    ax.set_ylim(0, 1.8)
    ax.set_ylabel("Revenue ($B)", fontsize=10)
    ax.set_title("Datacenter & Communications Revenue — AI-Driven Growth", fontsize=12,
                 fontweight="bold", pad=10)
    ax.yaxis.grid(True, zorder=0)
    ax.set_axisbelow(True)

    ax.text(0.01, -0.12,
            "Source: Coherent Corp Earnings Releases. Note: Segment reporting restructured from Q1 FY2026.",
            transform=ax.transAxes, fontsize=7, color=COHR_DARK)

    fig.tight_layout()
    save(fig, "cohr_chart10_datacenter.png")


# ═══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    os.makedirs(OUT_DIR, exist_ok=True)
    print("Generating COHR Q3 FY2026 charts...")
    chart1_revenue()
    chart2_eps()
    chart3_segments()
    chart4_margins()
    chart5_beat_miss()
    chart6_guidance()
    chart7_capex()
    chart8_leverage()
    chart9_yoy()
    chart10_datacenter()
    print("Done — 10 charts saved.")
