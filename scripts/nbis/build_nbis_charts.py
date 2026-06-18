#!/usr/bin/env python3
"""Charts for Nebius Group (NBIS) Q1 2026 Earnings Update.
All output PNGs go to output/NBIS/ at 150 dpi.
Labels kept in English so the same PNGs are shared by EN + CN reports.
"""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
import numpy as np

OUT = "/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis/output/NBIS"

plt.rcParams.update({
    "font.family":        "sans-serif",
    "font.sans-serif":    ["Arial Unicode MS", "STHeiti", "Hiragino Sans GB", "Times New Roman"],
    "axes.unicode_minus": False,
    "axes.edgecolor":     "#444444",
    "axes.linewidth":     0.8,
    "axes.titlesize":     13,
    "axes.titleweight":   "bold",
    "axes.labelsize":     11,
    "xtick.labelsize":    10,
    "ytick.labelsize":    10,
    "figure.dpi":         150,
    "savefig.dpi":        150,
})

# Brand palette
NB_DARK   = "#0B2545"   # deep navy
NB_BLUE   = "#1F6FEB"   # bright blue
NB_TEAL   = "#13B5B1"   # teal
NB_GREEN  = "#2BA84A"
NB_RED    = "#D1495B"
NB_GREY   = "#9AA5B1"
NB_GOLD   = "#E8A33D"

QTRS = ["Q1'25", "Q2'25", "Q3'25", "Q4'25", "Q1'26"]

def save(fig, name):
    path = os.path.join(OUT, name)
    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("saved", path)


# 1. Group revenue progression -------------------------------------------------
def chart_revenue():
    rev = [50.9, 105.1, 146.1, 228.0, 399.0]
    fig, ax = plt.subplots(figsize=(7.2, 4.2))
    colors = [NB_BLUE]*4 + [NB_DARK]
    bars = ax.bar(QTRS, rev, color=colors, width=0.62)
    for b, v in zip(bars, rev):
        ax.text(b.get_x()+b.get_width()/2, v+8, f"${v:,.0f}M",
                ha="center", va="bottom", fontsize=10, fontweight="bold")
    ax.set_title("Figure 1. Group Revenue by Quarter (US$M)")
    ax.set_ylabel("Revenue (US$M)")
    ax.set_ylim(0, 460)
    ax.annotate("+684% YoY\n+75% QoQ", xy=(4, 399), xytext=(2.6, 360),
                fontsize=9.5, color=NB_DARK, fontweight="bold",
                ha="center")
    ax.grid(axis="y", ls=":", alpha=0.4)
    save(fig, "nbis_chart1_revenue.png")


# 2. Core AI Cloud ARR progression ---------------------------------------------
def chart_arr():
    arr = [249, 430, 551, 1250, 1920]
    fig, ax = plt.subplots(figsize=(7.2, 4.2))
    bars = ax.bar(QTRS, arr, color=[NB_TEAL]*4+[NB_DARK], width=0.62)
    for b, v in zip(bars, arr):
        ax.text(b.get_x()+b.get_width()/2, v+30, f"${v/1000:,.2f}B" if v>=1000 else f"${v}M",
                ha="center", va="bottom", fontsize=10, fontweight="bold")
    ax.set_title("Figure 2. Core AI Cloud ARR (Annualized Run-Rate, US$M)")
    ax.set_ylabel("ARR (US$M)")
    ax.set_ylim(0, 2250)
    ax.annotate("+54% QoQ", xy=(4, 1920), xytext=(3.0, 2050),
                fontsize=10, color=NB_DARK, fontweight="bold", ha="center")
    ax.grid(axis="y", ls=":", alpha=0.4)
    save(fig, "nbis_chart2_arr.png")


# 3. Group Adjusted EBITDA & margin --------------------------------------------
def chart_ebitda():
    ebitda = [-53.7, -28.0, 3.0, 15.0, 129.5]   # Q2/Q3'25 interpolated estimate
    margin = [-106, -27, 2, 7, 32]
    fig, ax = plt.subplots(figsize=(7.2, 4.2))
    colors = [NB_RED if v < 0 else NB_GREEN for v in ebitda]
    bars = ax.bar(QTRS, ebitda, color=colors, width=0.58)
    for b, v in zip(bars, ebitda):
        off = 6 if v >= 0 else -6
        ax.text(b.get_x()+b.get_width()/2, v+off, f"${v:,.0f}M",
                ha="center", va="bottom" if v>=0 else "top",
                fontsize=9.5, fontweight="bold")
    ax.axhline(0, color="#444", lw=0.9)
    ax.set_title("Figure 3. Group Adjusted EBITDA & Margin")
    ax.set_ylabel("Adj. EBITDA (US$M)")
    ax.set_ylim(-90, 175)
    ax2 = ax.twinx()
    ax2.plot(QTRS, margin, color=NB_DARK, marker="o", lw=2, label="Adj. EBITDA margin")
    ax2.set_ylabel("Margin (%)")
    ax2.set_ylim(-130, 60)
    for x, m in zip(QTRS, margin):
        ax2.text(x, m+6, f"{m}%", ha="center", fontsize=8.5, color=NB_DARK)
    ax.text(0.01, -0.22, "Note: Q2'25 / Q3'25 Adj. EBITDA interpolated for trend illustration.",
            transform=ax.transAxes, fontsize=7.5, color="#666")
    ax.grid(axis="y", ls=":", alpha=0.35)
    save(fig, "nbis_chart3_ebitda.png")


# 4. Beat / Miss vs consensus --------------------------------------------------
def chart_beatmiss():
    fig, ax = plt.subplots(figsize=(7.2, 4.0))
    labels = ["Revenue\n(US$M)", "Adj. EBITDA\n(US$M)", "Adj. EPS\n(US$)"]
    cons = [391.6, 95.0, -0.78]
    act  = [399.0, 129.5, -0.39]
    x = np.arange(len(labels))
    w = 0.36
    b1 = ax.bar(x-w/2, cons, w, label="Consensus", color=NB_GREY)
    b2 = ax.bar(x+w/2, act,  w, label="Actual",    color=NB_BLUE)
    ax.set_title("Figure 4. Q1 2026 Actual vs. Consensus (Beat)")
    ax.set_xticks(x); ax.set_xticklabels(labels)
    ax.axhline(0, color="#444", lw=0.8)
    ax.legend(loc="upper right", frameon=False)
    for bars, vals in [(b1, cons), (b2, act)]:
        for b, v in zip(bars, vals):
            ax.text(b.get_x()+b.get_width()/2, v + (3 if v>=0 else -3),
                    f"{v:,.2f}" if abs(v) < 5 else f"{v:,.0f}",
                    ha="center", va="bottom" if v>=0 else "top", fontsize=8.5)
    ax.grid(axis="y", ls=":", alpha=0.35)
    save(fig, "nbis_chart4_beatmiss.png")


# 5. Contracted power capacity -------------------------------------------------
def chart_power():
    fig, ax = plt.subplots(figsize=(7.2, 4.0))
    cats = ["Q4'25\ncontracted", "Q1'26\ncontracted", "YE-2026\ntarget"]
    vals = [2.0, 3.5, 4.0]
    bars = ax.bar(cats, vals, color=[NB_GREY, NB_BLUE, NB_DARK], width=0.6)
    for b, v in zip(bars, vals):
        ax.text(b.get_x()+b.get_width()/2, v+0.06, f"{v:.1f} GW",
                ha="center", fontweight="bold", fontsize=10)
    ax.set_title("Figure 5. Contracted Power Capacity (GW)")
    ax.set_ylabel("Gigawatts (GW)")
    ax.set_ylim(0, 4.6)
    ax.grid(axis="y", ls=":", alpha=0.4)
    ax.text(0.01, -0.20, ">75% of capacity owned; new 1.2 GW owned AI factory in Pennsylvania.",
            transform=ax.transAxes, fontsize=8, color="#555")
    save(fig, "nbis_chart5_power.png")


# 6. Liquidity & capital raised ------------------------------------------------
def chart_liquidity():
    fig, ax = plt.subplots(figsize=(7.2, 4.0))
    cats = ["Cash\nDec-25", "Conv.\nnotes", "NVIDIA\nequity", "Op. cash\n& other", "Cash\nMar-26"]
    # waterfall
    start = 3678.1
    conv = 4337.5
    nvda = 2000.0
    other = 9298.2 - start - conv - nvda
    fig, ax = plt.subplots(figsize=(7.4, 4.2))
    running = start
    xs = range(5)
    # bar 1 baseline
    ax.bar(0, start, color=NB_DARK, width=0.6)
    ax.text(0, start+120, f"${start/1000:.1f}B", ha="center", fontsize=9, fontweight="bold")
    bottoms = [running]
    for i,(lab,delta,col) in enumerate([("Conv.\nnotes",conv,NB_TEAL),
                                         ("NVIDIA\nequity",nvda,NB_BLUE),
                                         ("Op. cash\n& other",other,NB_GREEN)], start=1):
        ax.bar(i, delta, bottom=running, color=col, width=0.6)
        ax.text(i, running+delta+120, f"+${delta/1000:.1f}B", ha="center", fontsize=9, fontweight="bold")
        running += delta
    ax.bar(4, 9298.2, color=NB_DARK, width=0.6)
    ax.text(4, 9298.2+120, f"${9298.2/1000:.1f}B", ha="center", fontsize=9, fontweight="bold")
    ax.set_xticks(range(5))
    ax.set_xticklabels(["Cash\nDec-25","Conv.\nnotes","NVIDIA\nequity","Op. cash\n& other","Cash\nMar-26"])
    ax.set_title("Figure 6. Liquidity Bridge: Cash Build to $9.3B (US$M)")
    ax.set_ylabel("US$M")
    ax.set_ylim(0, 10800)
    ax.grid(axis="y", ls=":", alpha=0.35)
    save(fig, "nbis_chart6_liquidity.png")


# 7. CapEx guidance raise ------------------------------------------------------
def chart_capex():
    fig, ax = plt.subplots(figsize=(7.0, 4.0))
    cats = ["Prior\nguidance", "Raised\nguidance"]
    lows = [16, 20]; highs = [20, 25]
    x = np.arange(2)
    for i in range(2):
        ax.bar(i, highs[i]-lows[i], bottom=lows[i], width=0.5,
               color=[NB_GREY, NB_DARK][i])
        ax.text(i, highs[i]+0.4, f"${lows[i]}–{highs[i]}B", ha="center",
                fontsize=11, fontweight="bold")
    ax.set_xticks(x); ax.set_xticklabels(cats)
    ax.set_title("Figure 7. FY2026 CapEx Guidance Raised")
    ax.set_ylabel("CapEx (US$B)")
    ax.set_ylim(0, 28)
    ax.grid(axis="y", ls=":", alpha=0.35)
    ax.text(0.01, -0.18, ">90% of current capex already secured by cash & contractual commitments.",
            transform=ax.transAxes, fontsize=8, color="#555")
    save(fig, "nbis_chart7_capex.png")


# 8. FY2026 guidance dashboard -------------------------------------------------
def chart_guidance():
    fig, ax = plt.subplots(figsize=(7.2, 4.0))
    metrics = ["ARR\nexit-2026", "Group revenue", "Adj. EBITDA\nmargin"]
    # show ranges as bars (revenue/ARR in $B, margin separate handled via text)
    cats = ["ARR exit-2026\n($B)", "Group revenue\n($B)"]
    lows = [7, 3.0]; highs = [9, 3.4]
    x = np.arange(2)
    for i in range(2):
        ax.bar(i, highs[i]-lows[i], bottom=lows[i], width=0.5, color=[NB_TEAL, NB_BLUE][i])
        ax.text(i, highs[i]+0.2, f"${lows[i]}–{highs[i]}B", ha="center", fontsize=11, fontweight="bold")
    ax.set_xticks(x); ax.set_xticklabels(cats)
    ax.set_ylim(0, 10.5)
    ax.set_ylabel("US$B")
    ax.set_title("Figure 8. FY2026 Guidance (Reiterated)")
    ax.grid(axis="y", ls=":", alpha=0.35)
    ax.text(0.5, 0.78, "Group Adj. EBITDA margin target: ~40%",
            transform=ax.transAxes, ha="center", fontsize=10.5,
            fontweight="bold", color=NB_DARK,
            bbox=dict(boxstyle="round,pad=0.4", fc="#EAF2FF", ec=NB_BLUE))
    save(fig, "nbis_chart8_guidance.png")


# 9. Nebius AI core EBITDA margin expansion ------------------------------------
def chart_coremargin():
    fig, ax = plt.subplots(figsize=(7.0, 4.0))
    cats = ["Q4'25", "Q1'26"]
    vals = [24, 45]
    bars = ax.bar(cats, vals, color=[NB_GREY, NB_GREEN], width=0.5)
    for b, v in zip(bars, vals):
        ax.text(b.get_x()+b.get_width()/2, v+1, f"{v}%", ha="center",
                fontsize=12, fontweight="bold")
    ax.set_title("Figure 9. Nebius AI Cloud Adj. EBITDA Margin Expansion")
    ax.set_ylabel("Margin (%)")
    ax.set_ylim(0, 55)
    ax.grid(axis="y", ls=":", alpha=0.35)
    ax.annotate("+21 pts", xy=(1, 45), xytext=(0.5, 38), ha="center",
                fontsize=10, fontweight="bold", color=NB_GREEN)
    save(fig, "nbis_chart9_coremargin.png")


# 10. Revenue mix --------------------------------------------------------------
def chart_mix():
    fig, ax = plt.subplots(figsize=(5.6, 4.4))
    sizes = [98, 2]
    labels = ["Nebius AI Cloud\n$390M (98%)", "Other\n(Avride, Toloka,\nTripleTen) $9M (2%)"]
    ax.pie(sizes, labels=labels, colors=[NB_BLUE, NB_GOLD],
           autopct="", startangle=90, wedgeprops=dict(width=0.42, edgecolor="white"),
           textprops=dict(fontsize=9.5))
    ax.set_title("Figure 10. Q1 2026 Revenue Mix")
    save(fig, "nbis_chart10_mix.png")


if __name__ == "__main__":
    chart_revenue()
    chart_arr()
    chart_ebitda()
    chart_beatmiss()
    chart_power()
    chart_liquidity()
    chart_capex()
    chart_guidance()
    chart_coremargin()
    chart_mix()
    print("All NBIS charts generated.")
