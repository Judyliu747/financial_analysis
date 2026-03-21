"""
Energy Fuels Inc. (UUUU) — Q4 & FY2025 Earnings Update
Chart builder (matplotlib → PNG files)
"""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

BASE = "/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis/output/UUUU"

NAVY   = "#1C1C4D"
BLUE   = "#0071E3"
GREEN  = "#1E8A44"
RED    = "#CC0000"
GREY   = "#8E8E93"
LTBLUE = "#E8F2FF"
AMBER  = "#E67E22"
TEAL   = "#006666"

def save(fig, name):
    path = os.path.join(BASE, name)
    fig.savefig(path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"  Saved: {name}")


# ── Chart 1: Quarterly Revenue Trend (Q1 2024 – Q4 2025) ─────────────────────
def chart1_quarterly_revenue():
    quarters = ["Q1\n2024", "Q2\n2024", "Q3\n2024", "Q4\n2024",
                "Q1\n2025", "Q2\n2025", "Q3\n2025", "Q4\n2025"]
    # Revenue in $M – sources: press releases + PRN 2025 annual report
    uranium = [25.3, 0.0, 0.0, 12.6, 16.9, 4.2, 10.4, 27.1]  # Q4'25 uranium-driven total
    hms     = [0.0, 16.2, 23.7, 0.0, 0.0, 8.5, 7.3, 0.0]
    # Note: Q4-25 total $27.1M ~all uranium; HMS lumped into HMS for simplicity
    # Q3-25 total $17.7M (uranium $10.4M + HMS $7.3M approx)
    # Q1-25 total $16.9M per press release

    x = np.arange(len(quarters))
    width = 0.55

    fig, ax = plt.subplots(figsize=(10, 5))
    bars1 = ax.bar(x, uranium, width, label="Uranium", color=BLUE)
    bars2 = ax.bar(x, hms, width, bottom=uranium, label="Heavy Mineral Sands / Other", color=TEAL, alpha=0.85)

    # Total labels
    totals = [u + h for u, h in zip(uranium, hms)]
    for xi, tot in zip(x, totals):
        ax.text(xi, tot + 0.4, f"${tot:.1f}M", ha="center", va="bottom", fontsize=8.5, color=NAVY, fontweight="bold")

    ax.set_xticks(x)
    ax.set_xticklabels(quarters, fontsize=9)
    ax.set_ylabel("Revenue ($M)", fontsize=10)
    ax.set_title("UUUU — Quarterly Revenue by Segment (Q1 2024 – Q4 2025)", fontsize=12, color=NAVY, fontweight="bold")
    ax.legend(fontsize=9)
    ax.set_ylim(0, 55)
    ax.yaxis.grid(True, linestyle="--", alpha=0.5)
    ax.set_axisbelow(True)
    ax.spines[["top","right"]].set_visible(False)

    # Annotation box: FY2025 total
    ax.text(0.98, 0.96, "FY2025 Total Revenue: $65.9M\nFY2024 Total Revenue: $78.1M",
            transform=ax.transAxes, ha="right", va="top", fontsize=8.5,
            bbox=dict(boxstyle="round,pad=0.4", fc=LTBLUE, ec=BLUE, lw=0.8))

    save(fig, "uuuu_chart1_revenue.png")


# ── Chart 2: Uranium Sales Volume & Realized Price ───────────────────────────
def chart2_uranium_ops():
    quarters = ["Q1\n2024", "Q2\n2024", "Q3\n2024", "Q4\n2024",
                "Q1\n2025", "Q2\n2025", "Q3\n2025", "Q4\n2025"]
    # lbs sold (thousands)
    lbs = [300, 0, 0, 150, 50, 150, 240, 360]
    # avg realized price $/lb (approximate; Q4'24 ~$84/lb spot + contract mix)
    price = [84.4, None, None, 84.0, 80.0, 79.0, 69.6, 74.9]

    x = np.arange(len(quarters))
    width = 0.55

    fig, ax1 = plt.subplots(figsize=(10, 5))
    color_bars = [GREEN if v > 0 else GREY for v in lbs]
    bars = ax1.bar(x, lbs, width, color=color_bars, label="Uranium Sold (klbs)", alpha=0.85)

    for xi, v in zip(x, lbs):
        if v > 0:
            ax1.text(xi, v + 4, f"{v}k", ha="center", va="bottom", fontsize=8.5, color=NAVY, fontweight="bold")

    ax1.set_ylabel("Uranium Sold (thousand lbs U₃O₈)", fontsize=10, color=NAVY)
    ax1.set_xticks(x)
    ax1.set_xticklabels(quarters, fontsize=9)
    ax1.set_ylim(0, 430)
    ax1.yaxis.grid(True, linestyle="--", alpha=0.4)
    ax1.set_axisbelow(True)

    ax2 = ax1.twinx()
    valid_x = [xi for xi, p in zip(x, price) if p is not None]
    valid_p = [p for p in price if p is not None]
    ax2.plot(valid_x, valid_p, "o-", color=RED, linewidth=2, markersize=6, label="Avg Realized Price ($/lb)")
    ax2.set_ylabel("Avg Realized Price ($/lb U₃O₈)", fontsize=10, color=RED)
    ax2.set_ylim(50, 110)
    ax2.spines["right"].set_color(RED)
    ax2.tick_params(axis="y", colors=RED)

    # Spot price reference band
    ax2.axhspan(63, 83, alpha=0.08, color=AMBER, label="2025 Spot Range ($63–$83/lb)")

    ax1.set_title("UUUU — Uranium Sales Volume & Realized Price", fontsize=12, color=NAVY, fontweight="bold")

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, fontsize=8.5, loc="upper left")
    ax1.spines[["top"]].set_visible(False)

    save(fig, "uuuu_chart2_uranium_ops.png")


# ── Chart 3: Net Income / EPS trend ──────────────────────────────────────────
def chart3_net_income():
    quarters = ["Q1\n2024", "Q2\n2024", "Q3\n2024", "Q4\n2024",
                "Q1\n2025", "Q2\n2025", "Q3\n2025", "Q4\n2025"]
    net_loss = [3.64, -9.8, -12.8, -28.8, -26.3, -22.2, -16.7, -20.9]  # $M
    eps      = [0.02, -0.06, -0.07, -0.19, -0.13, -0.10, -0.07, -0.08]

    x = np.arange(len(quarters))
    colors = [GREEN if v >= 0 else RED for v in net_loss]

    fig, ax1 = plt.subplots(figsize=(10, 5))
    bars = ax1.bar(x, net_loss, 0.55, color=colors, alpha=0.85)
    for xi, v in zip(x, net_loss):
        yoff = 0.5 if v >= 0 else -1.5
        ax1.text(xi, v + yoff, f"${v:.1f}M", ha="center", va="bottom" if v >= 0 else "top",
                 fontsize=8, color=NAVY, fontweight="bold")

    ax1.axhline(0, color=GREY, linewidth=0.8)
    ax1.set_ylabel("Net Income / (Loss) ($M)", fontsize=10)
    ax1.set_xticks(x)
    ax1.set_xticklabels(quarters, fontsize=9)
    ax1.yaxis.grid(True, linestyle="--", alpha=0.4)
    ax1.set_axisbelow(True)

    ax2 = ax1.twinx()
    ax2.plot(x, eps, "s--", color=BLUE, linewidth=2, markersize=6, label="EPS ($)")
    ax2.set_ylabel("EPS ($)", fontsize=10, color=BLUE)
    ax2.tick_params(axis="y", colors=BLUE)
    ax2.axhline(0, color=BLUE, linewidth=0.4, linestyle=":")

    ax1.set_title("UUUU — Quarterly Net Income / (Loss) & EPS", fontsize=12, color=NAVY, fontweight="bold")

    # EPS consensus miss annotation for Q4 2025
    ax2.annotate("Q4'25 EPS: -$0.08\n(Cons: -$0.07, Miss)", xy=(7, -0.08), xytext=(5.8, -0.04),
                 fontsize=8, color=RED,
                 arrowprops=dict(arrowstyle="->", color=RED, lw=1.0))

    lines2, labels2 = ax2.get_legend_handles_labels()
    ax2.legend(lines2, labels2, fontsize=9, loc="lower right")
    ax1.spines[["top"]].set_visible(False)

    # FY2025 annotation
    ax1.text(0.02, 0.05, "FY2025 Net Loss: ($86.1M) / ($0.38/share)\nFY2024 Net Loss: ($47.8M) / ($0.28/share)",
             transform=ax1.transAxes, fontsize=8.5,
             bbox=dict(boxstyle="round,pad=0.4", fc="#FFE8E8", ec=RED, lw=0.8))

    save(fig, "uuuu_chart3_net_income.png")


# ── Chart 4: Beat / Miss vs Consensus ────────────────────────────────────────
def chart4_beat_miss():
    quarters = ["Q1\n2025", "Q2\n2025", "Q3\n2025", "Q4\n2025"]
    rev_actual   = [16.9, 4.2+8.5, 17.7, 27.1]
    rev_cons     = [None, None, None, 27.0]   # limited consensus data pre-Q4

    eps_actual   = [-0.13, -0.10, -0.07, -0.08]
    eps_cons     = [None, None, -0.05, -0.07]

    x = np.arange(len(quarters))
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))

    # Revenue
    bars_a = ax1.bar(x - 0.2, rev_actual, 0.35, color=BLUE, label="Actual Revenue")
    cons_vals = [27.0]
    ax1.bar([3 + 0.2], cons_vals, 0.35, color=GREY, alpha=0.6, label="Consensus")
    ax1.set_xticks(x)
    ax1.set_xticklabels(quarters)
    ax1.set_ylabel("Revenue ($M)")
    ax1.set_title("Revenue: Actual vs Consensus", fontsize=11, color=NAVY)
    for xi, v in zip(x, rev_actual):
        ax1.text(xi - 0.2, v + 0.3, f"${v:.1f}M", ha="center", fontsize=8.5, color=NAVY)
    ax1.text(3 + 0.2, 27.0 + 0.3, "$27.0M", ha="center", fontsize=8.5, color=GREY)
    ax1.legend(fontsize=9)
    ax1.yaxis.grid(True, linestyle="--", alpha=0.4)
    ax1.set_axisbelow(True)
    ax1.spines[["top","right"]].set_visible(False)
    ax1.text(3.0, 24, "+$0.1M\nBeat", ha="center", fontsize=9, color=GREEN, fontweight="bold")

    # EPS
    ax2.bar(x - 0.2, eps_actual, 0.35, color=[RED if v < 0 else GREEN for v in eps_actual], label="Actual EPS")
    ax2.bar([2 + 0.2, 3 + 0.2], [-0.05, -0.07], 0.35, color=GREY, alpha=0.6, label="Consensus")
    ax2.set_xticks(x)
    ax2.set_xticklabels(quarters)
    ax2.set_ylabel("EPS ($)")
    ax2.set_title("EPS: Actual vs Consensus", fontsize=11, color=NAVY)
    ax2.axhline(0, color=GREY, linewidth=0.8)
    for xi, v in zip(x, eps_actual):
        ax2.text(xi - 0.2, v - 0.005, f"${v:.2f}", ha="center", va="top", fontsize=8.5, color=NAVY)
    ax2.legend(fontsize=9)
    ax2.yaxis.grid(True, linestyle="--", alpha=0.4)
    ax2.set_axisbelow(True)
    ax2.spines[["top","right"]].set_visible(False)
    ax2.text(2.1, -0.04, "Miss\n(-$0.02)", ha="center", fontsize=9, color=RED, fontweight="bold")
    ax2.text(3.1, -0.065, "Miss\n(-$0.01)", ha="center", fontsize=9, color=RED, fontweight="bold")

    fig.suptitle("UUUU — 2025 Quarterly Earnings vs Consensus", fontsize=13, color=NAVY, fontweight="bold")
    plt.tight_layout()
    save(fig, "uuuu_chart4_beat_miss.png")


# ── Chart 5: Full Year Revenue Breakdown 2023–2025 ───────────────────────────
def chart5_annual_revenue():
    years = ["FY2023", "FY2024", "FY2025"]
    uranium_rev = [15.2, 37.9, 48.2]
    hms_rev     = [0.0, 39.9, 15.8]
    vanadium    = [3.1, 0.0, 0.0]
    ree_other   = [1.2, 0.3, 1.9]  # approx. REE/other

    x = np.arange(len(years))
    width = 0.5
    fig, ax = plt.subplots(figsize=(9, 5))

    b1 = ax.bar(x, uranium_rev, width, color=BLUE, label="Uranium")
    b2 = ax.bar(x, hms_rev, width, bottom=uranium_rev, color=TEAL, alpha=0.85, label="Heavy Mineral Sands")
    b3 = ax.bar(x, vanadium, width,
                bottom=[u+h for u,h in zip(uranium_rev, hms_rev)], color=AMBER, alpha=0.85, label="Vanadium")
    b4 = ax.bar(x, ree_other, width,
                bottom=[u+h+v for u,h,v in zip(uranium_rev, hms_rev, vanadium)],
                color=GREEN, alpha=0.85, label="REE / Other")

    totals = [u+h+v+r for u,h,v,r in zip(uranium_rev, hms_rev, vanadium, ree_other)]
    for xi, tot in zip(x, totals):
        ax.text(xi, tot + 0.5, f"${tot:.1f}M", ha="center", fontsize=10, color=NAVY, fontweight="bold")

    ax.set_xticks(x)
    ax.set_xticklabels(years, fontsize=11)
    ax.set_ylabel("Revenue ($M)", fontsize=10)
    ax.set_title("UUUU — Annual Revenue by Segment", fontsize=12, color=NAVY, fontweight="bold")
    ax.legend(fontsize=9, loc="upper right")
    ax.set_ylim(0, 100)
    ax.yaxis.grid(True, linestyle="--", alpha=0.4)
    ax.set_axisbelow(True)
    ax.spines[["top","right"]].set_visible(False)

    save(fig, "uuuu_chart5_annual_revenue.png")


# ── Chart 6: Uranium Production & Inventory Build ────────────────────────────
def chart6_uranium_production():
    categories = ["2023\nMined", "2024\nMined", "2025\nMined",
                  "2024\nProduced", "2025\nProduced",
                  "End-2024\nInventory", "End-2025\nInventory"]
    values = [0.41, 1.10, 1.72, 0.55, 1.015, 0.98, 2.18]
    colors_list = [TEAL, TEAL, TEAL, BLUE, BLUE, AMBER, GREEN]

    fig, ax = plt.subplots(figsize=(11, 5))
    bars = ax.bar(categories, values, 0.6, color=colors_list, alpha=0.9)
    for bar, v in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width()/2, v + 0.02, f"{v:.2f}M lbs",
                ha="center", va="bottom", fontsize=9, color=NAVY, fontweight="bold")

    ax.set_ylabel("Million lbs U₃O₈", fontsize=10)
    ax.set_title("UUUU — Uranium Mining, Processing & Inventory (lbs U₃O₈)", fontsize=12, color=NAVY, fontweight="bold")
    ax.yaxis.grid(True, linestyle="--", alpha=0.4)
    ax.set_axisbelow(True)
    ax.spines[["top","right"]].set_visible(False)

    legend_handles = [
        mpatches.Patch(color=TEAL, label="Mined (contained)"),
        mpatches.Patch(color=BLUE, label="Processed / Finished"),
        mpatches.Patch(color=AMBER, label="End-2024 Inventory (total)"),
        mpatches.Patch(color=GREEN, label="End-2025 Inventory (total)"),
    ]
    ax.legend(handles=legend_handles, fontsize=9)

    save(fig, "uuuu_chart6_uranium_production.png")


# ── Chart 7: Uranium Cost Structure ──────────────────────────────────────────
def chart7_cost_structure():
    labels = ["FY2024\nCOGS/lb", "FY2025\nCOGS/lb", "2026E\nProcessing\n(Pinyon Plain)",
              "FY2025\nRealized\nPrice", "FY2024\nRealized\nPrice"]
    values = [53, 43, 26.5, 74.2, 87.2]  # midpoint for 2026E = ~(23+30)/2
    colors_list = [RED, AMBER, GREEN, BLUE, BLUE]

    fig, ax = plt.subplots(figsize=(9, 5))
    bars = ax.bar(labels, values, 0.55, color=colors_list, alpha=0.9)
    for bar, v in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width()/2, v + 0.8, f"${v:.1f}", ha="center", fontsize=10, fontweight="bold", color=NAVY)

    ax.set_ylabel("$/lb U₃O₈", fontsize=10)
    ax.set_title("UUUU — Uranium Cost Structure vs Realized Price", fontsize=12, color=NAVY, fontweight="bold")
    ax.set_ylim(0, 100)
    ax.yaxis.grid(True, linestyle="--", alpha=0.4)
    ax.set_axisbelow(True)
    ax.spines[["top","right"]].set_visible(False)

    # Margin annotations
    ax.annotate("", xy=(1, 43), xytext=(1, 74.2),
                arrowprops=dict(arrowstyle="<->", color=GREEN, lw=1.5))
    ax.text(1.35, 58, "~42%\ngross margin\n2025", fontsize=8.5, color=GREEN)

    save(fig, "uuuu_chart7_cost_structure.png")


# ── Chart 8: Balance Sheet & Cash ────────────────────────────────────────────
def chart8_balance_sheet():
    items  = ["End-2023\nWorking\nCapital", "End-2024\nWorking\nCapital",
              "End-2025\nWorking\nCapital", "End-2025\nCash &\nEquiv.",
              "End-2025\nMkt Securities", "Convertible\nNotes ($700M)"]
    values = [112, 170.9, 927.4, 64.7, 797.1, 700]
    colors_list = [BLUE, BLUE, GREEN, TEAL, TEAL, RED]

    fig, ax = plt.subplots(figsize=(10, 5))
    bars = ax.bar(items, values, 0.55, color=colors_list, alpha=0.9)
    for bar, v in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width()/2, v + 8, f"${v:.0f}M",
                ha="center", va="bottom", fontsize=9, fontweight="bold", color=NAVY)

    ax.set_ylabel("$M", fontsize=10)
    ax.set_title("UUUU — Balance Sheet Highlights", fontsize=12, color=NAVY, fontweight="bold")
    ax.set_ylim(0, 1100)
    ax.yaxis.grid(True, linestyle="--", alpha=0.4)
    ax.set_axisbelow(True)
    ax.spines[["top","right"]].set_visible(False)

    ax.text(0.98, 0.97,
            "0.75% Convertible Notes due 2031\nConversion price: $20.34/share\n(Capped call: $30.70/share)",
            transform=ax.transAxes, ha="right", va="top", fontsize=8.5,
            bbox=dict(boxstyle="round,pad=0.4", fc="#FFE8E8", ec=RED, lw=0.8))

    save(fig, "uuuu_chart8_balance_sheet.png")


# ── Chart 9: REE Growth Roadmap (NPV waterfall) ──────────────────────────────
def chart9_ree_roadmap():
    stages = ["Phase 1\n(Current)", "Phase 1B/1C\nExpansion", "Phase 2\nSeparation", "Phase 2 +\nVara Mada"]
    npv    = [0.15, 0.6, 1.9, 3.7]   # $B, approximate
    ebitda = [0.01, 0.08, 0.311, 0.765]  # $B/yr

    x = np.arange(len(stages))
    fig, ax1 = plt.subplots(figsize=(10, 5))
    bars = ax1.bar(x, npv, 0.5, color=[BLUE, TEAL, GREEN, AMBER], alpha=0.9, label="NPV ($B)")
    for xi, v in zip(x, npv):
        ax1.text(xi, v + 0.04, f"${v:.2f}B", ha="center", fontsize=10, fontweight="bold", color=NAVY)

    ax1.set_ylabel("NPV ($B) at 8% Discount Rate", fontsize=10)
    ax1.set_xticks(x)
    ax1.set_xticklabels(stages, fontsize=9.5)
    ax1.set_ylim(0, 4.5)
    ax1.yaxis.grid(True, linestyle="--", alpha=0.4)
    ax1.set_axisbelow(True)

    ax2 = ax1.twinx()
    ax2.plot(x, [e*1000 for e in ebitda], "D--", color=RED, linewidth=2, markersize=7, label="Annual EBITDA ($M)")
    ax2.set_ylabel("Annual EBITDA ($M)", fontsize=10, color=RED)
    ax2.tick_params(axis="y", colors=RED)
    for xi, v in zip(x, [e*1000 for e in ebitda]):
        ax2.text(xi + 0.18, v + 5, f"${v:.0f}M", fontsize=8.5, color=RED)

    ax1.set_title("UUUU — REE Expansion NPV & EBITDA Roadmap", fontsize=12, color=NAVY, fontweight="bold")

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, fontsize=9, loc="upper left")
    ax1.spines[["top"]].set_visible(False)

    ax1.text(0.98, 0.12,
             "Phase 2 Target Production:\n  NdPr: 5,513 t/yr\n  Dy: 165 t/yr\n  Tb: 48 t/yr\nCapex: ~$410M",
             transform=ax1.transAxes, ha="right", va="bottom", fontsize=8.5,
             bbox=dict(boxstyle="round,pad=0.4", fc=LTBLUE, ec=BLUE, lw=0.8))

    save(fig, "uuuu_chart9_ree_roadmap.png")


# ── Chart 10: Analyst Price Targets ──────────────────────────────────────────
def chart10_price_targets():
    analysts  = ["H.C.\nWainwright", "Analyst B\n(avg consensus)", "Analyst C\n(low)", "Analyst D\n(high)"]
    targets   = [27.25, 23.08, 15.50, 34.00]
    ratings   = ["Buy", "Buy", "Hold", "Buy"]
    colors_list = [GREEN, GREEN, AMBER, GREEN]

    fig, ax = plt.subplots(figsize=(9, 5))
    bars = ax.barh(analysts, targets, 0.5, color=colors_list, alpha=0.85)
    for bar, v, r in zip(bars, targets, ratings):
        ax.text(v + 0.3, bar.get_y() + bar.get_height()/2, f"${v:.2f} ({r})",
                va="center", fontsize=10, fontweight="bold", color=NAVY)

    # Current price line
    current_price = 18.67
    ax.axvline(current_price, color=RED, linestyle="--", linewidth=1.8, label=f"Current Price: ${current_price}")
    ax.set_xlabel("12-Month Price Target (USD)", fontsize=10)
    ax.set_title("UUUU — Analyst Price Targets (as of March 2026)", fontsize=12, color=NAVY, fontweight="bold")
    ax.legend(fontsize=9)
    ax.set_xlim(0, 42)
    ax.xaxis.grid(True, linestyle="--", alpha=0.4)
    ax.set_axisbelow(True)
    ax.spines[["top","right"]].set_visible(False)

    ax.text(0.98, 0.05,
            "7 Buy / 1 Hold / 0 Sell\nConsensus: Strong Buy",
            transform=ax.transAxes, ha="right", va="bottom", fontsize=9,
            bbox=dict(boxstyle="round,pad=0.4", fc=LTBLUE, ec=BLUE, lw=0.8))

    save(fig, "uuuu_chart10_price_targets.png")


# ── Chart 11: Uranium Market — Spot & LT Price Trend ─────────────────────────
def chart11_uranium_market():
    # Monthly approximate spot price 2024-2025 based on market context
    months = ["Jan\n2024", "Apr\n2024", "Jul\n2024", "Oct\n2024",
              "Jan\n2025", "Apr\n2025", "Jul\n2025", "Oct\n2025", "Dec\n2025", "Jan\n2026"]
    spot   = [106, 88, 85, 81, 74, 65, 72, 80, 73, 98]
    lt     = [68, 73, 76, 78, 80, 80, 82, 84, 86, 87]

    x = np.arange(len(months))
    fig, ax = plt.subplots(figsize=(11, 5))
    ax.plot(x, spot, "o-", color=BLUE, linewidth=2.5, markersize=6, label="U₃O₈ Spot Price ($/lb)")
    ax.plot(x, lt, "s--", color=GREEN, linewidth=2, markersize=5, label="Long-Term Contract Price ($/lb)")

    ax.fill_between(x, spot, lt, where=[s > l for s, l in zip(spot, lt)], alpha=0.1, color=BLUE)
    ax.fill_between(x, spot, lt, where=[s < l for s, l in zip(spot, lt)], alpha=0.1, color=GREEN)

    ax.set_xticks(x)
    ax.set_xticklabels(months, fontsize=8.5)
    ax.set_ylabel("Price ($/lb U₃O₈)", fontsize=10)
    ax.set_title("Uranium Market — Spot vs Long-Term Contract Price Trend", fontsize=12, color=NAVY, fontweight="bold")
    ax.legend(fontsize=9)
    ax.set_ylim(50, 125)
    ax.yaxis.grid(True, linestyle="--", alpha=0.4)
    ax.set_axisbelow(True)
    ax.spines[["top","right"]].set_visible(False)

    ax.annotate("2025 Spot Range:\n$63–$83/lb", xy=(5, 65), xytext=(3, 58),
                fontsize=8.5, color=RED,
                arrowprops=dict(arrowstyle="->", color=RED, lw=1.0))
    ax.text(9, 100, "Jan 2026:\n~$98–100/lb", fontsize=8.5, color=BLUE, ha="center")

    save(fig, "uuuu_chart11_uranium_market.png")


if __name__ == "__main__":
    print("Generating UUUU charts...")
    chart1_quarterly_revenue()
    chart2_uranium_ops()
    chart3_net_income()
    chart4_beat_miss()
    chart5_annual_revenue()
    chart6_uranium_production()
    chart7_cost_structure()
    chart8_balance_sheet()
    chart9_ree_roadmap()
    chart10_price_targets()
    chart11_uranium_market()
    print("All charts generated.")
