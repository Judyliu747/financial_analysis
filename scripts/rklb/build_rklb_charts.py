"""
Build RKLB (Rocket Lab USA) Q4 FY2025 Earnings Update Charts
Generates 10 PNG charts to output/RKLB/
"""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

OUT = "/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis/output/RKLB/"
os.makedirs(OUT, exist_ok=True)

# Color palette
BLUE  = "#1B3A6B"
LBLUE = "#4A90D9"
TEAL  = "#2E8B6A"
GRAY  = "#8C8C8C"
RED   = "#C0392B"
GREEN = "#27AE60"
GOLD  = "#F39C12"
BG    = "#F8F9FA"

plt.rcParams.update({
    "font.family": "Times New Roman",
    "font.size": 10,
    "axes.spines.top": False,
    "axes.spines.right": False,
})

# ─── DATA ─────────────────────────────────────────────────────────────────────

quarters_8 = ["Q1'24", "Q2'24", "Q3'24", "Q4'24", "Q1'25", "Q2'25", "Q3'25", "Q4'25"]
revenue_8  = [92.8, 106.0, 105.0, 132.4, 123.0, 144.0, 155.0, 179.7]

quarters_q4 = ["Q4'23", "Q4'24", "Q4'25"]
revenue_q4  = [59.9, 132.4, 179.7]

# Segment Q4 2025
seg_labels = ["Launch Services", "Space Systems"]
seg_q4     = [75.9, 103.8]

# Annual segment comparison
annual_years  = ["FY2024", "FY2025"]
annual_launch = [125.4, 199.0]
annual_space  = [310.8, 402.8]
annual_total  = [436.2, 601.8]

# Gross margin quarterly
gm_quarters = ["Q1'24", "Q2'24", "Q3'24", "Q4'24", "Q1'25", "Q2'25", "Q3'25", "Q4'25"]
gm_gaap     = [20.0, 24.0, 26.0, 29.0, 30.0, 33.0, 37.0, 38.0]
gm_nongaap  = [28.0, 32.0, 34.0, 37.0, 38.0, 41.0, 41.6, 44.0]

# Adj. EBITDA quarterly (loss = negative)
ebitda_qtrs = ["Q1'25", "Q2'25", "Q3'25", "Q4'25"]
ebitda_vals = [-33.0, -29.0, -26.3, -17.4]

# Launch cadence (Electron missions per quarter)
launch_qtrs = ["Q1'24", "Q2'24", "Q3'24", "Q4'24", "Q1'25", "Q2'25", "Q3'25", "Q4'25"]
launches    = [3, 4, 5, 4, 5, 5, 4, 7]

# Annual revenue for trend chart
annual_rev_years = ["2021", "2022", "2023", "2024", "2025"]
annual_rev_vals  = [29.5, 211.0, 245.0, 436.2, 601.8]

# YoY growth by quarter
yoy_growth = [69.0, 71.0, 55.0, 121.0, 33.0, 36.0, 48.0, 36.0]


# ─── CHART 1: Quarterly Revenue (8 quarters) ──────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5), facecolor=BG)
ax.set_facecolor(BG)
colors = [GRAY if q != "Q4'25" else BLUE for q in quarters_8]
bars = ax.bar(quarters_8, revenue_8, color=colors, edgecolor="white", width=0.6, zorder=3)
for bar, val in zip(bars, revenue_8):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 2,
            f"${val:.0f}M", ha="center", va="bottom", fontsize=9,
            fontweight="bold" if val == 179.7 else "normal",
            color=BLUE if val == 179.7 else "black")
ax.set_title("Quarterly Revenue (Q1 2024 – Q4 2025)", fontsize=13, fontweight="bold", color=BLUE, pad=10)
ax.set_ylabel("Revenue (US$M)", fontsize=10)
ax.set_ylim(0, 220)
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"${x:.0f}M"))
ax.axhline(179.7, color=BLUE, linewidth=0.8, linestyle="--", alpha=0.4, zorder=2)
ax.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
ax.text(0.98, 0.97, "Record Q4", transform=ax.transAxes, ha="right", va="top",
        fontsize=8, color=BLUE, style="italic")
plt.tight_layout()
plt.savefig(OUT + "rklb_chart1_quarterly_revenue.png", dpi=150, bbox_inches="tight")
plt.close()
print("Chart 1 saved.")

# ─── CHART 2: YoY Revenue Growth ──────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5), facecolor=BG)
ax.set_facecolor(BG)
bar_colors = [GREEN if g > 0 else RED for g in yoy_growth]
bars = ax.bar(quarters_8, yoy_growth, color=bar_colors, edgecolor="white", width=0.6, zorder=3)
for bar, val in zip(bars, yoy_growth):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
            f"{val:.0f}%", ha="center", va="bottom", fontsize=9)
ax.set_title("Year-over-Year Revenue Growth (%)", fontsize=13, fontweight="bold", color=BLUE, pad=10)
ax.set_ylabel("YoY Growth (%)", fontsize=10)
ax.set_ylim(0, 150)
ax.axhline(38, color=BLUE, linewidth=1.2, linestyle="--", alpha=0.7, zorder=2)
ax.text(7.4, 40, "FY2025\nAvg 38%", fontsize=7.5, color=BLUE, ha="right")
ax.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
plt.tight_layout()
plt.savefig(OUT + "rklb_chart2_yoy_growth.png", dpi=150, bbox_inches="tight")
plt.close()
print("Chart 2 saved.")

# ─── CHART 3: Q4 Segment Revenue Breakdown ───────────────────────────────────
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 5), facecolor=BG)
fig.patch.set_facecolor(BG)

# Pie chart
ax1.set_facecolor(BG)
wedge_colors = [LBLUE, BLUE]
wedges, texts, autotexts = ax1.pie(
    seg_q4, labels=seg_labels, autopct="%1.0f%%",
    colors=wedge_colors, startangle=90,
    wedgeprops=dict(edgecolor="white", linewidth=2),
    textprops=dict(fontsize=10)
)
for at in autotexts:
    at.set_fontsize(10)
    at.set_fontweight("bold")
    at.set_color("white")
ax1.set_title("Q4 2025 Revenue Mix\n($179.7M Total)", fontsize=11, fontweight="bold", color=BLUE)

# Bar chart by segment FY2024 vs FY2025
x = np.arange(2)
width = 0.35
ax2.set_facecolor(BG)
b1 = ax2.bar(x - width/2, annual_launch, width, label="Launch Services", color=LBLUE, edgecolor="white")
b2 = ax2.bar(x + width/2, annual_space, width, label="Space Systems", color=BLUE, edgecolor="white")
ax2.set_xticks(x); ax2.set_xticklabels(["FY2024", "FY2025"])
ax2.set_ylabel("Revenue (US$M)"); ax2.set_ylim(0, 500)
ax2.set_title("Annual Segment Revenue\n(FY2024 vs FY2025)", fontsize=11, fontweight="bold", color=BLUE)
for bar in list(b1) + list(b2):
    ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 5,
             f"${bar.get_height():.0f}M", ha="center", va="bottom", fontsize=8.5)
ax2.legend(fontsize=9)
ax2.grid(axis="y", linestyle="--", alpha=0.4)
ax2.spines["top"].set_visible(False); ax2.spines["right"].set_visible(False)

plt.suptitle("Revenue Segment Analysis", fontsize=13, fontweight="bold", color=BLUE, y=1.01)
plt.tight_layout()
plt.savefig(OUT + "rklb_chart3_segment_revenue.png", dpi=150, bbox_inches="tight")
plt.close()
print("Chart 3 saved.")

# ─── CHART 4: Beat / Miss Summary Table ──────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 4.5), facecolor=BG)
ax.set_facecolor(BG)
ax.axis("off")

table_data = [
    ["Metric", "Consensus / Guidance", "Actual", "vs. Estimate", "Result"],
    ["Revenue",       "~$177M",              "$179.7M",   "+1.6%",         "✔ BEAT"],
    ["GAAP EPS",      "($0.09)–($0.10)",     "($0.09)",   "In-line",       "✔ IN-LINE"],
    ["Adj. EBITDA",   "($23M)–($29M) loss",  "($17.4M)",  "+$5.6–11.6M",  "✔ BEAT"],
    ["Gross Margin",  "~35–37% GAAP",        "38% GAAP",  "+100–300bps",  "✔ BEAT"],
]
col_widths = [0.18, 0.24, 0.16, 0.20, 0.14]
row_colors = [
    [BLUE]*5,
    [BG, BG, BG, BG, "#D5F5E3"],
    [BG, BG, BG, BG, "#EBF5FB"],
    [BG, BG, BG, BG, "#D5F5E3"],
    [BG, BG, BG, BG, "#D5F5E3"],
]
tbl = ax.table(
    cellText=table_data, cellLoc="center", loc="center",
    bbox=[0.0, 0.0, 1.0, 1.0], colWidths=col_widths
)
tbl.auto_set_font_size(False)
tbl.set_fontsize(10)
for (r, c), cell in tbl.get_celld().items():
    cell.set_edgecolor("#CCCCCC")
    if r == 0:
        cell.set_facecolor(BLUE)
        cell.set_text_props(color="white", fontweight="bold")
        cell.set_height(0.22)
    else:
        cell.set_facecolor(row_colors[r][c])
        cell.set_height(0.18)
        if c == 4:
            cell.set_text_props(color=GREEN if "✔" in table_data[r][c] else RED, fontweight="bold")

ax.set_title("Q4 2025 Beat / Miss Summary", fontsize=13, fontweight="bold", color=BLUE, pad=15)
plt.tight_layout()
plt.savefig(OUT + "rklb_chart4_beat_miss.png", dpi=150, bbox_inches="tight")
plt.close()
print("Chart 4 saved.")

# ─── CHART 5: Gross Margin Trend ─────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5), facecolor=BG)
ax.set_facecolor(BG)
ax.plot(gm_quarters, gm_gaap, color=BLUE, marker="o", linewidth=2, markersize=7, label="GAAP Gross Margin", zorder=3)
ax.plot(gm_quarters, gm_nongaap, color=TEAL, marker="s", linewidth=2, markersize=7, linestyle="--",
        label="Non-GAAP Gross Margin", zorder=3)
for i, (g, ng) in enumerate(zip(gm_gaap, gm_nongaap)):
    ax.annotate(f"{g:.0f}%", (gm_quarters[i], g), textcoords="offset points",
                xytext=(0, 7), ha="center", fontsize=8.5, color=BLUE)
    ax.annotate(f"{ng:.0f}%", (gm_quarters[i], ng), textcoords="offset points",
                xytext=(0, -14), ha="center", fontsize=8.5, color=TEAL)
ax.fill_between(range(len(gm_quarters)), gm_gaap, alpha=0.1, color=BLUE)
ax.set_xticks(range(len(gm_quarters))); ax.set_xticklabels(gm_quarters)
ax.set_ylabel("Gross Margin (%)", fontsize=10)
ax.set_ylim(10, 55)
ax.set_title("Gross Margin Expansion (Q1 2024 – Q4 2025)", fontsize=13, fontweight="bold", color=BLUE, pad=10)
ax.legend(loc="lower right", fontsize=9)
ax.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
plt.tight_layout()
plt.savefig(OUT + "rklb_chart5_gross_margin.png", dpi=150, bbox_inches="tight")
plt.close()
print("Chart 5 saved.")

# ─── CHART 6: Adjusted EBITDA Trend ──────────────────────────────────────────
fig, ax = plt.subplots(figsize=(8, 5), facecolor=BG)
ax.set_facecolor(BG)
bar_cols = [RED if v < 0 else GREEN for v in ebitda_vals]
bars = ax.bar(ebitda_qtrs, ebitda_vals, color=bar_cols, edgecolor="white", width=0.5, zorder=3)
for bar, val in zip(bars, ebitda_vals):
    ax.text(bar.get_x() + bar.get_width()/2, val - 1.5,
            f"(${ abs(val):.1f}M)", ha="center", va="top", fontsize=9.5, color="white", fontweight="bold")
ax.axhline(0, color="black", linewidth=0.8)
ax.set_title("Adjusted EBITDA — 2025 Quarterly Trend", fontsize=13, fontweight="bold", color=BLUE, pad=10)
ax.set_ylabel("Adj. EBITDA (US$M)", fontsize=10)
ax.set_ylim(-45, 5)
ax.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
# Guidance band for Q4
ax.axhspan(-29, -23, alpha=0.15, color=GOLD, zorder=1)
ax.text(3, -26, "Q4 guidance\n($23M)–($29M)", ha="center", fontsize=8, color=GOLD)
plt.tight_layout()
plt.savefig(OUT + "rklb_chart6_adj_ebitda.png", dpi=150, bbox_inches="tight")
plt.close()
print("Chart 6 saved.")

# ─── CHART 7: Annual Revenue Trend ───────────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5), facecolor=BG)
ax.set_facecolor(BG)
bar_cols = [GRAY, GRAY, GRAY, GRAY, BLUE]
bars = ax.bar(annual_rev_years, annual_rev_vals, color=bar_cols, edgecolor="white", width=0.55, zorder=3)
for bar, val in zip(bars, annual_rev_vals):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 8,
            f"${val:.0f}M", ha="center", va="bottom", fontsize=9.5,
            fontweight="bold" if val == 601.8 else "normal")
ax.plot(annual_rev_years, annual_rev_vals, color=LBLUE, marker="o", linewidth=1.5, markersize=5, zorder=4)
ax.set_title("Annual Revenue Trend (2021 – 2025)", fontsize=13, fontweight="bold", color=BLUE, pad=10)
ax.set_ylabel("Revenue (US$M)", fontsize=10)
ax.set_ylim(0, 750)
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"${x:.0f}M"))
ax.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
ax.text(4, 630, "+38% YoY", ha="center", fontsize=9, color=GREEN, fontweight="bold")
plt.tight_layout()
plt.savefig(OUT + "rklb_chart7_annual_revenue.png", dpi=150, bbox_inches="tight")
plt.close()
print("Chart 7 saved.")

# ─── CHART 8: Q4 YoY Revenue Comparison ─────────────────────────────────────
fig, ax = plt.subplots(figsize=(7, 5), facecolor=BG)
ax.set_facecolor(BG)
bar_cols_q4 = [GRAY, GRAY, BLUE]
bars = ax.bar(quarters_q4, revenue_q4, color=bar_cols_q4, edgecolor="white", width=0.5, zorder=3)
for bar, val in zip(bars, revenue_q4):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 2,
            f"${val:.1f}M", ha="center", va="bottom", fontsize=10,
            fontweight="bold" if val == 179.7 else "normal")
ax.set_title("Q4 Revenue: Year-over-Year Comparison", fontsize=12, fontweight="bold", color=BLUE, pad=10)
ax.set_ylabel("Revenue (US$M)", fontsize=10)
ax.set_ylim(0, 230)
# Annotation arrows
ax.annotate("", xy=(1, 132.4), xytext=(0, 59.9),
            arrowprops=dict(arrowstyle="->", color=BLUE, lw=1.2))
ax.text(0.5, 95, "+121%", ha="center", fontsize=9, color=BLUE)
ax.annotate("", xy=(2, 179.7), xytext=(1, 132.4),
            arrowprops=dict(arrowstyle="->", color=GREEN, lw=1.2))
ax.text(1.5, 155, "+36%", ha="center", fontsize=9, color=GREEN)
ax.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
plt.tight_layout()
plt.savefig(OUT + "rklb_chart8_q4_yoy.png", dpi=150, bbox_inches="tight")
plt.close()
print("Chart 8 saved.")

# ─── CHART 9: Launch Cadence ─────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5), facecolor=BG)
ax.set_facecolor(BG)
bar_cols_l = [LBLUE if q.endswith("'24") else BLUE for q in launch_qtrs]
bars = ax.bar(launch_qtrs, launches, color=bar_cols_l, edgecolor="white", width=0.6, zorder=3)
for bar, val in zip(bars, launches):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.1,
            str(val), ha="center", va="bottom", fontsize=10, fontweight="bold")
ax.set_title("Electron Launch Cadence (Q1 2024 – Q4 2025)", fontsize=13, fontweight="bold", color=BLUE, pad=10)
ax.set_ylabel("Number of Launches", fontsize=10)
ax.set_ylim(0, 9)
ax.yaxis.set_major_locator(plt.MaxNLocator(integer=True))
ax.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
p1 = mpatches.Patch(color=LBLUE, label="FY2024 (16 total)")
p2 = mpatches.Patch(color=BLUE, label="FY2025 (21 total, record; Q4 incl. 1 HASTE)")
ax.legend(handles=[p1, p2], fontsize=9)
plt.tight_layout()
plt.savefig(OUT + "rklb_chart9_launch_cadence.png", dpi=150, bbox_inches="tight")
plt.close()
print("Chart 9 saved.")

# ─── CHART 10: Q1 2026 Guidance vs Recent Quarters ───────────────────────────
fig, ax = plt.subplots(figsize=(8, 5), facecolor=BG)
ax.set_facecolor(BG)
guide_qtrs = ["Q2'25", "Q3'25", "Q4'25", "Q1'26E"]
guide_rev  = [144.0, 155.0, 179.7, None]
guide_low_val  = 185.0
guide_high_val = 200.0
guide_mid_val  = 192.5

# Actual bars
actual_cols = [GRAY, GRAY, BLUE]
for i, (q, v) in enumerate(zip(guide_qtrs[:-1], guide_rev[:-1])):
    ax.bar(i, v, color=actual_cols[i], edgecolor="white", width=0.5, zorder=3)
    ax.text(i, v + 2, f"${v:.0f}M", ha="center", va="bottom", fontsize=10)

# Guidance range bar
ax.bar(3, guide_high_val - guide_low_val, bottom=guide_low_val,
       color=GOLD, edgecolor="white", width=0.5, alpha=0.8, zorder=3, label="Guidance Range")
ax.text(3, guide_mid_val + 0.5, f"$185–200M", ha="center", va="bottom", fontsize=9, fontweight="bold", color="#7D6608")

ax.set_xticks(range(4))
ax.set_xticklabels(guide_qtrs)
ax.set_title("Revenue Trend & Q1 2026 Guidance", fontsize=13, fontweight="bold", color=BLUE, pad=10)
ax.set_ylabel("Revenue (US$M)", fontsize=10)
ax.set_ylim(100, 240)
ax.legend(fontsize=9)
ax.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
plt.tight_layout()
plt.savefig(OUT + "rklb_chart10_guidance.png", dpi=150, bbox_inches="tight")
plt.close()
print("Chart 10 saved.")

print("\nAll 10 charts saved to:", OUT)
