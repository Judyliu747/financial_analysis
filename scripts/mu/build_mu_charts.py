"""
Micron Technology (MU) Q1 FY2026 Earnings Update — Chart Generator
Quarter ended: November 27, 2025 | Reported: December 17, 2025
"""

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import os

OUT = "/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis/output/MU"
os.makedirs(OUT, exist_ok=True)

plt.rcParams.update({
    "font.family":        "serif",
    "font.serif":         ["Times New Roman"],
    "axes.unicode_minus": False,
    "figure.dpi":         150,
    "axes.spines.top":    False,
    "axes.spines.right":  False,
})

TEAL   = "#1A6B8A"
GRAY   = "#A0A0A0"
ORANGE = "#D97B2B"
GREEN  = "#2E7D52"
RED    = "#C0392B"
LBLUE  = "#5BA3C9"
LGRAY  = "#E8E8E8"

# ─── Data ────────────────────────────────────────────────────────────────────

quarters  = ["Q1\nFY25", "Q2\nFY25", "Q3\nFY25", "Q4\nFY25", "Q1\nFY26"]
revenue   = [8.71,        7.74,        8.89,       11.32,       13.64]      # $B
dram_rev  = [6.39,        5.54,        6.72,        8.82,       10.80]      # $B
nand_rev  = [2.32,        2.20,        2.17,        2.50,        2.70]      # $B

# EPS (non-GAAP)
eps_actual = [1.79,  1.30,  1.89,  3.02,  4.78]
eps_cons   = [None,  None,  None,  None,  3.94]   # Q1 FY26 consensus only

# Gross margins (non-GAAP)
gm         = [22.6,  22.4,  35.3,  45.8,  56.8]   # %

# Business unit revenue Q1 FY26
bu_labels  = ["Cloud Memory", "Core Data Ctr", "Mobile & Client", "Auto & Embedded"]
bu_vals    = [5.28,            2.40,             4.30,              1.66]

# Beat/Miss summary vs consensus Q1 FY26
bm_metrics  = ["Revenue\n($B)", "Gross Margin\n(%)", "Non-GAAP EPS\n($)"]
bm_actual   = [13.64,            56.8,                4.78]
bm_est      = [13.00,            47.0,                3.94]

# Q2 FY26 Guidance
guide_labels = ["Revenue\n($B)", "Gross Margin\n(%)", "EPS ($)"]
guide_vals   = [18.70,            68.0,                8.42]
guide_prior  = [13.64,            56.8,                4.78]   # Q1 actuals for context

# HBM TAM forecast
tam_years    = [2025, 2026, 2027, 2028]
tam_vals     = [35,    52,    72,   100]  # $B  (~40% CAGR)

# Operating cash flow
ocf          = [2.45,  1.76,  3.12,  5.73,  8.41]

# Capex
capex        = [2.11,  2.46,  2.73,  3.37,  4.50]

# DRAM / NAND mix %
dram_pct     = [73, 72, 76, 78, 79]
nand_pct     = [27, 28, 24, 22, 20]   # other ~1%


# ─── Chart 1: Quarterly Revenue Trend ───────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5))
bars = ax.bar(quarters, revenue, color=[GRAY]*4 + [TEAL], width=0.55, zorder=3)
for bar, val in zip(bars, revenue):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.18,
            f"${val:.2f}B", ha="center", va="bottom", fontsize=9.5, fontweight="bold",
            color=TEAL if val == revenue[-1] else "#444")
ax.set_ylabel("Revenue ($B)", fontsize=10)
ax.set_title("Micron Technology — Quarterly Revenue", fontsize=13, fontweight="bold", pad=10)
ax.set_ylim(0, 16.5)
ax.axhline(0, color="#888", linewidth=0.6)
ax.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
ax.tick_params(axis="both", labelsize=9)
fig.text(0.5, -0.02, "Source: Micron Q1 FY2026 Earnings Release (Dec 17, 2025); SEC Form 10-Q",
         ha="center", fontsize=7.5, color="#666")
plt.tight_layout()
plt.savefig(os.path.join(OUT, "mu_chart1_revenue.png"), bbox_inches="tight")
plt.close()
print("Chart 1 saved.")

# ─── Chart 2: DRAM vs NAND Revenue ─────────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5))
x = np.arange(len(quarters))
w = 0.38
b1 = ax.bar(x - w/2, dram_rev, width=w, color=TEAL, label="DRAM", zorder=3)
b2 = ax.bar(x + w/2, nand_rev, width=w, color=ORANGE, label="NAND", zorder=3)
for bar, val in zip(b1, dram_rev):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.1,
            f"${val:.1f}B", ha="center", va="bottom", fontsize=8, color="#333")
for bar, val in zip(b2, nand_rev):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.1,
            f"${val:.1f}B", ha="center", va="bottom", fontsize=8, color="#333")
ax.set_xticks(x); ax.set_xticklabels(quarters, fontsize=9)
ax.set_ylabel("Revenue ($B)", fontsize=10)
ax.set_title("Micron — DRAM vs. NAND Revenue by Quarter", fontsize=13, fontweight="bold", pad=10)
ax.set_ylim(0, 13.5)
ax.legend(fontsize=9, frameon=False)
ax.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
fig.text(0.5, -0.02, "Source: Micron Q1 FY2026 Earnings Release; Form 10-Q (Dec 2025)",
         ha="center", fontsize=7.5, color="#666")
plt.tight_layout()
plt.savefig(os.path.join(OUT, "mu_chart2_dram_nand.png"), bbox_inches="tight")
plt.close()
print("Chart 2 saved.")

# ─── Chart 3: Non-GAAP EPS Progression ──────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5))
colors = [GRAY]*4 + [TEAL]
bars = ax.bar(quarters, eps_actual, color=colors, width=0.55, zorder=3)
for bar, val in zip(bars, eps_actual):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.06,
            f"${val:.2f}", ha="center", va="bottom", fontsize=9.5, fontweight="bold",
            color=TEAL if val == eps_actual[-1] else "#444")
# consensus dot
x_idx = len(quarters) - 1
ax.scatter(x_idx, eps_cons[-1], color=RED, s=80, zorder=5, label=f"Consensus Est. ${eps_cons[-1]:.2f}")
ax.legend(fontsize=9, frameon=False)
ax.set_ylabel("Non-GAAP EPS ($)", fontsize=10)
ax.set_title("Micron Technology — Non-GAAP Diluted EPS", fontsize=13, fontweight="bold", pad=10)
ax.set_ylim(0, 5.8)
ax.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
fig.text(0.5, -0.02, "Source: Micron Q1 FY2026 Earnings Release (Dec 17, 2025); Consensus: Bloomberg",
         ha="center", fontsize=7.5, color="#666")
plt.tight_layout()
plt.savefig(os.path.join(OUT, "mu_chart3_eps.png"), bbox_inches="tight")
plt.close()
print("Chart 3 saved.")

# ─── Chart 4: Gross Margin Trend ────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5))
ax.plot(quarters, gm, color=TEAL, linewidth=2.5, marker="o", markersize=8, zorder=3)
for i, (q, val) in enumerate(zip(quarters, gm)):
    color = TEAL if i == len(gm)-1 else "#444"
    ax.annotate(f"{val:.1f}%", (q, val), textcoords="offset points",
                xytext=(0, 10), ha="center", fontsize=9.5,
                fontweight="bold" if i == len(gm)-1 else "normal", color=color)
ax.fill_between(range(len(quarters)), gm, alpha=0.12, color=TEAL)
ax.set_ylabel("Non-GAAP Gross Margin (%)", fontsize=10)
ax.set_title("Micron Technology — Non-GAAP Gross Margin Trend", fontsize=13, fontweight="bold", pad=10)
ax.set_ylim(15, 65)
ax.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
ax.tick_params(axis="both", labelsize=9)
fig.text(0.5, -0.02, "Source: Micron Earnings Releases Q1 FY2025–Q1 FY2026",
         ha="center", fontsize=7.5, color="#666")
plt.tight_layout()
plt.savefig(os.path.join(OUT, "mu_chart4_gross_margin.png"), bbox_inches="tight")
plt.close()
print("Chart 4 saved.")

# ─── Chart 5: Business Unit Revenue — Q1 FY26 ───────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5))
bar_colors = [TEAL, LBLUE, ORANGE, GREEN]
bars = ax.barh(bu_labels, bu_vals, color=bar_colors, height=0.5, zorder=3)
for bar, val in zip(bars, bu_vals):
    ax.text(bar.get_width() + 0.06, bar.get_y() + bar.get_height()/2,
            f"${val:.2f}B", va="center", fontsize=9.5, fontweight="bold", color="#333")
ax.set_xlabel("Revenue ($B)", fontsize=10)
ax.set_title("Micron — Q1 FY2026 Revenue by Business Unit", fontsize=13, fontweight="bold", pad=10)
ax.set_xlim(0, 7.2)
ax.grid(axis="x", linestyle="--", alpha=0.4, zorder=0)
ax.tick_params(axis="both", labelsize=9)
fig.text(0.5, -0.02, "Source: Micron Q1 FY2026 Earnings Release (Dec 17, 2025)",
         ha="center", fontsize=7.5, color="#666")
plt.tight_layout()
plt.savefig(os.path.join(OUT, "mu_chart5_business_units.png"), bbox_inches="tight")
plt.close()
print("Chart 5 saved.")

# ─── Chart 6: Beat / Miss Summary ───────────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5))
x = np.arange(len(bm_metrics))
w = 0.35
b1 = ax.bar(x - w/2, bm_actual, width=w, color=TEAL, label="Actual", zorder=3)
b2 = ax.bar(x + w/2, bm_est,    width=w, color=RED,  label="Consensus Est.", zorder=3, alpha=0.8)
for bar, val in zip(b1, bm_actual):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.4,
            f"{val:.2f}", ha="center", va="bottom", fontsize=9.5, fontweight="bold", color=TEAL)
for bar, val in zip(b2, bm_est):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.4,
            f"{val:.2f}", ha="center", va="bottom", fontsize=9.5, color=RED)
ax.set_xticks(x); ax.set_xticklabels(bm_metrics, fontsize=10)
ax.set_title("Micron Q1 FY2026 — Results vs. Consensus (Beat Across the Board)",
             fontsize=12, fontweight="bold", pad=10)
ax.set_ylim(0, 65)
ax.legend(fontsize=9, frameon=False)
ax.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
fig.text(0.5, -0.02, "Source: Micron Q1 FY2026 Earnings Release; Bloomberg consensus as of Dec 17, 2025",
         ha="center", fontsize=7.5, color="#666")
plt.tight_layout()
plt.savefig(os.path.join(OUT, "mu_chart6_beat_miss.png"), bbox_inches="tight")
plt.close()
print("Chart 6 saved.")

# ─── Chart 7: Operating Cash Flow & CapEx ───────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5))
x = np.arange(len(quarters))
w = 0.38
b1 = ax.bar(x - w/2, ocf,   width=w, color=TEAL,   label="Operating Cash Flow", zorder=3)
b2 = ax.bar(x + w/2, capex, width=w, color=ORANGE, label="CapEx", zorder=3)
for bar, val in zip(b1, ocf):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.1,
            f"${val:.2f}B", ha="center", va="bottom", fontsize=8, color="#333")
for bar, val in zip(b2, capex):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.1,
            f"${val:.2f}B", ha="center", va="bottom", fontsize=8, color="#333")
ax.set_xticks(x); ax.set_xticklabels(quarters, fontsize=9)
ax.set_ylabel("$B", fontsize=10)
ax.set_title("Micron — Operating Cash Flow & Capital Expenditures", fontsize=13, fontweight="bold", pad=10)
ax.set_ylim(0, 10.5)
ax.legend(fontsize=9, frameon=False)
ax.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
fig.text(0.5, -0.02, "Source: Micron Earnings Releases Q1 FY2025–Q1 FY2026; Form 10-Q",
         ha="center", fontsize=7.5, color="#666")
plt.tight_layout()
plt.savefig(os.path.join(OUT, "mu_chart7_cash_flow.png"), bbox_inches="tight")
plt.close()
print("Chart 7 saved.")

# ─── Chart 8: HBM TAM Forecast ──────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5))
bars = ax.bar(tam_years, tam_vals, color=[GRAY, LBLUE, TEAL, "#0D3F55"], width=0.55, zorder=3)
for bar, val in zip(bars, tam_vals):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
            f"${val}B", ha="center", va="bottom", fontsize=10, fontweight="bold", color="#333")
ax.set_ylabel("HBM TAM ($B)", fontsize=10)
ax.set_title("Global HBM TAM Forecast — ~40% CAGR (2025–2028)", fontsize=13, fontweight="bold", pad=10)
ax.set_ylim(0, 120)
ax.set_xticks(tam_years)
ax.set_xticklabels([str(y) for y in tam_years], fontsize=10)
ax.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
ax.annotate("~40% CAGR", xy=(2027, 80), fontsize=11, color=TEAL, fontweight="bold",
            arrowprops=dict(arrowstyle="->", color=TEAL, lw=1.5),
            xytext=(2025.8, 100))
fig.text(0.5, -0.02, "Source: Micron Q1 FY2026 Earnings Call (Dec 17, 2025); Management commentary",
         ha="center", fontsize=7.5, color="#666")
plt.tight_layout()
plt.savefig(os.path.join(OUT, "mu_chart8_hbm_tam.png"), bbox_inches="tight")
plt.close()
print("Chart 8 saved.")

# ─── Chart 9: Q2 FY26 Guidance vs Q1 FY26 Actuals ──────────────────────────
fig, ax = plt.subplots(figsize=(9, 5))
x = np.arange(len(guide_labels))
w = 0.35
b1 = ax.bar(x - w/2, guide_prior, width=w, color=GRAY,  label="Q1 FY26 Actual", zorder=3)
b2 = ax.bar(x + w/2, guide_vals,  width=w, color=TEAL,  label="Q2 FY26 Guidance", zorder=3)
labels_prior = ["$13.64B", "56.8%", "$4.78"]
labels_guide = ["$18.70B", "~68%",  "$8.42"]
for bar, lbl in zip(b1, labels_prior):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
            lbl, ha="center", va="bottom", fontsize=9, color="#555")
for bar, lbl in zip(b2, labels_guide):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
            lbl, ha="center", va="bottom", fontsize=9.5, fontweight="bold", color=TEAL)
ax.set_xticks(x); ax.set_xticklabels(guide_labels, fontsize=10)
ax.set_title("Q2 FY2026 Guidance vs. Q1 FY2026 Actuals", fontsize=13, fontweight="bold", pad=10)
ax.set_ylim(0, 80)
ax.legend(fontsize=9, frameon=False)
ax.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
fig.text(0.5, -0.02, "Source: Micron Q1 FY2026 Earnings Call Guidance Commentary (Dec 17, 2025)",
         ha="center", fontsize=7.5, color="#666")
plt.tight_layout()
plt.savefig(os.path.join(OUT, "mu_chart9_guidance.png"), bbox_inches="tight")
plt.close()
print("Chart 9 saved.")

# ─── Chart 10: DRAM/NAND Product Mix % ──────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5))
x = np.arange(len(quarters))
ax.stackplot(x, dram_pct, nand_pct,
             labels=["DRAM", "NAND & Other"],
             colors=[TEAL, ORANGE], alpha=0.85)
ax.set_xticks(x); ax.set_xticklabels(quarters, fontsize=9)
ax.set_ylabel("Revenue Mix (%)", fontsize=10)
ax.set_title("Micron — Revenue Mix: DRAM vs. NAND & Other", fontsize=13, fontweight="bold", pad=10)
ax.set_ylim(0, 105)
ax.legend(loc="upper left", fontsize=9, frameon=False)
ax.grid(axis="y", linestyle="--", alpha=0.3, zorder=0)
for i, (d, n) in enumerate(zip(dram_pct, nand_pct)):
    ax.text(i, d/2, f"{d}%", ha="center", va="center", fontsize=9, color="white", fontweight="bold")
    ax.text(i, d + n/2, f"{n}%", ha="center", va="center", fontsize=9, color="white", fontweight="bold")
fig.text(0.5, -0.02, "Source: Micron Earnings Releases Q1 FY2025–Q1 FY2026",
         ha="center", fontsize=7.5, color="#666")
plt.tight_layout()
plt.savefig(os.path.join(OUT, "mu_chart10_mix.png"), bbox_inches="tight")
plt.close()
print("Chart 10 saved.")

print("\nAll 10 charts generated successfully.")
