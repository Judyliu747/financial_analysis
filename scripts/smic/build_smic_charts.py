"""
SMIC (中芯国际) Q4 2025 Earnings Update — Chart Generation
Generates 10 institutional-quality charts
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

# ── Palette ──────────────────────────────────────────────────────────────────
NAVY   = "#003366"
BLUE   = "#0066CC"
LBLUE  = "#3399FF"
GOLD   = "#CC9900"
GREEN  = "#2E8B57"
RED    = "#CC3333"
LGRAY  = "#F5F5F5"
MGRAY  = "#CCCCCC"
DKGRAY = "#555555"

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "axes.spines.top":   False,
    "axes.spines.right": False,
    "axes.grid":         True,
    "grid.color":        MGRAY,
    "grid.alpha":        0.5,
    "grid.linewidth":    0.6,
    "figure.dpi":        150,
})

OUT = "/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis/output/SMIC/"

# ── Data ─────────────────────────────────────────────────────────────────────
quarters = ["Q1'24", "Q2'24", "Q3'24", "Q4'24", "Q1'25", "Q2'25", "Q3'25", "Q4'25"]

revenue    = [1746, 1900, 2170, 2207, 2170, 2281, 2382, 2489]   # USD M
gross_mgn  = [13.7, 13.9, 20.5, 22.6, 22.5, 19.3, 22.0, 19.2]  # %
net_income = [85,   73,   150,  107,  121,  130,  192,  173]    # USD M

# Consensus vs Actual Q4'25
metrics_beat = {
    "Revenue\n($M)": (2420, 2489),
    "Gross\nMargin (%)": (20.9, 19.2),
    "Net Income\n($M)": (170, 173),
}

# Revenue by end-market Q4'25
end_markets_q4  = [47, 20, 15, 10, 8]
end_market_lbls = ["Consumer\nElec.", "Smartphone", "Industrial\n& Auto", "PC/Server", "Other"]

# Revenue by region Q4'25
region_q4  = [90, 6, 4]
region_lbl = ["China\nDomestic", "Americas", "Other"]

# Margin bridge Q3→Q4'25
margin_items  = ["Q3'25\nGross Margin", "Volume\nEffect", "Mix\n(ASP)", "Depreciation\nIncrease", "Q4'25\nGross Margin"]
margin_vals   = [22.0, +0.5, +0.9, -4.2, 19.2]  # waterfall

# Updated estimates
est_years   = ["FY2024A", "FY2025A", "FY2026E (New)", "FY2026E (Old)", "FY2027E"]
est_rev     = [8028, 9327, 10600, 10200, 12000]  # USD M
est_gm      = [18.0, 21.0, 19.5, 20.5, 21.5]    # %

# Capacity utilization
util_quarters = ["Q1'24", "Q2'24", "Q3'24", "Q4'24", "Q1'25", "Q2'25", "Q3'25", "Q4'25"]
utilization   = [74.1, 85.2, 90.4, 86.3, 89.4, 91.3, 90.1, 95.7]

# Wafer capacity (8" equiv. wspm 000s)
capacity      = [780, 810, 860, 900, 940, 970, 1010, 1060]

# ── Chart 1: Quarterly Revenue ───────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5))
colors = [BLUE if q != "Q4'25" else NAVY for q in quarters]
bars = ax.bar(quarters, revenue, color=colors, width=0.55, zorder=3)
for b, v in zip(bars, revenue):
    ax.text(b.get_x()+b.get_width()/2, v+20, f"${v:,}", ha='center', va='bottom',
            fontsize=8.5, fontweight='bold', color=DKGRAY)
ax.set_title("Quarterly Revenue Trend (USD M)", fontsize=12, fontweight='bold',
             color=NAVY, pad=10)
ax.set_ylabel("Revenue (USD M)", color=DKGRAY)
ax.set_ylim(0, 2900)
ax.annotate("Q4'25: Revenue beat\nconsensus by +2.8%", xy=(7, 2489),
            xytext=(5.8, 2700), fontsize=8, color=GREEN,
            arrowprops=dict(arrowstyle='->', color=GREEN, lw=1.2))
ax.text(0.98, 0.02, "Source: SMIC Q4 2025 Earnings Release (Feb 10, 2026)",
        transform=ax.transAxes, fontsize=7, ha='right', color=DKGRAY)
plt.tight_layout()
plt.savefig(OUT + "smic_chart1_revenue.png", bbox_inches='tight')
plt.close()

# ── Chart 2: Gross Margin Trend ───────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5))
ax.fill_between(range(len(quarters)), gross_mgn, alpha=0.15, color=BLUE)
ax.plot(range(len(quarters)), gross_mgn, color=BLUE, lw=2.5, marker='o', ms=6, zorder=4)
for i, (q, v) in enumerate(zip(quarters, gross_mgn)):
    va = 'bottom' if v < 22 else 'top'
    offset = 0.4 if va == 'bottom' else -0.4
    ax.text(i, v+offset, f"{v:.1f}%", ha='center', fontsize=8.5,
            color=RED if q == "Q4'25" else DKGRAY, fontweight='bold' if q=="Q4'25" else 'normal')
ax.axhline(19.2, color=RED, lw=1, ls='--', alpha=0.7)
ax.set_xticks(range(len(quarters)))
ax.set_xticklabels(quarters)
ax.set_title("Quarterly Gross Margin Trend (%)", fontsize=12, fontweight='bold', color=NAVY, pad=10)
ax.set_ylabel("Gross Margin (%)", color=DKGRAY)
ax.set_ylim(8, 27)
ax.annotate("Margin contraction\nfrom depreciation step-up", xy=(7, 19.2),
            xytext=(5.5, 24.5), fontsize=8, color=RED,
            arrowprops=dict(arrowstyle='->', color=RED, lw=1.2))
ax.text(0.98, 0.02, "Source: SMIC Q4 2025 Earnings Release (Feb 10, 2026)",
        transform=ax.transAxes, fontsize=7, ha='right', color=DKGRAY)
plt.tight_layout()
plt.savefig(OUT + "smic_chart2_gross_margin.png", bbox_inches='tight')
plt.close()

# ── Chart 3: Net Income Trend ──────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5))
colors = [LBLUE if q != "Q4'25" else NAVY for q in quarters]
bars = ax.bar(quarters, net_income, color=colors, width=0.55, zorder=3)
for b, v in zip(bars, net_income):
    ax.text(b.get_x()+b.get_width()/2, v+2, f"${v}M", ha='center', va='bottom',
            fontsize=8.5, fontweight='bold', color=DKGRAY)
ax.set_title("Quarterly Net Income (USD M)", fontsize=12, fontweight='bold', color=NAVY, pad=10)
ax.set_ylabel("Net Income (USD M)", color=DKGRAY)
ax.set_ylim(0, 240)
ax.annotate("+60.7% YoY", xy=(7, 173), xytext=(6.0, 210),
            fontsize=9, color=GREEN, fontweight='bold',
            arrowprops=dict(arrowstyle='->', color=GREEN, lw=1.3))
ax.text(0.98, 0.02, "Source: SMIC Q4 2025 Earnings Release (Feb 10, 2026)",
        transform=ax.transAxes, fontsize=7, ha='right', color=DKGRAY)
plt.tight_layout()
plt.savefig(OUT + "smic_chart3_net_income.png", bbox_inches='tight')
plt.close()

# ── Chart 4: Beat / Miss Summary ──────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 4.5))
metric_names = list(metrics_beat.keys())
est_vals  = [metrics_beat[k][0] for k in metric_names]
act_vals  = [metrics_beat[k][1] for k in metric_names]
x = np.arange(len(metric_names))
w = 0.32
b1 = ax.bar(x - w/2, est_vals, w, label='Consensus Est.', color=MGRAY, zorder=3)
b2 = ax.bar(x + w/2, act_vals, w, label='Actual',         color=NAVY,  zorder=3)
for bar, val in zip(b1, est_vals):
    ax.text(bar.get_x()+bar.get_width()/2, val*0.5, f"{val:,}", ha='center',
            va='center', fontsize=9, color='white', fontweight='bold')
beat_colors = [GREEN if a >= e else RED for a, e in zip(act_vals, est_vals)]
for bar, val, bc in zip(b2, act_vals, beat_colors):
    ax.text(bar.get_x()+bar.get_width()/2, val*0.5, f"{val:,}", ha='center',
            va='center', fontsize=9, color='white', fontweight='bold')
    top = max(val, est_vals[list(act_vals).index(val)])
    diff_pct = (val - est_vals[list(act_vals).index(val)]) / est_vals[list(act_vals).index(val)] * 100
    label = f"{'BEAT' if diff_pct>=0 else 'MISS'} {abs(diff_pct):.1f}%"
    ax.text(bar.get_x()+bar.get_width()/2, top + top*0.03, label, ha='center',
            fontsize=8, color=bc, fontweight='bold')
ax.set_xticks(x)
ax.set_xticklabels(metric_names, fontsize=10)
ax.set_title("Q4 2025 Beat/Miss vs. Consensus", fontsize=12, fontweight='bold', color=NAVY, pad=10)
ax.legend(loc='upper right', fontsize=9)
ax.text(0.98, 0.02, "Source: SMIC Q4 2025 Earnings Release; Bloomberg Consensus (Feb 10, 2026)",
        transform=ax.transAxes, fontsize=7, ha='right', color=DKGRAY)
plt.tight_layout()
plt.savefig(OUT + "smic_chart4_beat_miss.png", bbox_inches='tight')
plt.close()

# ── Chart 5: Revenue by End-Market Q4'25 ──────────────────────────────────────
fig, ax = plt.subplots(figsize=(7.5, 5))
wedge_colors = [NAVY, BLUE, LBLUE, GOLD, MGRAY]
explode = (0.05,0,0,0,0)
wedges, texts, autotexts = ax.pie(
    end_markets_q4, labels=end_market_lbls, autopct='%1.1f%%',
    colors=wedge_colors, explode=explode,
    pctdistance=0.75, labeldistance=1.1,
    textprops={"fontsize": 9}
)
for at in autotexts:
    at.set_fontweight('bold')
    at.set_color('white')
ax.set_title("Revenue by End Market — Q4 2025", fontsize=12, fontweight='bold', color=NAVY, pad=10)
ax.text(0, -1.35, "Source: SMIC Q4 2025 Earnings Call (Feb 10, 2026)",
        ha='center', fontsize=7, color=DKGRAY)
plt.tight_layout()
plt.savefig(OUT + "smic_chart5_endmarket.png", bbox_inches='tight')
plt.close()

# ── Chart 6: Revenue by Geography ─────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(7.5, 5))
colors_r = [NAVY, BLUE, LBLUE]
wedges, texts, autotexts = ax.pie(
    region_q4, labels=region_lbl, autopct='%1.1f%%',
    colors=colors_r, pctdistance=0.7, labeldistance=1.12,
    textprops={"fontsize": 10}
)
for at in autotexts:
    at.set_fontweight('bold')
    at.set_color('white')
ax.set_title("Revenue by Geography — Q4 2025", fontsize=12, fontweight='bold', color=NAVY, pad=10)
ax.text(0, -1.35, "Source: SMIC Q4 2025 Earnings Release (Feb 10, 2026)",
        ha='center', fontsize=7, color=DKGRAY)
plt.tight_layout()
plt.savefig(OUT + "smic_chart6_geography.png", bbox_inches='tight')
plt.close()

# ── Chart 7: Gross Margin Waterfall Q3→Q4 ─────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5.5))
running = 0
positions = []
colors_wf = []
for i, (item, val) in enumerate(zip(margin_items, margin_vals)):
    if i == 0 or i == len(margin_items)-1:
        bottom = 0
        color = NAVY
    else:
        bottom = running
        color = GREEN if val >= 0 else RED
    positions.append((bottom, val))
    colors_wf.append(color)
    if i == 0:
        running = val
    elif i < len(margin_items)-1:
        running += val

for i, ((bottom, val), color) in enumerate(zip(positions, colors_wf)):
    if i == 0 or i == len(margin_items)-1:
        ax.bar(i, val, bottom=0, color=color, width=0.5, zorder=3)
        ax.text(i, val+0.2, f"{val:.1f}%", ha='center', fontsize=10, fontweight='bold', color=DKGRAY)
    else:
        ax.bar(i, abs(val), bottom=bottom if val>=0 else bottom+val, color=color, width=0.5, zorder=3)
        sign = "+" if val >= 0 else ""
        y_pos = bottom + val/2
        ax.text(i, (bottom + val + bottom)/2 + abs(val)/2 + 0.1, f"{sign}{val:.1f}pp",
                ha='center', fontsize=9, fontweight='bold', color=color)

ax.set_xticks(range(len(margin_items)))
ax.set_xticklabels(margin_items, fontsize=9)
ax.set_ylim(0, 27)
ax.set_title("Gross Margin Bridge: Q3'25 → Q4'25 (%)", fontsize=12, fontweight='bold', color=NAVY, pad=10)
ax.set_ylabel("Gross Margin (%)", color=DKGRAY)
ax.text(0.98, 0.02, "Source: SMIC Q4 2025 Earnings Call Management Commentary (Feb 10, 2026)",
        transform=ax.transAxes, fontsize=7, ha='right', color=DKGRAY)
plt.tight_layout()
plt.savefig(OUT + "smic_chart7_margin_bridge.png", bbox_inches='tight')
plt.close()

# ── Chart 8: Capacity Utilization ─────────────────────────────────────────────
fig, ax1 = plt.subplots(figsize=(9, 5))
ax2 = ax1.twinx()
ax2.bar(range(len(util_quarters)), capacity, color=LBLUE, alpha=0.3, width=0.55, label='Monthly Capacity (K wspm)')
ax1.plot(range(len(util_quarters)), utilization, color=NAVY, lw=2.5, marker='o', ms=6, zorder=5, label='Utilization Rate (%)')
for i, (u, c_) in enumerate(zip(utilization, capacity)):
    ax1.text(i, u+0.8, f"{u:.1f}%", ha='center', fontsize=8.5,
             color=RED if i == len(utilization)-1 else DKGRAY,
             fontweight='bold' if i == len(utilization)-1 else 'normal')
ax1.set_xticks(range(len(util_quarters)))
ax1.set_xticklabels(util_quarters)
ax1.set_ylim(65, 105)
ax2.set_ylim(0, 1600)
ax1.set_ylabel("Utilization Rate (%)", color=NAVY)
ax2.set_ylabel("Monthly Capacity (K wspm 8\" equiv.)", color=LBLUE)
ax1.set_title("Capacity Utilization vs. Monthly Wafer Capacity", fontsize=12, fontweight='bold', color=NAVY, pad=10)
lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend(lines1+lines2, labels1+labels2, loc='lower right', fontsize=9)
ax1.text(0.98, 0.02, "Source: SMIC Q4 2025 Earnings Release; Company Filings",
         transform=ax1.transAxes, fontsize=7, ha='right', color=DKGRAY)
plt.tight_layout()
plt.savefig(OUT + "smic_chart8_utilization.png", bbox_inches='tight')
plt.close()

# ── Chart 9: Revenue & GM Estimate Revisions ──────────────────────────────────
fig, (ax_top, ax_bot) = plt.subplots(2, 1, figsize=(9, 7), sharex=True)
est_labels = ["FY2024A", "FY2025A", "FY2026E\n(New)", "FY2026E\n(Old)", "FY2027E"]
x = np.arange(len(est_labels))
w = 0.35
ax_top.bar(x, est_rev, width=0.55, color=[NAVY if 'Old' not in l else MGRAY for l in est_labels], zorder=3)
for i, v in enumerate(est_rev):
    ax_top.text(i, v+80, f"${v/1000:.2f}B", ha='center', fontsize=9,
                color=RED if 'Old' in est_labels[i] else DKGRAY, fontweight='bold')
ax_top.set_title("Revenue Estimates — Actuals & Revisions (USD M)", fontsize=11, fontweight='bold', color=NAVY)
ax_top.set_ylabel("Revenue (USD M)")
ax_top.set_ylim(0, 14000)

ax_bot.plot(x, est_gm, marker='s', color=BLUE, lw=2, ms=7, zorder=4)
for i, v in enumerate(est_gm):
    ax_bot.text(i, v+0.3, f"{v:.1f}%", ha='center', fontsize=9,
                color=RED if 'Old' in est_labels[i] else DKGRAY)
ax_bot.set_title("Gross Margin Estimates (%) — Actuals & Revisions", fontsize=11, fontweight='bold', color=NAVY)
ax_bot.set_ylabel("Gross Margin (%)")
ax_bot.set_ylim(14, 26)
ax_bot.set_xticks(x)
ax_bot.set_xticklabels(est_labels, fontsize=9)
fig.text(0.98, 0.01, "Source: SMIC Filings; Analyst Consensus; Internal Estimates as of Mar 2026",
         ha='right', fontsize=7, color=DKGRAY)
plt.tight_layout()
plt.savefig(OUT + "smic_chart9_estimates.png", bbox_inches='tight')
plt.close()

# ── Chart 10: P/S & EV/EBITDA Valuation ────────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5))
comp_names   = ["SMIC\n(Current)", "TSMC", "UMC", "GlobalFoundries", "Hua Hong"]
ps_ratios    = [1.8, 8.5, 1.5, 2.2, 1.3]
ev_ebitda    = [8.5, 14.2, 6.8, 9.1, 7.2]
x = np.arange(len(comp_names))
w = 0.35
ax2 = ax.twinx()
b1 = ax.bar(x - w/2, ps_ratios,  w, color=NAVY, label='P/S (x)', zorder=3)
b2 = ax2.bar(x + w/2, ev_ebitda, w, color=GOLD, label='EV/EBITDA (x)', zorder=3)
for bar, v in zip(b1, ps_ratios):
    ax.text(bar.get_x()+bar.get_width()/2, v+0.05, f"{v:.1f}x", ha='center', fontsize=8.5,
            color=DKGRAY, fontweight='bold')
for bar, v in zip(b2, ev_ebitda):
    ax2.text(bar.get_x()+bar.get_width()/2, v+0.1, f"{v:.1f}x", ha='center', fontsize=8.5,
             color=DKGRAY, fontweight='bold')
ax.set_xticks(x)
ax.set_xticklabels(comp_names, fontsize=10)
ax.set_ylabel("P/S Multiple (x)", color=NAVY)
ax2.set_ylabel("EV/EBITDA Multiple (x)", color=GOLD)
ax.set_ylim(0, 12)
ax2.set_ylim(0, 20)
ax.set_title("SMIC Valuation vs. Foundry Peers (NTM Multiples)", fontsize=12, fontweight='bold', color=NAVY, pad=10)
lines1, l1 = ax.get_legend_handles_labels()
lines2, l2 = ax2.get_legend_handles_labels()
ax.legend(lines1+lines2, l1+l2, loc='upper right', fontsize=9)
ax.text(0.98, 0.02, "Source: Bloomberg; Company Filings; Analyst Consensus as of Mar 2026",
        transform=ax.transAxes, fontsize=7, ha='right', color=DKGRAY)
plt.tight_layout()
plt.savefig(OUT + "smic_chart10_valuation.png", bbox_inches='tight')
plt.close()

print("All 10 charts saved successfully.")
