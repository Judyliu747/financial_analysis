#!/usr/bin/env python3
"""Generate charts for Cerebras Systems (CBRS) IPO & FY2025 Financial Review."""

import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import os

OUT = "/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis/output/CBRS"
os.makedirs(OUT, exist_ok=True)

plt.rcParams.update({
    "font.family":        "sans-serif",
    "font.sans-serif":    ["Arial Unicode MS", "STHeiti", "Hiragino Sans GB", "Times New Roman"],
    "axes.unicode_minus": False,
    "figure.dpi":         150,
    "savefig.dpi":        150,
    "savefig.bbox":       "tight",
    "axes.spines.top":    False,
    "axes.spines.right":  False,
})

NAVY   = "#1B2A4A"
BLUE   = "#2E5090"
LBLUE  = "#5B8DBE"
GOLD   = "#D4A843"
RED    = "#C0392B"
GREEN  = "#27AE60"
GRAY   = "#7F8C8D"
LGRAY  = "#ECF0F1"

# ==============================================================
# Chart 1: Annual Revenue Progression (FY2022-FY2025)
# ==============================================================
fig, ax = plt.subplots(figsize=(8, 5))
years = ["FY2022", "FY2023", "FY2024", "FY2025"]
rev   = [24.6, 78.7, 290.3, 510.0]
growth = [None, 220, 269, 76]

bars = ax.bar(years, rev, color=[LBLUE, LBLUE, BLUE, NAVY], width=0.55, edgecolor="white")
for i, (b, v) in enumerate(zip(bars, rev)):
    ax.text(b.get_x() + b.get_width()/2, b.get_height() + 8,
            f"${v:.0f}M", ha="center", va="bottom", fontsize=11, fontweight="bold", color=NAVY)
    if growth[i]:
        ax.text(b.get_x() + b.get_width()/2, b.get_height()/2,
                f"+{growth[i]}%", ha="center", va="center", fontsize=9, color="white", fontweight="bold")

ax.set_title("Cerebras Systems — 年度营收增长", fontsize=14, fontweight="bold", color=NAVY, pad=15)
ax.set_ylabel("营收（百万美元）", fontsize=11, color=NAVY)
ax.set_ylim(0, 600)
ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x:.0f}M"))
ax.tick_params(colors=GRAY)
fig.savefig(f"{OUT}/cbrs_chart1_revenue.png")
plt.close()

# ==============================================================
# Chart 2: Revenue by Segment (FY2025)
# ==============================================================
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 5))

segments = ["硬件系统\n$358M", "云服务及其他\n$152M"]
values   = [358.4, 151.6]
colors   = [NAVY, GOLD]
explode  = (0.03, 0.03)
wedges, texts, autotexts = ax1.pie(values, labels=segments, colors=colors, autopct="%1.1f%%",
                                    startangle=90, explode=explode, textprops={"fontsize": 10})
for at in autotexts:
    at.set_color("white")
    at.set_fontweight("bold")
ax1.set_title("FY2025 营收构成", fontsize=12, fontweight="bold", color=NAVY, pad=10)

hw_rev   = [None, None, 217.7, 358.4]
cl_rev   = [None, None,  72.6, 151.6]
x = np.arange(2)
xlabels = ["FY2024", "FY2025"]
w = 0.35
b1 = ax2.bar(x - w/2, [217.7, 358.4], w, label="硬件系统", color=NAVY)
b2 = ax2.bar(x + w/2, [72.6, 151.6],  w, label="云服务及其他", color=GOLD)
ax2.set_xticks(x)
ax2.set_xticklabels(xlabels)
ax2.set_ylabel("百万美元", fontsize=10, color=NAVY)
ax2.set_title("分部营收趋势", fontsize=12, fontweight="bold", color=NAVY, pad=10)
ax2.legend(fontsize=9)
for b in [b1, b2]:
    for bar in b:
        h = bar.get_height()
        ax2.text(bar.get_x() + bar.get_width()/2, h + 5, f"${h:.0f}M",
                 ha="center", fontsize=9, fontweight="bold", color=NAVY)
ax2.set_ylim(0, 430)
ax2.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x:.0f}"))
fig.suptitle("Cerebras Systems — 营收分部分析", fontsize=14, fontweight="bold", color=NAVY, y=1.02)
fig.tight_layout()
fig.savefig(f"{OUT}/cbrs_chart2_segment.png")
plt.close()

# ==============================================================
# Chart 3: Gross Margin Trend
# ==============================================================
fig, ax = plt.subplots(figsize=(8, 5))
years_gm = ["FY2022", "FY2023", "FY2024", "FY2025"]
gm       = [11.7, 33.5, 42.3, 39.0]
gp       = [2.9, 26.4, 122.7, 199.1]

ax2 = ax.twinx()
bars = ax2.bar(years_gm, gp, color=LBLUE, alpha=0.4, width=0.5, label="毛利润（百万美元）")
line = ax.plot(years_gm, gm, color=NAVY, marker="o", linewidth=2.5, markersize=8, label="毛利率（%）", zorder=5)

for i, (x_val, y_val) in enumerate(zip(years_gm, gm)):
    ax.annotate(f"{y_val:.1f}%", (x_val, y_val), textcoords="offset points",
                xytext=(0, 12), ha="center", fontsize=10, fontweight="bold", color=NAVY)

ax.set_ylabel("毛利率（%）", fontsize=11, color=NAVY)
ax2.set_ylabel("毛利润（百万美元）", fontsize=11, color=LBLUE)
ax.set_ylim(0, 55)
ax2.set_ylim(0, 280)
ax.set_title("Cerebras Systems — 毛利率与毛利润趋势", fontsize=14, fontweight="bold", color=NAVY, pad=15)
lines1, labels1 = ax.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax.legend(lines1 + lines2, labels1 + labels2, loc="upper left", fontsize=9)
ax.tick_params(colors=GRAY)
ax2.tick_params(colors=GRAY)
fig.savefig(f"{OUT}/cbrs_chart3_margin.png")
plt.close()

# ==============================================================
# Chart 4: Operating Expenses Breakdown
# ==============================================================
fig, ax = plt.subplots(figsize=(8, 5))
years_oe = ["FY2022", "FY2023", "FY2024", "FY2025"]
rd   = [155.4, 140.1, 158.2, 243.3]
sga  = [26.3,  20.2,  65.9, 101.6]

x = np.arange(len(years_oe))
w = 0.55
b1 = ax.bar(x, rd,  w, label="研发费用", color=BLUE)
b2 = ax.bar(x, sga, w, bottom=rd, label="销售及管理费用", color=GOLD)

for i in range(len(years_oe)):
    total = rd[i] + sga[i]
    ax.text(x[i], total + 5, f"${total:.0f}M", ha="center", fontsize=10, fontweight="bold", color=NAVY)

ax.set_xticks(x)
ax.set_xticklabels(years_oe)
ax.set_ylabel("百万美元", fontsize=11, color=NAVY)
ax.set_title("Cerebras Systems — 运营费用构成", fontsize=14, fontweight="bold", color=NAVY, pad=15)
ax.legend(fontsize=10)
ax.set_ylim(0, 420)
ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x:.0f}"))
ax.tick_params(colors=GRAY)
fig.savefig(f"{OUT}/cbrs_chart4_opex.png")
plt.close()

# ==============================================================
# Chart 5: Operating Loss Trend
# ==============================================================
fig, ax = plt.subplots(figsize=(8, 5))
years_ol = ["FY2022", "FY2023", "FY2024", "FY2025"]
op_loss  = [-178.8, -133.9, -101.4, -145.9]
op_mgn   = [-726.4, -170.1, -35.0, -28.6]

colors_bar = [RED if v < -130 else GOLD for v in op_loss]
bars = ax.bar(years_ol, op_loss, color=colors_bar, width=0.55, edgecolor="white")
for b, v in zip(bars, op_loss):
    ax.text(b.get_x() + b.get_width()/2, v - 8,
            f"(${abs(v):.0f}M)", ha="center", va="top", fontsize=10, fontweight="bold", color=NAVY)

ax.axhline(0, color=GRAY, linewidth=0.8)
ax.set_ylabel("百万美元", fontsize=11, color=NAVY)
ax.set_title("Cerebras Systems — 运营亏损趋势", fontsize=14, fontweight="bold", color=NAVY, pad=15)
ax.set_ylim(-220, 20)
ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x:.0f}"))
ax.tick_params(colors=GRAY)

ax2 = ax.twinx()
ax2.plot(years_ol, op_mgn, color=GREEN, marker="s", linewidth=2, markersize=7, label="运营利润率")
ax2.set_ylabel("运营利润率（%）", fontsize=11, color=GREEN)
ax2.set_ylim(-800, 50)
for x_val, y_val in zip(years_ol[1:], op_mgn[1:]):
    ax2.annotate(f"{y_val:.1f}%", (x_val, y_val), textcoords="offset points",
                 xytext=(0, 12), ha="center", fontsize=9, color=GREEN, fontweight="bold")
ax2.legend(loc="lower right", fontsize=9)
fig.savefig(f"{OUT}/cbrs_chart5_oploss.png")
plt.close()

# ==============================================================
# Chart 6: Customer Concentration (FY2025)
# ==============================================================
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 5))

labels_c = ["MBZUAI\n(62%)", "G42\n(24%)", "其他客户\n(14%)"]
sizes_c  = [62, 24, 14]
colors_c = [NAVY, BLUE, GOLD]
explode_c = (0.05, 0.02, 0.02)
wedges, texts, autotexts = ax1.pie(sizes_c, labels=labels_c, colors=colors_c,
                                    autopct="%1.0f%%", startangle=90, explode=explode_c,
                                    textprops={"fontsize": 10})
for at in autotexts:
    at.set_color("white")
    at.set_fontweight("bold")
ax1.set_title("FY2025 客户集中度", fontsize=12, fontweight="bold", color=NAVY, pad=10)

periods = ["H1 2024", "FY2025", "FY2026E"]
g42_pct = [87, 24, 5]
mbz_pct = [0, 62, 15]
oai_pct = [0, 0, 55]
oth_pct = [13, 14, 25]
x = np.arange(len(periods))
w = 0.55
ax2.bar(x, g42_pct, w, label="G42", color=BLUE)
ax2.bar(x, mbz_pct, w, bottom=g42_pct, label="MBZUAI", color=NAVY)
ax2.bar(x, oai_pct, w, bottom=[g+m for g,m in zip(g42_pct, mbz_pct)], label="OpenAI", color=GREEN)
ax2.bar(x, oth_pct, w, bottom=[g+m+o for g,m,o in zip(g42_pct, mbz_pct, oai_pct)], label="其他", color=GOLD)
ax2.set_xticks(x)
ax2.set_xticklabels(periods)
ax2.set_ylabel("营收占比（%）", fontsize=10, color=NAVY)
ax2.set_title("客户集中度演变", fontsize=12, fontweight="bold", color=NAVY, pad=10)
ax2.legend(fontsize=8, loc="upper right")
ax2.set_ylim(0, 115)

fig.suptitle("Cerebras Systems — 客户集中度分析", fontsize=14, fontweight="bold", color=NAVY, y=1.02)
fig.tight_layout()
fig.savefig(f"{OUT}/cbrs_chart6_customers.png")
plt.close()

# ==============================================================
# Chart 7: Net Income / EPS Trend
# ==============================================================
fig, ax = plt.subplots(figsize=(8, 5))
years_ni = ["FY2022", "FY2023", "FY2024", "FY2025"]
ni_gaap  = [-177.7, -127.2, -481.6, 237.8]
ni_adj   = [-177.7, -127.2, -481.6, -75.7]

x = np.arange(len(years_ni))
w = 0.35
b1 = ax.bar(x - w/2, ni_gaap, w, label="GAAP 净利润", color=NAVY)
b2 = ax.bar(x + w/2, ni_adj,  w, label="Non-GAAP 净利润（调整后）", color=GOLD)
ax.axhline(0, color=GRAY, linewidth=0.8)

for bars_group in [b1, b2]:
    for b in bars_group:
        h = b.get_height()
        offset = 8 if h >= 0 else -15
        ax.text(b.get_x() + b.get_width()/2, h + offset,
                f"${h:.0f}M", ha="center", fontsize=8, fontweight="bold", color=NAVY)

ax.set_xticks(x)
ax.set_xticklabels(years_ni)
ax.set_ylabel("百万美元", fontsize=11, color=NAVY)
ax.set_title("Cerebras Systems — GAAP vs. Non-GAAP 净利润", fontsize=14, fontweight="bold", color=NAVY, pad=15)
ax.legend(fontsize=9, loc="lower left")
ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x:.0f}"))
ax.tick_params(colors=GRAY)
fig.savefig(f"{OUT}/cbrs_chart7_netincome.png")
plt.close()

# ==============================================================
# Chart 8: R&D Investment as % of Revenue
# ==============================================================
fig, ax = plt.subplots(figsize=(8, 5))
years_rd = ["FY2022", "FY2023", "FY2024", "FY2025"]
rd_abs   = [155.4, 140.1, 158.2, 243.3]
rd_pct   = [631.3, 177.9, 54.5, 47.7]

ax2 = ax.twinx()
bars = ax.bar(years_rd, rd_abs, color=BLUE, width=0.55, alpha=0.7, label="研发费用（百万美元）")
for b, v in zip(bars, rd_abs):
    ax.text(b.get_x() + b.get_width()/2, v + 5,
            f"${v:.0f}M", ha="center", fontsize=10, fontweight="bold", color=NAVY)

ax2.plot(years_rd[1:], rd_pct[1:], color=RED, marker="o", linewidth=2.5, markersize=8, label="研发占收入比（%）")
for x_val, y_val in zip(years_rd[1:], rd_pct[1:]):
    ax2.annotate(f"{y_val:.1f}%", (x_val, y_val), textcoords="offset points",
                 xytext=(0, 12), ha="center", fontsize=10, fontweight="bold", color=RED)

ax.set_ylabel("研发费用（百万美元）", fontsize=11, color=NAVY)
ax2.set_ylabel("研发占收入比（%）", fontsize=11, color=RED)
ax.set_title("Cerebras Systems — 研发投入趋势", fontsize=14, fontweight="bold", color=NAVY, pad=15)
ax.set_ylim(0, 300)
ax2.set_ylim(0, 220)
lines1, labels1 = ax.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax.legend(lines1 + lines2, labels1 + labels2, loc="upper right", fontsize=9)
ax.tick_params(colors=GRAY)
fig.savefig(f"{OUT}/cbrs_chart8_rd.png")
plt.close()

# ==============================================================
# Chart 9: Valuation Comparison (P/S Ratio vs AI Chip Peers)
# ==============================================================
fig, ax = plt.subplots(figsize=(8, 5))
companies = ["NVIDIA\n(NVDA)", "AMD\n(AMD)", "Broadcom\n(AVGO)", "Marvell\n(MRVL)", "Cerebras\n(CBRS)"]
ps_ratios = [28.5, 9.2, 18.3, 15.6, 131.3]
colors_v  = [LBLUE, LBLUE, LBLUE, LBLUE, NAVY]

bars = ax.barh(companies, ps_ratios, color=colors_v, height=0.5, edgecolor="white")
for b, v in zip(bars, ps_ratios):
    ax.text(v + 2, b.get_y() + b.get_height()/2, f"{v:.1f}x",
            va="center", fontsize=11, fontweight="bold", color=NAVY)

ax.set_xlabel("市销率（P/S TTM）", fontsize=11, color=NAVY)
ax.set_title("Cerebras vs. AI芯片同行 — 市销率对比", fontsize=14, fontweight="bold", color=NAVY, pad=15)
ax.set_xlim(0, 160)
ax.tick_params(colors=GRAY)
ax.invert_yaxis()
fig.savefig(f"{OUT}/cbrs_chart9_valuation.png")
plt.close()

# ==============================================================
# Chart 10: IPO Price Performance & Capital Structure
# ==============================================================
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 5))

labels_p = ["IPO 定价\n$185", "首日开盘\n$350", "首日收盘\n$311", "次日收盘\n$280"]
prices   = [185, 350, 311, 280]
colors_p = [GRAY, GREEN, BLUE, GOLD]
bars = ax1.bar(labels_p, prices, color=colors_p, width=0.55, edgecolor="white")
for b, v in zip(bars, prices):
    ax1.text(b.get_x() + b.get_width()/2, v + 5,
             f"${v}", ha="center", fontsize=11, fontweight="bold", color=NAVY)
ax1.set_ylabel("股价（美元）", fontsize=10, color=NAVY)
ax1.set_title("IPO 定价与首日表现", fontsize=12, fontweight="bold", color=NAVY, pad=10)
ax1.set_ylim(0, 420)
ax1.tick_params(colors=GRAY)

fund_labels = ["IPO 募资\n$5.55B", "OpenAI 贷款\n$1.0B", "AWS 股权投资\n$0.27B"]
fund_vals   = [5.55, 1.0, 0.27]
colors_f    = [NAVY, GREEN, GOLD]
bars2 = ax2.bar(fund_labels, fund_vals, color=colors_f, width=0.55, edgecolor="white")
for b, v in zip(bars2, fund_vals):
    ax2.text(b.get_x() + b.get_width()/2, v + 0.1,
             f"${v:.2f}B", ha="center", fontsize=11, fontweight="bold", color=NAVY)
ax2.set_ylabel("十亿美元", fontsize=10, color=NAVY)
ax2.set_title("资金来源", fontsize=12, fontweight="bold", color=NAVY, pad=10)
ax2.set_ylim(0, 7)
ax2.tick_params(colors=GRAY)

fig.suptitle("Cerebras Systems — IPO 概览", fontsize=14, fontweight="bold", color=NAVY, y=1.02)
fig.tight_layout()
fig.savefig(f"{OUT}/cbrs_chart10_ipo.png")
plt.close()

print("All 10 charts generated successfully in:", OUT)
for i in range(1, 11):
    fname = [f for f in os.listdir(OUT) if f.startswith(f"cbrs_chart{i}")]
    print(f"  Chart {i}: {fname[0] if fname else 'MISSING'}")
