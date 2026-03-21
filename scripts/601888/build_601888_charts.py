"""
中国中免 (601888.SS / 1880.HK) Q4 2025 / FY2025 Earnings Update
Chart generation script — 10 PNGs → output/601888/
Data sources: preliminary FY2025 announcement (March 20, 2026),
              quarterly reports, Hainan customs data, Wind consensus
"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import os

OUT = "/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis/output/601888/"
os.makedirs(OUT, exist_ok=True)

# ── Brand Colors ──────────────────────────────────────────────────────────────
CDF_RED   = "#C41230"   # China Duty Free signature red
CDF_GOLD  = "#C9A84C"   # accent gold
CDF_NAVY  = "#1A2D5A"   # deep navy
CDF_GRAY  = "#8B8B8B"
CDF_GREEN = "#2E7D32"
CDF_LIGHT = "#F5F0E8"

plt.rcParams['font.family']       = 'Times New Roman'
plt.rcParams['axes.spines.top']   = False
plt.rcParams['axes.spines.right'] = False

# ── Data ──────────────────────────────────────────────────────────────────────
QTRS = ["Q1\n2024", "Q2\n2024", "Q3\n2024", "Q4\n2024",
        "Q1\n2025", "Q2\n2025", "Q3\n2025", "Q4\n2025"]

REV_DATA = [188.07, 124.58, 117.56, 134.53,
            167.46, 114.05, 117.11, 138.31]   # RMB 亿元

NP_DATA  = [23.06, 9.76, 6.36, 3.48,
             19.38, 6.62, 4.52, 5.34]          # 归母净利润 RMB 亿元

REV_YOY  = [-9.45, None, None, -19.46,
             -10.96, -8.45, -0.38, 2.81]       # %

NP_YOY   = [0.25, None, None, -76.93,
             -15.98, -32.21, -28.94, 53.49]    # %

# ─────────────────────────────────────────────────────────────────────────────
def save_footer(ax, text):
    ax.text(0.01, -0.12, text, transform=ax.transAxes,
            fontsize=7.5, color=CDF_GRAY)

# ─────────────────────────────────────────────────────────────────────────────
# Chart 1: Quarterly Revenue Trend
def chart1():
    fig, ax = plt.subplots(figsize=(10, 5.5))
    colors = [CDF_NAVY]*4 + [CDF_RED]*4
    bars = ax.bar(QTRS, REV_DATA, color=colors, edgecolor='white', linewidth=0.5, width=0.6)
    for bar, val in zip(bars, REV_DATA):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1.5,
                f'¥{val:.1f}B', ha='center', va='bottom', fontsize=9, fontweight='bold')
    ax.set_ylabel("Revenue (RMB Billion, 亿元)", fontsize=11)
    ax.set_title("China Tourism Group Duty Free (601888.SS)\nQuarterly Revenue Trend — Q1 2024 to Q4 2025",
                 fontsize=13, fontweight='bold', pad=12)
    ax.set_ylim(0, 235)
    ax.axvline(x=3.5, color=CDF_GOLD, linewidth=1.5, linestyle='--', alpha=0.8)
    ax.text(3.57, 215, 'FY2025 →', color=CDF_GOLD, fontsize=9, fontweight='bold')
    p1 = mpatches.Patch(color=CDF_NAVY, label='FY2024')
    p2 = mpatches.Patch(color=CDF_RED,  label='FY2025')
    ax.legend(handles=[p1, p2], loc='upper right', fontsize=9)
    save_footer(ax, "Source: Company quarterly reports; preliminary FY2025 announcement (March 20, 2026)")
    plt.tight_layout()
    plt.savefig(OUT + "cdf_chart1_quarterly_revenue.png", dpi=150, bbox_inches='tight')
    plt.close()
    print("Chart 1 saved: Quarterly Revenue")

# Chart 2: Quarterly Net Profit
def chart2():
    fig, ax = plt.subplots(figsize=(10, 5.5))
    colors = [CDF_NAVY]*4 + [CDF_RED]*4
    bars = ax.bar(QTRS, NP_DATA, color=colors, edgecolor='white', linewidth=0.5, width=0.6)
    for bar, val in zip(bars, NP_DATA):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.25,
                f'¥{val:.2f}B', ha='center', va='bottom', fontsize=9, fontweight='bold')
    ax.set_ylabel("Net Profit 归母 (RMB Billion, 亿元)", fontsize=11)
    ax.set_title("China Tourism Group Duty Free (601888.SS)\nQuarterly Net Profit (Attributable to Parent) — Q1 2024 to Q4 2025",
                 fontsize=13, fontweight='bold', pad=12)
    ax.set_ylim(0, 29)
    ax.axvline(x=3.5, color=CDF_GOLD, linewidth=1.5, linestyle='--', alpha=0.8)
    ax.text(3.57, 26.5, 'FY2025 →', color=CDF_GOLD, fontsize=9, fontweight='bold')
    ax.annotate('+53.5% YoY\nEarnings inflection\npoint', xy=(7, 5.34), xytext=(5.5, 13),
                arrowprops=dict(arrowstyle='->', color=CDF_GREEN, lw=1.5),
                fontsize=8.5, color=CDF_GREEN, fontweight='bold')
    p1 = mpatches.Patch(color=CDF_NAVY, label='FY2024')
    p2 = mpatches.Patch(color=CDF_RED,  label='FY2025')
    ax.legend(handles=[p1, p2], loc='upper right', fontsize=9)
    save_footer(ax, "Source: Company quarterly reports; preliminary FY2025 announcement (March 20, 2026)")
    plt.tight_layout()
    plt.savefig(OUT + "cdf_chart2_quarterly_netprofit.png", dpi=150, bbox_inches='tight')
    plt.close()
    print("Chart 2 saved: Quarterly Net Profit")

# Chart 3: Revenue YoY Growth — Recovery Trajectory
def chart3():
    fig, ax = plt.subplots(figsize=(10, 5.5))
    yoy_vals = [-9.45, None, None, -19.46, -10.96, -8.45, -0.38, 2.81]
    valid = [(i, v) for i, v in enumerate(yoy_vals) if v is not None]
    idx_v = [i for i, _ in valid]
    val_v = [v for _, v in valid]
    bar_colors = [CDF_GREEN if v >= 0 else CDF_RED for v in val_v]
    bars = ax.bar(idx_v, val_v, color=bar_colors, alpha=0.8, width=0.5)
    for bar, (i, v) in zip(bars, valid):
        y_pos = v + 0.5 if v >= 0 else v - 2.5
        ax.text(bar.get_x() + bar.get_width()/2, y_pos,
                f'{v:+.1f}%', ha='center', fontsize=9, fontweight='bold',
                color=CDF_GREEN if v >= 0 else CDF_RED)
    ax.set_xticks(range(8))
    ax.set_xticklabels(QTRS)
    ax.axhline(0, color='black', linewidth=0.8)
    ax.set_ylabel("Revenue YoY Growth (%)", fontsize=11)
    ax.set_title("China Tourism Group Duty Free (601888.SS)\nQuarterly Revenue YoY Growth — Recovery Trajectory",
                 fontsize=13, fontweight='bold', pad=12)
    ax.set_ylim(-28, 14)
    ax.annotate('First positive\nrevenue growth\nsince Q2 2023', xy=(7, 2.81), xytext=(5.2, 9),
                arrowprops=dict(arrowstyle='->', color=CDF_NAVY, lw=1.5),
                fontsize=8.5, color=CDF_NAVY, fontweight='bold')
    ax.axvline(x=3.5, color=CDF_GOLD, linewidth=1.5, linestyle='--', alpha=0.7)
    save_footer(ax, "Source: Company filings; note: Q2 2024 / Q3 2024 YoY growth not separately disclosed")
    plt.tight_layout()
    plt.savefig(OUT + "cdf_chart3_revenue_yoy.png", dpi=150, bbox_inches='tight')
    plt.close()
    print("Chart 3 saved: Revenue YoY")

# Chart 4: Gross Margin Trend
def chart4():
    fig, ax = plt.subplots(figsize=(10, 5.5))
    labels  = ["H1 2024\n(Est.)", "H1 2025", "9M 2025", "FY2025"]
    margins = [32.3,              31.5,       32.54,     33.0]
    ax.plot(range(len(labels)), margins, color=CDF_RED, marker='o', markersize=10,
            linewidth=2.5, markerfacecolor='white', markeredgewidth=2.5, markeredgecolor=CDF_RED)
    for i, (lbl, m) in enumerate(zip(labels, margins)):
        ax.text(i, m + 0.12, f'{m:.2f}%', ha='center', va='bottom',
                fontsize=10.5, fontweight='bold', color=CDF_RED)
    ax.fill_between(range(len(labels)), margins, 29, alpha=0.1, color=CDF_RED)
    ax.set_xticks(range(len(labels)))
    ax.set_xticklabels(labels)
    ax.set_ylim(29, 36)
    ax.set_ylabel("Gross Margin (%)", fontsize=11)
    ax.set_title("China Tourism Group Duty Free (601888.SS)\nGross Margin Trend — 2024 to FY2025",
                 fontsize=13, fontweight='bold', pad=12)
    ax.text(3, 34.5, "Q4 2025 alone: +4.12 ppts YoY\n(sharpest quarterly recovery)",
            ha='center', fontsize=9, color=CDF_GREEN, fontweight='bold',
            bbox=dict(boxstyle='round,pad=0.35', facecolor='white', edgecolor=CDF_GREEN, alpha=0.85))
    save_footer(ax, "Source: H1 2025 results announcement; 9M 2025 quarterly report; FY2025 preliminary announcement")
    plt.tight_layout()
    plt.savefig(OUT + "cdf_chart4_gross_margin.png", dpi=150, bbox_inches='tight')
    plt.close()
    print("Chart 4 saved: Gross Margin")

# Chart 5: Hainan Offshore Duty-Free Market
def chart5():
    fig, ax1 = plt.subplots(figsize=(10, 5.5))
    periods = ["H1\n2024", "H2\n2024\n(Est.)", "H1\n2025", "H2\n2025\n(Est.)"]
    total_sales = [182.5, 182.2, 167.6, 190.0]   # 亿元 (Hainan market total)
    avg_spend   = [10927, 11200, 6754, 7500]       # RMB per shopper trip
    x = np.arange(len(periods))
    bars = ax1.bar(x, total_sales, color=CDF_NAVY, alpha=0.75, width=0.5,
                   label='Hainan Total Duty-Free Sales (亿元)')
    for bar, val in zip(bars, total_sales):
        ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1.5,
                 f'{val:.1f}B', ha='center', fontsize=9, fontweight='bold')
    ax1.set_ylabel("Hainan Total Sales (RMB Billion, 亿元)", fontsize=10, color=CDF_NAVY)
    ax1.set_xticks(x)
    ax1.set_xticklabels(periods)
    ax1.set_ylim(0, 240)
    ax2 = ax1.twinx()
    ax2.plot(x, avg_spend, color=CDF_GOLD, marker='D', markersize=9, linewidth=2.5,
             markerfacecolor='white', markeredgewidth=2.5, label='Avg Spend per Shopper (RMB)')
    for xi, val in zip(x, avg_spend):
        ax2.text(xi, val + 150, f'¥{val:,}', ha='center', fontsize=9,
                 fontweight='bold', color=CDF_GOLD)
    ax2.set_ylabel("Avg Spend per Shopper (RMB)", fontsize=10, color=CDF_GOLD)
    ax2.tick_params(colors=CDF_GOLD)
    ax2.set_ylim(0, 16000)
    ax1.set_title("Hainan Offshore Duty-Free Market\nTotal Sales & Average Spend per Shopper",
                  fontsize=13, fontweight='bold', pad=12)
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper left', fontsize=9)
    save_footer(ax1, "Source: Hainan Island Customs; H1 2025 company report; H2 2025 estimated")
    plt.tight_layout()
    plt.savefig(OUT + "cdf_chart5_hainan_market.png", dpi=150, bbox_inches='tight')
    plt.close()
    print("Chart 5 saved: Hainan Market")

# Chart 6: FY2024 Revenue Segment Breakdown
def chart6():
    fig, ax = plt.subplots(figsize=(10, 5.5))
    segments = ["Hainan Duty-Free\n(离岛免税)", "Airport & Downtown\n(日上上海 etc.)",
                "Other Domestic\n& Overseas"]
    values  = [288.92, 160.35, 115.47]
    pcts    = [51.2, 28.4, 20.4]
    colors  = [CDF_RED, CDF_NAVY, CDF_GOLD]
    bars = ax.barh(segments, values, color=colors, edgecolor='white', height=0.5)
    for bar, val, pct in zip(bars, values, pcts):
        ax.text(bar.get_width() + 3, bar.get_y() + bar.get_height()/2,
                f'¥{val:.2f}B  ({pct:.1f}%)', va='center', fontsize=10, fontweight='bold')
    ax.set_xlabel("Revenue (RMB Billion, 亿元)", fontsize=11)
    ax.set_title("China Tourism Group Duty Free (601888.SS)\nFY2024 Revenue by Segment — Reference Baseline",
                 fontsize=13, fontweight='bold', pad=12)
    ax.set_xlim(0, 400)
    ax.text(5, -0.45, "FY2024 Total: ¥564.74B  |  Hainan remains >50% of revenue",
            fontsize=9, color=CDF_GRAY, style='italic')
    save_footer(ax, "Source: Company FY2024 Annual Report (March 2025)")
    plt.tight_layout()
    plt.savefig(OUT + "cdf_chart6_segment_revenue.png", dpi=150, bbox_inches='tight')
    plt.close()
    print("Chart 6 saved: Segment Revenue")

# Chart 7: Actual vs. Consensus — Beat/Miss Analysis
def chart7():
    fig, ax = plt.subplots(figsize=(10, 5.5))
    metrics   = ["FY2025\nRevenue", "FY2025\nNet Profit",
                 "Q4 2025\nRevenue", "Q4 2025\nNet Profit"]
    actual    = [536.94, 35.86, 138.31, 5.34]
    consensus = [536.5,  35.82, 132.0,  4.80]   # Wind/broker median estimates
    beat_pct  = [(a/c - 1)*100 for a, c in zip(actual, consensus)]
    x = np.arange(4)
    w = 0.35
    ax.bar(x - w/2, consensus, w, color=CDF_GRAY,  alpha=0.65, label='Consensus Estimate')
    bars2 = ax.bar(x + w/2, actual, w,
                   color=[CDF_GREEN if b >= 0 else CDF_RED for b in beat_pct], label='Actual Result')
    for i, (a, bp) in enumerate(zip(actual, beat_pct)):
        color = CDF_GREEN if bp >= 0 else CDF_RED
        ax.text(i + w/2, a + max(actual)*0.008, f'{bp:+.1f}%',
                ha='center', fontsize=9.5, fontweight='bold', color=color)
    ax.set_xticks(x)
    ax.set_xticklabels(metrics)
    ax.set_ylabel("RMB Billion (亿元)", fontsize=11)
    ax.set_title("China Tourism Group Duty Free (601888.SS)\nActual vs. Consensus — FY2025 & Q4 2025 Beat/Miss",
                 fontsize=13, fontweight='bold', pad=12)
    ax.legend(fontsize=9, loc='upper right')
    save_footer(ax, "Source: Wind consensus estimates (21 brokerages); preliminary FY2025 announcement (March 20, 2026)")
    plt.tight_layout()
    plt.savefig(OUT + "cdf_chart7_beat_miss.png", dpi=150, bbox_inches='tight')
    plt.close()
    print("Chart 7 saved: Beat/Miss")

# Chart 8: Annual Performance Comparison FY2023–FY2025
def chart8():
    fig, axes = plt.subplots(1, 2, figsize=(12, 5.5))
    years = ['FY2023', 'FY2024', 'FY2025A']
    rev   = [675.5, 564.74, 536.94]
    np_   = [67.1,  42.67,  35.86]

    colors_ann = [CDF_NAVY, CDF_NAVY, CDF_RED]
    bars1 = axes[0].bar(years, rev, color=colors_ann, edgecolor='white', width=0.5)
    for bar, val in zip(bars1, rev):
        axes[0].text(bar.get_x() + bar.get_width()/2, bar.get_height() + 4,
                     f'¥{val:.1f}B', ha='center', fontsize=11, fontweight='bold')
    axes[0].set_ylim(0, 850)
    axes[0].set_ylabel("Revenue (RMB Billion, 亿元)", fontsize=11)
    axes[0].set_title("Annual Revenue (FY2023–FY2025)", fontsize=12, fontweight='bold')
    yoy_r = [None, -16.38, -4.92]
    for i, yoy in enumerate(yoy_r):
        if yoy is not None:
            axes[0].text(i, rev[i]/2, f'{yoy:+.1f}%\nYoY', ha='center', fontsize=11,
                         color='white', fontweight='bold')

    bars2 = axes[1].bar(years, np_, color=colors_ann, edgecolor='white', width=0.5)
    for bar, val in zip(bars2, np_):
        axes[1].text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
                     f'¥{val:.2f}B', ha='center', fontsize=11, fontweight='bold')
    axes[1].set_ylim(0, 82)
    axes[1].set_ylabel("Net Profit 归母 (RMB Billion, 亿元)", fontsize=11)
    axes[1].set_title("Annual Net Profit (FY2023–FY2025)", fontsize=12, fontweight='bold')
    yoy_n = [None, -36.44, -15.97]
    for i, yoy in enumerate(yoy_n):
        if yoy is not None:
            axes[1].text(i, np_[i]/2, f'{yoy:+.1f}%\nYoY', ha='center', fontsize=11,
                         color='white', fontweight='bold')

    fig.suptitle("China Tourism Group Duty Free — Annual Performance Comparison",
                 fontsize=13, fontweight='bold', y=1.01)
    axes[0].text(0.01, -0.14, "Source: Company Annual Reports; FY2023 revenue derived from FY2024 YoY delta",
                 transform=axes[0].transAxes, fontsize=7.5, color=CDF_GRAY)
    plt.tight_layout()
    plt.savefig(OUT + "cdf_chart8_annual_comparison.png", dpi=150, bbox_inches='tight')
    plt.close()
    print("Chart 8 saved: Annual Comparison")

# Chart 9: FY2026E Net Profit Broker Estimates
def chart9():
    fig, ax = plt.subplots(figsize=(10, 5.5))
    brokers   = ["Huatai\nSecurities", "Wind\nConsensus", "Morgan\nStanley\n(implied)",
                 "Dongwu\nSecurities", "Galaxy\nSecurities"]
    estimates = [42.09, 44.17, 46.0, 57.6, 58.7]   # 亿元
    colors_b  = [CDF_GRAY, CDF_NAVY, CDF_GRAY, CDF_RED, CDF_RED]
    bars = ax.bar(brokers, estimates, color=colors_b, edgecolor='white', width=0.55)
    for bar, val in zip(bars, estimates):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.4,
                f'¥{val:.2f}B', ha='center', fontsize=10, fontweight='bold')
    ax.axhline(35.86, color=CDF_GOLD, linewidth=2, linestyle='--',
               label='FY2025A Actual: ¥35.86B')
    ax.text(4.35, 36.5, 'FY2025A\nActual', fontsize=8.5, color=CDF_GOLD, fontweight='bold')
    ax.set_ylabel("FY2026E Net Profit 归母 (RMB Billion, 亿元)", fontsize=11)
    ax.set_title("China Tourism Group Duty Free (601888.SS)\nFY2026E Net Profit Estimates by Broker",
                 fontsize=13, fontweight='bold', pad=12)
    ax.set_ylim(0, 75)
    ax.legend(fontsize=9, loc='upper left')
    save_footer(ax, "Source: Wind consensus; Huatai, Dongwu, Galaxy Securities; Morgan Stanley (H-share implied); March 2026")
    plt.tight_layout()
    plt.savefig(OUT + "cdf_chart9_broker_estimates.png", dpi=150, bbox_inches='tight')
    plt.close()
    print("Chart 9 saved: Broker Estimates")

# Chart 10: Revenue & Net Profit Recovery Forecast (FY2023–FY2027E)
def chart10():
    fig, ax1 = plt.subplots(figsize=(10, 5.5))
    years = ['FY2023', 'FY2024', 'FY2025A', 'FY2026E\n(Wind)', 'FY2027E\n(Est.)']
    rev   = [675.5, 564.74, 536.94, 610.0, 675.0]   # 亿元
    np_   = [67.1,  42.67,  35.86,  44.17, 52.0]    # 亿元
    x = np.arange(len(years))
    bar_colors = [CDF_NAVY]*2 + [CDF_RED] + [CDF_GOLD]*2
    bars = ax1.bar(x, rev, color=bar_colors, alpha=0.8, width=0.55, label='Revenue (亿元)')
    for bar, val in zip(bars, rev):
        ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 5,
                 f'{val:.0f}', ha='center', fontsize=9, fontweight='bold')
    ax1.set_ylabel("Revenue (RMB Billion, 亿元)", fontsize=11, color=CDF_NAVY)
    ax1.set_ylim(0, 870)
    ax2 = ax1.twinx()
    ax2.plot(x, np_, color=CDF_RED, marker='o', markersize=9, linewidth=2.5,
             markerfacecolor='white', markeredgewidth=2.5, label='Net Profit 归母 (亿元)')
    for xi, val in zip(x, np_):
        ax2.text(xi, val + 1.8, f'{val:.1f}', ha='center', fontsize=9,
                 fontweight='bold', color=CDF_RED)
    ax2.set_ylabel("Net Profit 归母 (RMB Billion, 亿元)", fontsize=11, color=CDF_RED)
    ax2.tick_params(colors=CDF_RED)
    ax2.set_ylim(0, 105)
    ax1.set_xticks(x)
    ax1.set_xticklabels(years)
    ax1.axvline(x=2.5, color=CDF_GRAY, linewidth=1.5, linestyle='--', alpha=0.7)
    ax1.text(2.52, 820, '← Actual  |  Forecast →', color=CDF_GRAY, fontsize=8.5)
    ax1.set_title("China Tourism Group Duty Free (601888.SS)\nRevenue & Net Profit — Actual and Recovery Forecast (FY2023–FY2027E)",
                  fontsize=13, fontweight='bold', pad=12)
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper left', fontsize=9)
    save_footer(ax1, "Source: Company filings; FY2026E/FY2027E based on Wind consensus estimates (March 2026)")
    plt.tight_layout()
    plt.savefig(OUT + "cdf_chart10_recovery_forecast.png", dpi=150, bbox_inches='tight')
    plt.close()
    print("Chart 10 saved: Recovery Forecast")

# ── Run all charts ────────────────────────────────────────────────────────────
if __name__ == "__main__":
    chart1()
    chart2()
    chart3()
    chart4()
    chart5()
    chart6()
    chart7()
    chart8()
    chart9()
    chart10()
    print("\nAll 10 charts generated successfully.")
