"""
ImmunityBio (IBRX) Q4 2025 Earnings Update — Chart Generator
Generates 10 institutional-quality charts for the earnings report
"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

OUT = "/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis/output/IBRX/"

NAVY   = "#003087"
RED    = "#C8102E"
GRAY   = "#5E6A71"
LGRAY  = "#D3D3D3"
GREEN  = "#1A7A4A"
AMBER  = "#E87722"

def style_ax(ax, title, ylabel="", xlabel=""):
    ax.set_title(title, fontsize=11, fontweight='bold', color=NAVY, pad=8)
    if ylabel: ax.set_ylabel(ylabel, fontsize=8, color=GRAY)
    if xlabel: ax.set_xlabel(xlabel, fontsize=8, color=GRAY)
    ax.tick_params(colors=GRAY, labelsize=8)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color(LGRAY)
    ax.spines['bottom'].set_color(LGRAY)
    ax.yaxis.grid(True, color=LGRAY, linewidth=0.5, linestyle='--')
    ax.set_axisbelow(True)

# ─────────────────────────────────────────────────
# Chart 1: Quarterly Net Product Revenue Trend
# ─────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(7, 4))
quarters = ['Q2\'24','Q3\'24','Q4\'24','Q1\'25','Q2\'25','Q3\'25','Q4\'25']
revenues = [2.5, 6.0, 7.2, 16.5, 26.4, 32.1, 38.3]
bars = ax.bar(quarters, revenues, color=[LGRAY]*4 + [NAVY, NAVY, RED], width=0.6, zorder=3)
for bar, val in zip(bars, revenues):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
            f'${val}M', ha='center', va='bottom', fontsize=8, fontweight='bold', color=NAVY)
ax.annotate('J-Code\nJan 2025', xy=('Q1\'25', 16.5), xytext=('Q1\'25', 28),
            arrowprops=dict(arrowstyle='->', color=RED, lw=1.2),
            fontsize=7, color=RED, ha='center')
style_ax(ax, 'ANKTIVA® Quarterly Net Product Revenue', 'Revenue ($M)')
ax.set_ylim(0, 50)
fig.tight_layout()
fig.savefig(OUT + "ibrx_chart1_revenue.png", dpi=150, bbox_inches='tight')
plt.close()
print("Chart 1 saved")

# ─────────────────────────────────────────────────
# Chart 2: YoY Revenue Growth by Quarter
# ─────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(7, 4))
qtrs = ['Q3\'24','Q4\'24','Q1\'25','Q2\'25','Q3\'25','Q4\'25']
growth = [None, 20, 129, 60, 22, 20]
# QoQ growth bars
colors = [NAVY if g and g > 0 else RED for g in growth if g is not None]
valid_q = ['Q4\'24','Q1\'25','Q2\'25','Q3\'25','Q4\'25']
valid_g = [20, 129, 60, 22, 20]
bars2 = ax.bar(valid_q, valid_g, color=[LGRAY, RED, NAVY, NAVY, NAVY], width=0.6, zorder=3)
for bar, val in zip(bars2, valid_g):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
            f'+{val}%', ha='center', va='bottom', fontsize=8, fontweight='bold', color=NAVY)
ax.axhline(0, color=GRAY, linewidth=0.8)
style_ax(ax, 'ANKTIVA® Revenue — QoQ Growth (%)', 'QoQ Growth (%)')
ax.set_ylim(0, 160)
fig.tight_layout()
fig.savefig(OUT + "ibrx_chart2_qoq_growth.png", dpi=150, bbox_inches='tight')
plt.close()
print("Chart 2 saved")

# ─────────────────────────────────────────────────
# Chart 3: Net Loss Trend (improving profitability)
# ─────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(7, 4))
qtrs3 = ['Q1\'25','Q2\'25','Q3\'25','Q4\'25']
net_loss = [-91.2, -87.8, -110.5, -61.9]
bars3 = ax.bar(qtrs3, net_loss, color=[RED, RED, RED, AMBER], width=0.6, zorder=3)
for bar, val in zip(bars3, net_loss):
    ax.text(bar.get_x() + bar.get_width()/2, val - 3,
            f'({abs(val)})', ha='center', va='top', fontsize=8, fontweight='bold', color='white')
ax.axhline(0, color=GRAY, linewidth=0.8)
style_ax(ax, 'Quarterly Net Loss ($M)', 'Net Loss ($M)')
ax.set_ylim(-130, 20)
ax.text(0.98, 0.05, 'Q4 includes $14M one-time write-off', transform=ax.transAxes,
        fontsize=7, color=GRAY, ha='right', style='italic')
fig.tight_layout()
fig.savefig(OUT + "ibrx_chart3_net_loss.png", dpi=150, bbox_inches='tight')
plt.close()
print("Chart 3 saved")

# ─────────────────────────────────────────────────
# Chart 4: Beat / Miss Summary (Q4 2025) — table style
# ─────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(7, 3))
ax.axis('off')
table_data = [
    ['Metric', 'Actual', 'Consensus Est.', 'Beat / Miss'],
    ['Net Product Revenue', '$38.3M', '~$35.0M', '+$3.3M  (+9.4%) ✓'],
    ['Net Loss (EPS)', '($0.06)', '($0.085)–($0.09)', '+$0.025  (+33%) ✓'],
    ['FY2025 Revenue', '$113.3M', '~$105M', '+$8.3M  (+7.9%) ✓'],
    ['Q4 QoQ Growth', '+20%', '—', 'Unbroken streak'],
]
tbl = ax.table(cellText=table_data[1:], colLabels=table_data[0],
               loc='center', cellLoc='center')
tbl.auto_set_font_size(False)
tbl.set_fontsize(9)
tbl.scale(1.2, 1.8)
for (row, col), cell in tbl.get_celld().items():
    if row == 0:
        cell.set_facecolor(NAVY)
        cell.set_text_props(color='white', fontweight='bold')
    elif col == 3:
        cell.set_facecolor('#e8f5e9')
        cell.set_text_props(color=GREEN, fontweight='bold')
    elif row % 2 == 0:
        cell.set_facecolor('#f7f9ff')
    cell.set_edgecolor(LGRAY)
ax.set_title('Q4 2025 Results vs. Consensus Estimates', fontsize=11,
             fontweight='bold', color=NAVY, pad=12)
fig.tight_layout()
fig.savefig(OUT + "ibrx_chart4_beat_miss.png", dpi=150, bbox_inches='tight')
plt.close()
print("Chart 4 saved")

# ─────────────────────────────────────────────────
# Chart 5: Annual Revenue Ramp (FY2024 → FY2026E)
# ─────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(6, 4))
years = ['FY2024A', 'FY2025A', 'FY2026E']
rev_ann = [15.7, 113.3, 195.0]
colors5 = [LGRAY, NAVY, AMBER]
bars5 = ax.bar(years, rev_ann, color=colors5, width=0.5, zorder=3)
for bar, val in zip(bars5, rev_ann):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 2,
            f'${val}M', ha='center', va='bottom', fontsize=9, fontweight='bold', color=NAVY)
# Growth arrows
ax.annotate('', xy=(1, 113.3), xytext=(0, 15.7),
            arrowprops=dict(arrowstyle='->', color=GREEN, lw=1.5))
ax.text(0.5, 75, '+621%', ha='center', fontsize=9, color=GREEN, fontweight='bold', rotation=60)
ax.annotate('', xy=(2, 195), xytext=(1, 113.3),
            arrowprops=dict(arrowstyle='->', color=AMBER, lw=1.5))
ax.text(1.5, 160, '+72%E', ha='center', fontsize=9, color=AMBER, fontweight='bold', rotation=40)
style_ax(ax, 'Annual Net Product Revenue Ramp', 'Revenue ($M)')
ax.set_ylim(0, 240)
fig.tight_layout()
fig.savefig(OUT + "ibrx_chart5_annual_rev.png", dpi=150, bbox_inches='tight')
plt.close()
print("Chart 5 saved")

# ─────────────────────────────────────────────────
# Chart 6: OpEx Breakdown FY2024 vs FY2025
# ─────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(7, 4))
cats = ['R&D Expense', 'SG&A Expense', 'Total OpEx']
fy24 = [190.2, 168.8, 359.0]
fy25 = [218.6, 150.0, 368.6]
x = np.arange(len(cats))
w = 0.35
b1 = ax.bar(x - w/2, fy24, width=w, color=LGRAY, label='FY2024A', zorder=3)
b2 = ax.bar(x + w/2, fy25, width=w, color=NAVY, label='FY2025A', zorder=3)
for bar, val in zip(b1, fy24):
    ax.text(bar.get_x()+bar.get_width()/2, val+3, f'${val:.0f}M', ha='center', fontsize=7.5, color=GRAY)
for bar, val in zip(b2, fy25):
    ax.text(bar.get_x()+bar.get_width()/2, val+3, f'${val:.0f}M', ha='center', fontsize=7.5, color=NAVY, fontweight='bold')
# Delta labels
deltas = ['+15%', '-11%', '+3%']
for i, d in enumerate(deltas):
    c = GREEN if '-' in d else RED
    ax.text(x[i], max(fy24[i], fy25[i]) + 18, d, ha='center', fontsize=8, color=c, fontweight='bold')
ax.legend(fontsize=8)
style_ax(ax, 'Operating Expense Breakdown: FY2024 vs. FY2025', '($M)')
ax.set_xticks(x); ax.set_xticklabels(cats, fontsize=9)
ax.set_ylim(0, 450)
fig.tight_layout()
fig.savefig(OUT + "ibrx_chart6_opex.png", dpi=150, bbox_inches='tight')
plt.close()
print("Chart 6 saved")

# ─────────────────────────────────────────────────
# Chart 7: Cash Runway Analysis
# ─────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(6, 4))
quarters_cr = ['Q4\'25A', 'Q1\'26E', 'Q2\'26E', 'Q3\'26E']
cash = [242.8, 195.0, 150.0, 110.0]
burn_rate_line = [242.8, 167.1, 91.4, 15.7]  # ~$75M/quarter burn
rev_proj = [38.3, 48.0, 55.0, 62.0]
ax.bar(quarters_cr, cash, color=[NAVY, AMBER, AMBER, RED], width=0.5, zorder=3, alpha=0.8, label='Cash Balance (Est.)')
ax.axhline(0, color=RED, linewidth=1.5, linestyle='--', label='Zero Cash')
for i, (q, c) in enumerate(zip(quarters_cr, cash)):
    ax.text(i, c + 5, f'${c:.0f}M', ha='center', fontsize=8, fontweight='bold', color=NAVY)
ax.set_xticks(range(len(quarters_cr))); ax.set_xticklabels(quarters_cr, fontsize=9)
style_ax(ax, 'Estimated Cash Runway (Flat Burn ~$75M/Q)', 'Cash ($M)')
ax.set_ylim(-30, 290)
ax.legend(fontsize=8)
ax.text(0.98, 0.55, 'Financing likely needed\nin H2 2026', transform=ax.transAxes,
        fontsize=8, color=RED, ha='right', fontweight='bold',
        bbox=dict(boxstyle='round,pad=0.3', facecolor='#fff0f0', edgecolor=RED))
fig.tight_layout()
fig.savefig(OUT + "ibrx_chart7_cash.png", dpi=150, bbox_inches='tight')
plt.close()
print("Chart 7 saved")

# ─────────────────────────────────────────────────
# Chart 8: Pipeline Progress (QUILT trials)
# ─────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(8, 4))
ax.axis('off')
trials = [
    ("QUILT 3.032 (BCG-Unresponsive CIS)", "FDA Approved\n(Dec 2023)", "sBLA Resubmitted\n(Mar 2026)", "Label Expansion\nPending", ""),
    ("QUILT 2.005 (BCG-Naïve NMIBC)", "Phase 3\nEnrolled (366 pts)", "84% vs 52% CR\n(Interim)", "BLA Q4 2026E", ""),
    ("QUILT 3.055 (NSCLC)", "Phase 2b\nComplete", "SFDA Approved\n(Saudi Arabia)", "EU Filing\n2026E", ""),
    ("rBCG (BCG Shortage)", "sBLA Resubmit\nExpected", "Global Access\n580 pts EAP", "", ""),
]
colors_p = [GREEN, NAVY, AMBER, RED]
y_positions = [0.85, 0.60, 0.35, 0.10]
stage_x = [0.02, 0.30, 0.58, 0.80]
for (name, s1, s2, s3, s4), yp, cp in zip(trials, y_positions, colors_p):
    ax.text(0.01, yp + 0.05, name, fontsize=8, fontweight='bold', color=cp, transform=ax.transAxes)
    for sx, st in zip(stage_x, [s1, s2, s3, s4]):
        if st:
            ax.text(sx, yp - 0.02, st, fontsize=7, color=NAVY, transform=ax.transAxes,
                    bbox=dict(boxstyle='round,pad=0.2', facecolor='#f0f4ff', edgecolor=cp, linewidth=1))
ax.set_title("ANKTIVA® Pipeline — Key Trial Status", fontsize=11, fontweight='bold', color=NAVY)
fig.tight_layout()
fig.savefig(OUT + "ibrx_chart8_pipeline.png", dpi=150, bbox_inches='tight')
plt.close()
print("Chart 8 saved")

# ─────────────────────────────────────────────────
# Chart 9: Estimate Revisions (Pre vs Post Earnings)
# ─────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(7, 4))
ests = ['FY2026E\nRevenue ($M)', 'FY2026E\nEPS (Loss)', 'FY2027E\nRevenue ($M)']
pre  = [165, -1.20, 280]
post = [195, -0.95, 340]
x = np.arange(len(ests))
w = 0.35
b_pre  = ax.bar(x - w/2, [abs(v) for v in pre],  width=w, color=LGRAY, label='Pre-Earnings', zorder=3)
b_post = ax.bar(x + w/2, [abs(v) for v in post], width=w, color=NAVY,  label='Post-Earnings', zorder=3)
labels_pre  = ['$165M', '($1.20)', '$280M']
labels_post = ['$195M', '($0.95)', '$340M']
for bar, lbl in zip(b_pre, labels_pre):
    ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+1, lbl, ha='center', fontsize=8, color=GRAY)
for bar, lbl in zip(b_post, labels_post):
    ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+1, lbl, ha='center', fontsize=8, color=NAVY, fontweight='bold')
deltas2 = ['+$30M\n(+18%)', '+$0.25\n(+21%)', '+$60M\n(+21%)']
for i, d in enumerate(deltas2):
    ax.text(x[i], max(abs(pre[i]), abs(post[i]))+10, d, ha='center', fontsize=7.5, color=GREEN, fontweight='bold')
ax.legend(fontsize=8)
style_ax(ax, 'Consensus Estimate Revisions (Pre vs. Post Q4 Earnings)', '($M or $)')
ax.set_xticks(x); ax.set_xticklabels(ests, fontsize=9)
ax.set_ylim(0, 430)
fig.tight_layout()
fig.savefig(OUT + "ibrx_chart9_revisions.png", dpi=150, bbox_inches='tight')
plt.close()
print("Chart 9 saved")

# ─────────────────────────────────────────────────
# Chart 10: Comparable Biotech Valuations (EV/Rev NTM)
# ─────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(8, 4))
comps = ['IBRX\n(ImmunityBio)', 'MRNA\n(Moderna)', 'BNTX\n(BioNTech)', 'SGEN\n(Seagen)', 'EXEL\n(Exelixis)', 'ALNY\n(Alnylam)']
ev_rev = [42.3, 5.1, 3.8, 8.2, 4.5, 11.7]
colors10 = [RED] + [LGRAY]*5
bars10 = ax.bar(comps, ev_rev, color=colors10, width=0.6, zorder=3)
for bar, val in zip(bars10, ev_rev):
    ax.text(bar.get_x()+bar.get_width()/2, val+0.3, f'{val}x', ha='center', fontsize=8.5,
            fontweight='bold', color=NAVY)
ax.axhline(np.mean(ev_rev), color=NAVY, linestyle='--', linewidth=1, label=f'Peer Avg: {np.mean(ev_rev):.1f}x')
style_ax(ax, 'EV / NTM Revenue Multiples — Comparable Oncology Biotechs', 'EV / NTM Revenue (x)')
ax.legend(fontsize=8)
ax.set_ylim(0, 55)
ax.text(0, 43, '← IBRX trades at significant premium\nreflecting high-growth ANKTIVA ramp expectations',
        fontsize=7, color=GRAY, style='italic')
fig.tight_layout()
fig.savefig(OUT + "ibrx_chart10_valuation.png", dpi=150, bbox_inches='tight')
plt.close()
print("Chart 10 saved")

print("\nAll 10 charts generated successfully.")
