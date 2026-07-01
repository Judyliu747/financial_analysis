"""
Shared data and chart helpers for Micron Technology Q3 FY2026 earnings update.
"""

from __future__ import annotations

from pathlib import Path
from typing import Iterable

from PIL import Image, ImageDraw, ImageFont


ROOT = Path("/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis")
OUT = ROOT / "output" / "MU"
OUT.mkdir(parents=True, exist_ok=True)

REPORT_DATE = "July 1, 2026"
REPORT_DATE_CN = "2026年7月1日"
QUARTER = "Q3 FY2026"
QUARTER_CN = "FY2026财年第三季度"
PERIOD_END = "May 28, 2026"
PERIOD_END_CN = "2026年5月28日"
RELEASE_DATE = "June 24, 2026"
FILING_DATE = "June 25, 2026"

PRESS_RELEASE_URL = "https://investors.micron.com/news-releases/news-release-details/micron-technology-inc-reports-record-results-third-quarter"
PRESENTATION_URL = "https://investors.micron.com/static-files/2354ecda-77a0-4ddd-8462-a631eb491356"
PREPARED_REMARKS_URL = "https://investors.micron.com/static-files/631b1a32-5537-46ae-8f40-82e42fc79dfe"
TEN_Q_URL = "https://investors.micron.com/sec-filings/sec-filing/10-q/0000723125-26-000015"
QUARTERLY_RESULTS_URL = "https://investors.micron.com/quarterly-results"
IBD_URL = "https://www.investors.com/news/technology/micron-stock-mu-fiscal-q3-2026-earnings/"
INVESTOPEDIA_URL = "https://www.investopedia.com/micron-earnings-q3-fy2026-memory-stock-soars-ai-demand-12006096"

LONG_BRIDGE_QUOTE = {
    "as_of": "June 30, 2026 close / July 1, 2026 pre-market",
    "last_close": 1154.29,
    "premarket": 1110.56,
    "market_cap_b": 1302.0,
    "source": "Longbridge CLI quote MU.US, pulled July 1, 2026 16:24 CST",
}

CONSENSUS = {
    "revenue_b": 35.91,
    "eps": 20.86,
    "q4_revenue_b": 43.58,
    "q4_eps": 25.72,
    "source": "FactSet consensus cited by Investor's Business Daily, June 24, 2026",
}

RESULTS = {
    "revenue_b": 41.456,
    "gaap_net_income_b": 28.243,
    "gaap_eps": 24.67,
    "non_gaap_net_income_b": 28.857,
    "non_gaap_eps": 25.11,
    "gaap_gross_margin_pct": 84.6,
    "non_gaap_gross_margin_pct": 84.9,
    "non_gaap_operating_income_b": 33.681,
    "non_gaap_operating_margin_pct": 81.2,
    "operating_cash_flow_b": 25.388,
    "capex_b": 7.084,
    "free_cash_flow_b": 18.304,
    "cash_investments_b": 30.155,
    "debt_b": 5.7,
}

GUIDANCE_Q4 = {
    "revenue_b_mid": 50.0,
    "revenue_b_low": 49.0,
    "revenue_b_high": 51.0,
    "non_gaap_gross_margin_pct": 86.0,
    "non_gaap_opex_b": 1.65,
    "non_gaap_eps_mid": 31.00,
    "non_gaap_eps_low": 30.00,
    "non_gaap_eps_high": 32.00,
    "capex_b": 10.0,
    "fy2026_capex_b": 27.0,
}

SCAS = {
    "signed_agreements": 16,
    "rpo_q3_b": 5.0,
    "rpo_signed_b": 100.0,
    "cash_commitments_b": 22.0,
    "cash_deposits_b": 18.0,
}

QUARTERS = ["Q3\nFY25", "Q4\nFY25", "Q1\nFY26", "Q2\nFY26", "Q3\nFY26"]
REVENUE_TREND = [9.301, 11.315, 13.639, 23.860, 41.456]
EPS_TREND = [1.91, 3.03, 4.78, 12.20, 25.11]
GROSS_MARGIN_TREND = [39.0, 45.8, 56.8, 74.9, 84.9]
OPERATING_MARGIN_TREND = [26.8, 39.0, 47.0, 69.0, 81.2]
DRAM_REV = [7.071, 8.82, 10.80, 18.768, 31.328]
NAND_REV = [2.155, 2.50, 2.70, 4.997, 9.943]
OCF_TREND = [4.609, 5.73, 8.41, 11.903, 25.388]
CAPEX_TREND = [2.90, 3.37, 4.50, 5.50, 7.084]

BUSINESS_UNITS = [
    ("Cloud Memory", 13.769, 83, 78, 307),
    ("Core Data Center", 11.524, 87, 103, 653),
    ("Mobile & Client", 11.521, 87, 49, 254),
    ("Auto & Embedded", 4.634, 79, 71, 311),
]

ESTIMATES = [
    ("Revenue ($B)", 112.0, 129.0, 175.0),
    ("Gross Margin", 78.0, 80.5, 82.0),
    ("Non-GAAP EPS", 60.0, 73.1, 115.0),
    ("Free Cash Flow ($B)", 42.0, 55.0, 75.0),
]

SOURCES = [
    ("Micron Q3 FY2026 earnings release", PRESS_RELEASE_URL),
    ("Micron Q3 FY2026 investor presentation", PRESENTATION_URL),
    ("Micron Q3 FY2026 prepared remarks/webcast materials", PREPARED_REMARKS_URL),
    ("Micron Q3 FY2026 Form 10-Q", TEN_Q_URL),
    ("Micron quarterly results page", QUARTERLY_RESULTS_URL),
    ("Investor's Business Daily / FactSet consensus", IBD_URL),
    ("Investopedia / Visible Alpha consensus", INVESTOPEDIA_URL),
]


TEAL = "#1A6B8A"
BLUE = "#0B3D5C"
LIGHT_BLUE = "#EAF4FA"
ORANGE = "#D97B2B"
GREEN = "#2E7D52"
RED = "#C0392B"
GRAY = "#6B7280"
LIGHT_GRAY = "#E5E7EB"
DARK = "#111827"
WHITE = "#FFFFFF"


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    candidates = [
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/System/Library/Fonts/Supplemental/Times New Roman Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Times New Roman.ttf",
        "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
    ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, size=size)
        except OSError:
            continue
    return ImageFont.load_default()


def _text_size(draw: ImageDraw.ImageDraw, text: str, fnt: ImageFont.ImageFont) -> tuple[int, int]:
    box = draw.textbbox((0, 0), text, font=fnt)
    return box[2] - box[0], box[3] - box[1]


def _hex(color: str) -> str:
    return color


def canvas(title: str, subtitle: str | None = None) -> tuple[Image.Image, ImageDraw.ImageDraw]:
    img = Image.new("RGB", (1350, 760), WHITE)
    d = ImageDraw.Draw(img)
    d.rectangle((0, 0, 1350, 72), fill=BLUE)
    d.text((42, 20), title, fill=WHITE, font=font(28, True))
    if subtitle:
        d.text((42, 82), subtitle, fill=GRAY, font=font(16))
    return img, d


def save(img: Image.Image, name: str, source: str) -> Path:
    d = ImageDraw.Draw(img)
    d.text((42, 724), f"Source: {source}", fill=GRAY, font=font(13))
    d.text((1170, 724), "Micron Q3 FY2026", fill=GRAY, font=font(13))
    path = OUT / name
    img.save(path)
    return path


def _plot_area(d: ImageDraw.ImageDraw, y_top: int = 140) -> tuple[int, int, int, int]:
    left, top, right, bottom = 100, y_top, 1265, 670
    d.line((left, bottom, right, bottom), fill="#9CA3AF", width=2)
    d.line((left, top, left, bottom), fill="#9CA3AF", width=2)
    return left, top, right, bottom


def bar_chart(name: str, title: str, labels: list[str], values: list[float], unit: str,
              source: str, highlight_last: bool = True, consensus: float | None = None,
              subtitle: str | None = None) -> Path:
    img, d = canvas(title, subtitle)
    left, top, right, bottom = _plot_area(d)
    max_v = max(values + ([consensus] if consensus else [])) * 1.18
    for i in range(6):
        y = bottom - (bottom - top) * i / 5
        d.line((left, y, right, y), fill=LIGHT_GRAY, width=1)
        d.text((35, y - 8), f"{max_v * i / 5:.0f}", fill=GRAY, font=font(13))
    slot = (right - left) / len(values)
    bw = slot * 0.48
    for i, (label, val) in enumerate(zip(labels, values)):
        x0 = left + slot * i + (slot - bw) / 2
        y0 = bottom - (val / max_v) * (bottom - top)
        color = TEAL if (highlight_last and i == len(values) - 1) else "#9CA3AF"
        d.rounded_rectangle((x0, y0, x0 + bw, bottom), radius=6, fill=color)
        d.text((x0 + bw / 2 - _text_size(d, f"{val:.1f}{unit}", font(15, True))[0] / 2, y0 - 26),
               f"{val:.1f}{unit}", fill=color, font=font(15, True))
        for j, part in enumerate(label.split("\n")):
            d.text((x0 + bw / 2 - _text_size(d, part, font(13))[0] / 2, bottom + 14 + j * 16),
                   part, fill=DARK, font=font(13))
    if consensus is not None:
        y = bottom - (consensus / max_v) * (bottom - top)
        d.line((left, y, right, y), fill=RED, width=3)
        d.text((right - 220, y - 26), f"Consensus {consensus:.1f}{unit}", fill=RED, font=font(15, True))
    return save(img, name, source)


def line_chart(name: str, title: str, labels: list[str], series: list[tuple[str, list[float], str]],
               unit: str, source: str, subtitle: str | None = None) -> Path:
    img, d = canvas(title, subtitle)
    left, top, right, bottom = _plot_area(d)
    max_v = max(max(vals) for _, vals, _ in series) * 1.15
    min_v = min(min(vals) for _, vals, _ in series) * 0.80
    for i in range(6):
        y = bottom - (bottom - top) * i / 5
        d.line((left, y, right, y), fill=LIGHT_GRAY, width=1)
        tick = min_v + (max_v - min_v) * i / 5
        d.text((35, y - 8), f"{tick:.0f}{unit}", fill=GRAY, font=font(13))
    slot = (right - left) / (len(labels) - 1)
    for label_i, label in enumerate(labels):
        x = left + slot * label_i
        for j, part in enumerate(label.split("\n")):
            d.text((x - _text_size(d, part, font(13))[0] / 2, bottom + 14 + j * 16), part, fill=DARK, font=font(13))
    for name_s, vals, color in series:
        pts = []
        for i, val in enumerate(vals):
            x = left + slot * i
            y = bottom - ((val - min_v) / (max_v - min_v)) * (bottom - top)
            pts.append((x, y))
        d.line(pts, fill=color, width=5)
        for x, y in pts:
            d.ellipse((x - 6, y - 6, x + 6, y + 6), fill=WHITE, outline=color, width=4)
        d.text((pts[-1][0] + 18, pts[-1][1] - 10), name_s, fill=color, font=font(15, True))
    return save(img, name, source)


def grouped_bar_chart(name: str, title: str, labels: list[str], group1: list[float],
                      group2: list[float], legend1: str, legend2: str, unit: str,
                      source: str, subtitle: str | None = None) -> Path:
    img, d = canvas(title, subtitle)
    left, top, right, bottom = _plot_area(d)
    max_v = max(group1 + group2) * 1.2
    for i in range(6):
        y = bottom - (bottom - top) * i / 5
        d.line((left, y, right, y), fill=LIGHT_GRAY, width=1)
        d.text((35, y - 8), f"{max_v * i / 5:.0f}", fill=GRAY, font=font(13))
    slot = (right - left) / len(labels)
    bw = slot * 0.22
    for i, label in enumerate(labels):
        center = left + slot * i + slot / 2
        for dx, val, color in [(-bw * 0.6, group1[i], TEAL), (bw * 0.6, group2[i], ORANGE)]:
            x0 = center + dx - bw / 2
            y0 = bottom - (val / max_v) * (bottom - top)
            d.rounded_rectangle((x0, y0, x0 + bw, bottom), radius=5, fill=color)
            d.text((x0 + bw / 2 - _text_size(d, f"{val:.1f}", font(12))[0] / 2, y0 - 19),
                   f"{val:.1f}", fill=color, font=font(12, True))
        for j, part in enumerate(label.split("\n")):
            d.text((center - _text_size(d, part, font(13))[0] / 2, bottom + 14 + j * 16), part, fill=DARK, font=font(13))
    d.rectangle((1010, 104, 1030, 124), fill=TEAL)
    d.text((1038, 101), legend1, fill=DARK, font=font(15))
    d.rectangle((1130, 104, 1150, 124), fill=ORANGE)
    d.text((1158, 101), legend2, fill=DARK, font=font(15))
    return save(img, name, source)


def horizontal_bar_chart(name: str, title: str, data: list[tuple[str, float, int, int, int]],
                         source: str, subtitle: str | None = None) -> Path:
    img, d = canvas(title, subtitle)
    left, top, right, bottom = 250, 155, 1235, 650
    max_v = max(v for _, v, _, _, _ in data) * 1.15
    for i, (label, val, gm, qoq, yoy) in enumerate(data):
        y = top + i * 105
        w = (val / max_v) * (right - left)
        d.text((42, y + 18), label, fill=DARK, font=font(18, True))
        d.text((42, y + 45), f"GM {gm}% | QoQ +{qoq}% | YoY +{yoy}%", fill=GRAY, font=font(14))
        d.rounded_rectangle((left, y + 20, left + w, y + 72), radius=8, fill=TEAL if i == 0 else "#4F9ABD")
        d.text((left + w + 16, y + 32), f"${val:.1f}B", fill=DARK, font=font(18, True))
    return save(img, name, source)


def estimate_chart() -> Path:
    img, d = canvas("Estimate Revisions and Price Target", "Illustrative model revisions after Q3 FY2026 beat")
    left, top, right, bottom = _plot_area(d)
    max_v = 180
    slot = (right - left) / len(ESTIMATES)
    bw = slot * 0.24
    for i in range(5):
        y = bottom - (bottom - top) * i / 4
        d.line((left, y, right, y), fill=LIGHT_GRAY, width=1)
    for i, (label, old, new, fy27) in enumerate(ESTIMATES):
        center = left + slot * i + slot / 2
        vals = [old, new, fy27]
        colors = ["#9CA3AF", TEAL, GREEN]
        names = ["Old FY26E", "New FY26E", "FY27E"]
        for j, val in enumerate(vals):
            x0 = center + (j - 1) * bw * 1.25 - bw / 2
            y0 = bottom - (val / max_v) * (bottom - top)
            d.rounded_rectangle((x0, y0, x0 + bw, bottom), radius=5, fill=colors[j])
            suffix = "%" if "Margin" in label else ""
            d.text((x0 + bw / 2 - _text_size(d, f"{val:.0f}{suffix}", font(12))[0] / 2, y0 - 18),
                   f"{val:.0f}{suffix}", fill=colors[j], font=font(12, True))
        d.text((center - _text_size(d, label, font(13, True))[0] / 2, bottom + 16), label, fill=DARK, font=font(13, True))
    for j, (name, color) in enumerate([("Old FY26E", "#9CA3AF"), ("New FY26E", TEAL), ("FY27E", GREEN)]):
        d.rectangle((930 + j * 130, 104, 950 + j * 130, 124), fill=color)
        d.text((958 + j * 130, 101), name, fill=DARK, font=font(14))
    d.rounded_rectangle((865, 520, 1235, 642), radius=10, fill=LIGHT_BLUE, outline=TEAL, width=2)
    d.text((890, 540), "Rating: BUY / Overweight", fill=BLUE, font=font(20, True))
    d.text((890, 570), "Price target: $1,250 -> $1,600", fill=BLUE, font=font(20, True))
    d.text((890, 600), "Implied upside vs. $1,154 close: ~39%", fill=BLUE, font=font(16))
    return save(img, "mu_chart10_estimates_valuation.png", "Company filings; FactSet/IBD consensus; Codex analysis")


def sca_chart() -> Path:
    labels = ["FQ3 RPO", "Signed SCA\nRPO", "Financial\nCommitments", "Cash\nDeposits"]
    values = [SCAS["rpo_q3_b"], SCAS["rpo_signed_b"], SCAS["cash_commitments_b"], SCAS["cash_deposits_b"]]
    return bar_chart(
        "mu_chart8_sca_visibility.png",
        "Strategic Customer Agreements Increase Visibility",
        labels,
        values,
        "B",
        "Micron Q3 FY2026 prepared remarks, June 24, 2026",
        subtitle="Remaining performance obligations and expected customer deposits from signed SCAs",
    )


def guidance_chart() -> Path:
    img, d = canvas("Q4 FY2026 Guidance vs. Street", "Management guided well above FactSet consensus")
    left, top, right, bottom = _plot_area(d)
    labels = ["Revenue\n($B)", "Gross Margin\n(%)", "Non-GAAP EPS\n($)"]
    actual = [RESULTS["revenue_b"], RESULTS["non_gaap_gross_margin_pct"], RESULTS["non_gaap_eps"]]
    guide = [GUIDANCE_Q4["revenue_b_mid"], GUIDANCE_Q4["non_gaap_gross_margin_pct"], GUIDANCE_Q4["non_gaap_eps_mid"]]
    street = [CONSENSUS["q4_revenue_b"], 0, CONSENSUS["q4_eps"]]
    max_v = 95
    slot = (right - left) / len(labels)
    bw = slot * 0.22
    for i in range(6):
        y = bottom - (bottom - top) * i / 5
        d.line((left, y, right, y), fill=LIGHT_GRAY, width=1)
    for i, label in enumerate(labels):
        center = left + slot * i + slot / 2
        vals = [actual[i], guide[i], street[i]]
        colors = ["#9CA3AF", TEAL, RED]
        for j, val in enumerate(vals):
            if val <= 0:
                continue
            x0 = center + (j - 1) * bw * 1.25 - bw / 2
            y0 = bottom - (val / max_v) * (bottom - top)
            d.rounded_rectangle((x0, y0, x0 + bw, bottom), radius=5, fill=colors[j])
            d.text((x0 + bw / 2 - _text_size(d, f"{val:.1f}", font(13))[0] / 2, y0 - 20),
                   f"{val:.1f}", fill=colors[j], font=font(13, True))
        for k, part in enumerate(label.split("\n")):
            d.text((center - _text_size(d, part, font(14, True))[0] / 2, bottom + 14 + k * 16), part, fill=DARK, font=font(14, True))
    for j, (name, color) in enumerate([("Q3 Actual", "#9CA3AF"), ("Q4 Guide", TEAL), ("Q4 Street", RED)]):
        d.rectangle((920 + j * 125, 104, 940 + j * 125, 124), fill=color)
        d.text((948 + j * 125, 101), name, fill=DARK, font=font(14))
    return save(img, "mu_chart9_guidance.png", "Micron Q3 FY2026 press release; FactSet consensus via IBD")


def build_all_charts() -> list[Path]:
    paths = [
        bar_chart(
            "mu_chart1_revenue.png",
            "Micron Quarterly Revenue",
            QUARTERS,
            REVENUE_TREND,
            "B",
            "Micron earnings releases and Q3 FY2026 Form 10-Q",
            consensus=CONSENSUS["revenue_b"],
            subtitle="Q3 FY2026 revenue of $41.46B beat FactSet consensus by $5.55B (+15%)",
        ),
        bar_chart(
            "mu_chart2_eps.png",
            "Non-GAAP EPS Progression",
            QUARTERS,
            EPS_TREND,
            "",
            "Micron earnings releases; FactSet consensus via IBD",
            consensus=CONSENSUS["eps"],
            subtitle="Q3 FY2026 non-GAAP EPS of $25.11 beat consensus by $4.25 (+20%)",
        ),
        line_chart(
            "mu_chart3_margins.png",
            "Margin Step-Change",
            QUARTERS,
            [
                ("Gross margin", GROSS_MARGIN_TREND, TEAL),
                ("Operating margin", OPERATING_MARGIN_TREND, ORANGE),
            ],
            "%",
            "Micron Q3 FY2026 press release and investor presentation",
            subtitle="Non-GAAP gross margin reached a company-record 84.9%",
        ),
        grouped_bar_chart(
            "mu_chart4_dram_nand.png",
            "DRAM and NAND Revenue",
            QUARTERS,
            DRAM_REV,
            NAND_REV,
            "DRAM",
            "NAND",
            "B",
            "Micron Q3 FY2026 investor presentation, slides 31-33",
            subtitle="Q3 DRAM revenue +343% YoY; NAND revenue +361% YoY",
        ),
        horizontal_bar_chart(
            "mu_chart5_business_units.png",
            "Q3 FY2026 Revenue by Business Unit",
            BUSINESS_UNITS,
            "Micron Q3 FY2026 press release and investor presentation, slides 22-24",
            subtitle="All four reporting units set revenue records",
        ),
        grouped_bar_chart(
            "mu_chart6_beat_miss.png",
            "Q3 Results vs. Consensus",
            ["Revenue\n($B)", "Non-GAAP\nEPS"],
            [RESULTS["revenue_b"], RESULTS["non_gaap_eps"]],
            [CONSENSUS["revenue_b"], CONSENSUS["eps"]],
            "Actual",
            "Consensus",
            "",
            "Micron Q3 FY2026 press release; FactSet consensus via IBD",
            subtitle="Beat across the two most-watched headline metrics",
        ),
        grouped_bar_chart(
            "mu_chart7_cash_flow_capex.png",
            "Cash Flow and Capital Intensity",
            QUARTERS,
            OCF_TREND,
            CAPEX_TREND,
            "Operating cash flow",
            "Net capex",
            "B",
            "Micron earnings releases and Q3 FY2026 prepared remarks",
            subtitle="Q3 adjusted FCF was a quarterly record at $18.3B",
        ),
        sca_chart(),
        guidance_chart(),
        estimate_chart(),
    ]
    return paths
