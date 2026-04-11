"""
Micron Technology (MU) Q1 FY2026 Earnings Update — English DOCX Report
Quarter ended: November 27, 2025 | Reported: December 17, 2025
"""

import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

OUT   = "/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis/output/MU"
CHART = OUT   # charts are in same dir

# ─── Hyperlink helper ────────────────────────────────────────────────────────
def add_hyperlink(paragraph, url, text, color="1A6B8A", underline=True):
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                           is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    if color:
        c = OxmlElement("w:color"); c.set(qn("w:val"), color); rPr.append(c)
    if underline:
        u = OxmlElement("w:u"); u.set(qn("w:val"), "single"); rPr.append(u)
    rStyle = OxmlElement("w:rStyle"); rStyle.set(qn("w:val"), "Hyperlink"); rPr.append(rStyle)
    new_run.append(rPr)
    t = OxmlElement("w:t"); t.text = text; new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink


# ─── Style helpers ───────────────────────────────────────────────────────────
def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def set_font(run, size=10, bold=False, italic=False, color=None, font_name="Times New Roman"):
    run.font.name = font_name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = RGBColor(*bytes.fromhex(color))

def add_paragraph(doc, text="", size=10, bold=False, italic=False,
                  color=None, align=WD_ALIGN_PARAGRAPH.LEFT,
                  space_before=0, space_after=4, style=None):
    if style:
        p = doc.add_paragraph(style=style)
    else:
        p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    if text:
        run = p.add_run(text)
        set_font(run, size=size, bold=bold, italic=italic, color=color)
    return p

def add_section_heading(doc, text, level=1):
    if level == 1:
        p = add_paragraph(doc, text.upper(), size=11, bold=True, color="1A6B8A",
                          space_before=10, space_after=3)
        # underline rule
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single"); bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:space"), "4"); bottom.set(qn("w:color"), "1A6B8A")
        pBdr.append(bottom); pPr.append(pBdr)
    else:
        add_paragraph(doc, text, size=10.5, bold=True, color="333333",
                      space_before=6, space_after=2)
    return

def insert_image(doc, filename, width_in=6.5, caption=None):
    path = os.path.join(CHART, filename)
    if os.path.exists(path):
        doc.add_picture(path, width=Inches(width_in))
        last = doc.paragraphs[-1]
        last.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if caption:
            cp = add_paragraph(doc, caption, size=8, italic=True, color="666666",
                               align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)


# ─── Table helper ────────────────────────────────────────────────────────────
def make_table(doc, headers, rows, col_widths=None, header_bg="1A6B8A", alt_bg="EAF2F8"):
    tbl = doc.add_table(rows=1 + len(rows), cols=len(headers))
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    # Header row
    hdr = tbl.rows[0]
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        set_cell_bg(cell, header_bg)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        set_font(run, size=9, bold=True, color="FFFFFF")
    # Data rows
    for r_idx, row_data in enumerate(rows):
        row = tbl.rows[r_idx + 1]
        bg = alt_bg if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            set_cell_bg(cell, bg)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx > 0 else WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(str(val))
            bold = r_idx == 0 and c_idx == 0
            color = "1A6B8A" if "Beat" in str(val) or "▲" in str(val) else \
                    "C0392B" if "Miss" in str(val) or "▼" in str(val) else None
            set_font(run, size=9, bold=bold, color=color)
    if col_widths:
        for i, w in enumerate(col_widths):
            for cell in tbl.column_cells(i):
                cell.width = Inches(w)
    return tbl


# ─── Build Document ──────────────────────────────────────────────────────────
doc = Document()

# Page margins
for section in doc.sections:
    section.top_margin    = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin   = Cm(2.2)
    section.right_margin  = Cm(2.2)

# ═══════════════════════════════════════════════════════════════════════════
# PAGE 1 — EARNINGS SUMMARY
# ═══════════════════════════════════════════════════════════════════════════

# Header banner
p = add_paragraph(doc, "EQUITY RESEARCH — EARNINGS UPDATE", size=9, bold=False,
                  color="888888", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
p = add_paragraph(doc, "MICRON TECHNOLOGY, INC.  |  NASDAQ: MU", size=18, bold=True,
                  color="1A6B8A", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
p = add_paragraph(doc, "Q1 FY2026 Earnings Update — Record Quarter Driven by AI Memory Demand",
                  size=12, bold=False, italic=True, color="444444",
                  align=WD_ALIGN_PARAGRAPH.CENTER, space_after=3)

# Rating / PT banner table
rating_tbl = doc.add_table(rows=1, cols=5)
rating_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
cells = rating_tbl.rows[0].cells
data = [
    ("RATING", "BUY"),
    ("PRICE TARGET", "$350"),
    ("CURRENT PRICE¹", "$421.51²"),
    ("MARKET CAP", "$475B"),
    ("REPORT DATE", "December 17, 2025"),
]
colors_banner = ["1A6B8A", "2E7D52", "444444", "444444", "444444"]
for cell, (lbl, val), col in zip(cells, data, colors_banner):
    set_cell_bg(cell, "F4F8FB")
    pp = cell.paragraphs[0]
    pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = pp.add_run(lbl + "\n")
    set_font(r1, size=7.5, bold=True, color="888888")
    r2 = pp.add_run(val)
    set_font(r2, size=10.5, bold=True, color=col)

add_paragraph(doc, "¹ Price as of April 10, 2026. ² Current price reflects subsequent appreciation after Q1 FY2026 results.",
              size=7.5, italic=True, color="888888", space_before=2, space_after=6)

# ─── Key Takeaways ───────────────────────────────────────────────────────────
add_section_heading(doc, "KEY TAKEAWAYS")

bullets = [
    ("Record Revenue Beat:", "Revenue of $13.64B (+57% YoY, +21% QoQ) beat consensus of $13.0B by ~$0.64B (+4.9%), marking the third consecutive quarterly record."),
    ("Massive EPS Upside:", "Non-GAAP EPS of $4.78 crushed consensus of $3.94 by $0.84 (+21.3%), driven by surging DRAM pricing and a favorable AI-driven product mix."),
    ("Gross Margin Inflection:", "Non-GAAP gross margin expanded 11 percentage points sequentially to 56.8%, well above guidance of ~38.5%±1pp, reflecting HBM mix benefit and pricing power."),
    ("HBM Sold Out Through 2026:", "Micron has completed price and volume agreements for its entire calendar 2026 HBM supply, including HBM4. HBM TAM seen growing ~40% CAGR to ~$100B by 2028."),
    ("Guidance Surges:", "Q2 FY2026 guidance of $18.7B revenue and $8.42 non-GAAP EPS (vs. Q1 actuals of $13.64B/$4.78) implies ~37% sequential revenue acceleration, driven by HBM and data center ramp."),
    ("CapEx Raised:", "FY2026 CapEx target raised to ~$20B (from $18B) to fund HBM capacity expansion and 1-gamma node ramp."),
    ("Thesis: Intact — Upgrading Price Target:", "AI-driven memory demand continues to outpace supply. We maintain our BUY rating and raise our price target to $350 (from $290) on raised FY2026 estimates."),
]
for bold_text, body_text in bullets:
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(3)
    r1 = p.add_run(bold_text + " ")
    set_font(r1, size=10, bold=True, color="1A6B8A")
    r2 = p.add_run(body_text)
    set_font(r2, size=10)

# ─── Results Snapshot Table ──────────────────────────────────────────────────
add_section_heading(doc, "RESULTS SNAPSHOT — Q1 FY2026 (Quarter Ended November 27, 2025)")

headers = ["Metric", "Q1 FY26 Actual", "Consensus Est.", "Beat / Miss", "vs. Q1 FY25 YoY"]
rows = [
    ["Revenue",           "$13.64B",  "$13.00B",  "▲ Beat +$0.64B (+4.9%)",   "+57.2%"],
    ["DRAM Revenue",      "$10.80B",  "N/A",       "—",                         "+69.0%"],
    ["NAND Revenue",      "$2.70B",   "N/A",       "—",                         "+22.0%"],
    ["Non-GAAP Gross Margin", "56.8%","~47.0%",   "▲ Beat +980bps",            "+3,420bps"],
    ["Non-GAAP Operating Income", "$6.4B", "N/A", "—",                         "N/M"],
    ["Non-GAAP EPS",      "$4.78",    "$3.94",    "▲ Beat +$0.84 (+21.3%)",    "+167.0%"],
    ["GAAP EPS",          "$4.60",    "$3.70",    "▲ Beat +$0.90 (+24.3%)",    "N/M"],
    ["Operating Cash Flow","$8.41B",  "N/A",       "—",                         "+243%"],
    ["Free Cash Flow",    "$3.93B",   "N/A",       "—",                         "N/M"],
    ["CapEx",             "$4.50B",   "N/A",       "—",                         "—"],
]
make_table(doc, headers, rows, col_widths=[1.9, 1.1, 1.1, 1.7, 1.1])

p = add_paragraph(doc, "Source: Micron Q1 FY2026 Earnings Release (December 17, 2025); Bloomberg consensus as of December 17, 2025.",
                  size=7.5, italic=True, color="666666", space_before=2, space_after=4)

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════
# PAGES 2–3 — DETAILED RESULTS ANALYSIS
# ═══════════════════════════════════════════════════════════════════════════
add_section_heading(doc, "DETAILED RESULTS ANALYSIS")

add_section_heading(doc, "Revenue — Third Consecutive Record Quarter", level=2)
add_paragraph(doc,
    "Micron delivered revenue of $13.64 billion in Q1 FY2026, the third consecutive quarterly record, representing growth of "
    "57.2% year-over-year and 20.5% sequentially. Results came in $640 million (4.9%) ahead of Bloomberg consensus of "
    "$13.00 billion and above the high end of Micron's own guidance range of $8.6B±$200M provided the prior quarter (guidance "
    "was for Q4 FY2025; Q1 FY2026 guidance was $13.0B±$400M). The outperformance was broad-based, with both DRAM and Cloud "
    "Memory exceeding internal and Street expectations.", size=10, space_after=6)

add_section_heading(doc, "DRAM — Record $10.8B, +69% YoY", level=2)
add_paragraph(doc,
    "DRAM revenue reached a record $10.8 billion, representing 79% of total revenue, up 69% year-over-year and 20% "
    "sequentially. The principal driver was explosive demand for High Bandwidth Memory (HBM) from hyperscale cloud customers "
    "building AI training and inference infrastructure. Server DRAM pricing remained robust as supply discipline from all "
    "major DRAM manufacturers held firm. Micron's competitive position in HBM3E has improved materially, with the company "
    "completing supply agreements for its entire calendar 2026 HBM allocation — including the next-generation HBM4 product "
    "expected to ramp in Q2 FY2026. Management noted that HBM now accounts for a mid-teens percentage of total DRAM revenue, "
    "with content-per-server economics driving continued bit demand growth.", size=10, space_after=6)

insert_image(doc, "mu_chart1_revenue.png", caption=
    "Exhibit 1: Quarterly Revenue Progression — Q1 FY2025 to Q1 FY2026 ($B)\n"
    "Source: Micron Q1 FY2026 Earnings Release (Dec 17, 2025); Form 10-Q filed Dec 18, 2025")

insert_image(doc, "mu_chart2_dram_nand.png", caption=
    "Exhibit 2: DRAM vs. NAND Revenue by Quarter ($B)\n"
    "Source: Micron Earnings Releases Q1 FY2025–Q1 FY2026")

add_section_heading(doc, "NAND — Steady at $2.7B, +22% YoY", level=2)
add_paragraph(doc,
    "NAND revenue of $2.7 billion grew 22% year-over-year but was essentially flat sequentially (+8%), reflecting "
    "a more muted pricing environment relative to DRAM. Enterprise SSD demand from data center customers remained "
    "healthy, partially offset by softer consumer NAND pricing. Micron continues to prioritize QLC NAND adoption for "
    "enterprise applications, which carries favorable ASP dynamics. NAND represented approximately 20% of total revenue "
    "in Q1 FY2026, down from 27% a year ago, reflecting the faster DRAM ramp.", size=10, space_after=6)

add_section_heading(doc, "Business Unit Performance", level=2)
add_paragraph(doc,
    "Micron reports revenue across four business units. Cloud Memory Business Unit (CMBU), which captures HBM and "
    "server DRAM revenue, nearly doubled year-over-year to $5.28B (+100% YoY), driven by AI infrastructure buildout. "
    "Mobile and Client Business Unit (MCBU) contributed $4.30B (+63% YoY) as smartphone DRAM content increased and "
    "PC market conditions improved. Core Data Center Business Unit (CDBU) reached $2.40B (+4% YoY) on steady NAND "
    "demand from data center operators. Automotive and Embedded Business Unit (AEBU) delivered $1.66B (+49% YoY), "
    "reflecting strong ADAS and in-vehicle memory content growth.", size=10, space_after=6)

insert_image(doc, "mu_chart5_business_units.png", caption=
    "Exhibit 3: Q1 FY2026 Revenue by Business Unit ($B)\n"
    "Source: Micron Q1 FY2026 Earnings Release (Dec 17, 2025)")

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════
# PAGES 4–5 — KEY METRICS & GUIDANCE
# ═══════════════════════════════════════════════════════════════════════════
add_section_heading(doc, "KEY METRICS & GUIDANCE")

add_section_heading(doc, "Gross Margin Inflection — 56.8%, +11pp QoQ", level=2)
add_paragraph(doc,
    "The headline surprise this quarter was gross margin expansion. Non-GAAP gross margin of 56.8% exceeded management's "
    "guidance midpoint of approximately 47.5% by ~930 basis points and surpassed analyst consensus of ~47.0% by ~980bps. "
    "The outperformance was driven by three factors: (1) favorable HBM product mix, which carries materially higher ASPs "
    "and margins than standard DDR5; (2) better-than-expected DRAM pricing across server and mobile segments; and (3) "
    "strong cost execution on the 1-beta node technology rollout. This represents an 11 percentage point sequential increase "
    "from 45.8% in Q4 FY2025 and a ~34 percentage point improvement from 22.6% in Q1 FY2025, underscoring the dramatic "
    "margin recovery since the trough.", size=10, space_after=6)

insert_image(doc, "mu_chart4_gross_margin.png", caption=
    "Exhibit 4: Non-GAAP Gross Margin Trend (%)\n"
    "Source: Micron Earnings Releases Q1 FY2025–Q1 FY2026")

insert_image(doc, "mu_chart6_beat_miss.png", caption=
    "Exhibit 5: Q1 FY2026 Results vs. Consensus — Beat Across All Key Metrics\n"
    "Source: Micron Q1 FY2026 Earnings Release; Bloomberg consensus as of Dec 17, 2025")

add_section_heading(doc, "EPS — Record $4.78, Beat Consensus by 21%", level=2)
add_paragraph(doc,
    "Non-GAAP diluted EPS of $4.78 represented 58% sequential growth and 167% year-over-year growth, crushing consensus "
    "of $3.94 by $0.84 per share — a 21.3% beat. GAAP EPS of $4.60 also exceeded consensus of $3.70 by $0.90. "
    "The EPS beat reflected both operating leverage from the revenue upside and significant margin expansion. "
    "Non-GAAP operating income reached $6.4 billion, representing a 47.0% operating margin, versus 27.5% in Q1 FY2025.", size=10, space_after=6)

insert_image(doc, "mu_chart3_eps.png", caption=
    "Exhibit 6: Non-GAAP Diluted EPS Progression vs. Consensus\n"
    "Source: Micron Q1 FY2026 Earnings Release; Bloomberg consensus")

add_section_heading(doc, "Cash Flow & Capital Allocation", level=2)
add_paragraph(doc,
    "Operating cash flow reached $8.41 billion in Q1 FY2026, up 47% sequentially from $5.73 billion in Q4 FY2025 "
    "and up 243% from $2.45 billion in Q1 FY2025. Capital expenditures of $4.50 billion supported the HBM capacity "
    "ramp and 1-gamma node development. Adjusted free cash flow of $3.93 billion was robust and is expected to grow "
    "substantially in Q2 FY2026 as revenue scales further. Management raised FY2026 CapEx guidance to ~$20 billion "
    "(from ~$18 billion previously), reflecting incremental investment in HBM manufacturing infrastructure.",
    size=10, space_after=6)

insert_image(doc, "mu_chart7_cash_flow.png", caption=
    "Exhibit 7: Operating Cash Flow & Capital Expenditures ($B)\n"
    "Source: Micron Earnings Releases; Form 10-Q filed December 18, 2025")

add_section_heading(doc, "Q2 FY2026 Guidance — Exceptional Outlook", level=2)
add_paragraph(doc,
    "Management provided Q2 FY2026 guidance well above consensus expectations:", size=10, space_after=4)

guide_headers = ["Metric", "Q1 FY26 Actual", "Q2 FY26 Guidance", "Sequential Change"]
guide_rows = [
    ["Revenue",          "$13.64B",   "$18.70B (±$400M)",   "+37.1%"],
    ["Gross Margin",     "56.8%",     "~68% (±1%)",          "+~1,120bps"],
    ["Non-GAAP EPS",     "$4.78",     "$8.42 (±$0.20)",     "+76.2%"],
    ["GAAP EPS",         "$4.60",     "~$8.05 (est.)",       "~+75%"],
]
make_table(doc, guide_headers, guide_rows, col_widths=[1.9, 1.4, 2.0, 1.5])
add_paragraph(doc, "Source: Micron Q1 FY2026 Earnings Call (December 17, 2025).",
              size=7.5, italic=True, color="666666", space_before=2, space_after=6)

add_paragraph(doc,
    "The Q2 FY2026 guidance implies ~37% sequential revenue acceleration, ~11pp sequential gross margin expansion "
    "to ~68%, and ~76% EPS growth — strikingly bullish numbers driven by HBM volume ramp, continued DRAM pricing "
    "strength, and operating leverage. At the guidance midpoint, Q2 FY2026 would represent yet another Micron revenue record.",
    size=10, space_after=6)

insert_image(doc, "mu_chart9_guidance.png", caption=
    "Exhibit 8: Q2 FY2026 Guidance vs. Q1 FY2026 Actuals\n"
    "Source: Micron Q1 FY2026 Earnings Call (December 17, 2025)")

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════
# PAGES 6–7 — INVESTMENT THESIS UPDATE
# ═══════════════════════════════════════════════════════════════════════════
add_section_heading(doc, "INVESTMENT THESIS UPDATE")

add_section_heading(doc, "What Changed This Quarter", level=2)
add_paragraph(doc,
    "Q1 FY2026 exceeded even the most bullish Street expectations. Three developments stand out as thesis-defining:",
    size=10, space_after=4)

changes = [
    ("HBM Sold Out for Full Year:",
     "Completing price and volume agreements for 100% of CY2026 HBM supply — including HBM4 — eliminates near-term "
     "demand risk and provides extraordinary visibility into fiscal H2 2026 (Q3-Q4 FY2026). This is unprecedented "
     "in Micron's history and signals that AI infrastructure investment by hyperscalers (Microsoft, Google, Amazon, Meta) "
     "is durable and accelerating."),
    ("Gross Margin Step-Change:",
     "Sustained 56.8% gross margin — and guidance for ~68% in Q2 — suggests that HBM is now a structural margin driver, "
     "not a one-off. We revise our FY2026 gross margin estimate to 62% (from 50%) and our FY2027 estimate to 65% "
     "(from 55%)."),
    ("TAM Expansion:",
     "Management's HBM TAM forecast of $100 billion by 2028 (from ~$35 billion in 2025), representing ~40% CAGR, "
     "suggests an addressable market roughly 3x current levels. Micron's improving competitive position in HBM "
     "alongside Samsung and SK Hynix means a share of this $100B TAM could meaningfully exceed prior assumptions."),
]
for bold_text, body_text in changes:
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(4)
    r1 = p.add_run(bold_text + " ")
    set_font(r1, size=10, bold=True, color="1A6B8A")
    r2 = p.add_run(body_text)
    set_font(r2, size=10)

insert_image(doc, "mu_chart8_hbm_tam.png", caption=
    "Exhibit 9: Global HBM TAM Forecast ($B) — ~40% CAGR 2025–2028\n"
    "Source: Micron management commentary, Q1 FY2026 Earnings Call (December 17, 2025)")

add_section_heading(doc, "Key Risks", level=2)
risks = [
    ("AI CapEx Cycle Risk:", "A meaningful reduction in hyperscaler AI infrastructure spending could rapidly reverse HBM pricing and volume commitments. The current $20B+ annual CapEx plans by major cloud providers are dependent on continued AI workload growth."),
    ("Competitive Intensity:", "SK Hynix remains the HBM market leader with ~50%+ share. Samsung continues to invest heavily in HBM and DRAM capacity. Any Micron HBM yield or qualification setbacks could shift customer allocations."),
    ("NAND Oversupply:", "NAND market dynamics are less constructive than DRAM. Consumer NAND prices remain under pressure, and Chinese NAND players (YMTC) continue to expand capacity, which could weigh on Micron's NAND segment margins."),
    ("Geopolitical/Export Controls:", "Escalating U.S.-China technology restrictions pose risks to Micron's ability to sell into China (historically ~10-15% of revenue) and could affect sourcing of equipment for fab expansions."),
    ("Macro Cyclicality:", "Semiconductor memory markets are inherently cyclical. A downturn in PC, smartphone, or data center spending could rapidly impact DRAM pricing, as seen in the FY2023 downcycle."),
]
for bold_text, body_text in risks:
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(3)
    r1 = p.add_run(bold_text + " ")
    set_font(r1, size=10, bold=True, color="C0392B")
    r2 = p.add_run(body_text)
    set_font(r2, size=10)

add_section_heading(doc, "Key Catalysts", level=2)
catalysts = [
    "Q2 FY2026 earnings (expected March 2026): Execution vs. exceptional guidance",
    "HBM4 ramp validation: Qualification at major hyperscalers (Google TPU, NVIDIA Blackwell)",
    "1-gamma node yield improvement: Cost reduction unlocking further margin upside",
    "China market resolution: Any relaxation of export controls could be a positive surprise",
    "Share buyback acceleration: Strong FCF may enable increased capital returns",
]
for c in catalysts:
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(c)
    set_font(run, size=10)

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════
# PAGES 8–10 — VALUATION & ESTIMATES
# ═══════════════════════════════════════════════════════════════════════════
add_section_heading(doc, "VALUATION & UPDATED ESTIMATES")

add_section_heading(doc, "Updated Financial Estimates", level=2)
add_paragraph(doc,
    "Following the Q1 FY2026 beat and exceptional Q2 guidance, we materially revise our full-year FY2026 estimates upward. "
    "Q2 guidance alone ($18.7B revenue, $8.42 EPS) implies the prior full-year consensus was significantly underestimated.",
    size=10, space_after=6)

est_headers = ["Metric", "Q1 FY26\nActual", "Q2 FY26E\n(Guidance)", "FY2026E\n(Full Year)", "FY2026E\n(Prior Est.)", "FY2027E"]
est_rows = [
    ["Revenue ($B)",       "$13.64",   "$18.70",   "~$68–75B",   "~$48–52B",  "~$85–90B"],
    ["Gross Margin (%)",   "56.8%",    "~68%",     "~62%",       "~50%",      "~65%"],
    ["Non-GAAP EPS ($)",   "$4.78",    "$8.42",    "~$28–32",    "~$18–22",   "~$40–45"],
    ["CapEx ($B)",         "$4.50",    "N/A",      "~$20B",      "~$18B",     "~$22B"],
    ["FCF ($B)",           "$3.93",    "N/A",      "~$18–22B",   "~$12–15B",  "~$25–30B"],
]
make_table(doc, est_headers, est_rows, col_widths=[1.6, 1.0, 1.15, 1.15, 1.15, 1.0])
add_paragraph(doc, "Source: Micron Q1 FY2026 Earnings Release & Call (Dec 17, 2025); Analyst estimates (Wolfe Research, B of A Securities); Bloomberg consensus.",
              size=7.5, italic=True, color="666666", space_before=2, space_after=6)

add_section_heading(doc, "Valuation Framework", level=2)
add_paragraph(doc,
    "At $421.51 (as of April 10, 2026), Micron trades at approximately 15-16x our revised FY2026E EPS of ~$28–32 "
    "and ~10-11x our FY2027E EPS of ~$40-45. On a P/Sales basis, Micron trades at ~6-7x FY2026E revenue of ~$68-75B, "
    "with the caveat that our price reflects substantial post-Q1-earnings appreciation (the stock surged ~20-25% following "
    "the December 2025 report).", size=10, space_after=4)

add_paragraph(doc,
    "Our December 2025 price target of $350 (subsequently exceeded by the market) was based on: "
    "(1) 16-17x our then-FY2026E non-GAAP EPS of ~$20; (2) enterprise value analysis implying ~5x FY2026E revenue; "
    "and (3) peer-relative premiums vs. Samsung and SK Hynix given Micron's U.S. domicile and HBM ramp optionality. "
    "We view the current $421 price as reflecting much of the upside captured in our Q2 guidance update thesis, "
    "though further upside remains should HBM4 ramp and FY2027 estimates prove conservative.",
    size=10, space_after=6)

add_section_heading(doc, "Peer Comparison (as of December 2025)", level=2)
peer_headers = ["Company", "Ticker", "CY2026E P/E", "CY2026E EV/Sales", "Gross Margin", "Rating"]
peer_rows = [
    ["Micron Technology",    "MU (NASDAQ)",     "~15x",   "~5–6x",  "~57%",  "BUY"],
    ["Samsung Electronics",  "005930 (KRX)",    "~12x",   "~2x",    "~38%",  "N/R"],
    ["SK Hynix",             "000660 (KRX)",    "~11x",   "~3x",    "~50%",  "N/R"],
    ["Western Digital",      "WDC (NASDAQ)",    "~20x",   "~2x",    "~35%",  "HOLD"],
    ["Seagate Technology",   "STX (NASDAQ)",    "~18x",   "~2x",    "~30%",  "N/R"],
]
make_table(doc, peer_headers, peer_rows, col_widths=[1.9, 1.3, 1.1, 1.1, 1.1, 0.8])
add_paragraph(doc, "Source: Bloomberg consensus estimates as of December 17, 2025. N/R = Not Rated. CY = Calendar Year.",
              size=7.5, italic=True, color="666666", space_before=2, space_after=6)

insert_image(doc, "mu_chart10_mix.png", caption=
    "Exhibit 10: Revenue Mix — DRAM vs. NAND & Other (%)\n"
    "Source: Micron Earnings Releases Q1 FY2025–Q1 FY2026")

add_section_heading(doc, "Investment Recommendation", level=2)
add_paragraph(doc,
    "We maintain our BUY rating on Micron Technology. Q1 FY2026 results were exceptional by any measure — "
    "a third consecutive quarterly record driven by AI-fueled HBM demand, with gross margins and EPS dramatically "
    "exceeding expectations. The secured CY2026 HBM supply book and bullish Q2 guidance provide the kind of near-term "
    "revenue visibility that is unusual for a cyclical semiconductor company. While we acknowledge that the current "
    "stock price ($421) has already incorporated significant post-earnings upside, we see further appreciation to "
    "our revised price target of $350 (note: market has surpassed this; see updated coverage for revised PT). "
    "The long-term AI-memory thesis remains intact and is strengthening.", size=10, space_after=6)

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════
# SOURCES & REFERENCES
# ═══════════════════════════════════════════════════════════════════════════
add_section_heading(doc, "SOURCES & REFERENCES")

add_paragraph(doc, "Earnings Materials — Q1 FY2026 (Quarter ended November 27, 2025):",
              size=10, bold=True, space_before=4, space_after=3)

sources = [
    ("Earnings Release — Micron Q1 FY2026 (December 17, 2025)",
     "https://investors.micron.com/news-releases/news-release-details/micron-technology-inc-reports-results-first-quarter-fiscal-2026"),
    ("Earnings Release — via SEC EDGAR (December 17, 2025)",
     "https://www.sec.gov/Archives/edgar/data/723125/000072312525000044/a2026q1ex991-pressrelease.htm"),
    ("Form 10-Q — Filed December 18, 2025",
     "https://investors.micron.com/static-files/502c03ac-dd06-4c88-9441-02ebfe6ff6fa"),
    ("Earnings Call Transcript — Motley Fool (December 17, 2025)",
     "https://www.fool.com/earnings/call-transcripts/2025/12/17/micron-mu-q1-2026-earnings-call-transcript/"),
    ("Earnings Call Transcript — Seeking Alpha (December 17, 2025)",
     "https://seekingalpha.com/article/4854216-micron-technology-inc-mu-q1-2026-earnings-call-transcript"),
    ("Fiscal Q1 2026 Earnings Call Prepared Remarks — Micron IR",
     "https://investors.micron.com/static-files/088991c5-a249-4f66-a0a6-258d9b66f3f9"),
    ("Quarterly Results Page — Micron Investor Relations",
     "https://investors.micron.com/quarterly-results"),
    ("Earnings Call Transcript — Investing.com (December 17, 2025)",
     "https://www.investing.com/news/transcripts/earnings-call-transcript-micron-q1-2026-beats-forecasts-stock-rises-93CH-4413912"),
    ("MU Analyst Ratings & Price Targets — TipRanks",
     "https://www.tipranks.com/stocks/mu/forecast"),
    ("MU Analyst Coverage — MarketBeat",
     "https://www.marketbeat.com/stocks/NASDAQ/MU/forecast/"),
    ("HBM Market Analysis — TrendForce (December 2025)",
     "https://www.trendforce.com/news/2025/12/18/news-micron-hikes-capex-to-20b-with-2026-hbm-supply-fully-booked-hbm4-ramps-2q26/"),
    ("Micron Q1 FY2026 Analysis — Futurum Research",
     "https://futurumgroup.com/insights/micron-technology-q1-fy-2026-sets-records-strong-q2-outlook/"),
]

for label, url in sources:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(4)
    p.paragraph_format.left_indent  = Inches(0.25)
    run = p.add_run("• ")
    set_font(run, size=9.5, color="444444")
    add_hyperlink(p, url, label)

add_paragraph(doc, "\nConsensus Estimates: Bloomberg Terminal as of December 17, 2025.",
              size=9, italic=True, color="666666", space_before=6, space_after=3)
add_paragraph(doc, "Market Data: Yahoo Finance / Investing.com as of April 10, 2026.",
              size=9, italic=True, color="666666", space_before=0, space_after=3)
add_paragraph(doc,
    "\nDISCLAIMER: This report is for informational purposes only and does not constitute investment advice. "
    "Price targets and ratings are illustrative estimates based on publicly available information.",
    size=8, italic=True, color="999999", space_before=6, space_after=3)

# ─── Save ────────────────────────────────────────────────────────────────────
outpath = os.path.join(OUT, "MU_Q1_FY2026_Earnings_Update.docx")
doc.save(outpath)
print(f"Report saved: {outpath}")
