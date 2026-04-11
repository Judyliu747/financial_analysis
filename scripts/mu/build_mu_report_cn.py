"""
Micron Technology (MU) Q1 FY2026 业绩更新报告 — 中文版
会计季度截止：2025年11月27日 | 发布日期：2025年12月17日
"""

import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT   = "/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis/output/MU"
CHART = OUT

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

def set_font(run, size=10, bold=False, italic=False, color=None,
             font_name="宋体", latin_font="Times New Roman"):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = RGBColor(*bytes.fromhex(color))

def add_paragraph(doc, text="", size=10, bold=False, italic=False,
                  color=None, align=WD_ALIGN_PARAGRAPH.LEFT,
                  space_before=0, space_after=4, heading=False):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    if text:
        run = p.add_run(text)
        fn = "黑体" if heading else "宋体"
        set_font(run, size=size, bold=bold, italic=italic, color=color, font_name=fn)
    return p

def add_section_heading(doc, text, level=1):
    if level == 1:
        p = add_paragraph(doc, text, size=11, bold=True, color="1A6B8A",
                          space_before=10, space_after=3, heading=True)
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single"); bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:space"), "4"); bottom.set(qn("w:color"), "1A6B8A")
        pBdr.append(bottom); pPr.append(pBdr)
    else:
        add_paragraph(doc, text, size=10.5, bold=True, color="333333",
                      space_before=6, space_after=2, heading=True)

def insert_image(doc, filename, width_in=6.5, caption=None):
    path = os.path.join(CHART, filename)
    if os.path.exists(path):
        doc.add_picture(path, width=Inches(width_in))
        last = doc.paragraphs[-1]
        last.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if caption:
            add_paragraph(doc, caption, size=8, italic=True, color="666666",
                          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)

def make_table(doc, headers, rows, col_widths=None, header_bg="1A6B8A", alt_bg="EAF2F8"):
    tbl = doc.add_table(rows=1 + len(rows), cols=len(headers))
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = tbl.rows[0]
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        set_cell_bg(cell, header_bg)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        set_font(run, size=9, bold=True, color="FFFFFF", font_name="黑体")
    for r_idx, row_data in enumerate(rows):
        row = tbl.rows[r_idx + 1]
        bg = alt_bg if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            set_cell_bg(cell, bg)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx > 0 else WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(str(val))
            color = "1A6B8A" if "超预期" in str(val) or "▲" in str(val) else \
                    "C0392B" if "低于" in str(val) or "▼" in str(val) else None
            set_font(run, size=9, color=color)
    if col_widths:
        for i, w in enumerate(col_widths):
            for cell in tbl.column_cells(i):
                cell.width = Inches(w)
    return tbl


# ─── Build Document ──────────────────────────────────────────────────────────
doc = Document()
for section in doc.sections:
    section.top_margin    = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin   = Cm(2.2)
    section.right_margin  = Cm(2.2)

# ═══════════════════════════════════════════════════════════════════════════
# 第一页 — 业绩摘要
# ═══════════════════════════════════════════════════════════════════════════
add_paragraph(doc, "股票研究 — 业绩更新报告", size=9, color="888888",
              align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
add_paragraph(doc, "美光科技（Micron Technology, Inc.）｜ 纳斯达克：MU", size=18, bold=True,
              color="1A6B8A", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2, heading=True)
add_paragraph(doc, "FY2026财年第一季度业绩更新 — 创纪录季度，AI存储需求驱动全面超预期",
              size=12, italic=True, color="444444",
              align=WD_ALIGN_PARAGRAPH.CENTER, space_after=3)

rating_tbl = doc.add_table(rows=1, cols=5)
rating_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
cells = rating_tbl.rows[0].cells
data = [
    ("评级", "买入"),
    ("目标价", "$350"),
    ("当前股价¹", "$421.51²"),
    ("市值", "$4,754亿"),
    ("报告日期", "2025年12月17日"),
]
for cell, (lbl, val) in zip(cells, data):
    set_cell_bg(cell, "F4F8FB")
    pp = cell.paragraphs[0]
    pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = pp.add_run(lbl + "\n")
    set_font(r1, size=7.5, bold=True, color="888888", font_name="黑体")
    r2 = pp.add_run(val)
    set_font(r2, size=10.5, bold=True, color="1A6B8A", font_name="黑体")

add_paragraph(doc, "¹ 股价截至2026年4月10日。² 当前股价反映FY2026Q1业绩公布后的市场上涨。",
              size=7.5, italic=True, color="888888", space_before=2, space_after=6)

# ─── 核心要点 ────────────────────────────────────────────────────────────────
add_section_heading(doc, "核心要点")

bullets = [
    ("营收创纪录，超预期：", "FY2026Q1营收136.4亿美元（同比+57.2%，环比+20.5%），超彭博一致预期130亿美元约6.4亿（+4.9%），连续第三个季度刷新公司纪录。"),
    ("EPS大幅超预期：", "Non-GAAP摊薄每股收益4.78美元，超一致预期3.94美元约0.84美元（+21.3%），主要得益于HBM需求推动的DRAM价格上涨及产品结构改善。"),
    ("毛利率跃升：", "Non-GAAP毛利率56.8%，环比提升约11个百分点，远超中期指引约47.5%约930个基点，创历史高点，反映HBM产品组合带来的结构性利润改善。"),
    ("HBM全年售罄：", "美光已完成2026全年HBM（含HBM4）全部产能的价格及数量协议，供需格局极为有利。管理层预计HBM全球TAM将以约40%复合年增长率扩张，从2025年约350亿美元增至2028年约1,000亿美元。"),
    ("Q2指引远超预期：", "FY2026Q2指引：营收187亿美元（±4亿）、Non-GAAP毛利率约68%（±1ppt）、Non-GAAP EPS 8.42美元（±0.20），环比增速分别为+37%、+11ppt、+76%。"),
    ("资本支出上调：", "FY2026全年资本支出从约180亿美元上调至约200亿美元，主要用于HBM产能扩张及1-gamma节点建设。"),
    ("投资逻辑完整，维持买入：", "AI驱动存储需求持续超预期，维持买入评级，目标价上调至350美元（原290美元）。"),
]
for bold_text, body_text in bullets:
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(3)
    r1 = p.add_run(bold_text + " ")
    set_font(r1, size=10, bold=True, color="1A6B8A", font_name="黑体")
    r2 = p.add_run(body_text)
    set_font(r2, size=10, font_name="宋体")

# ─── 业绩快照 ─────────────────────────────────────────────────────────────────
add_section_heading(doc, "业绩快照 — FY2026Q1（会计季度截止2025年11月27日）")

headers = ["指标", "FY2026Q1实际", "一致预期", "超/低预期", "同比变动"]
rows = [
    ["营收",             "136.4亿美元",  "130.0亿美元",  "▲ 超预期+6.4亿（+4.9%）",  "+57.2%"],
    ["DRAM营收",         "108.0亿美元",  "N/A",           "—",                         "+69.0%"],
    ["NAND营收",         "27.0亿美元",   "N/A",           "—",                         "+22.0%"],
    ["Non-GAAP毛利率",   "56.8%",        "约47.0%",       "▲ 超预期+980bps",           "+3,420bps"],
    ["Non-GAAP营业利润", "64亿美元",     "N/A",           "—",                         "显著改善"],
    ["Non-GAAP EPS",     "$4.78",        "$3.94",         "▲ 超预期+$0.84（+21.3%）", "+167.0%"],
    ["GAAP EPS",         "$4.60",        "$3.70",         "▲ 超预期+$0.90（+24.3%）", "显著改善"],
    ["经营现金流",        "84.1亿美元",   "N/A",           "—",                         "+243%"],
    ["自由现金流",        "39.3亿美元",   "N/A",           "—",                         "显著改善"],
    ["资本支出",          "45.0亿美元",   "N/A",           "—",                         "—"],
]
make_table(doc, headers, rows, col_widths=[1.9, 1.2, 1.0, 1.7, 1.0])
add_paragraph(doc, "资料来源：美光FY2026Q1业绩公告（2025年12月17日）；彭博一致预期（截至2025年12月17日）。",
              size=7.5, italic=True, color="666666", space_before=2, space_after=4)

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════
# 第二–三页 — 详细业绩分析
# ═══════════════════════════════════════════════════════════════════════════
add_section_heading(doc, "详细业绩分析")

add_section_heading(doc, "营收 — 连续第三个季度刷新纪录", level=2)
add_paragraph(doc,
    "美光FY2026Q1实现营收136.4亿美元，同比增长57.2%，环比增长20.5%，超公司自身指引上限约6.4亿美元，"
    "超彭博一致预期约4.9%。营收创第三个季度纪录，主要受益于AI基础设施建设带动的高带宽内存（HBM）需求爆发，"
    "以及服务器DRAM价格持续上行。", size=10, space_after=6)

add_section_heading(doc, "DRAM — 创纪录108亿美元，同比+69%", level=2)
add_paragraph(doc,
    "DRAM营收达108亿美元（占总营收79%），同比增长69%，环比增长20%，再创历史新高。核心驱动力在于：（1）超大型云服务商（微软、谷歌、亚马逊、Meta）AI训练和推理基础设施大规模扩建推动HBM需求；"
    "（2）服务器DDR5 DRAM供需紧张格局维持；（3）主要DRAM厂商普遍维持供给纪律。美光已完成2026全年HBM全部产能的价格及数量协议，包括下一代HBM4（预计FY2026Q2量产）。"
    "管理层表示，HBM目前约占DRAM营收中双位数百分比，AI服务器每台内存用量的持续提升将驱动比特需求长期增长。", size=10, space_after=6)

insert_image(doc, "mu_chart1_revenue.png", caption=
    "图1：美光季度营收走势（FY2025Q1–FY2026Q1，单位：十亿美元）\n"
    "资料来源：美光FY2026Q1业绩公告（2025年12月17日）；SEC Form 10-Q（2025年12月18日提交）")

insert_image(doc, "mu_chart2_dram_nand.png", caption=
    "图2：DRAM与NAND分产品营收（FY2025Q1–FY2026Q1，单位：十亿美元）\n"
    "资料来源：美光各季度业绩公告")

add_section_heading(doc, "NAND — 27亿美元，同比+22%", level=2)
add_paragraph(doc,
    "NAND营收27亿美元，同比增长22%，环比增长约8%，走势较DRAM温和。数据中心企业SSD需求保持稳定，"
    "消费级NAND价格承压，中国NAND厂商（长江存储）扩产持续。NAND占总营收比例从一年前约27%下降至20%，"
    "反映DRAM更快的增速。美光继续推进面向企业应用的QLC NAND，有助于改善产品组合ASP。",
    size=10, space_after=6)

add_section_heading(doc, "业务单元营收拆分", level=2)
add_paragraph(doc,
    "云存储业务单元（CMBU）：HBM及服务器DRAM，营收52.8亿美元，同比翻倍（+100%），毛利率高达66%，"
    "为公司核心利润引擎。核心数据中心业务单元（CDBU）：数据中心NAND及部分DRAM，营收24亿美元（+4% YoY）。"
    "移动与客户端业务单元（MCBU）：营收43亿美元（+63% YoY），受益于智能手机DRAM内存提升和PC市场复苏。"
    "汽车与嵌入式业务单元（AEBU）：营收16.6亿美元（+49% YoY），ADAS和车载存储需求强劲。",
    size=10, space_after=6)

insert_image(doc, "mu_chart5_business_units.png", caption=
    "图3：FY2026Q1各业务单元营收（单位：十亿美元）\n"
    "资料来源：美光FY2026Q1业绩公告（2025年12月17日）")

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════
# 第四–五页 — 核心指标与业绩指引
# ═══════════════════════════════════════════════════════════════════════════
add_section_heading(doc, "核心指标与业绩指引")

add_section_heading(doc, "毛利率 — 56.8%，环比+11个百分点", level=2)
add_paragraph(doc,
    "Non-GAAP毛利率56.8%是本季最大亮点，超公司指引中值约47.5%约930bps，超市场预期约47.0%约980bps。"
    "超预期原因：（1）HBM产品组合比例提升（HBM ASP显著高于标准DDR5）；（2）DRAM服务器及移动端价格好于预期；"
    "（3）1-beta节点成本控制优异。毛利率从FY2025Q1的22.6%大幅跃升至56.8%，约34个百分点的改善清晰呈现"
    "公司经历FY2023低谷后的盈利能力恢复轨迹。",
    size=10, space_after=6)

insert_image(doc, "mu_chart4_gross_margin.png", caption=
    "图4：Non-GAAP毛利率走势（FY2025Q1–FY2026Q1，%）\n"
    "资料来源：美光各季度业绩公告")

insert_image(doc, "mu_chart6_beat_miss.png", caption=
    "图5：FY2026Q1业绩 vs. 一致预期 — 全面超预期\n"
    "资料来源：美光FY2026Q1业绩公告；彭博一致预期（2025年12月17日）")

add_section_heading(doc, "EPS — 4.78美元，超预期21%", level=2)
add_paragraph(doc,
    "Non-GAAP摊薄EPS 4.78美元，环比增长58%，同比增长167%，超一致预期3.94美元0.84美元（+21.3%）。"
    "GAAP摊薄EPS 4.60美元，超一致预期3.70美元。Non-GAAP营业利润64亿美元，营业利润率47.0%（去年同期27.5%）。",
    size=10, space_after=6)

insert_image(doc, "mu_chart3_eps.png", caption=
    "图6：Non-GAAP摊薄EPS走势及一致预期对比\n"
    "资料来源：美光FY2026Q1业绩公告；彭博一致预期")

add_section_heading(doc, "现金流与资本开支", level=2)
add_paragraph(doc,
    "经营现金流84.1亿美元，环比增长47%，同比增长243%，创历史新高。资本支出45亿美元，自由现金流39.3亿美元。"
    "管理层将FY2026全年资本支出指引从约180亿美元上调至约200亿美元，增量主要用于HBM产能扩张及1-gamma节点研发。"
    "公司预计Q2及后续季度自由现金流将显著提升。",
    size=10, space_after=6)

insert_image(doc, "mu_chart7_cash_flow.png", caption=
    "图7：经营现金流与资本支出（FY2025Q1–FY2026Q1，单位：十亿美元）\n"
    "资料来源：美光各季度业绩公告；Form 10-Q（2025年12月18日提交）")

add_section_heading(doc, "FY2026Q2 业绩指引 — 大幅优于预期", level=2)
add_paragraph(doc, "管理层FY2026Q2指引全面超出市场预期：", size=10, space_after=4)

guide_headers = ["指标", "FY2026Q1实际", "FY2026Q2指引", "环比变动"]
guide_rows = [
    ["营收",          "136.4亿美元",   "187亿美元（±4亿）",   "+37.1%"],
    ["Non-GAAP毛利率","56.8%",         "约68%（±1ppt）",       "+约11ppt"],
    ["Non-GAAP EPS",  "$4.78",         "$8.42（±$0.20）",      "+76.2%"],
]
make_table(doc, guide_headers, guide_rows, col_widths=[1.9, 1.5, 2.0, 1.4])
add_paragraph(doc, "资料来源：美光FY2026Q1业绩电话会议（2025年12月17日）。",
              size=7.5, italic=True, color="666666", space_before=2, space_after=6)

add_paragraph(doc,
    "Q2指引意味着营收环比加速约37%、毛利率再升约11个百分点至68%、EPS环比大增76%，"
    "主要驱动力为HBM放量、数据中心DRAM持续提价以及规模效应。按指引中值计算，FY2026Q2将再次刷新美光营收历史纪录。",
    size=10, space_after=6)

insert_image(doc, "mu_chart9_guidance.png", caption=
    "图8：FY2026Q2指引 vs. FY2026Q1实际\n"
    "资料来源：美光FY2026Q1业绩电话会议（2025年12月17日）")

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════
# 第六–七页 — 投资逻辑更新
# ═══════════════════════════════════════════════════════════════════════════
add_section_heading(doc, "投资逻辑更新")

add_section_heading(doc, "本季度的关键变化", level=2)
add_paragraph(doc, "FY2026Q1业绩超出了街道最乐观的预期，以下三点具有定性意义：", size=10, space_after=4)

changes = [
    ("HBM全年售罄，能见度极高：",
     "完成2026全年HBM（含HBM4）全部产能的价格及数量协议，消除了近期需求不确定性，"
     "并为FY2026下半年（Q3-Q4）提供了极强的营收可预期性。这在美光历史上前所未有，"
     "表明超大规模云服务商AI基础设施投入具有高度确定性和持续性。"),
    ("毛利率结构性跃升：",
     "Q1 56.8%、Q2指引68%，呈现出HBM已成为利润结构改变因素的信号，而非一次性。"
     "我们将FY2026全年毛利率预测从50%上调至62%，FY2027从55%上调至65%。"),
    ("TAM显著扩张：",
     "管理层预计HBM全球TAM将从2025年约350亿美元增至2028年约1,000亿美元（约40% CAGR），"
     "约为当前市场规模3倍。美光在HBM领域竞争力的持续提升（HBM4领先）使其受益程度或超出此前预期。"),
]
for bold_text, body_text in changes:
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(4)
    r1 = p.add_run(bold_text + " ")
    set_font(r1, size=10, bold=True, color="1A6B8A", font_name="黑体")
    r2 = p.add_run(body_text)
    set_font(r2, size=10, font_name="宋体")

insert_image(doc, "mu_chart8_hbm_tam.png", caption=
    "图9：全球HBM TAM预测（十亿美元），约40% CAGR（2025–2028）\n"
    "资料来源：美光管理层评论，FY2026Q1业绩电话会议（2025年12月17日）")

add_section_heading(doc, "主要风险", level=2)
risks = [
    ("AI资本支出下行风险：", "云服务商AI基础设施投入若显著收缩，将迅速反转HBM定价与出货量趋势，进而对业绩产生非线性影响。"),
    ("竞争格局加剧：", "SK海力士仍为HBM市场领导者（约50%+市占率），三星持续加大HBM和DRAM产能投入。若美光HBM良率或客户认证出现问题，可能导致份额流失。"),
    ("NAND供过于求：", "NAND市场格局弱于DRAM，消费级NAND价格承压，中国NAND厂商（长江存储）持续扩产，可能拖累美光NAND业务利润率。"),
    ("地缘政治与出口管制：", "中美技术限制升级可能影响美光对华销售（历史上占营收10-15%），并可能波及设备采购和产能扩张计划。"),
    ("半导体行业周期性：", "存储半导体市场具有固有周期性，PC、智能手机或数据中心支出下滑均可能导致DRAM定价快速回落，正如FY2023下行周期所示。"),
]
for bold_text, body_text in risks:
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(3)
    r1 = p.add_run(bold_text + " ")
    set_font(r1, size=10, bold=True, color="C0392B", font_name="黑体")
    r2 = p.add_run(body_text)
    set_font(r2, size=10, font_name="宋体")

add_section_heading(doc, "潜在催化剂", level=2)
catalysts = [
    "FY2026Q2业绩（预计2026年3月公布）：指引兑现情况验证",
    "HBM4量产认证：谷歌TPU、英伟达Blackwell平台客户认证进展",
    "1-gamma节点良率提升：进一步推动成本下降和利润率扩张",
    "中国市场政策变化：若出口管制放宽将是超预期正面催化剂",
    "股票回购提速：强劲自由现金流为加大资本回报奠定基础",
]
for c in catalysts:
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(c)
    set_font(run, size=10, font_name="宋体")

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════
# 第八–十页 — 估值与预测
# ═══════════════════════════════════════════════════════════════════════════
add_section_heading(doc, "估值与盈利预测更新")

add_section_heading(doc, "更新后财务预测", level=2)
add_paragraph(doc,
    "综合FY2026Q1实际业绩及Q2指引，我们大幅上调全年预测。仅Q2指引（营收187亿、EPS 8.42美元）"
    "便已意味着此前全年一致预期存在大幅低估。",
    size=10, space_after=6)

est_headers = ["指标", "FY2026Q1\n实际", "FY2026Q2E\n（指引）", "FY2026E\n（全年）", "FY2026E\n（前次预测）", "FY2027E"]
est_rows = [
    ["营收（十亿美元）",    "$13.64",  "$18.70",  "~$68–75B",  "~$48–52B",  "~$85–90B"],
    ["Non-GAAP毛利率",      "56.8%",   "约68%",   "约62%",     "约50%",     "约65%"],
    ["Non-GAAP EPS",        "$4.78",   "$8.42",   "约$28–32",  "约$18–22",  "约$40–45"],
    ["资本支出（十亿美元）", "$4.50",  "N/A",     "约$20B",    "约$18B",    "约$22B"],
    ["自由现金流（十亿美元）","$3.93", "N/A",     "约$18–22B", "约$12–15B", "约$25–30B"],
]
make_table(doc, est_headers, est_rows, col_widths=[1.6, 1.0, 1.15, 1.15, 1.15, 1.0])
add_paragraph(doc, "资料来源：美光FY2026Q1业绩公告及电话会议（2025年12月17日）；Wolfe Research、美国银行证券分析师预测；彭博一致预期。",
              size=7.5, italic=True, color="666666", space_before=2, space_after=6)

add_section_heading(doc, "估值分析", level=2)
add_paragraph(doc,
    "按2026年4月10日收盘价421.51美元计算，美光当前对应约15-16倍FY2026E Non-GAAP EPS（约28-32美元），"
    "约10-11倍FY2027E EPS（约40-45美元）。市销率约6-7倍FY2026E营收（约680-750亿美元）。"
    "需要注意的是，当前股价已充分反映FY2026Q1业绩公布后的大幅上涨（约20-25%）。",
    size=10, space_after=4)

add_paragraph(doc,
    "我们2025年12月提出的350美元目标价基于：（1）约16-17倍FY2026E Non-GAAP EPS约20美元；"
    "（2）约5倍FY2026E营收的EV/Sales估值；（3）相对三星、SK海力士估值溢价（美国上市地位 + HBM弹性）。"
    "目前市场已超越该目标价，建议参阅最新覆盖报告获取修正后目标价及评级。",
    size=10, space_after=6)

add_section_heading(doc, "可比公司分析（截至2025年12月）", level=2)
peer_headers = ["公司", "股票代码", "CY2026E市盈率", "CY2026E EV/Sales", "毛利率", "评级"]
peer_rows = [
    ["美光科技",      "MU（纳斯达克）",   "约15倍",  "约5–6倍",  "约57%",  "买入"],
    ["三星电子",      "005930（韩交所）", "约12倍",  "约2倍",    "约38%",  "未评级"],
    ["SK海力士",      "000660（韩交所）", "约11倍",  "约3倍",    "约50%",  "未评级"],
    ["西部数据",      "WDC（纳斯达克）",  "约20倍",  "约2倍",    "约35%",  "持有"],
    ["希捷科技",      "STX（纳斯达克）",  "约18倍",  "约2倍",    "约30%",  "未评级"],
]
make_table(doc, peer_headers, peer_rows, col_widths=[1.7, 1.4, 1.1, 1.2, 1.0, 0.9])
add_paragraph(doc, "资料来源：彭博一致预期（截至2025年12月17日）。未评级 = 本报告未覆盖。CY = 日历年度。",
              size=7.5, italic=True, color="666666", space_before=2, space_after=6)

insert_image(doc, "mu_chart10_mix.png", caption=
    "图10：营收结构 — DRAM vs. NAND及其他（%）\n"
    "资料来源：美光各季度业绩公告（FY2025Q1–FY2026Q1）")

add_section_heading(doc, "投资建议", level=2)
add_paragraph(doc,
    "维持美光科技买入评级。FY2026Q1业绩从任何维度衡量均堪称卓越——AI驱动HBM需求爆发推动第三个季度营收破纪录，"
    "毛利率和EPS均大幅超预期。CY2026年HBM产能完全售罄及强劲的Q2指引为近期营收提供了极高能见度，"
    "这在美光历史上极为罕见。当前股价421美元已相当程度反映了后Q1时代的上行空间，"
    "但若HBM4顺利放量且FY2027预测证明偏保守，仍有进一步上升空间。长期AI存储需求主题完整且持续强化。",
    size=10, space_after=6)

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════
# 资料来源
# ═══════════════════════════════════════════════════════════════════════════
add_section_heading(doc, "资料来源与参考文献")

add_paragraph(doc, "FY2026Q1业绩相关资料（会计季度截止2025年11月27日）：",
              size=10, bold=True, space_before=4, space_after=3, heading=True)

sources = [
    ("FY2026Q1业绩公告 — 美光科技投资者关系（2025年12月17日）",
     "https://investors.micron.com/news-releases/news-release-details/micron-technology-inc-reports-results-first-quarter-fiscal-2026"),
    ("FY2026Q1业绩公告 — 美国SEC EDGAR（2025年12月17日）",
     "https://www.sec.gov/Archives/edgar/data/723125/000072312525000044/a2026q1ex991-pressrelease.htm"),
    ("Form 10-Q — 美国证交会（2025年12月18日提交）",
     "https://investors.micron.com/static-files/502c03ac-dd06-4c88-9441-02ebfe6ff6fa"),
    ("FY2026Q1业绩电话会议文字记录 — The Motley Fool（2025年12月17日）",
     "https://www.fool.com/earnings/call-transcripts/2025/12/17/micron-mu-q1-2026-earnings-call-transcript/"),
    ("FY2026Q1业绩电话会议文字记录 — Seeking Alpha（2025年12月17日）",
     "https://seekingalpha.com/article/4854216-micron-technology-inc-mu-q1-2026-earnings-call-transcript"),
    ("FY2026Q1电话会议准备发言稿 — 美光投资者关系",
     "https://investors.micron.com/static-files/088991c5-a249-4f66-a0a6-258d9b66f3f9"),
    ("季度业绩页面 — 美光投资者关系",
     "https://investors.micron.com/quarterly-results"),
    ("业绩电话会议记录 — Investing.com（2025年12月17日）",
     "https://www.investing.com/news/transcripts/earnings-call-transcript-micron-q1-2026-beats-forecasts-stock-rises-93CH-4413912"),
    ("分析师评级与目标价 — TipRanks",
     "https://www.tipranks.com/stocks/mu/forecast"),
    ("分析师覆盖 — MarketBeat",
     "https://www.marketbeat.com/stocks/NASDAQ/MU/forecast/"),
    ("HBM市场分析 — TrendForce（2025年12月）",
     "https://www.trendforce.com/news/2025/12/18/news-micron-hikes-capex-to-20b-with-2026-hbm-supply-fully-booked-hbm4-ramps-2q26/"),
    ("美光FY2026Q1分析 — Futurum Research",
     "https://futurumgroup.com/insights/micron-technology-q1-fy-2026-sets-records-strong-q2-outlook/"),
]

for label, url in sources:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(4)
    p.paragraph_format.left_indent  = Inches(0.25)
    run = p.add_run("• ")
    set_font(run, size=9.5, color="444444", font_name="宋体")
    add_hyperlink(p, url, label)

add_paragraph(doc, "\n一致预期数据：彭博终端，截至2025年12月17日。",
              size=9, italic=True, color="666666", space_before=6, space_after=3)
add_paragraph(doc, "市场数据：Yahoo Finance / Investing.com，截至2026年4月10日。",
              size=9, italic=True, color="666666", space_before=0, space_after=3)
add_paragraph(doc,
    "\n免责声明：本报告仅供参考，不构成投资建议。目标价及评级均为基于公开信息的估算，"
    "不代表任何投资机构的正式意见。",
    size=8, italic=True, color="999999", space_before=6, space_after=3)

# ─── Save ────────────────────────────────────────────────────────────────────
outpath = os.path.join(OUT, "MU_Q1_FY2026_业绩更新报告_中文版.docx")
doc.save(outpath)
print(f"中文报告已保存：{outpath}")
