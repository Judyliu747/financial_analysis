/**
 * SMIC (中芯国际) One-Page Strip Profile
 * Investment Banking format — 4:3, PptxGenJS
 */

import { createRequire } from 'module';
const require = createRequire(import.meta.url);
const pptxgen = require('/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis/node_modules/pptxgenjs/dist/pptxgen.cjs.js');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_4x3'; // 10" x 7.5"

// ── Brand Colors ──────────────────────────────────────────────────────────────
const NAVY   = '173C85';  // SMIC primary blue
const RED    = 'D71718';  // SMIC brand red
const ORANGE = 'E79722';  // SMIC brand orange
const BLACK  = '1A1A1A';
const DGRAY  = '444444';
const MGRAY  = 'CCCCCC';
const LGRAY  = 'F5F6FA';
const WHITE  = 'FFFFFF';

const slide = pptx.addSlide();

// ── Background ────────────────────────────────────────────────────────────────
slide.background = { color: WHITE };

// ══════════════════════════════════════════════════════════════════
// HEADER BANNER
// ══════════════════════════════════════════════════════════════════

// Navy header banner background
slide.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.52,
  fill: { color: NAVY },
  line: { type: 'none' }
});

// Red accent left strip on banner
slide.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 0.08, h: 0.52,
  fill: { color: RED },
  line: { type: 'none' }
});

// Company name
slide.addText('Semiconductor Manufacturing International Corporation (SMIC)  |  0981.HK  •  688981.SS', {
  x: 0.18, y: 0.03, w: 7.5, h: 0.36,
  fontSize: 13, bold: true, color: WHITE, fontFace: 'Arial', valign: 'middle'
});

// Rating badge
slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 7.8, y: 0.08, w: 1.1, h: 0.28,
  fill: { color: ORANGE },
  line: { type: 'none' },
  rectRadius: 0.03
});
slide.addText('OUTPERFORM', {
  x: 7.8, y: 0.08, w: 1.1, h: 0.28,
  fontSize: 7.5, bold: true, color: WHITE, fontFace: 'Arial', align: 'center', valign: 'middle'
});

// Target price badge
slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 9.0, y: 0.08, w: 0.88, h: 0.28,
  fill: { color: 'FFFFFF' },
  line: { color: ORANGE, pt: 0.8 },
  rectRadius: 0.03
});
slide.addText('TP: HKD 38', {
  x: 9.0, y: 0.08, w: 0.88, h: 0.28,
  fontSize: 7.5, bold: true, color: ORANGE, fontFace: 'Arial', align: 'center', valign: 'middle'
});

// ══════════════════════════════════════════════════════════════════
// SUBTITLE ROW — key stats bar
// ══════════════════════════════════════════════════════════════════
slide.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0.52, w: 10, h: 0.28,
  fill: { color: LGRAY },
  line: { type: 'none' }
});

const statItems = [
  { label: 'Mkt Cap', value: '~$38.8B' },
  { label: 'EV',       value: '~$37.8B' },
  { label: 'Price',    value: 'HKD 28.40' },
  { label: '52-Wk',   value: 'HKD 17.82–42.80' },
  { label: 'Sector',  value: 'Semiconductor Foundry' },
  { label: 'Report Date', value: 'Mar 16, 2026' },
];
const statW = 10 / statItems.length;
statItems.forEach((s, i) => {
  slide.addText([
    { text: s.label + ': ', options: { bold: true, color: NAVY } },
    { text: s.value, options: { color: DGRAY } }
  ], {
    x: i * statW, y: 0.53, w: statW, h: 0.24,
    fontSize: 7.5, fontFace: 'Arial', align: 'center', valign: 'middle'
  });
});

// ── Thin divider below stat bar ───────────────────────────────────
slide.addShape(pptx.shapes.LINE, {
  x: 0, y: 0.8, w: 10, h: 0,
  line: { color: MGRAY, pt: 0.5 }
});

// ══════════════════════════════════════════════════════════════════
// VERTICAL DIVIDER (center line)
// ══════════════════════════════════════════════════════════════════
slide.addShape(pptx.shapes.LINE, {
  x: 5.0, y: 0.82, w: 0, h: 3.1,
  line: { color: MGRAY, pt: 0.5 }
});

// ══════════════════════════════════════════════════════════════════
// HORIZONTAL DIVIDER (mid line)
// ══════════════════════════════════════════════════════════════════
slide.addShape(pptx.shapes.LINE, {
  x: 0, y: 3.95, w: 10, h: 0,
  line: { color: MGRAY, pt: 0.5 }
});

// ══════════════════════════════════════════════════════════════════
// QUADRANT 1 — COMPANY OVERVIEW (top-left)
// ══════════════════════════════════════════════════════════════════

// Accent bar
slide.addShape(pptx.shapes.RECTANGLE, {
  x: 0.25, y: 0.88, w: 0.07, h: 0.24,
  fill: { color: RED }, line: { type: 'none' }
});
// Header
slide.addText('Company Overview', {
  x: 0.38, y: 0.87, w: 4.5, h: 0.27,
  fontSize: 11, bold: true, color: NAVY, fontFace: 'Arial'
});

// Bullets
slide.addText([
  { text: 'HQ: ', options: { bold: true } }, { text: 'Shanghai, China; Founded: 2000; ~20,000+ employees', options: {} },
  { text: '\n' },
  { text: 'Exchange: ', options: { bold: true } }, { text: 'HKEX (0981.HK) + STAR Market (688981.SS)', options: {} },
  { text: '\n' },
  { text: 'Co-CEOs: ', options: { bold: true } }, { text: 'Dr. Zhao Haijun & Dr. Liang Mong-Song; Chairman: Gao Yonggang', options: {} },
  { text: '\n' },
  { text: 'Business: ', options: { bold: true } }, { text: "China's largest pure-play foundry; logic ICs in consumer, mobile, industrial & AI", options: {} },
  { text: '\n' },
  { text: 'Process Nodes: ', options: { bold: true } }, { text: '0.35µm → 7nm (N+2); developing 5nm (N+3) via multi-patterning DUV (no EUV)', options: {} },
  { text: '\n' },
  { text: 'Capacity: ', options: { bold: true } }, { text: '~1.06M 8" eq. wspm (end-2025); +111K wspm added in FY2025', options: {} },
  { text: '\n' },
  { text: 'Utilization: ', options: { bold: true } }, { text: '95.7% (Q4\'25) — highest in 2 years; target >90% through FY2026', options: {} },
  { text: '\n' },
  { text: 'Customers: ', options: { bold: true } }, { text: "Huawei (~35-40% rev.), Qualcomm China, HiSilicon, domestic fabless leaders", options: {} },
], {
  x: 0.28, y: 1.17, w: 4.6, h: 2.62,
  fontSize: 8.5, fontFace: 'Arial', color: DGRAY, valign: 'top',
  paraSpaceAfter: 3.5
});

// ══════════════════════════════════════════════════════════════════
// QUADRANT 2 — BUSINESS & POSITIONING (top-right)
// ══════════════════════════════════════════════════════════════════
slide.addShape(pptx.shapes.RECTANGLE, {
  x: 5.08, y: 0.88, w: 0.07, h: 0.24,
  fill: { color: ORANGE }, line: { type: 'none' }
});
slide.addText('Business & Positioning', {
  x: 5.22, y: 0.87, w: 4.5, h: 0.27,
  fontSize: 11, bold: true, color: NAVY, fontFace: 'Arial'
});

slide.addText([
  { text: 'End Markets (Q4\'25): ', options: { bold: true } }, { text: 'Consumer Elec. 47% | Smartphone 20% | Industrial/Auto 15% | PC/Server 10% | Other 8%', options: {} },
  { text: '\n' },
  { text: 'Geography: ', options: { bold: true } }, { text: 'China Domestic ~90% | Americas ~6% | Other ~4% — strongest domestic revenue share among global foundries', options: {} },
  { text: '\n' },
  { text: '7nm Ramp: ', options: { bold: true } }, { text: 'N+2 yield 60-70% (vs. <40% at launch H2\'23); ~20K wspm capacity; Huawei buys ~15K wspm for Kirin + Ascend 910C AI chips', options: {} },
  { text: '\n' },
  { text: '5nm Dev. (N+3): ', options: { bold: true } }, { text: 'Multi-patterning DUV; Kirin 9030 (2H26E); China 7nm/5nm output targeted ×5 in 2 years (TrendForce)', options: {} },
  { text: '\n' },
  { text: 'AI Tailwind: ', options: { bold: true } }, { text: 'Huawei Ascend 910C yield improved to ~40% (from ~20% 6 months ago) — now commercially profitable', options: {} },
  { text: '\n' },
  { text: 'Self-Sufficiency: ', options: { bold: true } }, { text: "Sole domestic-Chinese supplier of advanced nodes; government policy-backed; no EUV access = technology moat v. TSMC", options: {} },
  { text: '\n' },
  { text: 'Capex Cycle: ', options: { bold: true } }, { text: '~$6.5B invested in FY2025; FY2026 guided "roughly flat" — Beijing & Tianjin fabs ramping', options: {} },
], {
  x: 5.1, y: 1.17, w: 4.65, h: 2.62,
  fontSize: 8.5, fontFace: 'Arial', color: DGRAY, valign: 'top',
  paraSpaceAfter: 3.5
});

// ══════════════════════════════════════════════════════════════════
// QUADRANT 3 — KEY FINANCIALS (bottom-left)
// ══════════════════════════════════════════════════════════════════
slide.addShape(pptx.shapes.RECTANGLE, {
  x: 0.25, y: 4.02, w: 0.07, h: 0.24,
  fill: { color: RED }, line: { type: 'none' }
});
slide.addText('Key Financials & Valuation', {
  x: 0.38, y: 4.01, w: 4.5, h: 0.27,
  fontSize: 11, bold: true, color: NAVY, fontFace: 'Arial'
});

// Financial table
const hdrFill = { color: NAVY };
const hdrFont = { color: WHITE, bold: true };
const altFill = { color: 'EEF2FA' };
const numAlign = 'center';

const tblData = [
  [
    { text: 'Metric ($M)',    options: { bold: true, color: WHITE, fill: NAVY } },
    { text: 'FY2024A',        options: { bold: true, color: WHITE, fill: NAVY, align: 'center' } },
    { text: 'FY2025A',        options: { bold: true, color: WHITE, fill: NAVY, align: 'center' } },
    { text: 'FY2026E',        options: { bold: true, color: WHITE, fill: NAVY, align: 'center' } },
  ],
  [
    { text: 'Revenue',        options: {} },
    { text: '$8,028',         options: { align: numAlign } },
    { text: '$9,327',         options: { align: numAlign, bold: true, color: NAVY } },
    { text: '$10,600',        options: { align: numAlign, color: 'CC5500' } },
  ],
  [
    { text: 'YoY Growth',     options: { fill: 'F8F9FF' } },
    { text: '+0.2%',          options: { align: numAlign, fill: 'F8F9FF' } },
    { text: '+16.2%',         options: { align: numAlign, fill: 'F8F9FF', bold: true, color: '1A6B3A' } },
    { text: '+13.6%E',        options: { align: numAlign, fill: 'F8F9FF', color: 'CC5500' } },
  ],
  [
    { text: 'Gross Profit',   options: {} },
    { text: '$1,445',         options: { align: numAlign } },
    { text: '$1,959',         options: { align: numAlign } },
    { text: '$2,067E',        options: { align: numAlign, color: 'CC5500' } },
  ],
  [
    { text: 'Gross Margin',   options: { fill: 'F8F9FF' } },
    { text: '18.0%',          options: { align: numAlign, fill: 'F8F9FF' } },
    { text: '21.0%',          options: { align: numAlign, fill: 'F8F9FF', bold: true, color: '1A6B3A' } },
    { text: '19.5%E',         options: { align: numAlign, fill: 'F8F9FF', color: 'CC5500' } },
  ],
  [
    { text: 'EBITDA',         options: {} },
    { text: '$4,100',         options: { align: numAlign } },
    { text: '$4,850',         options: { align: numAlign } },
    { text: '$5,300E',        options: { align: numAlign, color: 'CC5500' } },
  ],
  [
    { text: 'Net Income',     options: { fill: 'F8F9FF' } },
    { text: '$490',           options: { align: numAlign, fill: 'F8F9FF' } },
    { text: '$666',           options: { align: numAlign, fill: 'F8F9FF' } },
    { text: '$720E',          options: { align: numAlign, fill: 'F8F9FF', color: 'CC5500' } },
  ],
  [
    { text: 'CapEx',          options: {} },
    { text: '$5,700',         options: { align: numAlign } },
    { text: '$6,500',         options: { align: numAlign } },
    { text: '~$6,500E',       options: { align: numAlign, color: 'CC5500' } },
  ],
  [
    { text: 'EV / EBITDA',    options: { fill: 'F8F9FF' } },
    { text: '9.2x',           options: { align: numAlign, fill: 'F8F9FF' } },
    { text: '7.8x',           options: { align: numAlign, fill: 'F8F9FF' } },
    { text: '7.1xE',          options: { align: numAlign, fill: 'F8F9FF', color: 'CC5500' } },
  ],
  [
    { text: 'P / S',          options: {} },
    { text: '2.2x',           options: { align: numAlign } },
    { text: '1.9x',           options: { align: numAlign } },
    { text: '1.7xE',          options: { align: numAlign, color: 'CC5500' } },
  ],
];

slide.addTable(tblData, {
  x: 0.28, y: 4.32, w: 4.58, h: 3.0,
  fontFace: 'Arial', fontSize: 9,
  border: { pt: 0.4, color: MGRAY },
  colW: [1.7, 0.96, 0.96, 0.96],
  rowH: 0.27,
  valign: 'middle'
});

// ══════════════════════════════════════════════════════════════════
// QUADRANT 4 — STOCK PERFORMANCE & OWNERSHIP (bottom-right)
// ══════════════════════════════════════════════════════════════════
slide.addShape(pptx.shapes.RECTANGLE, {
  x: 5.08, y: 4.02, w: 0.07, h: 0.24,
  fill: { color: ORANGE }, line: { type: 'none' }
});
slide.addText('Recent Developments & Shareholder Structure', {
  x: 5.22, y: 4.01, w: 4.55, h: 0.27,
  fontSize: 11, bold: true, color: NAVY, fontFace: 'Arial'
});

// Q4 2025 beat/miss summary mini-table
const q4Data = [
  [
    { text: 'Q4\'25 vs. Consensus', options: { bold: true, color: WHITE, fill: NAVY, colspan: 3 } }
  ],
  [
    { text: 'Revenue',        options: { bold: true } },
    { text: '$2,489M',        options: { align: 'center' } },
    { text: '▲ +2.8% BEAT',  options: { align: 'center', bold: true, color: '1A6B3A' } }
  ],
  [
    { text: 'Gross Margin',   options: { bold: true, fill: 'F8F9FF' } },
    { text: '19.2%',          options: { align: 'center', fill: 'F8F9FF' } },
    { text: '▼ MISS vs 20.9%',options: { align: 'center', bold: true, color: RED, fill: 'F8F9FF' } }
  ],
  [
    { text: 'Net Income',     options: { bold: true } },
    { text: '$173M',          options: { align: 'center' } },
    { text: '▲ +1.7% BEAT',  options: { align: 'center', bold: true, color: '1A6B3A' } }
  ],
  [
    { text: 'Q1\'26 Guide',   options: { bold: true, fill: 'F8F9FF' } },
    { text: 'Rev flat; GM 18-20%', options: { align: 'center', fill: 'F8F9FF', colspan: 2 } },
  ],
];

slide.addTable(q4Data, {
  x: 5.1, y: 4.32, w: 4.65, h: 1.42,
  fontFace: 'Arial', fontSize: 8.5,
  border: { pt: 0.4, color: MGRAY },
  colW: [1.4, 1.25, 2.0],
  rowH: 0.245,
  valign: 'middle'
});

// Shareholder structure sub-header
slide.addShape(pptx.shapes.RECTANGLE, {
  x: 5.1, y: 5.82, w: 0.06, h: 0.2,
  fill: { color: ORANGE }, line: { type: 'none' }
});
slide.addText('Top Shareholders', {
  x: 5.22, y: 5.81, w: 3.5, h: 0.22,
  fontSize: 9.5, bold: true, color: NAVY, fontFace: 'Arial'
});

const shData = [
  [
    { text: 'Shareholder', options: { bold: true, color: WHITE, fill: NAVY } },
    { text: 'Stake (%)',   options: { bold: true, color: WHITE, fill: NAVY, align: 'center' } },
    { text: 'Type',        options: { bold: true, color: WHITE, fill: NAVY, align: 'center' } },
  ],
  ['National IC Fund I & II (大基金)', { text: '~19.0%', options: { align: 'center', bold: true, color: NAVY } }, { text: 'State SOE', options: { align: 'center' } }],
  [{ text: 'Shanghai Govt. (国盛集团等)', options: { fill: 'F8F9FF' } }, { text: '~16.0%', options: { align: 'center', fill: 'F8F9FF' } }, { text: 'Municipal', options: { align: 'center', fill: 'F8F9FF' } }],
  ['CITIC Capital (中信资本)', { text: '~3.5%', options: { align: 'center' } }, { text: 'PE', options: { align: 'center' } }],
  [{ text: 'Public Float (HK + A-Share)', options: { fill: 'F8F9FF' } }, { text: '~61.5%', options: { align: 'center', fill: 'F8F9FF' } }, { text: 'Public', options: { align: 'center', fill: 'F8F9FF' } }],
];

slide.addTable(shData, {
  x: 5.1, y: 6.06, w: 4.65, h: 1.1,
  fontFace: 'Arial', fontSize: 8.5,
  border: { pt: 0.4, color: MGRAY },
  colW: [2.2, 1.05, 1.4],
  rowH: 0.2,
  valign: 'middle'
});

// ══════════════════════════════════════════════════════════════════
// FOOTER
// ══════════════════════════════════════════════════════════════════
slide.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 7.28, w: 10, h: 0.22,
  fill: { color: NAVY }, line: { type: 'none' }
});
slide.addText(
  'Source: SMIC Q4 2025 Earnings Release (Feb 10, 2026); Bloomberg Consensus; Company Filings; TrendForce; Analyst Estimates as of Mar 2026  |  Confidential — For Discussion Purposes Only',
  {
    x: 0.1, y: 7.29, w: 9.8, h: 0.2,
    fontSize: 6.5, color: 'CCDDFF', fontFace: 'Arial', valign: 'middle'
  }
);

// ══════════════════════════════════════════════════════════════════
// SAVE
// ══════════════════════════════════════════════════════════════════
const outPath = '/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis/output/SMIC/SMIC_OnePager_StripProfile.pptx';
await pptx.writeFile({ fileName: outPath });
console.log('Saved:', outPath);
