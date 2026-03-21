/**
 * Alibaba Group (BABA / 9988.HK) — One-Page Strip Profile
 * Investment Banking format — 4:3 (10" x 7.5"), PptxGenJS
 * Output: output/BABA/BABA_OnePager_StripProfile.pptx
 */

import { createRequire } from 'module';
const require = createRequire(import.meta.url);
const pptxgen = require('/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis/node_modules/pptxgenjs/dist/pptxgen.cjs.js');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_4x3'; // 10" x 7.5"

// ── Brand Colors ──────────────────────────────────────────────────────────────
const ORANGE  = 'FF6A00';   // Alibaba brand orange
const NAVY    = '1F2D5C';   // Alibaba deep navy
const CLOUD   = '0070C0';   // Cloud segment blue
const GREEN   = '00A651';   // positive / beat
const RED     = 'E8192C';   // miss / negative
const LGRAY   = 'F5F7FA';   // alternating row bg
const MGRAY   = 'CCCCCC';   // divider lines
const DGRAY   = '444444';   // secondary text
const WHITE   = 'FFFFFF';
const BLACK   = '1A1A1A';

const slide = pptx.addSlide();
slide.background = { color: WHITE };

// ══════════════════════════════════════════════════════════════════
// HEADER BANNER
// ══════════════════════════════════════════════════════════════════

// Navy header background
slide.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.52,
  fill: { color: NAVY }, line: { type: 'none' }
});

// Orange accent left strip
slide.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 0.08, h: 0.52,
  fill: { color: ORANGE }, line: { type: 'none' }
});

// Company name + tickers
slide.addText('Alibaba Group Holding Limited (NYSE: BABA  |  HKEX: 9988)', {
  x: 0.18, y: 0.03, w: 7.4, h: 0.36,
  fontSize: 13, bold: true, color: WHITE, fontFace: 'Arial', valign: 'middle'
});

// Rating badge
slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 7.7, y: 0.08, w: 0.9, h: 0.28,
  fill: { color: GREEN }, line: { type: 'none' }, rectRadius: 0.03
});
slide.addText('BUY', {
  x: 7.7, y: 0.08, w: 0.9, h: 0.28,
  fontSize: 8, bold: true, color: WHITE, fontFace: 'Arial', align: 'center', valign: 'middle'
});

// Price Target badge
slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 8.68, y: 0.08, w: 1.22, h: 0.28,
  fill: { color: WHITE }, line: { color: ORANGE, pt: 0.8 }, rectRadius: 0.03
});
slide.addText('PT: $120.00', {
  x: 8.68, y: 0.08, w: 1.22, h: 0.28,
  fontSize: 8, bold: true, color: ORANGE, fontFace: 'Arial', align: 'center', valign: 'middle'
});

// ══════════════════════════════════════════════════════════════════
// SUBTITLE STATS BAR
// ══════════════════════════════════════════════════════════════════

slide.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0.52, w: 10, h: 0.28,
  fill: { color: LGRAY }, line: { type: 'none' }
});

const stats = [
  { label: 'Price', val: '$122.41' },
  { label: 'Mkt Cap', val: '~$292B' },
  { label: '52W Range', val: '$95.73 – $192.67' },
  { label: 'Sector', val: 'Consumer Discretionary / Technology' },
  { label: 'Exchange', val: 'NYSE / HKEX' },
];
const statW = 10 / stats.length;
stats.forEach((s, i) => {
  slide.addText([
    { text: s.label + ': ', options: { color: DGRAY, fontSize: 7, bold: false } },
    { text: s.val, options: { color: BLACK, fontSize: 7, bold: true } },
  ], {
    x: i * statW, y: 0.52, w: statW, h: 0.28,
    align: 'center', valign: 'middle', fontFace: 'Arial'
  });
});

// ══════════════════════════════════════════════════════════════════
// MAIN CONTENT — 2×2 QUADRANTS
// Layout: y starts at 0.80 (after header + stats bar)
// Total remaining height: 7.5 - 0.80 = 6.70 (minus footer ~0.22 = 6.48)
// Left col: 0 – 4.85,  Right col: 5.05 – 10
// Row 1: y 0.80 – 4.00,  Row 2: y 4.02 – 7.28
// ══════════════════════════════════════════════════════════════════

const COL1X = 0.15;
const COL2X = 5.05;
const COL_W = 4.75;
const ROW1Y = 0.82;
const ROW1H = 3.12;
const ROW2Y = 4.00;
const ROW2H = 3.15;

// Center divider line
slide.addShape(pptx.shapes.LINE, {
  x: 4.95, y: 0.82, w: 0, h: 6.34,
  line: { color: MGRAY, width: 0.5 }
});
// Horizontal divider
slide.addShape(pptx.shapes.LINE, {
  x: 0, y: 3.95, w: 10, h: 0,
  line: { color: MGRAY, width: 0.5 }
});

// ── SECTION HEADER HELPER ──────────────────────────────────────────────────
function addSectionHeader(x, y, w, label, accentColor = ORANGE) {
  // Orange accent bar
  slide.addShape(pptx.shapes.RECTANGLE, {
    x, y, w, h: 0.22,
    fill: { color: accentColor }, line: { type: 'none' }
  });
  slide.addText(label, {
    x: x + 0.08, y: y, w: w - 0.08, h: 0.22,
    fontSize: 7.5, bold: true, color: WHITE,
    fontFace: 'Arial', valign: 'middle'
  });
}

// ══════════════════════════════════════════════════════════════════
// QUADRANT 1 — Company Overview (top-left)
// ══════════════════════════════════════════════════════════════════
addSectionHeader(COL1X, ROW1Y, COL_W, 'COMPANY OVERVIEW', NAVY);

const overviewBullets = [
  ['Founded / HQ', 'April 1999, Hangzhou, Zhejiang, China'],
  ['Employees',    '~240,000 (FY2025)'],
  ['CEO',          'Eddie Wu (since Sep 2023)'],
  ['Chairman',     'Joseph Tsai'],
  ['Listing',      'NYSE (BABA) since 2014 | HKEX (9988) since 2019'],
  ['Business',     'China\'s largest e-commerce & cloud ecosystem; global AI leader'],
  ['Qwen AI App',  '300M+ monthly active users (MAU) as of Q3 FY2026'],
  ['88VIP Members','59M+ premium members (double-digit YoY growth)'],
  ['AI Capex',     '¥380B (~$53B) over 3 years committed; AI infra build-out'],
  ['Buyback',      '$25B authorization; $11.9B repurchased in FY2025 (5.1% net reduction)'],
];

overviewBullets.forEach(([lbl, val], i) => {
  const yPos = ROW1Y + 0.26 + i * 0.275;
  slide.addText([
    { text: lbl + ': ', options: { bold: true, color: NAVY, fontSize: 7 } },
    { text: val,        options: { bold: false, color: BLACK, fontSize: 7 } },
  ], {
    x: COL1X + 0.08, y: yPos, w: COL_W - 0.1, h: 0.26,
    fontFace: 'Arial', valign: 'middle',
    fill: i % 2 === 0 ? { type: 'none' } : { color: LGRAY },
  });
});

// ══════════════════════════════════════════════════════════════════
// QUADRANT 2 — Business & Positioning (top-right)
// ══════════════════════════════════════════════════════════════════
addSectionHeader(COL2X, ROW1Y, COL_W, 'BUSINESS & STRATEGIC POSITIONING', ORANGE);

const posBullets = [
  { icon: '●', color: CLOUD,  text: 'CLOUD INTELLIGENCE — #1 cloud in China (~36% market share); Cloud revenue +36% YoY to ¥43.3B in Q3 FY26; AI product revenue: triple-digit growth for 10 consecutive quarters.' },
  { icon: '●', color: ORANGE, text: 'CHINA E-COMMERCE — Taobao & Tmall dominate domestic marketplace GMV; monetized via advertising, take-rate & logistics. Quick commerce (Taobao Instant) +56% YoY targeting ¥1T GMV by FY2028.' },
  { icon: '●', color: NAVY,   text: 'INTERNATIONAL COMMERCE (AIDC) — AliExpress (global), Lazada (SE Asia), Trendyol (Turkey); in deliberate quality-over-quantity pivot; losses narrowing materially YoY.' },
  { icon: '●', color: GREEN,  text: 'AI STRATEGY — Qwen frontier model family at 300M MAU; competing with OpenAI/Google globally; T-Head AI chips (470K+ shipped) reducing GPU dependency; AI cloud demand outpacing supply.' },
  { icon: '●', color: DGRAY,  text: 'INVESTMENT THESIS — Deliberate investment cycle compressing near-term EPS. Cloud re-acceleration to +36% and 88VIP at 59M validate long-term compounding. Trades at ~11x FY2027E Non-GAAP P/E.' },
];

posBullets.forEach((b, i) => {
  const yPos = ROW1Y + 0.26 + i * 0.555;
  slide.addText([
    { text: b.icon + '  ', options: { bold: true, color: b.color, fontSize: 8 } },
    { text: b.text,        options: { color: BLACK, fontSize: 7 } },
  ], {
    x: COL2X + 0.06, y: yPos, w: COL_W - 0.1, h: 0.52,
    fontFace: 'Arial', valign: 'top',
    fill: i % 2 === 0 ? { type: 'none' } : { color: LGRAY },
  });
});

// ══════════════════════════════════════════════════════════════════
// QUADRANT 3 — Key Financials (bottom-left)
// ══════════════════════════════════════════════════════════════════
addSectionHeader(COL1X, ROW2Y, COL_W, 'KEY FINANCIALS', NAVY);

// Financial table
const tblHeaders = ['Metric', 'FY2025A', 'FY2026E', 'FY2027E'];
const tblRows = [
  ['Revenue (¥B)',          '¥941B',  '~¥1,025B', '~¥1,140B'],
  ['Revenue ($B)',          '$134B',  '~$146B',   '~$163B'],
  ['Revenue YoY',           '+7%',    '+9%',      '+11%'],
  ['Cloud Revenue (¥B)',    '¥117B',  '~¥159B',   '~¥211B'],
  ['Cloud YoY',             '+20%',   '+36%',     '+33%'],
  ['Adj. EBITA (¥B)',       '~¥175B', '~¥90B',    '~¥155B'],
  ['Adj. EBITA Margin',     '18.6%',  '~8.8%',    '~13.6%'],
  ['Non-GAAP EPS/ADS (¥)', '¥50.0',  '~¥28.0',   '~¥52.0'],
  ['Non-GAAP EPS/ADS ($)', '$9.01',  '~$4.00',   '~$7.43'],
  ['Non-GAAP P/E',         '~11x',   '~22x',     '~11x'],
  ['88VIP Members',        '>50M',   '>59M',     '>70M est.'],
];

const tblW  = [1.75, 0.88, 0.92, 0.92];
const tblTotalW = tblW.reduce((a, b) => a + b, 0);
const tblStartX = COL1X + 0.05;
const tblStartY = ROW2Y + 0.26;
const rowH = 0.235;

// Header row
let curX = tblStartX;
tblHeaders.forEach((h, j) => {
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: curX, y: tblStartY, w: tblW[j], h: rowH,
    fill: { color: NAVY }, line: { color: WHITE, pt: 0.5 }
  });
  slide.addText(h, {
    x: curX, y: tblStartY, w: tblW[j], h: rowH,
    fontSize: 6.5, bold: true, color: WHITE, fontFace: 'Arial',
    align: 'center', valign: 'middle'
  });
  curX += tblW[j];
});

// Data rows
tblRows.forEach((row, i) => {
  const rowY = tblStartY + rowH * (i + 1);
  const bg   = i % 2 === 0 ? WHITE : LGRAY;
  curX = tblStartX;
  row.forEach((cell, j) => {
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: curX, y: rowY, w: tblW[j], h: rowH,
      fill: { color: bg }, line: { color: MGRAY, pt: 0.25 }
    });
    const isLabel = j === 0;
    slide.addText(cell, {
      x: curX + 0.03, y: rowY, w: tblW[j] - 0.03, h: rowH,
      fontSize: 6.5, bold: isLabel, color: isLabel ? NAVY : BLACK,
      fontFace: 'Arial', align: isLabel ? 'left' : 'center', valign: 'middle'
    });
    curX += tblW[j];
  });
});

// Note below table
const noteY = tblStartY + rowH * (tblRows.length + 1) + 0.02;
slide.addText('E = Analyst estimate. ¥/USD at ~7.0. FY ends March 31.', {
  x: COL1X + 0.05, y: noteY, w: COL_W - 0.1, h: 0.18,
  fontSize: 6, color: DGRAY, italic: true, fontFace: 'Arial', valign: 'middle'
});

// ══════════════════════════════════════════════════════════════════
// QUADRANT 4 — Shareholders & Recent Developments (bottom-right)
// ══════════════════════════════════════════════════════════════════
addSectionHeader(COL2X, ROW2Y, COL_W, 'SHAREHOLDERS & RECENT DEVELOPMENTS', ORANGE);

// ── Shareholders mini-chart (horizontal bars) ────────────────────
const shareholders = [
  { name: 'Jack Ma',     pct: 4.1,  color: ORANGE },
  { name: 'SoftBank',    pct: 3.8,  color: NAVY   },
  { name: 'Vanguard',    pct: 3.5,  color: CLOUD  },
  { name: 'BlackRock',   pct: 2.8,  color: DGRAY  },
  { name: 'Joseph Tsai', pct: 1.3,  color: GREEN  },
  { name: 'State Street',pct: 1.1,  color: MGRAY  },
];

const shdX = COL2X + 0.1;
const shdLblW = 1.05;
const shdBarMaxW = 2.2;
const shdPctW = 0.5;
const shdStartY = ROW2Y + 0.26;
const shdRowH = 0.27;

// Sub-header
slide.addText('Top Shareholders (Approximate)', {
  x: shdX, y: shdStartY - 0.01, w: COL_W - 0.12, h: 0.20,
  fontSize: 7, bold: true, color: NAVY, fontFace: 'Arial', valign: 'middle'
});

shareholders.forEach((sh, i) => {
  const rowY = shdStartY + 0.20 + i * shdRowH;
  const barW = (sh.pct / 6.0) * shdBarMaxW;

  slide.addText(sh.name, {
    x: shdX, y: rowY, w: shdLblW, h: shdRowH,
    fontSize: 6.5, color: BLACK, fontFace: 'Arial', valign: 'middle'
  });
  // Bar background
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: shdX + shdLblW, y: rowY + 0.05, w: shdBarMaxW, h: shdRowH - 0.1,
    fill: { color: LGRAY }, line: { type: 'none' }
  });
  // Bar fill
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: shdX + shdLblW, y: rowY + 0.05, w: barW, h: shdRowH - 0.1,
    fill: { color: sh.color }, line: { type: 'none' }
  });
  slide.addText(`${sh.pct.toFixed(1)}%`, {
    x: shdX + shdLblW + shdBarMaxW + 0.05, y: rowY,
    w: 0.4, h: shdRowH,
    fontSize: 6.5, bold: true, color: NAVY, fontFace: 'Arial', valign: 'middle'
  });
});

// ── Recent Developments ──────────────────────────────────────────
const devStartY = shdStartY + 0.20 + shareholders.length * shdRowH + 0.10;

slide.addText('Recent Developments', {
  x: COL2X + 0.1, y: devStartY, w: COL_W - 0.12, h: 0.20,
  fontSize: 7, bold: true, color: NAVY, fontFace: 'Arial', valign: 'middle'
});

const devs = [
  { date: 'Mar 2026', text: 'Q3 FY2026: Cloud +36% YoY; revenue miss (¥284.8B); EPS –67% on ¥380B AI capex cycle; stock –7.2%' },
  { date: 'Mar 2026', text: 'Committed ¥380B (~$53B) AI/cloud capex over 3 years; CEO notes pace "may be insufficient"' },
  { date: 'Feb 2026', text: 'Qwen AI consumer app reached 300M MAU; cloud external revenue crossed ¥100B for FY2026 YTD' },
  { date: 'Dec 2025', text: '88VIP premium members surpassed 59M; T-Head AI chips: 470K+ cumulative shipments' },
  { date: 'Nov 2025', text: 'Q2 FY2026: Cloud +34%; Adj. EBITA –78% YoY — deepest investment trough of the cycle' },
];

devs.forEach((d, i) => {
  const rowY = devStartY + 0.22 + i * 0.34;
  slide.addText([
    { text: d.date + '  ', options: { bold: true, color: ORANGE, fontSize: 6.5 } },
    { text: d.text,        options: { color: BLACK,  fontSize: 6.5 } },
  ], {
    x: COL2X + 0.1, y: rowY, w: COL_W - 0.15, h: 0.32,
    fontFace: 'Arial', valign: 'top',
    fill: i % 2 === 0 ? { type: 'none' } : { color: LGRAY }
  });
});

// ══════════════════════════════════════════════════════════════════
// FOOTER
// ══════════════════════════════════════════════════════════════════
slide.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 7.30, w: 10, h: 0.20,
  fill: { color: NAVY }, line: { type: 'none' }
});
slide.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 7.30, w: 0.08, h: 0.20,
  fill: { color: ORANGE }, line: { type: 'none' }
});
slide.addText(
  'Sources: Alibaba Group IR, BusinessWire, CNBC, SEC EDGAR, yfinance. Market data as of March 21, 2026. For informational purposes only — not investment advice.',
  {
    x: 0.18, y: 7.30, w: 7.8, h: 0.20,
    fontSize: 5.5, color: WHITE, fontFace: 'Arial', valign: 'middle'
  }
);
slide.addText('March 21, 2026', {
  x: 8.1, y: 7.30, w: 1.8, h: 0.20,
  fontSize: 5.5, color: WHITE, fontFace: 'Arial', align: 'right', valign: 'middle'
});

// ══════════════════════════════════════════════════════════════════
// SAVE
// ══════════════════════════════════════════════════════════════════
const OUT = '/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis/output/BABA/BABA_OnePager_StripProfile.pptx';
pptx.writeFile({ fileName: OUT })
  .then(() => console.log(`✓ One-pager saved: ${OUT}`))
  .catch(err => { console.error('Error:', err); process.exit(1); });
