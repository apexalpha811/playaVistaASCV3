/* Playa Vista ASC — Investor Deck generator
   Mirrors the V3 website (playaVistaASCV3) narrative + design system.
   Dark-throughout premium feel. Palette + Inter pulled from index.html. */

const pptxgen = require("pptxgenjs");

// ---- design tokens (from website :root) ----
const BG    = "0A0A0F";   // near-black
const CARD  = "12121A";   // card surface
const CARD2 = "16161F";   // slightly lifted
const BORDER= "262636";   // visible on dark (site uses 1E1E2E; nudged up for print contrast)
const LIME  = "B6C800";
const BLUE  = "1D00C6";
const BLUE_L= "5B6BFF";   // lighter blue for accents on dark
const TEXT  = "FFFFFF";
const SUB   = "A8A8B4";   // muted
const RED   = "FF7A8A";   // negative figures
const AMBER = "FFC850";   // breakeven / caution
const ICEB  = "CDD6FF";   // phase-2 label tint

const HEAD = "Georgia";   // header font with personality
const BODY = "Calibri";   // clean body
// Inter isn't guaranteed on every machine; Calibri/Georgia render reliably in PPT + PDF.

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";        // 13.3 x 7.5
const PW = 13.3, PH = 7.5;
pres.author = "Playa Vista ASC & Rejuvenation Clinic";
pres.title  = "Playa Vista ASC — Investor Pitch";

// ---------- helpers ----------
const shadow = () => ({ type: "outer", color: "000000", blur: 9, offset: 4, angle: 90, opacity: 0.45 });

function bg(slide){ slide.background = { color: BG }; }

// eyebrow label (lime, tracked, uppercase)
function eyebrow(slide, text, x, y, w){
  slide.addText(text.toUpperCase(), {
    x, y, w: w||6, h: 0.32, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: LIME,
    charSpacing: 3, align: "left", valign: "middle"
  });
}

function title(slide, text, x, y, w, size){
  slide.addText(text, {
    x, y, w: w||11.5, h: 0.9, margin: 0,
    fontFace: HEAD, fontSize: size||32, bold: true, color: TEXT,
    align: "left", valign: "top", lineSpacingMultiple: 1.0
  });
}

// card surface
function card(slide, x, y, w, h, opts){
  opts = opts || {};
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: 0.08,
    fill: { color: opts.fill || CARD },
    line: { color: opts.border || BORDER, width: opts.lw || 1 },
    shadow: opts.shadow === false ? undefined : shadow()
  });
}

// small pill badge
function badge(slide, text, x, y, fill, txtColor){
  const w = Math.max(0.9, 0.16 + text.length * 0.072);
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h: 0.3, rectRadius: 0.15,
    fill: { color: fill }, line: { type: "none" }
  });
  slide.addText(text.toUpperCase(), {
    x, y, w, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 9, bold: true, color: txtColor,
    charSpacing: 1, align: "center", valign: "middle"
  });
  return w;
}

// big stat: number + label inside a card region
function stat(slide, num, label, x, y, w, numColor, numSize){
  slide.addText(num, {
    x, y, w, h: 0.95, margin: 0,
    fontFace: HEAD, fontSize: numSize || 40, bold: true,
    color: numColor || LIME, align: "left", valign: "middle"
  });
  slide.addText(label, {
    x, y: y + 0.95, w, h: 0.85, margin: 0,
    fontFace: BODY, fontSize: 12.5, color: SUB, align: "left", valign: "top"
  });
}

function footer(slide, n){
  slide.addText("PLAYA VISTA ASC & REJUVENATION CLINIC", {
    x: 0.6, y: PH - 0.42, w: 7, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 8.5, color: "55555F", charSpacing: 1, valign: "middle"
  });
  slide.addText("Confidential · For Qualified Investors Only", {
    x: PW - 5.2, y: PH - 0.42, w: 4, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 8.5, color: "55555F", align: "right", valign: "middle"
  });
  if (n != null) slide.addText(String(n), {
    x: PW - 0.85, y: PH - 0.42, w: 0.3, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 8.5, color: LIME, align: "right", valign: "middle"
  });
}

// ============================================================
// SLIDE 1 — TITLE / HERO
// ============================================================
(() => {
  const s = pres.addSlide(); bg(s);
  // atmospheric corner glows (transparent ovals)
  s.addShape(pres.shapes.OVAL, { x: 8.6, y: -3.2, w: 7.5, h: 7.5, fill: { color: LIME, transparency: 90 }, line: { type: "none" } });
  s.addShape(pres.shapes.OVAL, { x: -3.2, y: 3.6, w: 8, h: 8, fill: { color: BLUE, transparency: 82 }, line: { type: "none" } });

  eyebrow(s, "Established 2026 · Playa Vista, CA", 0.85, 0.95, 8);
  s.addText([
    { text: "PLAYA VISTA ", options: { color: TEXT } },
    { text: "ASC", options: { color: LIME } },
    { text: "", options: { breakLine: true } },
    { text: "& REJUVENATION CLINIC", options: { color: TEXT } }
  ], { x: 0.8, y: 1.4, w: 11.6, h: 1.9, margin: 0, fontFace: HEAD, fontSize: 50, bold: true, lineSpacingMultiple: 0.98 });

  // lime rule
  s.addShape(pres.shapes.RECTANGLE, { x: 0.85, y: 3.35, w: 1.0, h: 0.06, fill: { color: LIME }, line: { type: "none" } });

  s.addText([
    { text: "Cash-flow positive on the clinic alone", options: { color: LIME, bold: true } },
    { text: " — 10+ months before the first OR ever opens.", options: { color: TEXT, bold: true } }
  ], { x: 0.85, y: 3.62, w: 11.4, h: 0.7, margin: 0, fontFace: BODY, fontSize: 19, valign: "middle" });

  s.addText("A vertically integrated surgical and wellness destination — where longevity programming accelerates surgical outcomes and cash-pay wellness funds the buildout.",
    { x: 0.85, y: 4.32, w: 10.6, h: 0.8, margin: 0, fontFace: BODY, fontSize: 13.5, color: SUB });

  // metric strip
  const my = 5.45, mh = 1.35;
  const metrics = [
    ["$43B+", "U.S. ASC market"],
    ["~67%", "Blended margin at stabilization"],
    ["~$6.3M", "Annual operating income\n(Projected · Base Case)"]
  ];
  const mw = 3.7, gap = 0.35; let mx = 0.85;
  s.addShape(pres.shapes.RECTANGLE, { x: 0.85, y: my - 0.18, w: 11.6, h: 0.02, fill: { color: BORDER }, line: { type: "none" } });
  metrics.forEach(([n,l]) => {
    s.addText(n, { x: mx, y: my, w: mw, h: 0.6, margin: 0, fontFace: HEAD, fontSize: 30, bold: true, color: LIME });
    s.addText(l, { x: mx, y: my + 0.62, w: mw, h: 0.7, margin: 0, fontFace: BODY, fontSize: 11, color: SUB });
    mx += mw + gap;
  });
  footer(s, null);
})();

// ============================================================
// SLIDE 2 — THE PROBLEM
// ============================================================
(() => {
  const s = pres.addSlide(); bg(s);
  eyebrow(s, "The Problem", 0.6, 0.55, 6);
  title(s, "A Fragmented, Slow Standard of Care", 0.6, 0.92, 12, 32);

  const items = [
    ["6–12 wk", "Surgical wait times patients endure for outpatient procedures.", LIME],
    ["0", "Integrated luxury ASC + wellness facilities on the Westside today.", LIME],
    ["3+", "Disconnected providers patients must navigate for surgery, recovery, and wellness.", LIME],
  ];
  const cw = 3.85, gap = 0.42, h = 3.5; let x = 0.6, y = 2.3;
  items.forEach(([n, l]) => {
    card(s, x, y, cw, h);
    s.addText(n, { x: x + 0.35, y: y + 0.4, w: cw - 0.7, h: 1.2, margin: 0, fontFace: HEAD, fontSize: 46, bold: true, color: LIME, valign: "middle" });
    s.addText(l, { x: x + 0.35, y: y + 1.7, w: cw - 0.7, h: 1.5, margin: 0, fontFace: BODY, fontSize: 14, color: SUB, valign: "top" });
    x += cw + gap;
  });
  footer(s, 2);
})();

// ============================================================
// SLIDE 3 — MISSION / VISION / PHILOSOPHY
// ============================================================
(() => {
  const s = pres.addSlide(); bg(s);
  eyebrow(s, "Who We Are", 0.6, 0.55, 6);
  title(s, "Mission · Vision · Philosophy", 0.6, 0.92, 12, 32);

  const cols = [
    ["Mission", "Deliver world-class outpatient surgical care and evidence-based rejuvenation under one roof — premium healthcare made accessible, efficient, and transformative."],
    ["Vision", "Become the leading vertically integrated ambulatory surgery and longevity destination on LA's Westside — known for clinical excellence, outcomes, and innovation."],
    ["Philosophy", "Surgical outcomes improve when patients are optimized before procedures and supported with evidence-based recovery, regenerative, and longevity protocols afterward."],
  ];
  const cw = 3.85, gap = 0.42, h = 3.6; let x = 0.6, y = 2.3;
  cols.forEach(([head, body]) => {
    card(s, x, y, cw, h);
    // left lime accent
    s.addShape(pres.shapes.RECTANGLE, { x: x, y: y + 0.18, w: 0.07, h: h - 0.36, fill: { color: LIME }, line: { type: "none" } });
    s.addText(head, { x: x + 0.38, y: y + 0.35, w: cw - 0.7, h: 0.5, margin: 0, fontFace: HEAD, fontSize: 20, bold: true, color: TEXT });
    s.addText(body, { x: x + 0.38, y: y + 1.0, w: cw - 0.72, h: 2.4, margin: 0, fontFace: BODY, fontSize: 13.5, color: SUB, valign: "top", lineSpacingMultiple: 1.05 });
    x += cw + gap;
  });
  footer(s, 3);
})();

// ============================================================
// SLIDE 4 — THE OPPORTUNITY
// ============================================================
(() => {
  const s = pres.addSlide(); bg(s);
  eyebrow(s, "Why Now, Why Westside", 0.6, 0.55, 8);
  title(s, "The Opening", 0.6, 0.92, 12, 32);

  const cards = [
    ["Unserved Westside", "No luxury ASC on the Westside integrates wellness into surgical care."],
    ["Dual-Stream Unlock", "Insurance-backed surgical revenue and cash-pay wellness revenue, combined under one roof."],
    ["Early Cash Flow", "Phase 1 clinic revenue starts 10+ months before OR activation — capital works while accreditation runs."],
  ];
  const cw = 3.85, gap = 0.42, h = 2.85; let x = 0.6, y = 2.15;
  cards.forEach(([head, body]) => {
    card(s, x, y, cw, h);
    s.addText(head, { x: x + 0.35, y: y + 0.32, w: cw - 0.7, h: 0.6, margin: 0, fontFace: HEAD, fontSize: 18, bold: true, color: LIME });
    s.addText(body, { x: x + 0.35, y: y + 1.0, w: cw - 0.7, h: 1.6, margin: 0, fontFace: BODY, fontSize: 13.5, color: SUB, valign: "top", lineSpacingMultiple: 1.05 });
    x += cw + gap;
  });

  // geography chips
  eyebrow(s, "Target Geography", 0.6, 5.35, 6);
  let gx = 0.6; const chips = ["Playa Vista", "Culver City", "Marina del Rey", "Santa Monica"];
  chips.forEach(c => {
    const w = 0.4 + c.length * 0.105;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: gx, y: 5.75, w, h: 0.46, rectRadius: 0.23, fill: { color: CARD }, line: { color: BORDER, width: 1 } });
    s.addText(c, { x: gx, y: 5.75, w, h: 0.46, margin: 0, fontFace: BODY, fontSize: 12.5, bold: true, color: SUB, align: "center", valign: "middle" });
    gx += w + 0.25;
  });
  footer(s, 4);
})();

// ============================================================
// SLIDE 5 — THREE REVENUE STREAMS
// ============================================================
(() => {
  const s = pres.addSlide(); bg(s);
  eyebrow(s, "Service Portfolio", 0.6, 0.55, 8);
  title(s, "Three Revenue Streams · One Facility", 0.6, 0.92, 12.2, 32);
  s.addText("6,000–6,500 SF in Playa Vista. Phase 1 launches at Month 3–6; Phase 2 ORs activate at Month 16.",
    { x: 0.6, y: 1.78, w: 11.5, h: 0.4, margin: 0, fontFace: BODY, fontSize: 13.5, color: SUB });

  const streams = [
    ["Phase 1 · Mo 3–6", LIME, BG, "Wellness + Aesthetics", "$419K/mo", "73% margin · cash-pay · funds the OR build", false],
    ["Phase 2 · Mo 16", BLUE, TEXT, "ASC Surgical Core", "$767K/mo", "61% margin · insurance-backed", false],
    ["Mo 16+", LIME, BG, "Combined Run-Rate", "$1.19M/mo", "at stabilization", true],
  ];
  const cw = 3.85, gap = 0.42, h = 3.55; let x = 0.6, y = 2.45;
  streams.forEach(([btxt, bfill, btxtc, head, val, meta, hl]) => {
    card(s, x, y, cw, h, { border: hl ? LIME : BORDER, lw: hl ? 2 : 1 });
    badge(s, btxt, x + 0.35, y + 0.32, bfill, btxtc);
    s.addText(head, { x: x + 0.35, y: y + 0.85, w: cw - 0.7, h: 0.55, margin: 0, fontFace: HEAD, fontSize: 18, bold: true, color: TEXT });
    s.addText(val, { x: x + 0.35, y: y + 1.5, w: cw - 0.7, h: 0.9, margin: 0, fontFace: HEAD, fontSize: 36, bold: true, color: LIME });
    s.addText(meta, { x: x + 0.35, y: y + 2.55, w: cw - 0.7, h: 0.8, margin: 0, fontFace: BODY, fontSize: 12.5, color: SUB, valign: "top" });
    x += cw + gap;
  });
  footer(s, 5);
})();

// ============================================================
// SLIDE 6 — SERVICE CATEGORIES
// ============================================================
(() => {
  const s = pres.addSlide(); bg(s);
  eyebrow(s, "Service Categories", 0.6, 0.5, 8);
  title(s, "6 Categories · 21 Revenue Lines", 0.6, 0.87, 12, 30);

  const cats = [
    ["Phase 1", "Regenerative", "PRP · Hormone/Peptide/GLP · IV Therapy (NAD+)", "$162.5K/mo", false],
    ["Phase 1", "Wellness & Aesthetics", "Botox & Fillers · Laser · Morpheus8 · Emsculpt · CoolSculpting", "$166.8K/mo", false],
    ["Phase 1", "Recovery", "Hyperbaric O₂ · Compression/NeuFit · Red Light", "$49.4K/mo", false],
    ["Phase 1", "Diagnostics", "Advanced Biometrics · DNA · DUTCH · DEXA", "$26.5K/mo", false],
    ["Phase 1", "Aescape Robots", "2× robotic massage stations · 85% margin · zero physician dependency", "$14K/mo", false],
    ["Phase 2", "Surgery", "Pain Mgmt · Podiatry · Orthopedic · Spine (MIS)", "$767K/mo", true],
  ];
  const cw = 3.85, gap = 0.42, ch = 1.95, rgap = 0.3;
  let x0 = 0.6, y0 = 1.95;
  cats.forEach((c, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = x0 + col * (cw + gap), y = y0 + row * (ch + rgap);
    const isP2 = c[4];
    card(s, x, y, cw, ch, { border: isP2 ? BLUE_L : BORDER, lw: isP2 ? 1.5 : 1 });
    s.addText(c[0].toUpperCase(), { x: x + 0.3, y: y + 0.2, w: cw - 0.6, h: 0.28, margin: 0, fontFace: BODY, fontSize: 9.5, bold: true, color: isP2 ? ICEB : SUB, charSpacing: 2 });
    s.addText(c[1], { x: x + 0.3, y: y + 0.46, w: cw - 0.6, h: 0.4, margin: 0, fontFace: HEAD, fontSize: 16, bold: true, color: TEXT });
    s.addText(c[2], { x: x + 0.3, y: y + 0.9, w: cw - 0.6, h: 0.7, margin: 0, fontFace: BODY, fontSize: 10.5, color: SUB, valign: "top", lineSpacingMultiple: 1.0 });
    s.addText(c[3], { x: x + 0.3, y: y + ch - 0.42, w: cw - 0.6, h: 0.34, margin: 0, fontFace: HEAD, fontSize: 14, bold: true, color: LIME });
  });
  footer(s, 6);
})();

// ============================================================
// SLIDE 7 — PHASE 1 STANDALONE
// ============================================================
(() => {
  const s = pres.addSlide(); bg(s);
  eyebrow(s, "Phase 1 Standalone — Clinic Only, Zero ORs", 0.6, 0.5, 9);
  title(s, "Profitable Before a Single Surgery", 0.6, 0.87, 12, 30);
  s.addText("The de-risk. The clinic alone — no OR, no surgical license, no payer contracts — clears $3.65M net annually.",
    { x: 0.6, y: 1.72, w: 12, h: 0.4, margin: 0, fontFace: BODY, fontSize: 13.5, color: SUB });

  const rows = [
    ["Category", "Monthly Gross", "Monthly Net", "Annual Gross", "Annual Net"],
    ["Regenerative", "$162,500", "$121,170", "$1,950,000", "$1,454,040"],
    ["Aesthetics", "$166,775", "$114,214", "$2,001,300", "$1,370,574"],
    ["Recovery", "$49,400", "$39,748", "$592,800", "$476,976"],
    ["Diagnostics", "$26,500", "$17,405", "$318,000", "$208,860"],
    ["Aescape Robots", "$14,000", "$11,900", "$168,000", "$142,800"],
    ["PHASE 1 TOTAL", "$419,175", "$304,437", "$5,030,100", "$3,653,250"],
  ];
  const tableData = rows.map((r, ri) => r.map((cell, ci) => {
    if (ri === 0) return { text: cell, options: { fill: { color: CARD2 }, color: LIME, bold: true, fontSize: 12, align: ci === 0 ? "left" : "right", valign: "middle" } };
    const isTotal = ri === rows.length - 1;
    return { text: cell, options: {
      fill: { color: isTotal ? "1C1C12" : CARD },
      color: isTotal ? LIME : (ci === 0 ? TEXT : SUB),
      bold: isTotal || ci === 0, fontSize: isTotal ? 12.5 : 12,
      align: ci === 0 ? "left" : "right", valign: "middle"
    } };
  }));
  s.addTable(tableData, {
    x: 0.6, y: 2.35, w: 12.1, colW: [3.0, 2.275, 2.275, 2.275, 2.275],
    rowH: 0.5, border: { pt: 0.5, color: BORDER },
    fontFace: BODY, valign: "middle", margin: [4, 8, 4, 8]
  });
  footer(s, 7);
})();

// ============================================================
// SLIDE 8 — 3-YEAR RAMP (native column chart)
// ============================================================
(() => {
  const s = pres.addSlide(); bg(s);
  eyebrow(s, "3-Year Ramp Model", 0.6, 0.5, 8);
  title(s, "Path to Stabilization", 0.6, 0.87, 12, 30);

  const labels = ["Mo 0–3","Mo 3–6","Mo 6–12","Mo 12–16","Mo 16–20","Mo 20–28","Mo 28–36"];
  const vals = [-273, -121, -47, 0, 212, 346, 522]; // op income $K/mo
  s.addChart(pres.charts.BAR, [{ name: "Op Income / Mo ($K)", labels, values: vals }], {
    x: 0.6, y: 2.1, w: 12.1, h: 3.7, barDir: "col",
    chartColors: ["B6C800"],
    chartColorsOpacity: [100],
    chartArea: { fill: { color: BG } }, plotArea: { fill: { color: BG } },
    catAxisLabelColor: SUB, catAxisLabelFontFace: BODY, catAxisLabelFontSize: 11,
    valAxisLabelColor: SUB, valAxisLabelFontFace: BODY, valAxisLabelFontSize: 10,
    valAxisLabelFormatCode: '$#,##0"K"',
    valGridLine: { color: "23232F", size: 0.5 }, catGridLine: { style: "none" },
    valAxisLineColor: BORDER, catAxisLineColor: BORDER,
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TEXT,
    dataLabelFontFace: BODY, dataLabelFontSize: 10, dataLabelFormatCode: '$#,##0"K"',
    showLegend: false, showTitle: false, barGapWidthPct: 55
  });
  s.addText([
    { text: "Breakeven Month ~14", options: { color: AMBER, bold: true } },
    { text: " — before the second OR is licensed.", options: { color: SUB } }
  ], { x: 0.6, y: 5.95, w: 12, h: 0.4, margin: 0, fontFace: BODY, fontSize: 13.5, align: "center", valign: "middle" });
  footer(s, 8);
})();

// ============================================================
// SLIDE 9 — FULL P&L
// ============================================================
(() => {
  const s = pres.addSlide(); bg(s);
  eyebrow(s, "Full P&L at Stabilization", 0.6, 0.5, 8);
  title(s, "Unit Economics", 0.6, 0.87, 12, 30);

  // P&L card (left)
  const px = 0.6, py = 1.95, pw = 7.1, ph = 4.85;
  card(s, px, py, pw, ph);
  const lines = [
    ["Gross Revenue", "$1,186,175", "100%", "norm"],
    ["Variable Costs / COGS (~33%)", "–$391,438", "", "neg"],
    ["Monthly Net after COGS", "$794,737", "67%", "sub"],
    ["Rent (6,500 SF NNN)", "–$44,000", "", "neg"],
    ["Clinical Staff", "–$110,000", "", "neg"],
    ["Admin / Management", "–$28,000", "", "neg"],
    ["Marketing & CAC", "–$35,000", "", "neg"],
    ["Malpractice + Liability", "–$12,000", "", "neg"],
    ["Equipment Leases", "–$22,000", "", "neg"],
    ["Regulatory / Accreditation", "–$8,000", "", "neg"],
    ["Utilities / Software", "–$14,000", "", "neg"],
    ["Total Fixed Overhead", "–$273,000", "23%", "sub"],
  ];
  let ly = py + 0.28; const lh = 0.305;
  lines.forEach(([l, v, pct, kind]) => {
    if (kind === "sub") {
      s.addShape(pres.shapes.RECTANGLE, { x: px + 0.0, y: ly - 0.02, w: pw, h: lh + 0.04, fill: { color: CARD2 }, line: { type: "none" } });
    }
    s.addText(l, { x: px + 0.35, y: ly, w: pw - 2.6, h: lh, margin: 0, fontFace: BODY, fontSize: 11.5, bold: kind === "sub", color: kind === "sub" ? TEXT : SUB, valign: "middle" });
    s.addText(v, { x: px + pw - 2.35, y: ly, w: 1.6, h: lh, margin: 0, fontFace: BODY, fontSize: 11.5, bold: kind === "sub", color: kind === "neg" ? RED : (kind === "sub" ? LIME : TEXT), align: "right", valign: "middle" });
    if (pct) s.addText(pct, { x: px + pw - 0.7, y: ly, w: 0.5, h: lh, margin: 0, fontFace: BODY, fontSize: 9.5, color: SUB, align: "right", valign: "middle" });
    ly += lh + 0.045;
  });
  // operating income band
  s.addShape(pres.shapes.RECTANGLE, { x: px, y: ly + 0.02, w: pw, h: 0.02, fill: { color: LIME }, line: { type: "none" } });
  s.addText([
    { text: "OPERATING INCOME / MO  ", options: { color: TEXT, bold: true } },
    { text: "Projected · Base Case", options: { color: SUB, bold: false, fontSize: 8.5 } }
  ], { x: px + 0.35, y: ly + 0.12, w: pw - 2.4, h: 0.42, margin: 0, fontFace: BODY, fontSize: 13.5, valign: "middle" });
  s.addText("$521,737", { x: px + pw - 2.35, y: ly + 0.12, w: 1.6, h: 0.42, margin: 0, fontFace: HEAD, fontSize: 15, bold: true, color: LIME, align: "right", valign: "middle" });
  s.addText("44%", { x: px + pw - 0.7, y: ly + 0.12, w: 0.5, h: 0.42, margin: 0, fontFace: BODY, fontSize: 9.5, color: SUB, align: "right", valign: "middle" });

  // right callouts
  const rx = 8.1, rw = 4.6;
  card(s, rx, py, rw, 2.3, { fill: CARD });
  s.addShape(pres.shapes.RECTANGLE, { x: rx, y: py + 0.18, w: 0.07, h: 2.3 - 0.36, fill: { color: LIME }, line: { type: "none" } });
  s.addText("~67%", { x: rx + 0.4, y: py + 0.3, w: rw - 0.7, h: 0.9, margin: 0, fontFace: HEAD, fontSize: 44, bold: true, color: LIME });
  s.addText("Net margin after COGS — before fixed overhead.", { x: rx + 0.4, y: py + 1.25, w: rw - 0.7, h: 0.9, margin: 0, fontFace: BODY, fontSize: 13, color: SUB, valign: "top" });

  card(s, rx, py + 2.55, rw, 2.3, { fill: CARD });
  s.addShape(pres.shapes.RECTANGLE, { x: rx, y: py + 2.73, w: 0.07, h: 2.3 - 0.36, fill: { color: LIME }, line: { type: "none" } });
  s.addText("~$6.26M", { x: rx + 0.4, y: py + 2.85, w: rw - 0.7, h: 0.9, margin: 0, fontFace: HEAD, fontSize: 40, bold: true, color: LIME });
  s.addText("Annualized operating income at stabilization (base case).", { x: rx + 0.4, y: py + 3.8, w: rw - 0.7, h: 0.9, margin: 0, fontFace: BODY, fontSize: 13, color: SUB, valign: "top" });
  footer(s, 9);
})();

// ============================================================
// SLIDE 10 — SCENARIO BAND
// ============================================================
(() => {
  const s = pres.addSlide(); bg(s);
  eyebrow(s, "Scenario Band", 0.6, 0.5, 8);
  title(s, "Low · Base · Upside", 0.6, 0.87, 12, 30);
  s.addText("Modeled operating-income outcomes across patient-volume and pricing assumptions.",
    { x: 0.6, y: 1.72, w: 12, h: 0.4, margin: 0, fontFace: BODY, fontSize: 13.5, color: SUB });

  const sc = [
    ["Low", "$1.07–1.57M", "annual operating income", [["Gross/mo","$371,780"],["Net/mo","$247,562"],["Op income/mo","$89–131K"]], RED, false],
    ["Base · Expected", "~$6.3M", "annual operating income", [["Gross/mo","$1,186,175"],["Net/mo","$794,737"],["Op income/mo","~$522K"]], LIME, true],
    ["Upside", "$9.4–11.3M+", "annual operating income", [["Gross/mo","$1,588,600+"],["Net/mo","$1,060,594+"],["Op income/mo","$787–944K+"]], BLUE_L, false],
  ];
  const cw = 3.85, gap = 0.42, h = 4.0; let x = 0.6, y = 2.35;
  sc.forEach(([tag, hero, sub, rows, color, isBase]) => {
    card(s, x, y, cw, h, { border: isBase ? LIME : BORDER, lw: isBase ? 2 : 1.2 });
    if (isBase) badge(s, "EXPECTED", x + cw - 1.5, y - 0.15, LIME, BG);
    s.addText(tag.toUpperCase(), { x: x + 0.35, y: y + 0.3, w: cw - 0.7, h: 0.3, margin: 0, fontFace: BODY, fontSize: 11, bold: true, color: isBase ? LIME : SUB, charSpacing: 2 });
    s.addText(hero, { x: x + 0.35, y: y + 0.62, w: cw - 0.7, h: 0.7, margin: 0, fontFace: HEAD, fontSize: 28, bold: true, color: color });
    s.addText(sub, { x: x + 0.35, y: y + 1.3, w: cw - 0.7, h: 0.35, margin: 0, fontFace: BODY, fontSize: 11.5, color: SUB });
    let ry = y + 1.85;
    rows.forEach(([k, v]) => {
      s.addText(k, { x: x + 0.35, y: ry, w: (cw-0.7)/2, h: 0.4, margin: 0, fontFace: BODY, fontSize: 12, color: SUB, valign: "middle" });
      s.addText(v, { x: x + 0.35 + (cw-0.7)/2, y: ry, w: (cw-0.7)/2, h: 0.4, margin: 0, fontFace: BODY, fontSize: 12, bold: true, color: TEXT, align: "right", valign: "middle" });
      s.addShape(pres.shapes.RECTANGLE, { x: x + 0.35, y: ry + 0.43, w: cw - 0.7, h: 0.008, fill: { color: BORDER }, line: { type: "none" } });
      ry += 0.55;
    });
    x += cw + gap;
  });
  footer(s, 10);
})();

// ============================================================
// SLIDE 11 — TIMELINE
// ============================================================
(() => {
  const s = pres.addSlide(); bg(s);
  eyebrow(s, "Phased Timeline", 0.6, 0.5, 8);
  title(s, "From Permits to Full ASC", 0.6, 0.87, 12, 30);

  const steps = [
    ["Month 0–3", "Pre-Open / Permits", "Licensing, accreditation pre-app, lease, design."],
    ["Month 3–6", "Phase 1 Clinic Launch", "16 non-surgical procedures active."],
    ["Month 3–16", "OR Buildout (concurrent)", "2 ORs + 1 Procedure Room under construction."],
    ["Month 16", "Phase 2 Full ASC", "All 21 procedures active."],
    ["Month 16–36", "Stabilization", "Physician ramp, payer contracting, volume build."],
  ];
  const n = steps.length, cw = 2.28, gap = 0.18; let x = 0.6; const y = 2.5;
  // connecting line
  s.addShape(pres.shapes.RECTANGLE, { x: 0.85, y: y + 0.12, w: 11.6, h: 0.03, fill: { color: BLUE_L }, line: { type: "none" } });
  steps.forEach(([when, what, desc], i) => {
    const cx = x + i * (cw + gap);
    // dot
    s.addShape(pres.shapes.OVAL, { x: cx + 0.05, y: y, w: 0.28, h: 0.28, fill: { color: LIME }, line: { color: BG, width: 2 } });
    s.addText(when, { x: cx, y: y + 0.45, w: cw, h: 0.35, margin: 0, fontFace: HEAD, fontSize: 14, bold: true, color: LIME });
    s.addText(what, { x: cx, y: y + 0.85, w: cw, h: 0.7, margin: 0, fontFace: BODY, fontSize: 13, bold: true, color: TEXT, valign: "top" });
    s.addText(desc, { x: cx, y: y + 1.6, w: cw, h: 1.3, margin: 0, fontFace: BODY, fontSize: 11.5, color: SUB, valign: "top", lineSpacingMultiple: 1.05 });
  });
  footer(s, 11);
})();

// ============================================================
// SLIDE 12 — FACILITY LAYOUT
// ============================================================
(() => {
  const s = pres.addSlide(); bg(s);
  eyebrow(s, "Facility Layout", 0.6, 0.5, 8);
  title(s, "Class A · 6,500 SF · Playa Vista, CA", 0.6, 0.87, 12, 30);

  const spaces = [
    ["2 Operating Rooms", "Surgical-grade, activated Month 16."],
    ["1 Procedure Room", "Dedicated, activated Month 16."],
    ["Exam + Recovery", "Exam rooms + post-op recovery bays."],
    ["2× Aescape Robots", "Robotic massage stations (85% margin)."],
    ["IV + Regenerative", "IV therapy lounge + regenerative suite."],
    ["Med Spa Suite", "Dedicated aesthetics & injectables suite."],
    ["2× Ammortal + HBOT", "Recovery chambers + hyperbaric oxygen."],
    ["Support Spaces", "Kitchen, restrooms, clinical/admin offices."],
  ];
  const cols = 4, cw = 2.92, gap = 0.27, ch = 1.7, rgap = 0.3;
  let x0 = 0.6, y0 = 2.0;
  spaces.forEach((sp, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = x0 + col * (cw + gap), y = y0 + row * (ch + rgap);
    card(s, x, y, cw, ch);
    s.addShape(pres.shapes.RECTANGLE, { x: x + 0.3, y: y, w: 0.55, h: 0.05, fill: { color: LIME }, line: { type: "none" } });
    s.addText(sp[0], { x: x + 0.3, y: y + 0.25, w: cw - 0.6, h: 0.55, margin: 0, fontFace: HEAD, fontSize: 14.5, bold: true, color: TEXT, valign: "top" });
    s.addText(sp[1], { x: x + 0.3, y: y + 0.82, w: cw - 0.6, h: 0.75, margin: 0, fontFace: BODY, fontSize: 11, color: SUB, valign: "top", lineSpacingMultiple: 1.0 });
  });
  footer(s, 12);
})();

// ============================================================
// SLIDE 13 — THE ASK
// ============================================================
(() => {
  const s = pres.addSlide(); bg(s);
  s.addShape(pres.shapes.OVAL, { x: 9.2, y: -3, w: 7, h: 7, fill: { color: LIME, transparency: 90 }, line: { type: "none" } });
  eyebrow(s, "The Ask", 0.6, 0.5, 6);

  // ask block
  const ax = 0.6, ay = 1.0, aw = 7.3, ah = 5.5;
  card(s, ax, ay, aw, ah, { fill: "11140A", border: LIME, lw: 2 });
  s.addText("RAISING", { x: ax + 0.5, y: ay + 0.4, w: 4, h: 0.3, margin: 0, fontFace: BODY, fontSize: 12, bold: true, color: LIME, charSpacing: 3 });
  s.addText("$5M", { x: ax + 0.5, y: ay + 0.7, w: aw - 1, h: 1.6, margin: 0, fontFace: HEAD, fontSize: 80, bold: true, color: LIME });
  s.addText("To open the Phase 1 clinic, build out two ORs and a procedure room, and carry working capital through Month 16 license activation.",
    { x: ax + 0.5, y: ay + 2.35, w: aw - 1, h: 0.9, margin: 0, fontFace: BODY, fontSize: 14, color: TEXT, valign: "top", lineSpacingMultiple: 1.05 });
  s.addShape(pres.shapes.RECTANGLE, { x: ax + 0.5, y: ay + 3.35, w: aw - 1, h: 0.015, fill: { color: BORDER }, line: { type: "none" } });
  const facts = [["Instrument","SAFE · milestone-tranched"],["Valuation Cap","Targeting $15M–$25M"],["First Revenue","Month 3–6 (clinic)"]];
  let fx = ax + 0.5; const fw = (aw - 1) / 3;
  facts.forEach(([k,v]) => {
    s.addText(k.toUpperCase(), { x: fx, y: ay + 3.6, w: fw - 0.1, h: 0.3, margin: 0, fontFace: BODY, fontSize: 9.5, bold: true, color: SUB, charSpacing: 1.5 });
    s.addText(v, { x: fx, y: ay + 3.92, w: fw - 0.1, h: 0.9, margin: 0, fontFace: BODY, fontSize: 13, bold: true, color: TEXT, valign: "top" });
    fx += fw;
  });
  s.addText("Converts at the next priced round. Cap and discount finalized with lead investor.",
    { x: ax + 0.5, y: ay + ah - 0.5, w: aw - 1, h: 0.35, margin: 0, fontFace: BODY, fontSize: 9.5, italic: true, color: SUB, valign: "middle" });

  // use of funds (right)
  const ux = 8.25, uw = 4.45;
  eyebrow(s, "Use of Funds", ux, 1.0, 4);
  const uses = [
    ["Phase 1 Clinic Buildout", "$500K–$900K"],
    ["Phase 2 OR Construction + Equipment", "$2.5M–$3.25M"],
    ["Aesthetic & Recovery Equipment", "$400K–$800K"],
    ["Working Capital (to Month 16)", "$500K–$750K"],
    ["Pre-Opening (legal, licensing)", "$500K–$600K"],
  ];
  let uy = 1.5; const urh = 0.84;
  uses.forEach(([k,v]) => {
    card(s, ux, uy, uw, urh - 0.14, { shadow: false });
    s.addText(k, { x: ux + 0.28, y: uy, w: uw - 1.7, h: urh - 0.14, margin: 0, fontFace: BODY, fontSize: 11, color: TEXT, valign: "middle" });
    s.addText(v, { x: ux + uw - 1.55, y: uy, w: 1.3, h: urh - 0.14, margin: 0, fontFace: BODY, fontSize: 11.5, bold: true, color: LIME, align: "right", valign: "middle" });
    uy += urh;
  });
  card(s, ux, uy, uw, 0.7, { fill: "1C1C12", border: LIME, lw: 1.2, shadow: false });
  s.addText("DEPLOYABLE RANGE", { x: ux + 0.28, y: uy, w: uw - 1.7, h: 0.7, margin: 0, fontFace: BODY, fontSize: 11, bold: true, color: TEXT, valign: "middle", charSpacing: 1 });
  s.addText("$4.4M–$6.3M", { x: ux + uw - 1.85, y: uy, w: 1.6, h: 0.7, margin: 0, fontFace: HEAD, fontSize: 14, bold: true, color: LIME, align: "right", valign: "middle" });
  footer(s, 13);
})();

// ============================================================
// SLIDE 14 — RAISE STRUCTURE
// ============================================================
(() => {
  const s = pres.addSlide(); bg(s);
  eyebrow(s, "Raise Structure", 0.6, 0.5, 8);
  title(s, "Priced on Proof, Not Promise", 0.6, 0.87, 12, 30);
  s.addText("A SAFE released in two milestone-tied tranches — capital deploys against demonstrated progress, and the larger tranche prices after the clinic proves the model.",
    { x: 0.6, y: 1.72, w: 12.2, h: 0.6, margin: 0, fontFace: BODY, fontSize: 13.5, color: SUB, lineSpacingMultiple: 1.05 });

  // two tranche cards
  const tw = 5.95, gap = 0.4, th = 1.85; let x = 0.6, y = 2.5;
  const tr = [
    ["Tranche 1 — Clinic & Pre-Opening", "Funds the Phase 1 buildout, accreditation, and working capital. The earliest, highest-risk capital — converts at the cap."],
    ["Tranche 2 — OR Construction", "Releases at the clinic-open milestone, once revenue is flowing. De-risked capital funding the Phase 2 surgical core."],
  ];
  tr.forEach(([h, b], i) => {
    const cx = x + i * (tw + gap);
    card(s, cx, y, tw, th);
    s.addShape(pres.shapes.RECTANGLE, { x: cx, y: y + 0.16, w: 0.07, h: th - 0.32, fill: { color: LIME }, line: { type: "none" } });
    s.addText(h, { x: cx + 0.4, y: y + 0.28, w: tw - 0.7, h: 0.5, margin: 0, fontFace: HEAD, fontSize: 16, bold: true, color: TEXT });
    s.addText(b, { x: cx + 0.4, y: y + 0.85, w: tw - 0.7, h: 0.9, margin: 0, fontFace: BODY, fontSize: 12.5, color: SUB, valign: "top", lineSpacingMultiple: 1.05 });
  });

  // three dial cards
  const dw = 3.85, dgap = 0.42, dh = 1.95; let dx = 0.6, dy = 4.6;
  const dl = [
    ["Physician Co-Ownership Pool", "A defined equity pool reserved up front for surgeon partners under California ASC rules — aligning surgical volume without surprise dilution."],
    ["Investor Protections", "Valuation cap, conversion discount, and pro-rata rights for the lead — standard early-stage terms, calibrated in the term sheet."],
    ["Financing Layers", "Equipment leases, NMTC, and SBA 7(a) considered for OR buildout — reducing the equity required inside the $5M."],
  ];
  dl.forEach(([h, b]) => {
    card(s, dx, dy, dw, dh);
    s.addText(h, { x: dx + 0.3, y: dy + 0.22, w: dw - 0.6, h: 0.55, margin: 0, fontFace: HEAD, fontSize: 13.5, bold: true, color: LIME, valign: "top" });
    s.addText(b, { x: dx + 0.3, y: dy + 0.78, w: dw - 0.6, h: 1.05, margin: 0, fontFace: BODY, fontSize: 10.5, color: SUB, valign: "top", lineSpacingMultiple: 1.0 });
    dx += dw + dgap;
  });
  s.addText("Final structure subject to lead-investor term sheet and counsel. Not an offer to sell securities.",
    { x: 0.6, y: 6.72, w: 12, h: 0.3, margin: 0, fontFace: BODY, fontSize: 9, italic: true, color: "55555F" });
  footer(s, 14);
})();

// ============================================================
// SLIDE 15 — EXPANSION & VALUE CREATION
// ============================================================
(() => {
  const s = pres.addSlide(); bg(s);
  eyebrow(s, "Expansion & Value Creation", 0.6, 0.5, 8);
  title(s, "This Is a Platform Bet", 0.6, 0.87, 12, 30);
  s.addText("Playa Vista is the proof site. The model is built to replicate — same clinic-first sequencing, same payer mix, same wellness flywheel — across affluent metro pockets where Westside dynamics translate.",
    { x: 0.6, y: 1.7, w: 12.3, h: 0.6, margin: 0, fontFace: BODY, fontSize: 12.5, color: SUB, lineSpacingMultiple: 1.05 });

  const ex = [
    ["01 · Geographic Replication", "Westside Playbook, Other Metro Pockets", "Demand profile maps to Newport Beach, Manhattan Beach, Scottsdale, Austin, Miami. Lower buildout risk on site #2 onward."],
    ["02 · Clinic-First Sequencing", "Each Site Cash-Flows Before It Cuts", "Clinic opens Month 3–6, generating $300–400K/mo to fund the OR buildout. Capital efficiency compounds per location."],
    ["03 · Recurring Revenue Layer", "Membership & Programs", "Longevity, hormone, GLP-1, and IV/peptide therapy support subscription structures — episodic visits become retained patients."],
    ["04 · Proprietary Tech Moat", "Aescape, Ammortal, HBOT", "Robotic massage, recovery chambers, and HBOT at 80–85% margins with near-zero physician dependency. Margin defense at scale."],
    ["05 · Surgical-Wellness Loop", "Pre-Op Optimization → Recovery", "Wellness drives the surgical pipeline; surgical patients flow back into wellness. The integrated loop is the network effect."],
    ["06 · Strategic Optionality", "Exit Routes Open at Each Stage", "Multi-site footprint opens hospital systems, PE-backed ASC roll-ups, and longevity platforms — optionality preserved throughout."],
  ];
  const cols = 3, cw = 3.95, gap = 0.33, ch = 2.05, rgap = 0.25;
  let x0 = 0.6, y0 = 2.4;
  ex.forEach((e, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = x0 + col * (cw + gap), y = y0 + row * (ch + rgap);
    card(s, x, y, cw, ch);
    s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 0.9, h: 0.05, fill: { color: LIME }, line: { type: "none" } });
    s.addText(e[0].toUpperCase(), { x: x + 0.28, y: y + 0.2, w: cw - 0.56, h: 0.28, margin: 0, fontFace: BODY, fontSize: 9, bold: true, color: SUB, charSpacing: 1.5 });
    s.addText(e[1], { x: x + 0.28, y: y + 0.48, w: cw - 0.56, h: 0.6, margin: 0, fontFace: HEAD, fontSize: 13, bold: true, color: TEXT, valign: "top" });
    s.addText(e[2], { x: x + 0.28, y: y + 1.05, w: cw - 0.56, h: 0.9, margin: 0, fontFace: BODY, fontSize: 10, color: SUB, valign: "top", lineSpacingMultiple: 1.0 });
  });
  footer(s, 15);
})();

// ============================================================
// SLIDE 16 — TEAM + NEXT STEPS + CTA
// ============================================================
(() => {
  const s = pres.addSlide(); bg(s);
  eyebrow(s, "Team & Next Steps", 0.6, 0.5, 8);
  title(s, "Who Builds This — and What's Next", 0.6, 0.87, 12.5, 30);

  // team (left)
  const team = [
    ["Founder / Operator — Laura Brody", "In Place", LIME, BG],
    ["Medical Director (MD/DO)", "Recruiting", "333350", TEXT],
    ["Wellness Clinical Director", "Recruiting", "333350", TEXT],
    ["Aesthetics / Injector (NP/PA)", "Recruiting", "333350", TEXT],
    ["Chief AI Officer", "Recruiting", "333350", TEXT],
    ["Surgeon Partners (Ortho/Spine)", "Recruiting", "333350", TEXT],
    ["Anesthesia (MD vs. CRNA)", "TBD", "2A2A3A", SUB],
  ];
  const tx = 0.6, tw = 6.0; let ty = 2.0; const trh = 0.5;
  team.forEach(([role, st, fill, txtc]) => {
    card(s, tx, ty, tw, trh - 0.08, { shadow: false });
    s.addText(role, { x: tx + 0.28, y: ty, w: tw - 1.9, h: trh - 0.08, margin: 0, fontFace: BODY, fontSize: 11.5, color: role.indexOf("Laura") >= 0 ? LIME : TEXT, bold: role.indexOf("Laura") >= 0, valign: "middle" });
    badge(s, st, tx + tw - 1.55, ty + 0.04, fill, txtc);
    ty += trh;
  });

  // next steps (right)
  eyebrow(s, "The Path Forward", 7.0, 1.7, 5);
  const steps = [
    "Finalize raise structure, close commitments.",
    "Execute Playa Vista Class A lease.",
    "Begin Phase 1 clinic buildout (Mo 3–6 open).",
    "File AAAHC pre-app + California SURGC.",
    "Procure Aescape, Ammortal, HBOT, lasers.",
    "Hire clinical + medical directors, injector.",
    "Commence OR construction (Mo 16 target).",
    "Onboard surgeons + execute payer contracts.",
  ];
  let sy = 2.12; const srh = 0.4;
  steps.forEach((st, i) => {
    s.addText(String(i+1).padStart(2,"0"), { x: 7.0, y: sy, w: 0.5, h: srh, margin: 0, fontFace: HEAD, fontSize: 14, bold: true, color: LIME, valign: "middle" });
    s.addText(st, { x: 7.55, y: sy, w: 5.2, h: srh, margin: 0, fontFace: BODY, fontSize: 11.5, color: TEXT, valign: "middle" });
    sy += srh;
  });

  // CTA strip
  card(s, 0.6, 5.5, 12.1, 1.3, { fill: "11140A", border: LIME, lw: 1.5 });
  s.addText("LET'S BUILD THE WESTSIDE'S FIRST INTEGRATED ASC", { x: 0.9, y: 5.7, w: 8.5, h: 0.5, margin: 0, fontFace: HEAD, fontSize: 20, bold: true, color: TEXT, valign: "middle" });
  s.addText("Vertically integrated surgical care and longevity — under one roof, in Playa Vista.", { x: 0.9, y: 6.18, w: 8.3, h: 0.45, margin: 0, fontFace: BODY, fontSize: 12, color: SUB, valign: "middle" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 9.65, y: 5.95, w: 2.8, h: 0.5, rectRadius: 0.25, fill: { color: LIME }, line: { type: "none" } });
  s.addText("Laura@culvercitysurgical.com", { x: 9.65, y: 5.95, w: 2.8, h: 0.5, margin: 0, fontFace: BODY, fontSize: 11, bold: true, color: BG, align: "center", valign: "middle" });
  footer(s, 16);
})();

// ---- write ----
const out = process.argv[2] || "Playa-Vista-ASC-Investor-Deck.pptx";
pres.writeFile({ fileName: out }).then(f => console.log("WROTE " + f)).catch(e => { console.error(e); process.exit(1); });
