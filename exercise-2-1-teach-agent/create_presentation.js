const PptxGenJS = require("pptxgenjs");
const path = require("path");

const pptx = new PptxGenJS();

// ── Prosus Brand System ────────────────────────────────────────
// Source: https://www.prosus.com/stylesheets/global.css
const C = {
  cobaltDark:    "060F76",
  cobaltPrime:   "1136A8",
  cobaltLight:   "1D5EDC",
  graphiteDark:  "191919",
  graphitePrime: "3F3F3F",
  graphiteLight: "666666",
  chloridePrime: "00AC8E",
  chlorideLight: "00D3AF",
  argonPrime:    "AF015F",
  argonLight:    "D90276",
  oxidePrime:    "FCB828",
  oxideLight:    "FCDA28",
  lithiumPrime:  "5715B0",
  lithiumLight:  "7729E0",
  whiteDark:     "F1F3F4",
  whitePrime:    "F6F7F9",
  white:         "FFFFFF",
  textColor:     "4D4D4D",
};

// Slide dimensions (widescreen 10" x 5.625")
const SW = 10, SH = 5.625;

// Font — Neuzeit Grotesk is Typekit-only; Gill Sans is the closest system match
const FONT = "Gill Sans MT";
const FONT_BOLD = "Gill Sans MT";

const LOGO_PATH = path.join(__dirname, "prosus_logo.png");

// ── Helpers ────────────────────────────────────────────────────

function addLogo(slide) {
  // Top-right, 1.4" wide, proportional (567:152 ≈ 3.73:1)
  const w = 1.4, h = w / (567 / 152);
  slide.addImage({
    path: LOGO_PATH,
    x: SW - w - 0.25,
    y: 0.18,
    w,
    h,
  });
}

// Cobalt gradient bar across the full width at given y
function addCobaltBar(slide, y, h = 0.06) {
  // Simulate gradient with two overlapping rects
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y, w: SW, h,
    fill: { type: "solid", color: C.cobaltDark },
    line: { color: C.cobaltDark },
  });
  // lighter half
  slide.addShape(pptx.ShapeType.rect, {
    x: SW / 2, y, w: SW / 2, h,
    fill: { type: "solid", color: C.cobaltLight },
    line: { color: C.cobaltLight },
  });
}

// Accent rule (argon magenta) under a heading
function addAccentRule(slide, x, y, w) {
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w, h: 0.045,
    fill: { color: C.argonLight }, line: { color: C.argonLight },
  });
}

// Section title
function addSectionTitle(slide, text, y = 0.28) {
  slide.addText(text.toUpperCase(), {
    x: 0.38, y, w: 7.2, h: 0.48,
    fontSize: 18, bold: true,
    color: C.cobaltDark,
    fontFace: FONT_BOLD,
    charSpacing: 1.5,
  });
  addAccentRule(slide, 0.38, y + 0.5, 9.35);
}

// Metric card on white bg with cobalt top border
function metricCard(slide, x, y, w, h, label, value, sub, accentColor) {
  const border = accentColor || C.cobaltPrime;
  // card shadow illusion
  slide.addShape(pptx.ShapeType.rect, {
    x: x + 0.04, y: y + 0.04, w, h,
    fill: { color: "D8D8D8" }, line: { color: "D8D8D8" },
  });
  // white card
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w, h,
    fill: { color: C.white }, line: { color: "E0E0E0", width: 0.5 },
  });
  // top accent bar
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w, h: 0.055,
    fill: { color: border }, line: { color: border },
  });
  slide.addText([
    { text: value + "\n", options: { fontSize: 22, bold: true, color: border, breakLine: true } },
    { text: label + "\n", options: { fontSize: 8.5, bold: true, color: C.graphitePrime, breakLine: true } },
    { text: sub, options: { fontSize: 8, color: C.graphiteLight } },
  ], { x: x + 0.15, y: y + 0.15, w: w - 0.3, h: h - 0.2, valign: "top", align: "left" });
}

// Segment card — cobalt-dark bg, white text
function segmentCard(slide, x, y, w, h, title, lines, accent) {
  const ac = accent || C.oxidePrime;
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w, h,
    fill: { color: C.cobaltDark }, line: { color: C.cobaltPrime, width: 0.5 },
  });
  // left accent stripe
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w: 0.045, h,
    fill: { color: ac }, line: { color: ac },
  });
  const rows = [
    { text: title + "\n", options: { fontSize: 9.5, bold: true, color: ac, breakLine: true } },
    ...lines.map(l => ({ text: l + "\n", options: { fontSize: 8.5, color: C.white, breakLine: true } })),
  ];
  slide.addText(rows, { x: x + 0.12, y: y + 0.1, w: w - 0.2, h: h - 0.2, valign: "top" });
}

// ══════════════════════════════════════════════════════════════
// SLIDE 1 — TITLE
// ══════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  s.background = { color: C.white };

  // Full-bleed cobalt hero panel (left 55%)
  s.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 5.8, h: SH,
    fill: { color: C.cobaltDark }, line: { color: C.cobaltDark },
  });
  // subtle lighter cobalt strip
  s.addShape(pptx.ShapeType.rect, {
    x: 4.8, y: 0, w: 1.0, h: SH,
    fill: { color: C.cobaltPrime }, line: { color: C.cobaltPrime },
  });

  // Oxide gold accent bar (top)
  s.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 5.8, h: 0.07,
    fill: { color: C.oxidePrime }, line: { color: C.oxidePrime },
  });

  // Logo — top-right on white section
  addLogo(s);

  // PROSUS wordmark area
  s.addText("PROSUS N.V.", {
    x: 0.45, y: 1.1, w: 5.0, h: 0.8,
    fontSize: 44, bold: true, color: C.white, fontFace: FONT_BOLD,
    charSpacing: 2,
  });
  s.addText("FY2025 Annual Report", {
    x: 0.45, y: 1.95, w: 5.0, h: 0.45,
    fontSize: 18, color: C.oxideLight, fontFace: FONT,
  });
  s.addText("Management Overview", {
    x: 0.45, y: 2.42, w: 5.0, h: 0.35,
    fontSize: 13, color: "AAC4F0", fontFace: FONT,
  });

  // Argon rule
  s.addShape(pptx.ShapeType.rect, {
    x: 0.45, y: 2.85, w: 3.2, h: 0.045,
    fill: { color: C.argonLight }, line: { color: C.argonLight },
  });

  s.addText("Year ended 31 March 2025  ·  All figures in USD unless stated", {
    x: 0.45, y: 2.95, w: 5.0, h: 0.28,
    fontSize: 9, color: "7A9ED4", fontFace: FONT,
  });

  // 4 KPI cards on white side
  const kpis = [
    { label: "Revenue", value: "$6.2bn", sub: "+21% local currency", ac: C.cobaltLight },
    { label: "Group aEBIT", value: "$179m", sub: "+100% year-on-year", ac: C.chloridePrime },
    { label: "Core Headline Earnings", value: "$7.4bn", sub: "+47% year-on-year", ac: C.argonLight },
    { label: "Free Cash Flow", value: "$1.0bn", sub: "vs $422m in FY24", ac: C.oxidePrime },
  ];
  kpis.forEach(({ label, value, sub, ac }, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    metricCard(s, 6.1 + col * 1.85, 1.0 + row * 1.6, 1.72, 1.35, label, value, sub, ac);
  });

  // Bottom strip — dark bar with leadership names
  s.addShape(pptx.ShapeType.rect, {
    x: 0, y: SH - 0.52, w: SW, h: 0.52,
    fill: { color: C.graphiteDark }, line: { color: C.graphiteDark },
  });
  s.addText("CEO: Fabricio Bloisi   ·   CFO: Nico Marais   ·   Chair: Koos Bekker   ·   Euronext Amsterdam  ·  JSE", {
    x: 0.4, y: SH - 0.48, w: 9.2, h: 0.42,
    fontSize: 8.5, color: "AAAAAA", fontFace: FONT, align: "center", valign: "middle",
  });
}

// ══════════════════════════════════════════════════════════════
// SLIDE 2 — FINANCIAL PERFORMANCE
// ══════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  s.background = { color: C.whitePrime };

  addCobaltBar(s, 0, 0.72);
  addLogo(s);

  s.addText("Financial Performance", {
    x: 0.38, y: 0.12, w: 7.5, h: 0.5,
    fontSize: 22, bold: true, color: C.white, fontFace: FONT_BOLD, charSpacing: 0.5,
  });
  s.addText("FY2025 · Year ended 31 March 2025", {
    x: 0.38, y: 0.48, w: 7.5, h: 0.22,
    fontSize: 9, color: "AAC4F0", fontFace: FONT,
  });

  // ── aEBIT Bar Chart ──────────────────────────────────────────
  const chartX = 0.38, chartY = 0.88, chartW = 4.6, chartH = 2.35;
  const years  = ["FY22", "FY23", "FY24", "FY25"];
  const values = [-586,   -537,   -118,    179];
  const maxAbs = 640;
  const barW   = 0.72;
  const gap    = (chartW - years.length * barW) / (years.length + 1);
  const zeroY  = chartY + chartH * (maxAbs / (maxAbs * 2));

  // chart bg
  s.addShape(pptx.ShapeType.rect, {
    x: chartX, y: chartY, w: chartW, h: chartH,
    fill: { color: C.white }, line: { color: "DEDEDE", width: 0.5 },
  });
  // chart label
  s.addText("GROUP aEBIT TRAJECTORY  (US$M)", {
    x: chartX, y: chartY - 0.22, w: chartW, h: 0.2,
    fontSize: 7.5, bold: true, color: C.graphitePrime, charSpacing: 1, fontFace: FONT,
  });
  // zero axis
  s.addShape(pptx.ShapeType.line, {
    x: chartX, y: zeroY, w: chartW, h: 0,
    line: { color: C.graphiteLight, width: 0.75 },
  });

  years.forEach((yr, i) => {
    const bx    = chartX + gap + i * (barW + gap);
    const pct   = Math.abs(values[i]) / maxAbs;
    const bh    = pct * (chartH / 2);
    const isNeg = values[i] < 0;
    const by    = isNeg ? zeroY : zeroY - bh;
    // use chloride for positive, argon for negative
    const clr   = isNeg ? C.argonPrime : C.chloridePrime;

    s.addShape(pptx.ShapeType.rect, {
      x: bx, y: by, w: barW, h: bh,
      fill: { color: clr }, line: { color: clr },
    });
    s.addText(String(values[i]), {
      x: bx, y: isNeg ? by + bh + 0.03 : by - 0.22, w: barW, h: 0.2,
      fontSize: 8.5, bold: true, color: isNeg ? C.argonPrime : C.chloridePrime,
      align: "center", fontFace: FONT_BOLD,
    });
    s.addText(yr, {
      x: bx, y: zeroY + (isNeg ? -0.22 : 0.04), w: barW, h: 0.2,
      fontSize: 8, color: C.graphitePrime, align: "center", fontFace: FONT,
    });
  });

  // ── Revenue Mix (right panel) ────────────────────────────────
  const rx = 5.28;
  s.addText("REVENUE MIX  (US$6.2BN)", {
    x: rx, y: chartY - 0.22, w: 4.4, h: 0.2,
    fontSize: 7.5, bold: true, color: C.graphitePrime, charSpacing: 1, fontFace: FONT,
  });
  s.addShape(pptx.ShapeType.rect, {
    x: rx, y: chartY, w: 4.4, h: chartH,
    fill: { color: C.white }, line: { color: "DEDEDE", width: 0.5 },
  });

  const revItems = [
    { label: "Etail (goods)",          pct: 38, val: "$2.3bn", color: C.cobaltDark },
    { label: "Payments & Fintech",     pct: 21, val: "$1.3bn", color: C.cobaltLight },
    { label: "Food Delivery",          pct: 20, val: "$1.3bn", color: C.chloridePrime },
    { label: "Classifieds",            pct: 12, val: "$717m",  color: C.oxidePrime },
    { label: "Other",                  pct:  9, val: "$0.6bn", color: C.lithiumLight },
  ];
  const BAR_TOTAL_W = 3.8;
  revItems.forEach(({ label, pct, val, color }, i) => {
    const ry2 = chartY + 0.18 + i * 0.42;
    const bw  = (pct / 100) * BAR_TOTAL_W;
    s.addShape(pptx.ShapeType.rect, {
      x: rx + 0.3, y: ry2 + 0.06, w: bw, h: 0.22,
      fill: { color }, line: { color },
    });
    s.addText(`${label}`, {
      x: rx + 0.3, y: ry2, w: 2.4, h: 0.18,
      fontSize: 8, color: C.graphitePrime, fontFace: FONT,
    });
    s.addText(`${val}  (${pct}%)`, {
      x: rx + 0.3 + bw + 0.1, y: ry2 + 0.05, w: 1.4, h: 0.22,
      fontSize: 8, bold: true, color: color, fontFace: FONT_BOLD, valign: "middle",
    });
  });

  // ── Bottom metric row ────────────────────────────────────────
  const mY = 3.38;
  const metrics = [
    { label: "Revenue Growth", value: "+13%", sub: "+21% local currency", ac: C.cobaltLight },
    { label: "SG&A Costs",     value: "$2.5bn", sub: "Staff costs +1% only", ac: C.graphitePrime },
    { label: "Net Cash",       value: "$2.6bn", sub: "FCF $1.0bn", ac: C.chloridePrime },
    { label: "Dividend/share", value: "€0.20", sub: "+100% YoY increase",  ac: C.oxidePrime },
    { label: "NAV/Share gain", value: "+11%", sub: "Since buyback Jun 2022", ac: C.argonLight },
  ];
  metrics.forEach(({ label, value, sub, ac }, i) => {
    metricCard(s, 0.38 + i * 1.88, mY, 1.76, 1.74, label, value, sub, ac);
  });
}

// ══════════════════════════════════════════════════════════════
// SLIDE 3 — BUSINESS SEGMENTS
// ══════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  s.background = { color: C.cobaltDark };

  // Subtle lighter overlay at top
  s.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: SW, h: 0.72,
    fill: { color: C.cobaltPrime }, line: { color: C.cobaltPrime },
  });
  s.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: SW, h: 0.06,
    fill: { color: C.oxidePrime }, line: { color: C.oxidePrime },
  });

  addLogo(s);
  s.addText("Business Segments", {
    x: 0.38, y: 0.1, w: 7.5, h: 0.42,
    fontSize: 22, bold: true, color: C.white, fontFace: FONT_BOLD, charSpacing: 0.5,
  });
  s.addText("FY2025 Performance Highlights", {
    x: 0.38, y: 0.46, w: 7.5, h: 0.22,
    fontSize: 9, color: "AAC4F0", fontFace: FONT,
  });

  const accents = [C.oxidePrime, C.chlorideLight, C.argonLight, C.oxideLight, C.lithiumLight, C.chloridePrime];
  const segments = [
    {
      name: "Food Delivery — iFood",
      lines: ["Revenue $1.3bn  (+30% local FX)", "aEBIT $226m  •  EBITDA margin 19.1%", "120m+ orders/month  •  56m annual users", "AI: 56% customer support automated"],
    },
    {
      name: "Classifieds — OLX",
      lines: ["Revenue $788m  (+18% local FX)", "aEBIT $273m  •  35% margin  (+11pp)", "63.6m monthly app users  •  9 markets", "Motors +24%  •  Real estate +23%"],
    },
    {
      name: "Payments & Fintech — PayU",
      lines: ["Revenue $1.1bn  (+34% local FX)", "iyzico (Turkey) aEBIT $18m  •  6% margin", "India loan book $558m  •  +19% YoY", "Mindgate acquisition $68m (real-time pmts)"],
    },
    {
      name: "Etail — eMAG",
      lines: ["Revenue $2.5bn  (+12%)", "aEBIT $14m  (turnaround from -$26m)", "Romania GMV +15%  •  EBITDA $84m", "Sameday courier revenue +38%"],
    },
    {
      name: "Edtech",
      lines: ["Revenue $170m  (+15%)", "aEBIT loss -$33m  (FY24: -$98m  ↑65m)", "Near breakeven at cash flow level", "US & India workforce/K-12 focus"],
    },
    {
      name: "Tencent — 23.5% Stake",
      lines: ["Equity earnings $5.7bn  (+103%)", "Tencent non-IFRS profit +41% YoY", "WeChat 1.39bn MAU  •  Dividend $1.0bn", "Planned 2025 buyback HKD80bn+"],
    },
  ];

  const cols = 3, cw = 3.05, ch = 1.42;
  const startX = 0.38, startY = 0.83;
  segments.forEach(({ name, lines }, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    segmentCard(s,
      startX + col * (cw + 0.12),
      startY + row * (ch + 0.1),
      cw, ch, name, lines, accents[i]
    );
  });

  // Bottom label
  s.addShape(pptx.ShapeType.rect, {
    x: 0, y: SH - 0.38, w: SW, h: 0.38,
    fill: { color: "040B5A" }, line: { color: "040B5A" },
  });
  s.addText("Total permanent staff: 23,323  ·  33,246 total employees  ·  23,000+ Toqan AI users daily", {
    x: 0.4, y: SH - 0.36, w: 9.2, h: 0.32,
    fontSize: 8, color: "7A9ED4", align: "center", valign: "middle", fontFace: FONT,
  });
}

// ══════════════════════════════════════════════════════════════
// SLIDE 4 — STRATEGY & OUTLOOK
// ══════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  s.background = { color: C.white };

  addCobaltBar(s, 0, 0.72);
  addLogo(s);

  s.addText("Strategy & Outlook", {
    x: 0.38, y: 0.1, w: 7.5, h: 0.42,
    fontSize: 22, bold: true, color: C.white, fontFace: FONT_BOLD, charSpacing: 0.5,
  });
  s.addText("AI-First · Regional Ecosystems · Shareholder Value", {
    x: 0.38, y: 0.48, w: 7.5, h: 0.22,
    fontSize: 9, color: "AAC4F0", fontFace: FONT,
  });

  // Three strategy pillars
  const pillars = [
    {
      icon: "AI-FIRST\nBUSINESSES",
      body: [
        "Toqan AI: 20,000+ daily users, +11% avg productivity",
        "Large Commerce Model (LCM) in development",
        "iFood: 56% support automated, 40% cost reduction",
        "30+ portfolio companies deploying AI in production",
        "156 domains + 59 trademarks + 1 patent filed FY25",
      ],
      color: C.cobaltDark,
      accent: C.cobaltLight,
    },
    {
      icon: "REGIONAL\nECOSYSTEMS",
      body: [
        "LatAm: >100m customers, >$25bn GMV post-Despegar",
        "iFood: 60m customers, 120m+ monthly orders",
        "India: Swiggy, PayU, Meesho, Urban Company, Rapido",
        "Europe: eMAG, OLX, iyzico + Just Eat (pending, €4.1bn)",
        "Digital credit CAGR ~27% FY23–30 (India market)",
      ],
      color: C.chloridePrime,
      accent: C.chlorideLight,
    },
    {
      icon: "SHAREHOLDER\nVALUE",
      body: [
        "Buyback: NAV/share +11% since June 2022",
        "Free float reduced by >27%",
        "Dividend doubled to €0.20/share (FY24: €0.10)",
        "Despegar acquired $1.7bn (closed May 2025)",
        "Just Eat conditional €4.1bn — 17 markets, 61m users",
      ],
      color: C.argonPrime,
      accent: C.argonLight,
    },
  ];

  pillars.forEach(({ icon, body, color, accent }, i) => {
    const px = 0.38 + i * 3.15;
    const pw = 3.0;

    // header
    s.addShape(pptx.ShapeType.rect, {
      x: px, y: 0.82, w: pw, h: 0.72,
      fill: { color }, line: { color },
    });
    s.addText(icon, {
      x: px, y: 0.82, w: pw, h: 0.72,
      fontSize: 11.5, bold: true, color: C.white, align: "center", valign: "middle",
      fontFace: FONT_BOLD, charSpacing: 1,
    });
    // accent bottom border on header
    s.addShape(pptx.ShapeType.rect, {
      x: px, y: 0.82 + 0.72 - 0.045, w: pw, h: 0.045,
      fill: { color: accent }, line: { color: accent },
    });

    // body card
    s.addShape(pptx.ShapeType.rect, {
      x: px, y: 1.54, w: pw, h: 2.58,
      fill: { color: C.whitePrime }, line: { color: "DEDEDE", width: 0.5 },
    });
    const rows = body.map(l => ({
      text: "– " + l + "\n",
      options: { fontSize: 8.8, color: C.textColor, breakLine: true },
    }));
    s.addText(rows, {
      x: px + 0.12, y: 1.62, w: pw - 0.22, h: 2.42, valign: "top", fontFace: FONT,
    });
  });

  // Bottom acquisitions strip
  s.addShape(pptx.ShapeType.rect, {
    x: 0, y: 4.25, w: SW, h: 1.0,
    fill: { color: C.graphiteDark }, line: { color: C.graphiteDark },
  });
  s.addText("KEY ACQUISITIONS  FY2025", {
    x: 0.38, y: 4.3, w: 3, h: 0.25,
    fontSize: 7.5, bold: true, color: C.oxidePrime, charSpacing: 1, fontFace: FONT_BOLD,
  });
  const acqs = [
    { name: "Despegar", detail: "LatAm #1 online travel  ·  US$1.7bn  ·  Closed May 2025", color: C.oxideLight },
    { name: "Just Eat Takeaway", detail: "EU food delivery, 17 markets, 61m customers  ·  €4.1bn  ·  Regulatory pending", color: C.chlorideLight },
    { name: "Mindgate Solutions", detail: "Real-time payments tech for PayU India  ·  US$68m", color: C.cobaltLight },
  ];
  acqs.forEach(({ name, detail, color }, i) => {
    const ay = 4.58 + i * 0.19;
    s.addShape(pptx.ShapeType.rect, {
      x: 0.38, y: ay + 0.04, w: 0.07, h: 0.12,
      fill: { color }, line: { color },
    });
    s.addText([
      { text: name + "  ", options: { bold: true, color: C.white, fontSize: 8.5 } },
      { text: detail, options: { color: "AAAAAA", fontSize: 8.5 } },
    ], { x: 0.55, y: ay, w: 9.1, h: 0.2, valign: "middle", fontFace: FONT });
  });
}

// ── Save ───────────────────────────────────────────────────────
const outPath = "Prosus_FY2025_Overview.pptx";
pptx.writeFile({ fileName: outPath }).then(() => {
  console.log(`✓ Saved: ${outPath}`);
}).catch(err => {
  console.error("Error:", err);
  process.exit(1);
});
