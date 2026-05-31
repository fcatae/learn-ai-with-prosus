const PptxGenJS = require("pptxgenjs");
const path = require("path");

const pptx = new PptxGenJS();

// ── Prosus Brand System ────────────────────────────────────────
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

const SW = 10, SH = 5.625;
const FONT = "Gill Sans MT";
const FONT_BOLD = "Gill Sans MT";
const LOGO_PATH = path.join(__dirname, "prosus_logo.png");

// ── Helpers ────────────────────────────────────────────────────

function addLogo(slide) {
  const w = 1.4, h = w / (567 / 152);
  slide.addImage({
    path: LOGO_PATH,
    x: SW - w - 0.25,
    y: 0.18,
    w,
    h,
  });
}

function addCobaltBar(slide, y, h = 0.06) {
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y, w: SW, h,
    fill: { type: "solid", color: C.cobaltDark },
    line: { color: C.cobaltDark },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: SW / 2, y, w: SW / 2, h,
    fill: { type: "solid", color: C.cobaltLight },
    line: { color: C.cobaltLight },
  });
}

function addAccentRule(slide, x, y, w) {
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w, h: 0.045,
    fill: { color: C.argonLight }, line: { color: C.argonLight },
  });
}

function metricCard(slide, x, y, w, h, label, value, sub, accentColor) {
  const border = accentColor || C.cobaltPrime;
  slide.addShape(pptx.ShapeType.rect, {
    x: x + 0.04, y: y + 0.04, w, h,
    fill: { color: "D8D8D8" }, line: { color: "D8D8D8" },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w, h,
    fill: { color: C.white }, line: { color: "E0E0E0", width: 0.5 },
  });
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

function segmentCard(slide, x, y, w, h, title, lines, accent) {
  const ac = accent || C.oxidePrime;
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w, h,
    fill: { color: C.cobaltDark }, line: { color: C.cobaltPrime, width: 0.5 },
  });
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

  s.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 5.8, h: SH,
    fill: { color: C.cobaltDark }, line: { color: C.cobaltDark },
  });
  s.addShape(pptx.ShapeType.rect, {
    x: 4.8, y: 0, w: 1.0, h: SH,
    fill: { color: C.cobaltPrime }, line: { color: C.cobaltPrime },
  });
  s.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 5.8, h: 0.07,
    fill: { color: C.chloridePrime }, line: { color: C.chloridePrime },
  });

  addLogo(s);

  s.addText("PROSUS N.V.", {
    x: 0.45, y: 1.1, w: 5.0, h: 0.8,
    fontSize: 44, bold: true, color: C.white, fontFace: FONT_BOLD,
    charSpacing: 2,
  });
  s.addText("Q1 FY2026 Trading Update", {
    x: 0.45, y: 1.95, w: 5.0, h: 0.45,
    fontSize: 18, color: C.chlorideLight, fontFace: FONT,
  });
  s.addText("Quarterly Performance Review", {
    x: 0.45, y: 2.42, w: 5.0, h: 0.35,
    fontSize: 13, color: "AAC4F0", fontFace: FONT,
  });

  s.addShape(pptx.ShapeType.rect, {
    x: 0.45, y: 2.85, w: 3.2, h: 0.045,
    fill: { color: C.chlorideLight }, line: { color: C.chlorideLight },
  });

  s.addText("Quarter ended 30 June 2025  ·  All figures in USD unless stated", {
    x: 0.45, y: 2.95, w: 5.0, h: 0.28,
    fontSize: 9, color: "7A9ED4", fontFace: FONT,
  });

  const kpis = [
    { label: "Revenue",        value: "$1.7bn",  sub: "+23% local currency",   ac: C.cobaltLight },
    { label: "Group aEBIT",    value: "$62m",    sub: "+QoQ momentum",         ac: C.chloridePrime },
    { label: "Orders (Food)",  value: "395m",    sub: "+18% vs Q1 FY25",       ac: C.argonLight },
    { label: "Free Cash Flow", value: "$258m",   sub: "vs $201m Q1 FY25",      ac: C.oxidePrime },
  ];
  kpis.forEach(({ label, value, sub, ac }, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    metricCard(s, 6.1 + col * 1.85, 1.0 + row * 1.6, 1.72, 1.35, label, value, sub, ac);
  });

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
  s.addText("Q1 FY2026 · Quarter ended 30 June 2025", {
    x: 0.38, y: 0.48, w: 7.5, h: 0.22,
    fontSize: 9, color: "AAC4F0", fontFace: FONT,
  });

  // ── aEBIT Quarterly Trend ──────────────────────────────────
  const chartX = 0.38, chartY = 0.88, chartW = 4.6, chartH = 2.35;
  const quarters = ["Q1 FY25", "Q2 FY25", "Q3 FY25", "Q4 FY25", "Q1 FY26"];
  const values   = [  28,        42,         53,         56,        62   ];
  const maxVal   = 80;
  const barW     = 0.68;
  const gap      = (chartW - quarters.length * barW) / (quarters.length + 1);
  const baseY    = chartY + chartH;

  s.addShape(pptx.ShapeType.rect, {
    x: chartX, y: chartY, w: chartW, h: chartH,
    fill: { color: C.white }, line: { color: "DEDEDE", width: 0.5 },
  });
  s.addText("GROUP aEBIT QUARTERLY TREND  (US$M)", {
    x: chartX, y: chartY - 0.22, w: chartW, h: 0.2,
    fontSize: 7.5, bold: true, color: C.graphitePrime, charSpacing: 1, fontFace: FONT,
  });

  quarters.forEach((qr, i) => {
    const bx  = chartX + gap + i * (barW + gap);
    const pct = values[i] / maxVal;
    const bh  = pct * chartH * 0.88;
    const by  = baseY - bh;
    const clr = i === quarters.length - 1 ? C.chloridePrime : C.cobaltLight;

    s.addShape(pptx.ShapeType.rect, {
      x: bx, y: by, w: barW, h: bh,
      fill: { color: clr }, line: { color: clr },
    });
    s.addText(String(values[i]), {
      x: bx, y: by - 0.22, w: barW, h: 0.2,
      fontSize: 8.5, bold: true, color: clr,
      align: "center", fontFace: FONT_BOLD,
    });
    s.addText(qr, {
      x: bx, y: baseY + 0.04, w: barW, h: 0.2,
      fontSize: 7.5, color: C.graphitePrime, align: "center", fontFace: FONT,
    });
  });

  // ── Revenue Mix (right panel) ──────────────────────────────
  const rx = 5.28;
  s.addText("Q1 FY2026 REVENUE MIX  (US$1.7BN)", {
    x: rx, y: chartY - 0.22, w: 4.4, h: 0.2,
    fontSize: 7.5, bold: true, color: C.graphitePrime, charSpacing: 1, fontFace: FONT,
  });
  s.addShape(pptx.ShapeType.rect, {
    x: rx, y: chartY, w: 4.4, h: chartH,
    fill: { color: C.white }, line: { color: "DEDEDE", width: 0.5 },
  });

  const revItems = [
    { label: "Etail (eMAG)",            pct: 37, val: "$628m",  color: C.cobaltDark },
    { label: "Payments & Fintech",      pct: 22, val: "$374m",  color: C.cobaltLight },
    { label: "Food Delivery",           pct: 21, val: "$357m",  color: C.chloridePrime },
    { label: "Classifieds (OLX)",       pct: 13, val: "$221m",  color: C.oxidePrime },
    { label: "Other",                   pct:  7, val: "$119m",  color: C.lithiumLight },
  ];
  const BAR_TOTAL_W = 3.8;
  revItems.forEach(({ label, pct, val, color }, i) => {
    const ry2 = chartY + 0.18 + i * 0.42;
    const bw  = (pct / 100) * BAR_TOTAL_W;
    s.addShape(pptx.ShapeType.rect, {
      x: rx + 0.3, y: ry2 + 0.06, w: bw, h: 0.22,
      fill: { color }, line: { color },
    });
    s.addText(label, {
      x: rx + 0.3, y: ry2, w: 2.4, h: 0.18,
      fontSize: 8, color: C.graphitePrime, fontFace: FONT,
    });
    s.addText(`${val}  (${pct}%)`, {
      x: rx + 0.3 + bw + 0.1, y: ry2 + 0.05, w: 1.4, h: 0.22,
      fontSize: 8, bold: true, color, fontFace: FONT_BOLD, valign: "middle",
    });
  });

  // ── Bottom metric row ──────────────────────────────────────
  const mY = 3.38;
  const metrics = [
    { label: "Revenue Growth",  value: "+16%",   sub: "+23% local currency",  ac: C.cobaltLight },
    { label: "aEBIT Margin",    value: "3.6%",   sub: "vs 1.9% Q1 FY25",     ac: C.chloridePrime },
    { label: "Net Cash",        value: "$2.4bn", sub: "Post Despegar close",  ac: C.graphitePrime },
    { label: "Orders Growth",   value: "+18%",   sub: "Food delivery Q1 FY26",ac: C.oxidePrime },
    { label: "Buyback Q1",      value: "$320m",  sub: "Cumulative $8.1bn",    ac: C.argonLight },
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

  s.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: SW, h: 0.72,
    fill: { color: C.cobaltPrime }, line: { color: C.cobaltPrime },
  });
  s.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: SW, h: 0.06,
    fill: { color: C.chloridePrime }, line: { color: C.chloridePrime },
  });

  addLogo(s);
  s.addText("Business Segments", {
    x: 0.38, y: 0.1, w: 7.5, h: 0.42,
    fontSize: 22, bold: true, color: C.white, fontFace: FONT_BOLD, charSpacing: 0.5,
  });
  s.addText("Q1 FY2026 Performance Highlights  ·  Quarter ended 30 June 2025", {
    x: 0.38, y: 0.46, w: 7.5, h: 0.22,
    fontSize: 9, color: "AAC4F0", fontFace: FONT,
  });

  const accents = [C.oxidePrime, C.chlorideLight, C.argonLight, C.oxideLight, C.lithiumLight, C.chloridePrime];
  const segments = [
    {
      name: "Food Delivery — iFood",
      lines: ["Revenue $357m  (+28% local FX)", "aEBIT $68m  •  EBITDA margin 21%", "395m orders  •  +18% YoY growth", "Despegar travel bundle pilot launched"],
    },
    {
      name: "Classifieds — OLX",
      lines: ["Revenue $221m  (+19% local FX)", "aEBIT $81m  •  37% margin  (+3pp)", "67m monthly app users  •  9 markets", "AI-powered listings: 38% uplift in CTR"],
    },
    {
      name: "Payments & Fintech — PayU",
      lines: ["Revenue $374m  (+31% local FX)", "iyzico Turkey: aEBIT $5m  •  Q1 seasonal", "India loan book $612m  •  +10% QoQ", "Mindgate integration on track"],
    },
    {
      name: "Etail — eMAG",
      lines: ["Revenue $628m  (+14%)", "aEBIT $8m  (continued turnaround)", "Romania logistics: Sameday +42% revenue", "Bulgaria & Hungary GMV +19%"],
    },
    {
      name: "Travel — Despegar",
      lines: ["First full quarter post-acquisition", "Revenue $98m  ·  GMV $1.2bn", "LatAm #1 OTA  ·  12 markets", "iFood x Despegar bundle synergy live"],
    },
    {
      name: "Tencent — 23.5% Stake",
      lines: ["Equity earnings est. ~$1.5bn Q1", "WeChat 1.40bn MAU  •  +0.7% QoQ", "Non-IFRS profit momentum continues", "Buyback HKD80bn+ programme ongoing"],
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

  s.addShape(pptx.ShapeType.rect, {
    x: 0, y: SH - 0.38, w: SW, h: 0.38,
    fill: { color: "040B5A" }, line: { color: "040B5A" },
  });
  s.addText("Just Eat Takeaway integration pending regulatory approval  ·  Despegar fully consolidated from May 2025  ·  Toqan AI: 25,000+ daily users", {
    x: 0.4, y: SH - 0.36, w: 9.2, h: 0.32,
    fontSize: 8, color: "7A9ED4", align: "center", valign: "middle", fontFace: FONT,
  });
}

// ══════════════════════════════════════════════════════════════
// SLIDE 4 — OUTLOOK & PRIORITIES
// ══════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  s.background = { color: C.white };

  addCobaltBar(s, 0, 0.72);
  addLogo(s);

  s.addText("Outlook & Priorities", {
    x: 0.38, y: 0.1, w: 7.5, h: 0.42,
    fontSize: 22, bold: true, color: C.white, fontFace: FONT_BOLD, charSpacing: 0.5,
  });
  s.addText("FY2026 Roadmap  ·  AI-First · Ecosystem Integration · Capital Returns", {
    x: 0.38, y: 0.48, w: 7.5, h: 0.22,
    fontSize: 9, color: "AAC4F0", fontFace: FONT,
  });

  const pillars = [
    {
      icon: "AI-FIRST\nSCALE-UP",
      body: [
        "Toqan AI: target 40,000 daily users by Q3 FY26",
        "Large Commerce Model (LCM): beta Q2 FY26",
        "iFood: 70%+ support automation target FY26",
        "OLX AI listings: roll-out to 6 additional markets",
        "PayU: real-time fraud detection via Mindgate AI",
      ],
      color: C.cobaltDark,
      accent: C.cobaltLight,
    },
    {
      icon: "ECOSYSTEM\nINTEGRATION",
      body: [
        "Despegar x iFood: LatAm super-app bundle",
        "Just Eat Takeaway: EU approval expected Q2 FY26",
        "PayU x Despegar: travel payment checkout",
        "OLX Autos: expansion into 3 new LatAm markets",
        "eMAG x Sameday: same-day pan-Europe pilot",
      ],
      color: C.chloridePrime,
      accent: C.chlorideLight,
    },
    {
      icon: "CAPITAL\nRETURNS",
      body: [
        "Buyback target: >$1.5bn in FY26",
        "Dividend maintained at ≥€0.20/share",
        "NAV/share target: +15% by FY26 year-end",
        "Tencent dividend reinvestment: $1.0bn deployed",
        "Non-core asset review ongoing (EdTech optionality)",
      ],
      color: C.argonPrime,
      accent: C.argonLight,
    },
  ];

  pillars.forEach(({ icon, body, color, accent }, i) => {
    const px = 0.38 + i * 3.15;
    const pw = 3.0;

    s.addShape(pptx.ShapeType.rect, {
      x: px, y: 0.82, w: pw, h: 0.72,
      fill: { color }, line: { color },
    });
    s.addText(icon, {
      x: px, y: 0.82, w: pw, h: 0.72,
      fontSize: 11.5, bold: true, color: C.white, align: "center", valign: "middle",
      fontFace: FONT_BOLD, charSpacing: 1,
    });
    s.addShape(pptx.ShapeType.rect, {
      x: px, y: 0.82 + 0.72 - 0.045, w: pw, h: 0.045,
      fill: { color: accent }, line: { color: accent },
    });

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

  // Bottom guidance strip
  s.addShape(pptx.ShapeType.rect, {
    x: 0, y: 4.25, w: SW, h: 1.0,
    fill: { color: C.graphiteDark }, line: { color: C.graphiteDark },
  });
  s.addText("FY2026 GUIDANCE", {
    x: 0.38, y: 4.3, w: 3, h: 0.25,
    fontSize: 7.5, bold: true, color: C.oxidePrime, charSpacing: 1, fontFace: FONT_BOLD,
  });
  const guidance = [
    { label: "Full-year Revenue",   detail: "~$7.5–7.7bn  ·  +21–24% local currency growth",                   color: C.oxideLight },
    { label: "Full-year aEBIT",     detail: "~$280–310m  ·  continued margin expansion across all segments",   color: C.chlorideLight },
    { label: "Just Eat Takeaway",   detail: "Integration proceeds post-regulatory close  ·  €4.1bn deal value", color: C.cobaltLight },
  ];
  guidance.forEach(({ label, detail, color }, i) => {
    const ay = 4.58 + i * 0.19;
    s.addShape(pptx.ShapeType.rect, {
      x: 0.38, y: ay + 0.04, w: 0.07, h: 0.12,
      fill: { color }, line: { color },
    });
    s.addText([
      { text: label + "  ", options: { bold: true, color: C.white, fontSize: 8.5 } },
      { text: detail, options: { color: "AAAAAA", fontSize: 8.5 } },
    ], { x: 0.55, y: ay, w: 9.1, h: 0.2, valign: "middle", fontFace: FONT });
  });
}

// ── Save ───────────────────────────────────────────────────────
const outPath = "Prosus_Q1_FY2026.pptx";
pptx.writeFile({ fileName: outPath }).then(() => {
  console.log(`✓ Saved: ${outPath}`);
}).catch(err => {
  console.error("Error:", err);
  process.exit(1);
});
