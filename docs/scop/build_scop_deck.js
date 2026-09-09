/**
 * SCOP Status Deck Generator
 * ===========================
 * Generates the 3-slide SCOP deck (QuickBase Status, Pathwave SCOP, Overall
 * SCOP Status) using pptxgenjs. Consumes the view objects produced by
 * scop_calculations.js — this is the second half of that pipeline:
 *
 *   dataset = buildScopDataset(rawRows, settings, asOfDate)
 *   qbView      = computeQuickBaseView(dataset)
 *   pwView      = computePathwaveView(dataset)
 *   overallView = computeOverallView(dataset)
 *   buildScopDeck({ dataset, qbView, pwView, overallView, generatedDate }, outputPath)
 *
 * Every layout number, color, and ordering choice here was iterated on directly
 * with CJ and approved — see SCOP_View_Build_Spec.md §4 for the reasoning
 * behind each one (why "In Progress" is amber not red, why OAD is carved out
 * on slide 2 but not slide 3, why the proportional bars show all HOPs not just
 * the tracked subset, etc.). Treat this file as the approved reference
 * implementation, not a rough draft to redesign from scratch.
 */

const pptxgen = require("pptxgenjs");

const NAVY = "124191";
const TEAL = "00A0B0";
const DARK = "1A1A1A";
const GRAY = "6B7280";
const LIGHT_BG = "F7F9FB";
const GREEN = "1E8449";
const AMBER = "B7791F";
const WHITE = "FFFFFF";
const BORDER = "E2E8F0";
const OAD_COLOR = NAVY;
const BAR_GREY = "D0D5DD";

function kpiCard(slide, x, y, w, h, value, label, color, sub) {
  slide.addShape("roundRect", { x, y, w, h, rectRadius: 0.07, fill: { color: LIGHT_BG }, line: { color: BORDER, width: 1 } });
  slide.addText(String(value), {
    x, y: y + 0.1, w, h: h - (sub ? 0.75 : 0.45), align: "center", valign: "bottom",
    fontFace: "Calibri", fontSize: 30, bold: true, color, isTextBox: true, margin: 0,
  });
  slide.addText(label, {
    x, y: y + h - (sub ? 0.68 : 0.38), w, h: 0.3, align: "center",
    fontFace: "Calibri", fontSize: 11, bold: true, color: DARK, isTextBox: true, margin: 0,
  });
  if (sub) {
    slide.addText(sub, {
      x: x + 0.15, y: y + h - 0.36, w: w - 0.3, h: 0.32, align: "center",
      fontFace: "Calibri", fontSize: 8.5, color: GRAY, isTextBox: true, margin: 0,
    });
  }
}

function footer(slide, text) {
  slide.addText(text, { x: 0.5, y: 7.1, w: 12.3, h: 0.3, fontFace: "Calibri", fontSize: 9, color: GRAY, isTextBox: true, margin: 0 });
}

function formatDate(d) {
  return d.toLocaleDateString("en-US", { month: "2-digit", day: "2-digit", year: "numeric" });
}

/**
 * @param {Object} data
 * @param {number} data.totalNokiaHops
 * @param {number} data.constructionComplete
 * @param {number} data.notYetConstructed
 * @param {Object} data.qbView      - output of computeQuickBaseView()
 * @param {Object} data.pwView      - output of computePathwaveView()
 * @param {Object} data.overallView - output of computeOverallView()
 * @param {Date}   data.generatedDate
 * @param {string} outputPath - where to write the .pptx file
 */
async function buildScopDeck(data, outputPath) {
  const { totalNokiaHops, constructionComplete, notYetConstructed, qbView, pwView, overallView, generatedDate } = data;
  const dateStr = formatDate(generatedDate);

  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  const cardW = 3.9, gap = 0.25;

  // ============ SLIDE 1: QuickBase Status ============
  const s1 = pres.addSlide();
  s1.background = { color: WHITE };
  s1.addText("QuickBase Status", {
    x: 0.5, y: 0.32, w: 12.3, h: 0.55, fontFace: "Calibri", fontSize: 26, bold: true, color: NAVY, isTextBox: true, margin: 0,
  });
  s1.addText(
    `${constructionComplete} of ${totalNokiaHops} HOPs have completed construction and are tracked here — ${notYetConstructed} are still under construction (Pending HOP Completion)`,
    { x: 0.5, y: 0.83, w: 12.3, h: 0.35, fontFace: "Calibri", fontSize: 12, color: GRAY, isTextBox: true, margin: 0 }
  );

  kpiCard(s1, 0.5, 1.5, cardW, 1.5, qbView.complete, "Complete", GREEN, "All QuickBase deliverables uploaded");
  kpiCard(s1, 0.5 + cardW + gap, 1.5, cardW, 1.5, qbView.pending, "In Progress", AMBER, "Construction complete, QuickBase not yet uploaded");
  kpiCard(s1, 0.5 + 2 * (cardW + gap), 1.5, cardW, 1.5, notYetConstructed, "Pending HOP Completion", GRAY, "Not yet construction complete");

  s1.addText(
    `${qbView.percentComplete}% of construction-complete sites (${qbView.complete} of ${constructionComplete}) have all QuickBase deliverables uploaded`,
    { x: 0.5, y: 3.35, w: 12.3, h: 0.4, fontFace: "Calibri", fontSize: 15, bold: true, color: DARK, isTextBox: true, margin: 0 }
  );

  {
    const barY = 3.95, barX = 0.5, barW = 12.3, barH = 0.55;
    const wComp = (barW * qbView.complete) / totalNokiaHops;
    const wPend = (barW * qbView.pending) / totalNokiaHops;
    const wNotBuilt = (barW * notYetConstructed) / totalNokiaHops;
    s1.addShape("rect", { x: barX, y: barY, w: wComp, h: barH, fill: { color: GREEN }, line: { type: "none" } });
    s1.addShape("rect", { x: barX + wComp, y: barY, w: wPend, h: barH, fill: { color: AMBER }, line: { type: "none" } });
    s1.addShape("rect", { x: barX + wComp + wPend, y: barY, w: wNotBuilt, h: barH, fill: { color: BAR_GREY }, line: { type: "none" } });
    s1.addText(`Full program view — all ${totalNokiaHops} HOPs`, {
      x: 0.5, y: barY + barH + 0.1, w: 12.3, h: 0.3, fontFace: "Calibri", fontSize: 10, italic: true, color: GRAY, isTextBox: true, margin: 0,
    });
  }

  footer(s1, `Generated ${dateStr}  |  Denominator = ${constructionComplete} construction-complete HOPs, not the full ${totalNokiaHops}`);

  // ============ SLIDE 2: Pathwave SCOP — GC Breakout ============
  const s2 = pres.addSlide();
  s2.background = { color: WHITE };
  s2.addText("Pathwave SCOP — Contractor Action Items", {
    x: 0.5, y: 0.28, w: 12.3, h: 0.5, fontFace: "Calibri", fontSize: 24, bold: true, color: NAVY, isTextBox: true, margin: 0,
  });
  s2.addText(
    `${constructionComplete} of ${totalNokiaHops} HOPs have completed construction — ${notYetConstructed} are Pending HOP Completion (still under construction)`,
    { x: 0.5, y: 0.76, w: 12.3, h: 0.28, fontFace: "Calibri", fontSize: 11.5, color: GRAY, isTextBox: true, margin: 0 }
  );

  const card4W = 2.925, card4Gap = 0.2;
  kpiCard(s2, 0.5, 1.16, card4W, 0.95, pwView.complete, "Complete & Approved", GREEN);
  kpiCard(s2, 0.5 + (card4W + card4Gap), 1.16, card4W, 0.95, pwView.inProgress, "In Progress — GC Action", AMBER);
  kpiCard(s2, 0.5 + 2 * (card4W + card4Gap), 1.16, card4W, 0.95, pwView.oad, "OAD — Tracked Separately", OAD_COLOR);
  kpiCard(s2, 0.5 + 3 * (card4W + card4Gap), 1.16, card4W, 0.95, notYetConstructed, "Pending HOP Completion", GRAY);

  {
    // full-scope bar, 4 segments in card order
    const barY2 = 2.28, barX2 = 0.5, barW2 = 12.3, barH2 = 0.38;
    const segs = [
      [pwView.complete, GREEN],
      [pwView.inProgress, AMBER],
      [pwView.oad, OAD_COLOR],
      [notYetConstructed, BAR_GREY],
    ];
    let xCursor = barX2;
    segs.forEach(([val, color]) => {
      const w = (barW2 * val) / totalNokiaHops;
      s2.addShape("rect", { x: xCursor, y: barY2, w, h: barH2, fill: { color }, line: { type: "none" } });
      xCursor += w;
    });
    s2.addText(`Full program view — all ${totalNokiaHops} HOPs`, {
      x: 0.5, y: barY2 + barH2 + 0.04, w: 12.3, h: 0.2, fontFace: "Calibri", fontSize: 9.5, italic: true, color: GRAY, isTextBox: true, margin: 0,
    });
  }

  // Two side-by-side tables: by-GC breakdown (left) + OAD sites (right)
  const tableTop = 3.0;
  s2.addText("Pathwave Items by GC", {
    x: 0.5, y: tableTop, w: 5.9, h: 0.24, fontFace: "Calibri", fontSize: 12, bold: true, color: DARK, isTextBox: true, margin: 0,
  });
  s2.addText("OAD Sites (Awaiting OAD, Not a GC Item)", {
    x: 6.7, y: tableTop, w: 6.1, h: 0.24, fontFace: "Calibri", fontSize: 12, bold: true, color: DARK, isTextBox: true, margin: 0,
  });

  const byGcEntries = Object.entries(pwView.byGC).sort((a, b) => b[1].inProgress - a[1].inProgress);
  const gcTableRows = [
    [
      { text: "GC", options: { bold: true, color: WHITE, fill: { color: NAVY }, fontSize: 9.5 } },
      { text: "In Prog.", options: { bold: true, color: WHITE, fill: { color: NAVY }, fontSize: 9.5, align: "center" } },
      { text: "Complete", options: { bold: true, color: WHITE, fill: { color: NAVY }, fontSize: 9.5, align: "center" } },
    ],
  ];
  byGcEntries.forEach(([gc, v], i) => {
    const rowFill = i % 2 === 0 ? WHITE : LIGHT_BG;
    gcTableRows.push([
      { text: gc, options: { fontSize: 9, fill: { color: rowFill }, color: DARK } },
      { text: String(v.inProgress), options: { fontSize: 9, align: "center", fill: { color: rowFill }, color: v.inProgress > 0 ? AMBER : GREEN, bold: true } },
      { text: String(v.complete), options: { fontSize: 9, align: "center", fill: { color: rowFill }, color: GREEN } },
    ]);
  });
  s2.addTable(gcTableRows, {
    x: 0.5, y: tableTop + 0.3, w: 5.9,
    colW: [2.9, 1.5, 1.5],
    rowH: [0.22, ...Array(byGcEntries.length).fill(0.175)],
    border: { type: "solid", color: BORDER, pt: 0.75 },
    autoPage: false, fontFace: "Calibri", margin: [1, 3, 1, 3], valign: "middle",
  });

  const oadTableRows = [
    [
      { text: "HOP", options: { bold: true, color: WHITE, fill: { color: OAD_COLOR }, fontSize: 9.5 } },
      { text: "GC", options: { bold: true, color: WHITE, fill: { color: OAD_COLOR }, fontSize: 9.5 } },
      { text: "Note", options: { bold: true, color: WHITE, fill: { color: OAD_COLOR }, fontSize: 9.5 } },
    ],
  ];
  (pwView.oadSites || []).forEach((s, i) => {
    const rowFill = i % 2 === 0 ? WHITE : "EEF2FA";
    oadTableRows.push([
      { text: s.hop, options: { fontSize: 8.5, fill: { color: rowFill }, color: DARK } },
      { text: s.gc, options: { fontSize: 8.5, fill: { color: rowFill }, color: DARK } },
      { text: s.note, options: { fontSize: 8.5, fill: { color: rowFill }, color: DARK } },
    ]);
  });
  s2.addTable(oadTableRows, {
    x: 6.7, y: tableTop + 0.3, w: 6.1,
    colW: [2.5, 1.5, 2.1],
    rowH: [0.22, ...Array((pwView.oadSites || []).length).fill(0.175)],
    border: { type: "solid", color: BORDER, pt: 0.75 },
    autoPage: false, fontFace: "Calibri", margin: [1, 3, 1, 3], valign: "middle",
  });

  footer(s2, `Generated ${dateStr}  |  OAD sites tracked separately — not counted as missing GC checklist items`);

  // ============ SLIDE 3: Overall SCOP Status ============
  const s3 = pres.addSlide();
  s3.background = { color: WHITE };
  s3.addText("Overall SCOP Status — Program Summary", {
    x: 0.5, y: 0.3, w: 12.3, h: 0.5, fontFace: "Calibri", fontSize: 24, bold: true, color: NAVY, isTextBox: true, margin: 0,
  });
  s3.addText(`Pathwave and QuickBase completion across the ${constructionComplete} construction-complete HOPs`, {
    x: 0.5, y: 0.78, w: 12.3, h: 0.3, fontFace: "Calibri", fontSize: 12, color: GRAY, isTextBox: true, margin: 0,
  });

  const rowY1 = 1.25, halfW = 5.9, halfH = 1.35;

  s3.addText("PATHWAVE", { x: 0.5, y: rowY1, w: halfW, h: 0.3, fontFace: "Calibri", fontSize: 13, bold: true, color: TEAL, isTextBox: true, margin: 0 });
  kpiCard(s3, 0.5, rowY1 + 0.35, halfW / 2 - 0.1, halfH, overallView.pathwaveComplete, "Complete", GREEN);
  kpiCard(s3, 0.5 + halfW / 2 + 0.1, rowY1 + 0.35, halfW / 2 - 0.1, halfH, overallView.pathwavePending, "Pending", AMBER);

  s3.addText("QUICKBASE", { x: 6.9, y: rowY1, w: halfW, h: 0.3, fontFace: "Calibri", fontSize: 13, bold: true, color: NAVY, isTextBox: true, margin: 0 });
  kpiCard(s3, 6.9, rowY1 + 0.35, halfW / 2 - 0.1, halfH, overallView.quickbaseComplete, "Complete", GREEN);
  kpiCard(s3, 6.9 + halfW / 2 + 0.1, rowY1 + 0.35, halfW / 2 - 0.1, halfH, overallView.quickbasePending, "Pending", AMBER);

  s3.addText("BOTH COMPLETE — FULL SCOP CLOSE-OUT", {
    x: 0.5, y: 4.35, w: 12.3, h: 0.3, fontFace: "Calibri", fontSize: 13, bold: true, color: DARK, isTextBox: true, margin: 0,
  });
  kpiCard(s3, 0.5, 4.7, 5.9, 1.5, overallView.fullyComplete, "100% Complete", GREEN, "Both Pathwave and QuickBase done");
  kpiCard(s3, 6.7, 4.7, 5.9, 1.5, overallView.notFullyComplete, "Not Yet Fully Complete", AMBER, "Still pending Pathwave and/or QuickBase items");

  footer(s3, `Generated ${dateStr}  |  ${overallView.trackedTotal} construction-complete HOPs (${notYetConstructed} still under construction, excluded from this view)`);

  await pres.writeFile({ fileName: outputPath });
  return outputPath;
}

module.exports = { buildScopDeck };
