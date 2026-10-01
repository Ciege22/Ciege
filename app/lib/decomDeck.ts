// Decom Dashboard slide deck (3 slides: Program Overview, Top Aging Sites,
// GC Decom Accountability). Extracted verbatim from app/reports/page.tsx so
// Deck Builder V2 can reuse the exact same slide-building code (colors,
// layout, everything) instead of a browser download — the Reports page's
// own "Export Decom Dashboard" button keeps working unchanged via
// downloadDecomSlides below.
//
// Runs client-side: pptxgenjs writeFile() triggers a browser download, same
// as the SCOP deck (app/lib/scopDeck.ts).

import pptxgen from 'pptxgenjs'
import { DecomRow, MissingDecomSite, summarizeDecomByGc, uniqueDecomGcNames, computeDecomFunnel, fmtGapLabel } from './decom'

const NAVY = '124191'
const TEAL = '00A0B0'
const WHITE = 'FFFFFF'
const LIGHT_GRAY = 'F4F6FA'
const RED = 'C0392B'
const AMBER = 'E67E22'
const GREEN = '27AE60'
const DARK_GRAY = '2C3E50'
const MID_GRAY = '7F8C8D'

function fmtShortDate(d: Date | null): string {
  if (!d) return '—'
  const yy = String(d.getFullYear()).slice(-2)
  return `${d.getMonth() + 1}/${d.getDate()}/${yy}`
}

function addSlideHeader(slide: pptxgen.Slide, title: string, subtitle: string) {
  slide.addShape('rect', { x: 0, y: 0, w: 10, h: 1.3, fill: { color: NAVY }, line: { type: 'none' } })
  slide.addText(title, { x: 0.35, y: 0.16, w: 7.6, h: 0.55, fontSize: 24, bold: true, color: WHITE, isTextBox: true })
  slide.addText(subtitle, { x: 0.35, y: 0.74, w: 7.6, h: 0.4, fontSize: 11, color: 'AED6F1', isTextBox: true })
  slide.addText('Nokia', { x: 7.9, y: 0.35, w: 1.75, h: 0.5, align: 'right', fontSize: 18, bold: true, color: TEAL, isTextBox: true })
}

function addSlideFooter(slide: pptxgen.Slide) {
  slide.addShape('rect', { x: 0, y: 5.38, w: 10, h: 0.24, fill: { color: TEAL }, line: { type: 'none' } })
  slide.addText('Nokia Confidential · Viaero MW Decom Program · Data sourced from Ciege PM Platform', { x: 0, y: 5.38, w: 10, h: 0.24, align: 'center', valign: 'middle', fontSize: 7, color: WHITE, isTextBox: true })
}

function renderDecomDeck(decomRows: DecomRow[], missingDecom: MissingDecomSite[], extraDroppedOff: number): pptxgen {
  const pres = new pptxgen()
  pres.layout = 'LAYOUT_16x9'

  const decomFunnel = computeDecomFunnel(decomRows, extraDroppedOff)
  // Same all-GC derivation as the on-page breakdown and the Excel dashboard —
  // every unique GC in the decom file itself, not the GC_CONFIG roster.
  const gcNames = uniqueDecomGcNames(decomRows)
  const gcSummary = summarizeDecomByGc(decomRows, gcNames, missingDecom)
    .sort((a, b) => (b.outstanding + b.pending) - (a.outstanding + a.pending))
  const todayStr = new Date().toLocaleDateString('en-US')

  // ---- Slide 1 — Program Overview ----
  const slide1 = pres.addSlide()
  addSlideHeader(slide1, 'Decom Status — Program Overview', `As of ${todayStr} · Viaero MW Program`)

  const startX = 0.35, boxW = 1.9, boxH = 1.55, boxGap = 0.42, boxY = 1.55
  const boxX = [0, 1, 2, 3].map(i => startX + i * (boxW + boxGap))
  const boxDefs = [
    { color: NAVY, label: 'Total Decom\nSites Tracked', value: decomFunnel.totalTracked, sub: 'CX Complete confirmed' },
    { color: TEAL, label: 'Dropped Off\nto Warehouse', value: decomFunnel.droppedOff, sub: 'DECOM Drop Off confirmed' },
    { color: TEAL, label: 'POD in\nPathwave', value: decomFunnel.podPathwave, sub: 'Signed POD confirmed' },
    { color: TEAL, label: 'POD in\nQuickBase', value: decomFunnel.podQuickBase, sub: 'QB approved' },
  ]
  const boxGaps: (number | null)[] = [null, decomFunnel.gap1, decomFunnel.gap2, decomFunnel.gap3]

  boxDefs.forEach((box, i) => {
    const x = boxX[i]
    slide1.addShape('rect', { x: x + 0.04, y: boxY + 0.04, w: boxW, h: boxH, fill: { color: 'D0D8E8' }, line: { type: 'none' } })
    slide1.addShape('rect', { x, y: boxY, w: boxW, h: boxH, fill: { color: box.color }, line: { type: 'none' } })
    slide1.addText(box.label, { x, y: boxY + 0.14, w: boxW, h: 0.5, align: 'center', fontSize: 10.5, bold: true, color: WHITE, isTextBox: true })
    slide1.addText(String(box.value), { x, y: boxY + 0.64, w: boxW, h: 0.55, align: 'center', fontSize: 26, bold: true, color: WHITE, isTextBox: true })
    slide1.addText(box.sub, { x, y: boxY + 1.2, w: boxW, h: 0.3, align: 'center', fontSize: 8, italic: true, color: WHITE, isTextBox: true })

    const gap = boxGaps[i]
    if (gap !== null) {
      const gapZoneX = x - boxGap
      slide1.addText('→', { x: gapZoneX, y: boxY + 0.42, w: boxGap, h: 0.4, align: 'center', fontSize: 18, bold: true, color: MID_GRAY, isTextBox: true })
      slide1.addText(fmtGapLabel(gap), { x: gapZoneX - 0.3, y: boxY + 0.82, w: boxGap + 0.6, h: 0.3, align: 'center', fontSize: 8, bold: true, color: gap > 0 ? RED : GREEN, isTextBox: true })
    }
  })

  slide1.addShape('rect', { x: 0.35, y: 3.3, w: 9.3, h: 0.58, fill: { color: 'EBF5FB' }, line: { color: TEAL, width: 1 } })
  slide1.addText(
    [
      { text: `📋 Sites Completed Within Last 7 Days — Pending Decom Entry: ${missingDecom.length} sites`, options: { bold: true, color: NAVY, fontSize: 10, breakLine: true } },
      { text: `Sites with recent construction completion being added to decom tracking. Expected to reflect on next week's report.`, options: { italic: true, color: MID_GRAY, fontSize: 8 } },
    ],
    { x: 0.5, y: 3.36, w: 9.0, h: 0.48, isTextBox: true, valign: 'top' }
  )

  slide1.addShape('rect', { x: 0.35, y: 4.05, w: 9.3, h: 1.35, fill: { color: LIGHT_GRAY }, line: { type: 'none' } })
  slide1.addText('Key Talking Points', { x: 0.5, y: 4.1, w: 9.0, h: 0.25, bold: true, color: NAVY, fontSize: 9.5, isTextBox: true })

  const pctDropped = decomFunnel.totalTracked > 0 ? Math.round((decomFunnel.droppedOff / decomFunnel.totalTracked) * 100) : 0
  const talkingPoints = [
    `${decomFunnel.totalTracked} sites confirmed construction complete and eligible for decom tracking`,
    `${decomFunnel.droppedOff} sites (${pctDropped}%) have completed physical drop-off to the warehouse`,
    `${decomFunnel.gap1} sites still outstanding — GCs must coordinate immediate drop-off`,
    `${decomFunnel.gap2} sites dropped off but pending POD confirmation in Pathwave — paperwork gap`,
    `${missingDecom.length} recently completed sites being onboarded to decom tracking — will reflect next week`,
  ]
  slide1.addText(
    talkingPoints.map(text => ({ text, options: { bullet: { type: 'bullet' as const, indent: 0.18 }, fontSize: 9, color: DARK_GRAY, breakLine: true } })),
    { x: 0.5, y: 4.38, w: 9.0, h: 0.95, isTextBox: true, valign: 'top' }
  )

  addSlideFooter(slide1)

  // ---- Slide 2 — Top Aging Sites ----
  const slide2 = pres.addSlide()
  addSlideHeader(slide2, 'Top Aging Sites — Drop Off Pending', 'Sites with longest outstanding decom — prioritized for immediate GC action')

  const truncate = (s: string, n: number) => (s.length > n ? s.slice(0, n) + '...' : s)
  const topAging = decomRows
    .filter(r => r.status === 'outstanding' || r.status === 'pending')
    .sort((a, b) => (b.aging ?? -1) - (a.aging ?? -1))
    .slice(0, 10)
    .map(r => ({
      site: truncate(r.siteName || r.hop || '—', 28),
      gc: r.gc || '—',
      cxComplete: fmtShortDate(r.cxComplete),
      aging: r.aging ?? 0,
    }))
  const agingColor = (aging: number) => (aging >= 45 ? RED : aging >= 21 ? AMBER : '2ECC71')

  slide2.addChart('bar', [{ name: 'Days Outstanding', labels: topAging.map(t => t.site), values: topAging.map(t => t.aging) }], {
    x: 0.35, y: 1.42, w: 5.6, h: 3.85,
    barDir: 'bar',
    chartColors: topAging.map(t => agingColor(t.aging)),
    showLegend: false,
    showTitle: false,
    showValue: true,
    dataLabelPosition: 'outEnd',
    dataLabelFontSize: 9,
    catAxisLineShow: false,
    catGridLine: { style: 'none' },
    valGridLine: { style: 'none' },
  })

  const tableRows2: pptxgen.TableRow[] = [
    [
      { text: 'Site', options: { bold: true, color: WHITE, fill: { color: NAVY } } },
      { text: 'GC', options: { bold: true, color: WHITE, fill: { color: NAVY } } },
      { text: 'CX Complete', options: { bold: true, color: WHITE, fill: { color: NAVY } } },
      { text: 'Days', options: { bold: true, color: WHITE, fill: { color: NAVY } } },
    ],
    ...topAging.map((t, i) => {
      const rowFill = i % 2 === 1 ? LIGHT_GRAY : WHITE
      const daysFill = t.aging >= 45 ? 'FDECEA' : t.aging >= 21 ? 'FEF9E7' : 'EAFAF1'
      const daysColor = agingColor(t.aging)
      return [
        { text: t.site, options: { fill: { color: rowFill }, fontSize: 8 } },
        { text: t.gc, options: { fill: { color: rowFill }, fontSize: 8 } },
        { text: t.cxComplete, options: { fill: { color: rowFill }, fontSize: 8 } },
        { text: String(t.aging), options: { fill: { color: daysFill }, fontSize: 8, bold: true, color: daysColor } },
      ]
    }),
  ]
  slide2.addTable(tableRows2, { x: 6.1, y: 1.42, colW: [1.5, 0.7, 0.85, 0.7], rowH: 0.355, fontSize: 8, border: { type: 'solid', color: 'DDDDDD', pt: 0.5 } })

  slide2.addText('⚠ GC action required on all red sites — drop-off overdue by 45+ days', { x: 0.35, y: 5.08, w: 5.6, h: 0.28, fontSize: 9, bold: true, color: RED, isTextBox: true })

  slide2.addShape('rect', { x: 6.1, y: 5.1, w: 0.14, h: 0.14, fill: { color: RED }, line: { type: 'none' } })
  slide2.addText('45+ days (Critical)', { x: 6.28, y: 5.06, w: 1.7, h: 0.22, fontSize: 8, color: DARK_GRAY, isTextBox: true })
  slide2.addShape('rect', { x: 7.9, y: 5.1, w: 0.14, h: 0.14, fill: { color: AMBER }, line: { type: 'none' } })
  slide2.addText('21-44 days (Urgent)', { x: 8.08, y: 5.06, w: 1.8, h: 0.22, fontSize: 8, color: DARK_GRAY, isTextBox: true })

  addSlideFooter(slide2)

  // ---- Slide 3 — GC Decom Accountability ----
  const slide3 = pres.addSlide()
  addSlideHeader(slide3, 'GC Decom Accountability — Status by Contractor', 'Tracking drop-off completion and POD confirmation per General Contractor')

  slide3.addChart('bar', [
    { name: 'Complete', labels: gcSummary.map(s => s.gc), values: gcSummary.map(s => s.complete) },
    { name: 'POD Gap', labels: gcSummary.map(s => s.gc), values: gcSummary.map(s => s.podGap) },
    { name: 'Outstanding', labels: gcSummary.map(s => s.gc), values: gcSummary.map(s => s.outstanding + s.pending) },
  ], {
    x: 0.35, y: 1.45, w: 5.8, h: 3.7,
    barDir: 'bar',
    barGrouping: 'stacked',
    chartColors: [GREEN, AMBER, RED],
    dataLabelPosition: 'ctr',
    showValue: true,
    dataLabelColor: WHITE,
    dataLabelFontSize: 8,
    showLegend: true,
    legendPos: 'b',
    legendFontSize: 9,
    showTitle: false,
  })

  const gcTotals = gcSummary.reduce((t, s) => ({
    complete: t.complete + s.complete,
    podGap: t.podGap + s.podGap,
    pending: t.pending + s.outstanding + s.pending,
  }), { complete: 0, podGap: 0, pending: 0 })

  const tableRows3: pptxgen.TableRow[] = [
    [
      { text: 'GC', options: { bold: true, color: WHITE, fill: { color: NAVY } } },
      { text: 'Done', options: { bold: true, color: WHITE, fill: { color: NAVY } } },
      { text: 'POD Gap', options: { bold: true, color: WHITE, fill: { color: NAVY } } },
      { text: 'Pend.', options: { bold: true, color: WHITE, fill: { color: NAVY } } },
    ],
    ...gcSummary.map((s, i) => {
      const rowFill = i % 2 === 1 ? LIGHT_GRAY : WHITE
      const pending = s.outstanding + s.pending
      return [
        { text: s.gc, options: { fill: { color: rowFill }, fontSize: 8 } },
        { text: String(s.complete), options: { fill: { color: rowFill }, fontSize: 8 } },
        { text: String(s.podGap), options: { fill: { color: s.podGap > 0 ? 'FFF3CD' : rowFill }, fontSize: 8 } },
        { text: String(pending), options: { fill: { color: pending > 0 ? 'FDECEA' : rowFill }, fontSize: 8, bold: pending > 0, color: pending > 0 ? RED : DARK_GRAY } },
      ]
    }),
    [
      { text: 'Total', options: { bold: true, color: WHITE, fill: { color: NAVY } } },
      { text: String(gcTotals.complete), options: { bold: true, color: WHITE, fill: { color: NAVY } } },
      { text: String(gcTotals.podGap), options: { bold: true, color: WHITE, fill: { color: NAVY } } },
      { text: String(gcTotals.pending), options: { bold: true, color: WHITE, fill: { color: NAVY } } },
    ],
  ]
  slide3.addTable(tableRows3, { x: 6.35, y: 1.48, colW: [1.5, 0.7, 0.7, 0.8], rowH: 0.38, fontSize: 8, border: { type: 'solid', color: 'DDDDDD', pt: 0.5 } })

  slide3.addText('Green = Complete · Amber = POD Gap · Red = Outstanding Drop-Off', { x: 0.35, y: 5.18, w: 5.8, h: 0.25, align: 'center', fontSize: 8, color: MID_GRAY, isTextBox: true })

  addSlideFooter(slide3)

  return pres
}

export async function downloadDecomSlides(decomRows: DecomRow[], missingDecom: MissingDecomSite[], extraDroppedOff: number) {
  const pres = renderDecomDeck(decomRows, missingDecom, extraDroppedOff)
  const today = new Date()
  const mm = String(today.getMonth() + 1).padStart(2, '0')
  const dd = String(today.getDate()).padStart(2, '0')
  const yyyy = today.getFullYear()
  await pres.writeFile({ fileName: `Decom Tracker Slides - ${mm}-${dd}-${yyyy}.pptx` })
}

// Returns the raw .pptx bytes instead of downloading — used by Deck Builder
// V2 to fold these 3 slides into its own build without a second download
// popping up in the browser.
export async function buildDecomDeckBlob(decomRows: DecomRow[], missingDecom: MissingDecomSite[], extraDroppedOff: number): Promise<Blob> {
  const pres = renderDecomDeck(decomRows, missingDecom, extraDroppedOff)
  return (await pres.write({ outputType: 'blob' })) as Blob
}
