'use client'

export const dynamic = 'force-dynamic'

import { useState, useCallback, useEffect } from 'react'
import { useRouter } from 'next/navigation'
import * as XLSX from 'xlsx'
import pptxgen from 'pptxgenjs'
import { supabase, loadTrackerSnapshot } from '../lib/supabase'
import { GC_CONFIG, matches, SPO_VENDOR_COL_IN_MASTER, CR_SUPPLIER_COL_IN_MASTER } from '../lib/gcConfig'
import BackToDashboard from '../components/BackToDashboard'
import { saveChunkedReport, loadChunkedReport } from '../lib/reportChunks'
import {
  parseDecomRows, summarizeDecomByGc, decomRowsForGc, DecomRow,
  parseTrackerHopsForDecom, findMissingDecom, MissingDecomSite, STATUS_DISPLAY_LABEL, TrackerHop,
  countDroppedOffWithoutCxComplete, uniqueDecomGcNames,
} from '../lib/decom'

interface ReportSnapshot {
  filename: string
  uploaded_at: string
  row_count: number
}

const SPO_COL_IDX = [7, 8, 33, 40, 41, 43, 47, 48, 49, 50, 51]
const SPO_HEADERS = ['Customer Site ID', 'Name', 'SOG Name', 'SPO Number', 'SPO Creation Date', 'SPO Vendor', 'SPO Value', 'IA Date', 'IA User', 'GR Date', 'GR Number']

const CR_COL_IDX = [0, 1, 2, 4, 5, 6, 7, 8, 9, 10, 11, 16, 18, 19, 25, 26, 27, 28, 29, 30, 31]
const CR_HEADERS = ['Requestor', 'Supplier Name', 'Path ID', 'Site Name', 'Site #', 'Network Site Name', 'Risk Budget', 'Materials or Labor', 'Reason for CR Details', 'Viaero Operation CR Filed', 'CR Type', 'Sellable to Who', 'PM Status', 'PM Status Owner', 'SPO Cost', 'GC Quote Shared', 'CQT Package', 'SPO #', 'SPO Issued Date', 'SPO IA/GR', 'CQT ID']

const NAVY = '124191'
const TEAL = '00A0B0'
const WHITE = 'FFFFFF'
const LIGHT_GRAY = 'F4F6FA'
const RED = 'C0392B'
const AMBER = 'E67E22'
const GREEN = '27AE60'
const DARK_GRAY = '2C3E50'
const MID_GRAY = '7F8C8D'

// Store only the columns Ciege actually reads (SPO_COL_IDX / CR_COL_IDX above —
// confirmed to be the full set referenced anywhere in the app or backend). Raw
// tracker exports run 50-90+ columns wide, and storing every column was pushing
// the Supabase upsert payload large enough to time out. Rows are truncated to a
// sparse array (nulled-out gaps, original index positions preserved, everything
// past the last needed column dropped entirely) rather than compacted or turned
// into an object — every existing row[N] lookup throughout the app and the
// Python backend (which indexes with a plain int, not a string key) keeps
// working unchanged, since this stays a real array both in JS and after
// json.loads() in Python.
function stripRow(row: unknown[], keepCols: number[]): unknown[] {
  const maxIdx = Math.max(...keepCols)
  const stripped: unknown[] = new Array(maxIdx + 1).fill(null)
  keepCols.forEach(i => { stripped[i] = row[i] ?? null })
  return stripped
}

function fmtDate(val: unknown): string {
  if (!val) return ''
  if (val instanceof Date) return val.toLocaleDateString('en-US')
  const d = new Date(String(val))
  return isNaN(d.getTime()) ? String(val) : d.toLocaleDateString('en-US')
}

interface DecomFunnel {
  totalTracked: number
  droppedOff: number
  podPathwave: number
  podQuickBase: number
  gap1: number
  gap2: number
  gap3: number
}

// Shared by the Reports page funnel section and Tab 1 of the Decom Dashboard Excel,
// so the two views can never drift out of sync with each other.
// `extraDroppedOff` folds in sites that have a real Drop Off date but haven't
// had CX Complete logged yet — parseDecomRows excludes those rows entirely
// (they aren't decom-eligible), so their drop-off would otherwise be invisible
// even though it genuinely happened. Only Dropped Off counts them — Total
// Tracked stays scoped to CX-Complete-confirmed rows, so gap1 can legitimately
// go negative (more dropped off than the tracked denominator) when it does.
function computeDecomFunnel(decomRows: DecomRow[], extraDroppedOff: number = 0): DecomFunnel {
  const totalTracked = decomRows.length
  const droppedOff = decomRows.filter(r => !!r.dropOffDate).length + extraDroppedOff
  const podPathwave = decomRows.filter(r => r.podPathwave).length
  const podQuickBase = decomRows.filter(r => r.podQuickBase).length
  return {
    totalTracked, droppedOff, podPathwave, podQuickBase,
    gap1: totalTracked - droppedOff,
    gap2: droppedOff - podPathwave,
    gap3: podPathwave - podQuickBase,
  }
}

// "Gap: -{N}" only reads correctly when gap is positive (sites stuck). Once
// Dropped Off can exceed Total Tracked, gap1 can go negative or zero — this
// picks the correct sign instead of rendering a literal double negative
// ("Gap: --2").
function fmtGapLabel(gap: number): string {
  if (gap > 0) return `Gap: -${gap}`
  if (gap < 0) return `Gap: +${Math.abs(gap)}`
  return 'Gap: 0'
}

function downloadGCReport(rows: unknown[][], colIdx: number[], headers: string[], vendorColInMaster: number, matchList: string[], filename: string) {
  const gcRows = rows.filter(row => matches(row[vendorColInMaster], matchList))

  const wb = XLSX.utils.book_new()
  const sheetData = [
    headers,
    ...gcRows.map(row => colIdx.map(i => {
      const val = row[i]
      if (val instanceof Date) return val.toLocaleDateString('en-US')
      return val ?? ''
    }))
  ]

  const ws = XLSX.utils.aoa_to_sheet(sheetData)

  // Style header row
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1')
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: 0, c })]
    if (cell) {
      cell.s = {
        font: { bold: true, color: { rgb: 'FFFFFF' } },
        fill: { fgColor: { rgb: NAVY } },
        alignment: { horizontal: 'center', wrapText: true }
      }
    }
  }

  // Auto column widths
  ws['!cols'] = headers.map((h, i) => {
    const maxLen = Math.max(h.length, ...sheetData.slice(1).map(r => String(r[i] || '').length))
    return { wch: Math.min(Math.max(maxLen + 2, 10), 40) }
  })

  XLSX.utils.book_append_sheet(wb, ws, headers[0].includes('SPO') || filename.includes('SPO') ? 'SPO Report' : 'CR Tracker')
  XLSX.writeFile(wb, filename)
}

const DECOM_HEADERS = ['HOP', 'Path ID', 'Site Name', 'Site Number', 'CM', 'CX Start', 'CX Complete', 'Aging (days)', 'DECOM Drop Off Date', 'DECOM Comments', 'POD Pathwave', 'POD QuickBase']

function decomRowToSheetRow(r: DecomRow): (string | number)[] {
  return [
    r.hop, r.pathId, r.siteName, r.siteNumber, r.cm,
    fmtDate(r.cxStart), fmtDate(r.cxComplete),
    r.aging ?? '',
    fmtDate(r.dropOffDate), r.comment,
    r.podPathwave ? 'Yes' : 'No', r.podQuickBase ? 'Yes' : 'No',
  ]
}

function styleDecomSheet(ws: XLSX.WorkSheet) {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1')
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: 0, c })]
    if (cell) {
      cell.s = {
        font: { bold: true, color: { rgb: 'FFFFFF' } },
        fill: { fgColor: { rgb: NAVY } },
        alignment: { horizontal: 'center', wrapText: true }
      }
    }
  }
  ws['!cols'] = DECOM_HEADERS.map(h => ({ wch: Math.max(h.length + 2, 12) }))
}

function downloadGcDecomReport(gcRows: DecomRow[], filename: string) {
  const outstandingPending = gcRows
    .filter(r => r.status === 'outstanding' || r.status === 'pending')
    .sort((a, b) => (b.aging ?? -1) - (a.aging ?? -1))
  const podGap = gcRows.filter(r => r.status === 'pod_gap')

  const wb = XLSX.utils.book_new()

  const ws1 = XLSX.utils.aoa_to_sheet([DECOM_HEADERS, ...outstandingPending.map(decomRowToSheetRow)])
  styleDecomSheet(ws1)
  XLSX.utils.book_append_sheet(wb, ws1, 'Pending Decom Drop Off')

  const ws2 = XLSX.utils.aoa_to_sheet([DECOM_HEADERS, ...podGap.map(decomRowToSheetRow)])
  styleDecomSheet(ws2)
  XLSX.utils.book_append_sheet(wb, ws2, 'Pending POD in Pathwave')

  XLSX.writeFile(wb, filename)
}

// Decom rows don't carry Nokia PM (the decom file has no such column) — cross-reference
// against the tracker to attribute one, Path ID first then HOP, same precedence findMissingDecom uses.
function nokiaPmLookup(trackerHops: TrackerHop[]) {
  const byPathId = new Map<string, string>()
  const byHop = new Map<string, string>()
  trackerHops.forEach(t => {
    if (t.pathId) byPathId.set(t.pathId, t.nokiaPm)
    if (t.hop) byHop.set(t.hop, t.nokiaPm)
  })
  return { byPathId, byHop }
}

function nokiaPmFor(r: DecomRow, lookup: ReturnType<typeof nokiaPmLookup>): string {
  if (r.pathId && lookup.byPathId.has(r.pathId)) return lookup.byPathId.get(r.pathId) || ''
  if (r.hop && lookup.byHop.has(r.hop)) return lookup.byHop.get(r.hop) || ''
  return ''
}

function downloadDecomDashboard(decomRows: DecomRow[], missingSites: MissingDecomSite[], trackerHops: TrackerHop[], extraDroppedOff: number) {
  // Every unique GC found in the decom file itself — not the GC_CONFIG roster —
  // so the breakdown always sums to the funnel's Total Decom Sites Tracked,
  // even for GCs not in the config list.
  const gcNames = uniqueDecomGcNames(decomRows)
  const summary = summarizeDecomByGc(decomRows, gcNames, missingSites)
  const lookup = nokiaPmLookup(trackerHops)
  const todayStr = new Date().toLocaleDateString('en-US')

  const wb = XLSX.utils.book_new()

  // ---- Tab 1: Decom Summary (funnel + GC breakdown) ----
  const sumHeaders = ['GC', 'Total Sites', 'Complete', 'Pending Drop Off', 'Pending POD Pathwave', 'Pending POD QuickBase', 'Missing', 'Avg Aging']
  const cols = sumHeaders.length
  const totals = summary.reduce((t, s) => ({
    total: t.total + s.total,
    complete: t.complete + s.complete,
    pendingDropOff: t.pendingDropOff + (s.outstanding + s.pending),
    pendingPathwave: t.pendingPathwave + s.pendingPathwave,
    pendingQuickBase: t.pendingQuickBase + s.pendingQuickBase,
    missing: t.missing + s.missing,
  }), { total: 0, complete: 0, pendingDropOff: 0, pendingPathwave: 0, pendingQuickBase: 0, missing: 0 })
  const allAging = decomRows.filter(r => r.status === 'outstanding' || r.status === 'pending').map(r => r.aging).filter((a): a is number => a !== null)
  const totalAvgAging = allAging.length > 0 ? Math.round(allAging.reduce((s, a) => s + a, 0) / allAging.length) : null

  // Same funnel math as the Reports page section — kept in one shared function
  // so the on-screen view and this export can never drift out of sync.
  const funnel = computeDecomFunnel(decomRows, extraDroppedOff)
  const boxLabels = ['Total Decom Sites Tracked', 'Total Sites Dropped Off', 'Total Sites POD in Pathwave', 'Total POD in QuickBase']
  const boxCounts = [funnel.totalTracked, funnel.droppedOff, funnel.podPathwave, funnel.podQuickBase]
  const boxGaps: (number | null)[] = [null, funnel.gap1, funnel.gap2, funnel.gap3]
  const boxColStart = [0, 2, 4, 6] // 4 boxes across 8 columns, 2 cols each

  const boxRow = new Array(cols).fill('')
  const countRow = new Array(cols).fill('')
  const gapRow = new Array(cols).fill('')
  boxColStart.forEach((c, i) => {
    boxRow[c] = boxLabels[i]
    countRow[c] = boxCounts[i]
    if (boxGaps[i] !== null) gapRow[c] = fmtGapLabel(boxGaps[i] as number)
  })

  const missingLineRowIdx = 7
  const headerRowIdx = 9
  const sumAoA: (string | number)[][] = [
    ['Viaero MW Program — Decom Status Dashboard'],
    [`As of ${todayStr}`],
    [],
    boxRow,
    countRow,
    gapRow,
    [],
    [`Missing From Tracker: ${missingSites.length} sites with no decom entry`],
    [],
    sumHeaders,
    ...summary.map(s => [s.gc, s.total, s.complete, s.outstanding + s.pending, s.pendingPathwave, s.pendingQuickBase, s.missing, s.avgAging ?? '']),
    ['Total', totals.total, totals.complete, totals.pendingDropOff, totals.pendingPathwave, totals.pendingQuickBase, totals.missing, totalAvgAging ?? ''],
  ]

  const ws1 = XLSX.utils.aoa_to_sheet(sumAoA)
  ws1['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: cols - 1 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: cols - 1 } },
    { s: { r: missingLineRowIdx, c: 0 }, e: { r: missingLineRowIdx, c: cols - 1 } },
    ...boxColStart.flatMap(c => [
      { s: { r: 3, c }, e: { r: 3, c: c + 1 } },
      { s: { r: 4, c }, e: { r: 4, c: c + 1 } },
      { s: { r: 5, c }, e: { r: 5, c: c + 1 } },
    ]),
  ]
  const titleCell = ws1[XLSX.utils.encode_cell({ r: 0, c: 0 })]
  if (titleCell) titleCell.s = { font: { bold: true, color: { rgb: NAVY }, sz: 16 } }
  const dateCell = ws1[XLSX.utils.encode_cell({ r: 1, c: 0 })]
  if (dateCell) dateCell.s = { font: { italic: true, color: { rgb: '555555' } } }

  boxColStart.forEach((c, i) => {
    const boxColor = i === 0 ? NAVY : TEAL
    const labelCell = ws1[XLSX.utils.encode_cell({ r: 3, c })]
    if (labelCell) labelCell.s = { font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: boxColor } }, alignment: { horizontal: 'center', wrapText: true } }
    const countCell = ws1[XLSX.utils.encode_cell({ r: 4, c })]
    if (countCell) countCell.s = { font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 18 }, fill: { fgColor: { rgb: boxColor } }, alignment: { horizontal: 'center' } }
    const gap = boxGaps[i]
    if (gap !== null) {
      const gapCell = ws1[XLSX.utils.encode_cell({ r: 5, c })]
      if (gapCell) gapCell.s = { font: { bold: true, color: { rgb: gap > 0 ? 'C00000' : '006100' } }, alignment: { horizontal: 'center' } }
    }
  })

  const missingLineCell = ws1[XLSX.utils.encode_cell({ r: missingLineRowIdx, c: 0 })]
  if (missingLineCell) missingLineCell.s = { font: { bold: true, color: { rgb: missingSites.length > 0 ? 'C00000' : '006100' } } }

  for (let c = 0; c < cols; c++) {
    const cell = ws1[XLSX.utils.encode_cell({ r: headerRowIdx, c })]
    if (cell) cell.s = { font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: NAVY } }, alignment: { horizontal: 'center' } }
  }

  summary.forEach((s, i) => {
    const r = headerRowIdx + 1 + i
    const cellAt = (c: number) => ws1[XLSX.utils.encode_cell({ r, c })]
    const completeCell = cellAt(2)
    if (completeCell && s.complete > 0) completeCell.s = { fill: { fgColor: { rgb: 'C6EFCE' } }, font: { color: { rgb: '006100' } } }
    const dropOffCell = cellAt(3)
    if (dropOffCell && (s.outstanding + s.pending) > 0) dropOffCell.s = { fill: { fgColor: { rgb: 'FFEB9C' } }, font: { color: { rgb: '9C6500' } } }
    const pathwaveCell = cellAt(4)
    if (pathwaveCell && s.pendingPathwave > 0) pathwaveCell.s = { fill: { fgColor: { rgb: 'FFEB9C' } }, font: { color: { rgb: '9C6500' } } }
    const qbCell = cellAt(5)
    if (qbCell && s.pendingQuickBase > 0) qbCell.s = { fill: { fgColor: { rgb: 'FFEB9C' } }, font: { color: { rgb: '9C6500' } } }
    const missingCell = cellAt(6)
    if (missingCell && s.missing > 0) missingCell.s = { fill: { fgColor: { rgb: 'FFC7CE' } }, font: { color: { rgb: '9C0006' } } }
  })

  const totalsRowIdx = headerRowIdx + 1 + summary.length
  for (let c = 0; c < cols; c++) {
    const cell = ws1[XLSX.utils.encode_cell({ r: totalsRowIdx, c })]
    if (cell) cell.s = { font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: NAVY } } }
  }

  ws1['!rows'] = [{}, {}, {}, { hpt: 30 }, { hpt: 24 }, {}]
  ws1['!cols'] = sumHeaders.map(h => ({ wch: Math.max(h.length + 4, 14) }))
  XLSX.utils.book_append_sheet(wb, ws1, 'Decom Summary')

  // ---- Tab 2: Site Detail ----
  const detailHeaders = ['HOP', 'Path ID', 'Site Name', 'GC', 'CM', 'Nokia PM', 'CX Complete', 'DECOM Drop Off Date', 'DECOM Comments', 'POD Pathwave', 'POD QuickBase', 'Status', 'Aging (days)']
  const detailRows = decomRows.map(r => ({
    row: [
      r.hop, r.pathId, r.siteName, r.gc, r.cm, nokiaPmFor(r, lookup),
      fmtDate(r.cxComplete), fmtDate(r.dropOffDate), r.comment,
      r.podPathwave ? 'Yes' : 'No', r.podQuickBase ? 'Yes' : 'No',
      STATUS_DISPLAY_LABEL[r.status], r.aging ?? '',
    ] as (string | number)[],
    status: r.status,
    aging: r.aging,
  }))

  const ws2 = XLSX.utils.aoa_to_sheet([detailHeaders, ...detailRows.map(d => d.row)])
  for (let c = 0; c < detailHeaders.length; c++) {
    const cell = ws2[XLSX.utils.encode_cell({ r: 0, c })]
    if (cell) cell.s = { font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: TEAL } }, alignment: { horizontal: 'center', wrapText: true } }
  }
  detailRows.forEach((d, i) => {
    const r = i + 1
    const isAgedPendingDropOff = STATUS_DISPLAY_LABEL[d.status] === 'Pending Decom Drop Off' && d.aging !== null && d.aging >= 7
    const rowFill = isAgedPendingDropOff ? 'FFC7CE' : (i % 2 === 1 ? 'F2F2F2' : 'FFFFFF')
    for (let c = 0; c < detailHeaders.length; c++) {
      const cell = ws2[XLSX.utils.encode_cell({ r, c })]
      if (cell) cell.s = { fill: { fgColor: { rgb: rowFill } } }
    }
  })
  ws2['!cols'] = detailHeaders.map(h => ({ wch: Math.max(h.length + 2, 12) }))
  XLSX.utils.book_append_sheet(wb, ws2, 'Site Detail')

  // ---- Tab 3: Missing From Tracker ----
  const missingHeaders = ['HOP', 'Path ID', 'GC', 'CM', 'Nokia PM', 'CX Complete', 'Days Since Complete']
  const sortedMissing = [...missingSites].sort((a, b) => b.daysElapsed - a.daysElapsed)

  const missingAoA: (string | number)[][] = [
    ['Sites With No Decom Tracking Entry'],
    [],
    missingHeaders,
  ]
  if (sortedMissing.length === 0) {
    missingAoA.push(['All completed sites are tracked in decom file ✅'])
  } else {
    sortedMissing.forEach(m => missingAoA.push([m.hop, m.pathId, m.gc, m.cm, m.nokiaPm, fmtDate(m.ms16a), m.daysElapsed]))
  }

  const ws3 = XLSX.utils.aoa_to_sheet(missingAoA)
  ws3['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: missingHeaders.length - 1 } }]
  const missingTitleCell = ws3[XLSX.utils.encode_cell({ r: 0, c: 0 })]
  if (missingTitleCell) missingTitleCell.s = { font: { bold: true, color: { rgb: NAVY }, sz: 14 } }

  for (let c = 0; c < missingHeaders.length; c++) {
    const cell = ws3[XLSX.utils.encode_cell({ r: 2, c })]
    if (cell) cell.s = { font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: 'C00000' } }, alignment: { horizontal: 'center' } }
  }
  if (sortedMissing.length === 0) {
    ws3['!merges'].push({ s: { r: 3, c: 0 }, e: { r: 3, c: missingHeaders.length - 1 } })
    const emptyCell = ws3[XLSX.utils.encode_cell({ r: 3, c: 0 })]
    if (emptyCell) emptyCell.s = { font: { italic: true, color: { rgb: '006100' } }, alignment: { horizontal: 'center' } }
  }
  ws3['!cols'] = missingHeaders.map(h => ({ wch: Math.max(h.length + 2, 14) }))
  XLSX.utils.book_append_sheet(wb, ws3, 'Missing From Tracker')

  XLSX.writeFile(wb, `Viaero_Decom_Dashboard_${new Date().toISOString().slice(0, 10)}.xlsx`)
}

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

async function downloadDecomSlides(decomRows: DecomRow[], missingDecom: MissingDecomSite[], extraDroppedOff: number) {
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

  const today = new Date()
  const mm = String(today.getMonth() + 1).padStart(2, '0')
  const dd = String(today.getDate()).padStart(2, '0')
  const yyyy = today.getFullYear()
  await pres.writeFile({ fileName: `Decom Tracker Slides - ${mm}-${dd}-${yyyy}.pptx` })
}

interface UploadBoxProps {
  type: 'spo' | 'cr' | 'decom'
  info: ReportSnapshot | null
  label: string
  uploading: string | null
  onUpload: (file: File, type: 'spo' | 'cr' | 'decom') => void
}

function UploadBox({ type, info, label, uploading, onUpload }: UploadBoxProps) {
  return (
    <div
      className="border-2 border-dashed border-gray-600 rounded-xl p-5 cursor-pointer hover:border-blue-500 transition-colors"
      onDragOver={(e) => e.preventDefault()}
      onDrop={(e) => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) onUpload(f, type) }}
      onClick={() => document.getElementById(`${type}-upload`)?.click()}
    >
      <input id={`${type}-upload`} type="file" accept=".xlsx" className="hidden"
        onChange={(e) => { const f = e.target.files?.[0]; if (f) onUpload(f, type) }} />
      {uploading === type
        ? <p className="text-blue-400 text-sm text-center">⏳ Processing...</p>
        : info
        ? <div>
            <p className="text-green-400 text-sm font-semibold">✅ {label} loaded</p>
            <p className="text-gray-500 text-xs mt-1">{info.filename} · {info.row_count} rows · {new Date(info.uploaded_at).toLocaleDateString('en-US')} at {new Date(info.uploaded_at).toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })}</p>
          </div>
        : <p className="text-gray-400 text-sm text-center">📂 Upload {label}</p>
      }
    </div>
  )
}

export default function ReportsPage() {
  const router = useRouter()
  const [spoRows, setSpoRows] = useState<unknown[][]>([])
  const [crRows, setCrRows] = useState<unknown[][]>([])
  const [decomRawRows, setDecomRawRows] = useState<unknown[][]>([])
  const [trackerRawRows, setTrackerRawRows] = useState<unknown[][]>([])
  const [spoInfo, setSpoInfo] = useState<ReportSnapshot | null>(null)
  const [crInfo, setCrInfo] = useState<ReportSnapshot | null>(null)
  const [decomInfo, setDecomInfo] = useState<ReportSnapshot | null>(null)
  const [uploading, setUploading] = useState<string | null>(null)
  const [activeTab, setActiveTab] = useState<'spo-cr-decom' | 'decom-tracker'>('spo-cr-decom')
  const [generatingSlides, setGeneratingSlides] = useState(false)
  const today = new Date().toLocaleDateString('en-US').replace(/\//g, '-')

  const decomRows = parseDecomRows(decomRawRows)
  console.log('[decom-render] decomRawRows:', decomRawRows.length, 'rows (incl. header) → decomRows parsed:', decomRows.length)
  const trackerHops = parseTrackerHopsForDecom(trackerRawRows)
  const missingDecomSites = findMissingDecom(decomRows, trackerHops)
  // Sites with a real Drop Off date but no CX Complete logged yet — excluded
  // from decomRows entirely (not decom-eligible), so their drop-off has to be
  // folded in separately or it's invisible everywhere, including here.
  const extraDroppedOff = countDroppedOffWithoutCxComplete(decomRawRows)
  const decomFunnel = computeDecomFunnel(decomRows, extraDroppedOff)
  // Every unique GC found in the decom file itself — not the GC_CONFIG roster —
  // so the breakdown always sums to Total Decom Sites Tracked, including GCs
  // not in the config list.
  const decomGcNames = uniqueDecomGcNames(decomRows)

  useEffect(() => {
    const load = async () => {
      const { data: spoSnap } = await supabase.from('report_snapshots').select('*').eq('id', 'spo').single()
      if (spoSnap) { setSpoRows(JSON.parse(spoSnap.data)); setSpoInfo({ filename: spoSnap.filename, uploaded_at: spoSnap.uploaded_at, row_count: JSON.parse(spoSnap.data).length }) }

      const crReport = await loadChunkedReport('cr')
      if (crReport) { setCrRows(crReport.rows); setCrInfo({ filename: crReport.filename, uploaded_at: crReport.uploaded_at, row_count: crReport.rows.length }) }

      const decomReport = await loadChunkedReport('decom')
      console.log('[decom-load] loadChunkedReport("decom") returned', decomReport ? decomReport.rows.length : 0, 'rows')
      if (decomReport) { setDecomRawRows(decomReport.rows); setDecomInfo({ filename: decomReport.filename, uploaded_at: decomReport.uploaded_at, row_count: decomReport.rows.length }) }

      const trackerSnap = await loadTrackerSnapshot()
      if (trackerSnap) setTrackerRawRows(trackerSnap.data)
    }
    load()
  }, [])

  const handleUpload = useCallback(async (file: File, type: 'spo' | 'cr' | 'decom') => {
    setUploading(type)
    const reader = new FileReader()
    reader.onload = async (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer)
        const wb = XLSX.read(data, { type: 'array', cellDates: true })

        let ws, headerRowIndex
        if (type === 'spo') {
          ws = wb.Sheets['NDPd_NAM_003_Viaero_SPO_Report'] || wb.Sheets[wb.SheetNames[0]]
          headerRowIndex = 0
        } else if (type === 'cr') {
          ws = wb.Sheets['CR Tracker'] || wb.Sheets[wb.SheetNames[0]]
          headerRowIndex = 1
        } else {
          const DECOM_SHEET_NAME = 'R&R - Link Drop Off Status'
          const matchedSheetName = wb.SheetNames.find(n => n.trim() === DECOM_SHEET_NAME)
          ws = matchedSheetName ? wb.Sheets[matchedSheetName] : undefined
          if (!ws) {
            console.error(`[report-upload] Decom upload failed: sheet "${DECOM_SHEET_NAME}" not found. Available sheets:`, wb.SheetNames)
            alert(`Could not find the "${DECOM_SHEET_NAME}" tab in this file — check the tab name and try again.`)
            setUploading(null)
            return
          }
          headerRowIndex = 0
        }

        const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as unknown[][]

        let allRows = rawRows
        if (type === 'cr' || type === 'decom') {
          let lastDataRow = rawRows.length - 1
          while (lastDataRow >= 0 && !(rawRows[lastDataRow] || []).some(c => c !== null && c !== undefined && String(c).trim() !== '')) {
            lastDataRow--
          }
          const rows = rawRows.slice(0, lastDataRow + 1)
          console.log(`[report-upload] ${type.toUpperCase()} raw rows:`, rawRows.length, '→ trimmed:', rows.length)
          allRows = rows
        }

        const dataRows = allRows.slice(headerRowIndex + 1).filter(row => row.some(v => v !== null))

        if (type === 'spo' || type === 'cr') {
          const keepCols = type === 'spo' ? SPO_COL_IDX : CR_COL_IDX
          const totalCols = dataRows.reduce((max, row) => Math.max(max, row.length), 0)
          const strippedRows = dataRows.map(row => stripRow(row, keepCols))
          console.log(`[report-upload] ${type.toUpperCase()} columns: kept ${keepCols.length} of ${totalCols} (dropped ${totalCols - keepCols.length})`)

          if (type === 'spo') {
            console.log('[report-upload] upserting id="spo"', 'filename=', file.name)
            const { error } = await supabase.from('report_snapshots').upsert({
              id: 'spo',
              filename: file.name,
              uploaded_at: new Date().toISOString(),
              data: JSON.stringify(strippedRows)
            })
            if (error) {
              console.error('[report-upload] SPO upsert failed:', error)
              alert('SPO upload failed to save — check console for details')
              setUploading(null)
              return
            }
            setSpoRows(strippedRows)
            setSpoInfo({ filename: file.name, uploaded_at: new Date().toISOString(), row_count: strippedRows.length })
          } else {
            console.log('[report-upload] chunking id="cr"', 'filename=', file.name, 'rows=', strippedRows.length)
            const { error } = await saveChunkedReport('cr', file.name, strippedRows)
            if (error) {
              console.error('[report-upload] CR chunked save failed:', error)
              alert('CR upload failed to save — check console for details')
              setUploading(null)
              return
            }
            setCrRows(strippedRows)
            setCrInfo({ filename: file.name, uploaded_at: new Date().toISOString(), row_count: strippedRows.length })
          }
        } else {
          // Decom isn't column-stripped — the parser resolves columns by
          // header name at load time, so the header row travels with the
          // stored data instead of being dropped like SPO/CR.
          const decomStoreRows = [allRows[headerRowIndex], ...dataRows]
          console.log('[report-upload] chunking id="decom"', 'filename=', file.name, 'rows=', decomStoreRows.length)
          const { error } = await saveChunkedReport('decom', file.name, decomStoreRows)
          if (error) {
            console.error('[report-upload] Decom chunked save failed:', error)
            alert('Decom upload failed to save — check console for details')
            setUploading(null)
            return
          }
          setDecomRawRows(decomStoreRows)
          setDecomInfo({ filename: file.name, uploaded_at: new Date().toISOString(), row_count: dataRows.length })
        }
      } catch (err) {
        console.error('Upload error:', err)
        alert('Upload failed — check file format')
      }
      setUploading(null)
    }
    reader.readAsArrayBuffer(file)
  }, [])

  return (
    <div className="min-h-screen bg-gray-950 text-white p-6">
      <div className="max-w-6xl mx-auto">

        <BackToDashboard />

        <div className="mb-6">
          <h1 className="text-3xl font-bold">Reports</h1>
          <p className="text-gray-400 mt-1">Upload SPO and CR master reports — download GC-specific filtered Excel files on demand</p>
        </div>

        {/* Tab Bar */}
        <div className="flex gap-2 mb-6">
          <button
            onClick={() => setActiveTab('spo-cr-decom')}
            className={`px-5 py-2 rounded-lg text-sm font-semibold transition-all ${activeTab === 'spo-cr-decom' ? 'bg-blue-600 text-white shadow-lg' : 'bg-gray-800 text-gray-300 hover:bg-gray-700'}`}>
            SPO / CR / Decom Reports
          </button>
          <button
            onClick={() => setActiveTab('decom-tracker')}
            className={`px-5 py-2 rounded-lg text-sm font-semibold transition-all ${activeTab === 'decom-tracker' ? 'bg-blue-600 text-white shadow-lg' : 'bg-gray-800 text-gray-300 hover:bg-gray-700'}`}>
            ♻️ Decom Tracker
          </button>
          <button
            onClick={() => router.push('/gr-tracker')}
            className="px-5 py-2 rounded-lg text-sm font-semibold bg-gray-800 text-gray-300 hover:bg-gray-700">
            💰 GR Tracker
          </button>
        </div>

        {activeTab === 'spo-cr-decom' && (
          <>
            {/* Upload Section */}
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-8">
              <div>
                <p className="text-sm font-semibold text-gray-400 mb-2">SPO MASTER REPORT</p>
                <UploadBox type="spo" info={spoInfo} label="SPO Master Report" uploading={uploading} onUpload={handleUpload} />
              </div>
              <div>
                <p className="text-sm font-semibold text-gray-400 mb-2">CR CLAIMS TRACKER</p>
                <UploadBox type="cr" info={crInfo} label="CR Claims Tracker" uploading={uploading} onUpload={handleUpload} />
              </div>
              <div>
                <p className="text-sm font-semibold text-gray-400 mb-2">DECOM TRACKER</p>
                <UploadBox type="decom" info={decomInfo} label="Decom Tracker" uploading={uploading} onUpload={handleUpload} />
              </div>
            </div>
          </>
        )}

        {activeTab === 'decom-tracker' && (
          <>
        {/* Decom Funnel */}
        {decomRows.length > 0 && (
          <div className="mb-8">
            <h2 className="text-lg font-bold mb-4">Decom Funnel</h2>
            <div className="flex items-stretch gap-0 overflow-x-auto pb-1">
              {[
                { label: 'Total Decom Sites Tracked', count: decomFunnel.totalTracked, gap: decomFunnel.gap1 },
                { label: 'Dropped Off', count: decomFunnel.droppedOff, gap: decomFunnel.gap2 },
                { label: 'POD Pathwave', count: decomFunnel.podPathwave, gap: decomFunnel.gap3 },
                { label: 'POD QuickBase', count: decomFunnel.podQuickBase, gap: null as number | null },
              ].map((box, i, arr) => (
                <div key={box.label} className="flex items-stretch">
                  <div className="min-w-[170px] bg-gray-900 rounded-xl border border-gray-700 overflow-hidden">
                    <div className="px-3 py-2 text-xs font-bold text-white text-center" style={{ backgroundColor: `#${NAVY}` }}>{box.label}</div>
                    <div className="px-3 py-3 text-center text-3xl font-bold" style={{ color: `#${TEAL}` }}>{box.count}</div>
                    {box.gap !== null && (
                      <div className={`px-3 pb-2 text-center text-xs font-bold ${box.gap > 0 ? 'text-red-400' : 'text-green-400'}`}>
                        {fmtGapLabel(box.gap)}
                      </div>
                    )}
                  </div>
                  {i < arr.length - 1 && <span className="text-gray-500 text-2xl px-3 self-center">→</span>}
                </div>
              ))}
            </div>
            <div className="mt-4 space-y-1">
              <p className="text-sm">
                <span className="text-red-400 font-semibold">Missing From Tracker: {missingDecomSites.length} sites</span>
                <span className="text-gray-400"> — completed HOPs with no decom entry</span>
              </p>
              <p className="text-xs text-gray-500 italic">
                Partial Returns Note: Sites with partial drop-offs may not be included in drop count
              </p>
            </div>
          </div>
        )}

        {/* Decom Summary */}
        {decomRows.length > 0 && (
          <div className="mb-8">
            <h2 className="text-lg font-bold mb-4">Decom Summary by GC</h2>
            <div className="overflow-x-auto bg-gray-900 rounded-xl border border-gray-700">
              <table className="w-full text-xs">
                <thead>
                  <tr className="bg-gray-800 text-gray-400">
                    <th className="text-left p-2">GC</th>
                    <th className="text-left p-2">Total Sites</th>
                    <th className="text-left p-2">Complete</th>
                    <th className="text-left p-2">Pending Decom Drop Off</th>
                    <th className="text-left p-2">Pending POD Pathwave</th>
                    <th className="text-left p-2">Pending POD QuickBase</th>
                    <th className="text-left p-2">Missing</th>
                    <th className="text-left p-2">Avg Aging (days)</th>
                  </tr>
                </thead>
                <tbody>
                  {(() => {
                    const summary = summarizeDecomByGc(decomRows, decomGcNames, missingDecomSites).sort((a, b) => (b.outstanding + b.pending) - (a.outstanding + a.pending))
                    const totals = summary.reduce((t, s) => ({
                      total: t.total + s.total,
                      complete: t.complete + s.complete,
                      outstanding: t.outstanding + s.outstanding,
                      pending: t.pending + s.pending,
                      pendingPathwave: t.pendingPathwave + s.pendingPathwave,
                      pendingQuickBase: t.pendingQuickBase + s.pendingQuickBase,
                      missing: t.missing + s.missing,
                    }), { total: 0, complete: 0, outstanding: 0, pending: 0, pendingPathwave: 0, pendingQuickBase: 0, missing: 0 })
                    const allAging = decomRows.filter(r => r.status === 'outstanding' || r.status === 'pending').map(r => r.aging).filter((a): a is number => a !== null)
                    const totalAvgAging = allAging.length > 0 ? Math.round(allAging.reduce((s, a) => s + a, 0) / allAging.length) : null
                    return (
                      <>
                        {summary.map(s => (
                          <tr key={s.gc} className="border-t border-gray-800">
                            <td className="p-2 font-semibold text-white whitespace-nowrap">{s.gc}</td>
                            <td className="p-2 text-gray-300">{s.total}</td>
                            <td className="p-2 text-green-400">{s.complete}</td>
                            <td className="p-2 text-red-400 font-bold">{s.outstanding + s.pending}</td>
                            <td className="p-2 text-orange-400">{s.pendingPathwave}</td>
                            <td className="p-2 text-blue-400">{s.pendingQuickBase}</td>
                            <td className="p-2 text-red-400">{s.missing}</td>
                            <td className="p-2 text-gray-300">{s.avgAging ?? '—'}</td>
                          </tr>
                        ))}
                        <tr className="border-t border-gray-700 bg-gray-800 font-bold">
                          <td className="p-2 text-white">Total</td>
                          <td className="p-2 text-gray-200">{totals.total}</td>
                          <td className="p-2 text-green-400">{totals.complete}</td>
                          <td className="p-2 text-red-400">{totals.outstanding + totals.pending}</td>
                          <td className="p-2 text-orange-400">{totals.pendingPathwave}</td>
                          <td className="p-2 text-blue-400">{totals.pendingQuickBase}</td>
                          <td className="p-2 text-red-400">{totals.missing}</td>
                          <td className="p-2 text-gray-200">{totalAvgAging ?? '—'}</td>
                        </tr>
                      </>
                    )
                  })()}
                </tbody>
              </table>
            </div>

            {missingDecomSites.length > 0 && (
              <div className="mt-6">
                <h3 className="text-md font-bold mb-2 text-red-400">Missing From Tracker ({missingDecomSites.length})</h3>
                <div className="overflow-x-auto bg-gray-900 rounded-xl border border-red-900">
                  <table className="w-full text-xs">
                    <thead>
                      <tr className="bg-gray-800 text-gray-400">
                        <th className="text-left p-2">HOP</th>
                        <th className="text-left p-2">Path ID</th>
                        <th className="text-left p-2">GC</th>
                        <th className="text-left p-2">Nokia PM</th>
                        <th className="text-left p-2">CX Complete</th>
                        <th className="text-left p-2">Days Since Complete</th>
                      </tr>
                    </thead>
                    <tbody>
                      {[...missingDecomSites].sort((a, b) => b.daysElapsed - a.daysElapsed).map(m => (
                        <tr key={`${m.pathId || m.hop}-${m.siteName}`} className="border-t border-gray-800">
                          <td className="p-2 text-white whitespace-nowrap">{m.hop}</td>
                          <td className="p-2 text-gray-300">{m.pathId}</td>
                          <td className="p-2 text-gray-300">{m.gc}</td>
                          <td className="p-2 text-gray-300">{m.nokiaPm}</td>
                          <td className="p-2 text-gray-300">{fmtDate(m.ms16a)}</td>
                          <td className="p-2 text-red-400 font-bold">{m.daysElapsed}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            <div className="mt-6 flex gap-3">
              <button
                onClick={() => downloadDecomDashboard(decomRows, missingDecomSites, trackerHops, extraDroppedOff)}
                className="px-4 py-2 bg-blue-600 hover:bg-blue-500 rounded-lg text-sm font-semibold"
              >
                ⬇️ Download Decom Dashboard
              </button>
              <button
                onClick={async () => {
                  if (decomRows.length === 0) { alert('Upload decom tracker first'); return }
                  setGeneratingSlides(true)
                  try {
                    await downloadDecomSlides(decomRows, missingDecomSites, extraDroppedOff)
                  } finally {
                    setGeneratingSlides(false)
                  }
                }}
                disabled={generatingSlides}
                className="px-4 py-2 bg-purple-600 hover:bg-purple-500 disabled:opacity-50 disabled:cursor-not-allowed rounded-lg text-sm font-semibold"
              >
                {generatingSlides ? '⏳ Generating...' : '⬇️ Download Decom Slides'}
              </button>
            </div>
          </div>
        )}
          </>
        )}

        {activeTab === 'spo-cr-decom' && (
          <>
        {/* GC Download Grid */}
        {(spoRows.length > 0 || crRows.length > 0 || decomRows.length > 0) && (
          <div>
            <h2 className="text-lg font-bold mb-4">GC-Specific Reports — Download on Demand</h2>
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
              {GC_CONFIG.map(cfg => {
                const spoCount = spoRows.filter(row => matches(row[SPO_VENDOR_COL_IN_MASTER], cfg.spo_match)).length
                const crCount = crRows.filter(row => matches(row[CR_SUPPLIER_COL_IN_MASTER], cfg.cr_match)).length
                const gcDecomRows = decomRowsForGc(decomRows, cfg.gc)

                return (
                  <div key={cfg.gc} className="bg-gray-900 rounded-xl border border-gray-700 p-4">
                    <h3 className="font-bold text-white text-base mb-1">{cfg.gc}</h3>
                    <div className="flex gap-2 text-xs text-gray-500 mb-3">
                      <span>{spoCount} SPO rows</span>
                      <span>·</span>
                      <span>{crCount} CR rows</span>
                      <span>·</span>
                      <span>{gcDecomRows.length} Decom rows</span>
                    </div>
                    <div className="flex flex-col gap-2">
                      <button
                        onClick={() => downloadGCReport(spoRows, SPO_COL_IDX, SPO_HEADERS, SPO_VENDOR_COL_IN_MASTER, cfg.spo_match, `${cfg.spo_label}_-_SPO_Report_-_${today}.xlsx`)}
                        disabled={spoRows.length === 0}
                        className="bg-blue-700 hover:bg-blue-600 disabled:opacity-40 disabled:cursor-not-allowed text-white text-xs px-3 py-2 rounded font-semibold">
                        📥 Download SPO Report
                      </button>
                      <button
                        onClick={() => downloadGCReport(crRows, CR_COL_IDX, CR_HEADERS, CR_SUPPLIER_COL_IN_MASTER, cfg.cr_match, `${cfg.cr_label}_-_CR_Tracker_-_${today}.xlsx`)}
                        disabled={crRows.length === 0}
                        className="bg-teal-700 hover:bg-teal-600 disabled:opacity-40 disabled:cursor-not-allowed text-white text-xs px-3 py-2 rounded font-semibold">
                        📥 Download CR Tracker
                      </button>
                      <button
                        onClick={() => downloadGcDecomReport(gcDecomRows, `${cfg.gc}_-_Decom_Report_-_${today}.xlsx`)}
                        disabled={gcDecomRows.length === 0}
                        className="bg-purple-700 hover:bg-purple-600 disabled:opacity-40 disabled:cursor-not-allowed text-white text-xs px-3 py-2 rounded font-semibold">
                        📥 Download Decom Report
                      </button>
                    </div>
                  </div>
                )
              })}
            </div>
          </div>
        )}

        {spoRows.length === 0 && crRows.length === 0 && decomRows.length === 0 && (
          <div className="bg-gray-900 rounded-xl border border-gray-700 p-12 text-center">
            <p className="text-gray-400 text-xl">📂 Upload your master reports above to get started</p>
          </div>
        )}
          </>
        )}

      </div>
    </div>
  )
}
