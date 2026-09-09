// SCOP Excel workbook builders — the Master Report (9 tabs) and the per-GC
// weekly report (Pathwave + OAD). Pure data-in / WorkBook-out, same shape as
// the Decom report builders in app/reports/page.tsx.
//
// NOTE: this project uses the community `xlsx` build, which ignores cell `.s`
// styles on write. Column widths (`!cols`) do work. Landscape / fit-to-width
// page setup (spec §5) isn't portable from the openpyxl original — SheetJS
// community doesn't emit it — so tabs print portrait; that's a known gap, not
// a regression.
//
// Excel formula-injection hazard (spec §5): never write a cell value that
// starts with `=`. The Filter Cheat Sheet prefixes descriptive text instead.

import * as XLSX from 'xlsx'
import {
  ScopRow, ScopGcReport, ScopCalcSettings,
  computeQuickBaseView, computePathwaveView, computeOverallView,
  buildItemizedMissingItems, pathwaveGcItems, QUICKBASE_ALL_ITEMS,
} from './scop'

function fmtDate(d: Date | null | undefined): string {
  if (!d) return ''
  return d instanceof Date && !isNaN(d.getTime()) ? d.toLocaleDateString('en-US') : ''
}

function sheetFromAoA(aoa: (string | number)[][], colWidths?: number[]): XLSX.WorkSheet {
  const ws = XLSX.utils.aoa_to_sheet(aoa)
  if (colWidths) {
    ws['!cols'] = colWidths.map(w => ({ wch: w }))
  } else if (aoa.length > 0) {
    ws['!cols'] = aoa[0].map((_, c) => {
      const maxLen = Math.max(...aoa.map(r => String(r[c] ?? '').length))
      return { wch: Math.min(Math.max(maxLen + 2, 10), 50) }
    })
  }
  return ws
}

// Simple-list tabs (Outstanding, Pathwave Complete, QuickBase Complete, Fully Complete)
const SIMPLE_HEADERS = ['HOP', 'Path ID', 'GC', 'CM', 'Near Site A', 'Far Site B']
function simpleListAoA(rows: ScopRow[]): (string | number)[][] {
  return [
    SIMPLE_HEADERS,
    ...rows.map(r => [r.hop, r.pathId, r.gc ?? '', r.cm, r.nearSiteA, r.farSiteB]),
  ]
}

// Itemized tabs (Pathwave Pending, QuickBase Pending)
const ITEMIZED_HEADERS = ['HOP', 'Path ID', 'GC', 'CM', 'Near Site A', 'Far Site B', 'Site', 'Days Since Complete', 'Missing Items']
function itemizedAoA(rows: ReturnType<typeof buildItemizedMissingItems>): (string | number)[][] {
  return [
    ITEMIZED_HEADERS,
    ...rows.map(r => [
      r.hop, r.pathId, r.gc ?? '', r.cm, r.nearSiteA, r.farSiteB,
      r.site, r.agingDays ?? '', r.missingItems,
    ]),
  ]
}

/**
 * The 9-tab Master SCOP Report. `dataset` is the full classified set (all
 * Nokia rows, not just GC-facing). Cross-foots to the spec §10 baseline.
 */
export function buildScopMasterWorkbook(dataset: ScopRow[], s: ScopCalcSettings): XLSX.WorkBook {
  const wb = XLSX.utils.book_new()
  const qb = computeQuickBaseView(dataset)
  const pw = computePathwaveView(dataset)
  const ov = computeOverallView(dataset)
  const todayStr = fmtDate(s.asOfDate)

  // ── Tab 1: Summary ────────────────────────────────────────────────────────
  const summaryAoA: (string | number)[][] = [
    [`SCOP Master Report — as of ${todayStr}`],
    [],
    ['Metric', 'Count'],
    ['Total Nokia HOPs', dataset.length],
    ['Construction Complete (tracked)', ov.trackedTotal],
    ['Pending HOP Completion', pw.pendingHopCompletion],
    [],
    ['PATHWAVE', ''],
    ['Complete & Approved', pw.complete],
    ['In Progress — GC Action', pw.inProgress],
    ['OAD — Tracked Separately', pw.oad],
    ['Pathwave Pending (In Progress + OAD, View 3)', ov.pathwavePending],
    [],
    ['QUICKBASE', ''],
    ['Complete', qb.complete],
    ['Pending', qb.pending],
    ['% Complete (of tracked)', `${qb.percentComplete}%`],
    [],
    ['OVERALL', ''],
    ['Fully Complete (both, within tracked)', ov.fullyComplete],
    ['Not Yet Fully Complete (within tracked)', ov.notFullyComplete],
    [],
    ['Cross-foot check', ''],
    ['Pathwave 4-way = Total?', pw.complete + pw.inProgress + pw.oad + pw.pendingHopCompletion === dataset.length ? 'OK' : 'MISMATCH'],
    ['QuickBase C+P = Tracked?', qb.complete + qb.pending === ov.trackedTotal ? 'OK' : 'MISMATCH'],
    ['Pathwave C+Pending = Tracked?', ov.pathwaveComplete + ov.pathwavePending === ov.trackedTotal ? 'OK' : 'MISMATCH'],
  ]
  XLSX.utils.book_append_sheet(wb, sheetFromAoA(summaryAoA, [42, 12]), 'Summary')

  // ── Tab 2: Outstanding List ──────────────────────────────────────────────
  const outstanding = dataset.filter(r => !r.isConstructionComplete)
  XLSX.utils.book_append_sheet(wb, sheetFromAoA(simpleListAoA(outstanding)), 'Outstanding List')

  // ── Tab 3: Pathwave Pending (itemized, excludes OAD) ──────────────────────
  const pathwaveInProgress = dataset.filter(r => r.pathwaveStatus === 'IN_PROGRESS')
  const pathwavePendingItemized = buildItemizedMissingItems(pathwaveInProgress, pathwaveGcItems(s))
  XLSX.utils.book_append_sheet(wb, sheetFromAoA(itemizedAoA(pathwavePendingItemized)), 'Pathwave Pending')

  // ── Tab 4: Pathwave Complete ─────────────────────────────────────────────
  XLSX.utils.book_append_sheet(
    wb, sheetFromAoA(simpleListAoA(dataset.filter(r => r.pathwaveStatus === 'COMPLETE'))), 'Pathwave Complete',
  )

  // ── Tab 4a: Pathwave OAD Sites ───────────────────────────────────────────
  const oadAoA: (string | number)[][] = [
    ['HOP', 'Path ID', 'GC', 'CM', 'Near Site A', 'Far Site B', 'Days Since Complete', 'Note (One and Done, verbatim)'],
    ...pw.oadSites.map(r => [r.hop, r.pathId, r.gc ?? '', r.cm, '', '', r.agingDays ?? '', r.note]),
  ]
  // oadSites doesn't carry Near/Far — pull them from the dataset by HOP.
  const siteByHop = new Map(dataset.map(r => [r.hop, r]))
  for (let i = 1; i < oadAoA.length; i++) {
    const src = siteByHop.get(String(oadAoA[i][0]))
    if (src) { oadAoA[i][4] = src.nearSiteA; oadAoA[i][5] = src.farSiteB }
  }
  XLSX.utils.book_append_sheet(wb, sheetFromAoA(oadAoA, [28, 16, 14, 14, 22, 22, 16, 48]), 'Pathwave OAD Sites')

  // ── Tab 5: QuickBase Pending (Nokia-internal, all 10 items, no OAD carve-out)
  const qbPending = dataset.filter(r => r.quickbaseStatus === 'PENDING')
  const qbPendingItemized = buildItemizedMissingItems(qbPending, QUICKBASE_ALL_ITEMS)
  const qbSheet = sheetFromAoA([
    ['QuickBase — Pending Items (Nokia-Internal, NOT a GC Action Item — do not send to a GC)'],
    [],
    ...itemizedAoA(qbPendingItemized),
  ])
  XLSX.utils.book_append_sheet(wb, qbSheet, 'QuickBase Pending')

  // ── Tab 6: QuickBase Complete ───────────────────────────────────────────
  XLSX.utils.book_append_sheet(
    wb, sheetFromAoA(simpleListAoA(dataset.filter(r => r.quickbaseStatus === 'COMPLETE'))), 'QuickBase Complete',
  )

  // ── Tab 7: Fully Complete ───────────────────────────────────────────────
  XLSX.utils.book_append_sheet(
    wb, sheetFromAoA(simpleListAoA(dataset.filter(r => r.fullyComplete))), 'Fully Complete',
  )

  // ── Tab 8: Filter Cheat Sheet ──────────────────────────────────────────
  const cheat: (string | number)[][] = [
    ['SCOP Filter Cheat Sheet — how each list above is derived'],
    [],
    ['List', 'Column', 'Condition'],
    ['Scope', 'Service provider', 'Exact value "Nokia" (case-insensitive, trimmed). Resolve by NAME not letter — position drifts.'],
    ['Cancelled excluded', 'any checklist column', 'Row dropped if any checklist cell contains "CXLD".'],
    ['Master gate', 'Construction Complete Actual', 'Has a real calendar date (parsed; a bare 00:00:00 / pre-1990 value counts as blank).'],
    ['Pending HOP Completion', 'Construction Complete Actual', 'No valid date.'],
    ['Pathwave Complete', 'One and Done', 'Value starts with "complete" (case-insensitive).'],
    ['Pathwave OAD', 'One and Done', 'Construction complete, not "complete", and text contains "oad" (substring, case-insensitive).'],
    ['Pathwave In Progress', 'One and Done', 'Construction complete, not "complete", not OAD.'],
    ['QuickBase Complete', 'Nokia Quickbase Deliverable Status', 'Value starts with "complete" (case-insensitive).'],
    ['QuickBase Pending', 'Nokia Quickbase Deliverable Status', 'Construction complete, value does not start with "complete". OAD carve-out does NOT apply here.'],
    ['Fully Complete', 'both status fields', 'Construction complete AND Pathwave Complete AND QuickBase Complete.'],
    ['Days Since Complete', 'Construction Complete Actual', '(as-of date) minus the date, in whole days. Sort itemized lists oldest-first.'],
    ['GC-owned Pathwave items', 'n/a', 'Install Photos, Decom Photos NQR, Install Photos NQR, Red-Line CD, Decom Asset Form, POD. Asset Form + Packing Slip are Nokia-owned.'],
    ['GC-owned QuickBase items', 'n/a', 'None — all QuickBase items are Nokia-owned. QuickBase never appears on a GC-facing report or email.'],
  ]
  XLSX.utils.book_append_sheet(wb, sheetFromAoA(cheat, [22, 30, 90]), 'Filter Cheat Sheet')

  return wb
}

/**
 * One GC's weekly report: Pathwave GC-owned items (excl. OAD) + an OAD section
 * only when that GC has OAD HOPs. No QuickBase section, ever (spec §7 / §9a).
 */
export function buildScopGcReportWorkbook(report: ScopGcReport): XLSX.WorkBook {
  const wb = XLSX.utils.book_new()
  const dateStr = fmtDate(report.generatedDate)

  const pwHeaders = ['HOP', 'Path ID', 'Near Site (A)', 'Far Site (B)', 'Site CM', 'Site', 'Days Since Complete', 'Missing Items']
  const pwAoA: (string | number)[][] = [
    [`${report.gc} — Weekly SCOP Action Items — ${dateStr}`],
    ['Pathwave GC-owned items only. Sorted oldest-first — clean up the top rows first. OAD sites are tracked separately.'],
    [],
    [`Total HOPs: ${report.summary.totalHOPs}    Completed: ${report.summary.completed}    Outstanding: ${report.summary.outstanding}    Oldest: ${report.summary.oldestOutstandingDays ?? '—'} days`],
    [],
    ['PATHWAVE — GC-Owned Items'],
    pwHeaders,
    ...report.sections.pathwave.map(r => [
      r.hop, r.pathId, r.nearSiteA, r.farSiteB, r.cm, r.site, r.agingDays ?? '', r.missingItems,
    ]),
  ]
  XLSX.utils.book_append_sheet(wb, sheetFromAoA(pwAoA, [28, 16, 22, 22, 16, 8, 16, 40]), 'Pathwave')

  if (report.sections.oad.length > 0) {
    const oadHeaders = ['HOP', 'Path ID', 'Near Site (A)', 'Far Site (B)', 'Site CM', 'Days Since Complete', 'One and Done (verbatim)']
    const oadAoA: (string | number)[][] = [
      [`${report.gc} — OAD — Awaiting OAD (Not a GC Checklist Item) — ${dateStr}`],
      [],
      oadHeaders,
      ...report.sections.oad.map(r => [
        r.hop, r.pathId, r.nearSiteA, r.farSiteB, r.cm, r.agingDays ?? '', r.note,
      ]),
    ]
    XLSX.utils.book_append_sheet(wb, sheetFromAoA(oadAoA, [28, 16, 22, 22, 16, 16, 48]), 'OAD')
  }

  return wb
}

export function downloadWorkbook(wb: XLSX.WorkBook, filename: string) {
  XLSX.writeFile(wb, filename)
}
