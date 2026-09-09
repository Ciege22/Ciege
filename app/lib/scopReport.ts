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
function simpleListAoA(title: string, rows: ScopRow[]): (string | number)[][] {
  return [
    [title], [],
    SIMPLE_HEADERS,
    ...rows.map(r => [r.hop, r.pathId, r.gc ?? '', r.cm, r.nearSiteA, r.farSiteB]),
  ]
}

// Itemized tabs (Pathwave Pending, QuickBase Pending)
const ITEMIZED_HEADERS = ['HOP', 'Path ID', 'GC', 'CM', 'Near Site A', 'Far Site B', 'Site', 'Days Since Complete', 'Missing Items']
function itemizedAoA(title: string, subtitle: string, rows: ReturnType<typeof buildItemizedMissingItems>): (string | number)[][] {
  return [
    [title], [subtitle], [],
    ITEMIZED_HEADERS,
    ...rows.map(r => [
      r.hop, r.pathId, r.gc ?? '', r.cm, r.nearSiteA, r.farSiteB,
      r.site, r.agingDays ?? '', r.missingItems,
    ]),
  ]
}

/**
 * The 9-tab Master SCOP Report — tab order, column sets, titles and the
 * Filter Cheat Sheet mirror docs/scop/build_scop_master_report.py (the tested
 * openpyxl reference). Cross-foots to the spec §10 baseline. Extra cross-foot
 * OK/MISMATCH rows are appended to the Summary tab as a build-check aid.
 */
export function buildScopMasterWorkbook(dataset: ScopRow[], s: ScopCalcSettings): XLSX.WorkBook {
  const wb = XLSX.utils.book_new()
  const qb = computeQuickBaseView(dataset)
  const pw = computePathwaveView(dataset)
  const ov = computeOverallView(dataset)
  const todayStr = fmtDate(s.asOfDate)

  // ── Tab 1: Summary ────────────────────────────────────────────────────────
  const summaryAoA: (string | number)[][] = [
    ['SCOP Master Report — Summary'],
    [`Generated ${todayStr}`],
    [],
    ['Metric', 'Count'],
    ['Total Nokia HOPs', dataset.length],
    ['Construction Complete (tracked)', ov.trackedTotal],
    ['Outstanding (not yet construction complete)', dataset.length - ov.trackedTotal],
    ['Pathwave — Complete', pw.complete],
    ['Pathwave — In Progress (true GC action, excl. OAD)', pw.inProgress],
    ['Pathwave — OAD (tracked separately)', pw.oad],
    ['QuickBase — Complete', qb.complete],
    ['QuickBase — Pending', qb.pending],
    ['Fully Complete (both Pathwave + QuickBase)', ov.fullyComplete],
    [],
    ['Cross-foot check', ''],
    ['Pathwave 4-way = Total', pw.complete + pw.inProgress + pw.oad + pw.pendingHopCompletion === dataset.length ? 'OK' : 'MISMATCH'],
    ['QuickBase Complete + Pending = Tracked', qb.complete + qb.pending === ov.trackedTotal ? 'OK' : 'MISMATCH'],
    ['Pathwave Complete + Pending = Tracked', ov.pathwaveComplete + ov.pathwavePending === ov.trackedTotal ? 'OK' : 'MISMATCH'],
  ]
  XLSX.utils.book_append_sheet(wb, sheetFromAoA(summaryAoA, [46, 14]), 'Summary')

  // ── Tab 2: Outstanding List ──────────────────────────────────────────────
  const outstanding = dataset.filter(r => !r.isConstructionComplete)
  XLSX.utils.book_append_sheet(
    wb, sheetFromAoA(simpleListAoA('Outstanding — Not Yet Construction Complete', outstanding)), 'Outstanding List',
  )

  // ── Tab 3: Pathwave Pending (itemized, excludes OAD) ──────────────────────
  const pathwaveInProgress = dataset.filter(r => r.pathwaveStatus === 'IN_PROGRESS')
  const pathwavePendingItemized = buildItemizedMissingItems(pathwaveInProgress, pathwaveGcItems(s))
  XLSX.utils.book_append_sheet(wb, sheetFromAoA(itemizedAoA(
    'Pathwave — Pending Items (GC-Owned, Action Needed)',
    'Sorted oldest-first (highest days since construction complete) — clean up the top rows first',
    pathwavePendingItemized,
  )), 'Pathwave Pending')

  // ── Tab 4: Pathwave Complete ─────────────────────────────────────────────
  XLSX.utils.book_append_sheet(wb, sheetFromAoA(simpleListAoA(
    'Pathwave — Complete (Pending Viaero Approval)',
    dataset.filter(r => r.pathwaveStatus === 'COMPLETE'),
  )), 'Pathwave Complete')

  // ── Tab 4a: Pathwave OAD Sites ───────────────────────────────────────────
  const oadHeaders = ['HOP', 'Path ID', 'GC', 'CM', 'Near Site A', 'Far Site B', 'Days Since Complete', 'Note']
  const siteByHop = new Map(dataset.map(r => [r.hop, r]))
  const oadAoA: (string | number)[][] = [
    ['Pathwave — OAD Sites (Awaiting OAD, Not a GC Checklist Item)'],
    ['Sorted oldest-first. These sites are blocked on OAD, not on GC-submitted checklist items.'],
    [],
    oadHeaders,
    ...pw.oadSites.map(r => {
      const src = siteByHop.get(r.hop)
      return [r.hop, r.pathId, r.gc ?? '', r.cm, src?.nearSiteA ?? '', src?.farSiteB ?? '', r.agingDays ?? '', r.note]
    }),
  ]
  XLSX.utils.book_append_sheet(wb, sheetFromAoA(oadAoA, [28, 16, 14, 12, 22, 22, 16, 44]), 'Pathwave OAD Sites')

  // ── Tab 5: QuickBase Pending (Nokia-internal, all QuickBase items, no OAD carve-out)
  const qbPending = dataset.filter(r => r.quickbaseStatus === 'PENDING')
  const qbPendingItemized = buildItemizedMissingItems(qbPending, QUICKBASE_ALL_ITEMS)
  XLSX.utils.book_append_sheet(wb, sheetFromAoA(itemizedAoA(
    'QuickBase — Pending Items (Nokia-Internal, NOT a GC Action Item)',
    'All QuickBase items are Nokia-owned. This list is for internal engineering tracking only — never send to a GC. Sorted oldest-first.',
    qbPendingItemized,
  )), 'QuickBase Pending')

  // ── Tab 6: QuickBase Complete ───────────────────────────────────────────
  XLSX.utils.book_append_sheet(wb, sheetFromAoA(simpleListAoA(
    'QuickBase — Complete', dataset.filter(r => r.quickbaseStatus === 'COMPLETE'),
  )), 'QuickBase Complete')

  // ── Tab 7: Fully Complete ───────────────────────────────────────────────
  XLSX.utils.book_append_sheet(wb, sheetFromAoA(simpleListAoA(
    'Fully Complete — Both Pathwave & QuickBase Done (Ready for Viaero Approval)',
    dataset.filter(r => r.fullyComplete),
  )), 'Fully Complete')

  // ── Tab 8: Filter Cheat Sheet ──────────────────────────────────────────
  // Wording mirrors the openpyxl reference. No cell value starts with "=".
  const cheat: (string | number)[][] = [
    ['Filter Cheat Sheet — How to Rebuild These Lists Manually'],
    [],
    ['List', 'Filter On', 'Condition'],
    ['Scope (which HOPs are yours)', 'Column "Service provider" (position may shift — filter by column NAME, not letter)', 'Exact value "Nokia" (case-insensitive)'],
    ['Outstanding (not built yet)', 'Column P — Construction Complete Actual', 'Blank / no valid date'],
    ['Construction-complete (tracked universe)', 'Column P — Construction Complete Actual', 'Has a valid date'],
    ['Pathwave — OAD (separate from Pending)', 'Column S — One and Done', 'Contains "OAD" anywhere, case-insensitive AND Column P has a date. Check BEFORE "Pending" below.'],
    ['Pathwave — Pending (true GC action)', 'Column S — One and Done', 'Does NOT start with "Complete" AND does NOT contain "OAD" AND Column P has a date'],
    ['Pathwave — Complete', 'Column S — One and Done', 'Starts with "Complete" AND Column P has a date'],
    ['QuickBase — Pending (Nokia-internal only)', 'Column BT — Nokia Quickbase Deliverable Status', 'Does NOT start with "Complete" AND Column P has a date. ALL QuickBase items are Nokia-owned — never GC-facing.'],
    ['QuickBase — Complete', 'Column BT — Nokia Quickbase Deliverable Status', 'Starts with "Complete" AND Column P has a date'],
    ['Fully Complete (approvals-ready)', 'Columns S + BT together', 'Both start with Complete AND Column P has a date'],
    ['Pathwave GC action items (weekly report)', 'Hard-coded item list', '6 of 8 Pathwave items: excludes Asset Form & Packing Slip (Nokia-owned).'],
    ['QuickBase GC action items', 'N/A', 'NONE. All QuickBase items are Nokia-owned. Never appears on a GC report or email.'],
  ]
  XLSX.utils.book_append_sheet(wb, sheetFromAoA(cheat, [32, 40, 56]), 'Filter Cheat Sheet')

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
