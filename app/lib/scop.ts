// SCOP (Site Close-Out Package) calculation engine.
//
// Faithful TypeScript port of docs/scop/scop_calculations.js — the business
// rules there are tested and validated against the real live tracker (09/08/2026
// corrected run: 205 Nokia HOPs, Pathwave 94/42/19/50, QuickBase 114/41). Do not
// change the classification/aggregation logic without re-validating against that
// baseline. The only deliberate deviations from the JS are the three places the
// settings schema (docs/scop/scop_settings_schema.js) says a value must be
// settings-driven rather than hard-coded — OAD keywords, the Pathwave GC-owned
// item set, and the aging-color thresholds — and every default there reproduces
// the JS's hard-coded behaviour exactly.
//
// See docs/scop/SCOP_View_Build_Spec.md for the "why" behind every rule,
// including the OAD carve-out (§4a) and two real data-alignment incidents (§10).

import { GC_CONFIG } from './gcConfig'
import type { ScopSettings } from './settings'
import { DEFAULT_SCOP } from './settings'

// ─────────────────────────────────────────────────────────────────────────────
// 1. COLUMN NAME CONSTANTS
// Always resolve columns by NAME, never by position — the "Service provider"
// column has drifted between exports and caused a 208-vs-205 miscount.
// ─────────────────────────────────────────────────────────────────────────────

export const SCOP_COLUMNS = {
  SCOPE: 'Service provider', // filter value: exact "Nokia" (case-insensitive)
  PATH_ID: 'Path ID',
  HOP: 'HOP',
  GENERAL_CONTRACTOR: 'General Contractor',
  SITE_CM: 'Site CM',
  NEAR_SITE_A: 'Near Site Name A',
  FAR_SITE_B: 'Far Site Name B',
  CONSTRUCTION_COMPLETE_ACTUAL: 'Construction Complete Actual', // the master gate
  FINAL_CUTOVER: 'Final Cutover', // display-only, not used for gating
  ONE_AND_DONE: 'One and Done', // authoritative Pathwave status field
  QB_DELIVERABLE_STATUS: 'Nokia Quickbase Deliverable Status', // authoritative QB status field
} as const

export interface ScopChecklistItem {
  label: string
  colA: string
  colB: string
}

// Pathwave checklist — the full 8 items per site. The 6 GC-owned ones are
// selected at runtime from ScopSettings.pathwaveGcOwnedItems (default = the
// 6 below minus Asset Form / Packing Slip, which are Nokia-owned).
export const PATHWAVE_ALL_ITEMS: ScopChecklistItem[] = [
  { label: 'Install Photos', colA: 'Viaero Install Photos Status Site A', colB: 'Viaero Install Photo Status B' },
  { label: 'Decom Photos NQR', colA: 'Site A Decom Photos NQR', colB: 'Site B Decom Photos NQR' },
  { label: 'Install Photos NQR', colA: 'Site A Install Photos NQR', colB: 'Site B Install Photos NQR' },
  { label: 'Red-Line CD', colA: 'Site A Red-Line CD', colB: 'Site B Red Line CDs' },
  { label: 'Decom Asset Form', colA: 'Site A Decom Asset Form', colB: 'Site B Decom Asset Form' },
  { label: 'POD', colA: 'Site A POD', colB: 'Site B POD' },
  { label: 'Asset Form', colA: 'Site A Asset Form', colB: 'Site B Asset Form' },
  { label: 'Packing Slip', colA: 'Site A Packing Slip', colB: 'Site B Packing Slip' },
]

// QuickBase checklist — ALL Nokia-owned, confirmed. Used only for the
// internal engineering breakdown, never a GC-facing report.
export const QUICKBASE_ALL_ITEMS: ScopChecklistItem[] = [
  { label: 'B2B Test', colA: 'QB Site A B2B Test', colB: 'Site B B2B Test' },
  { label: 'RFC2455 Test', colA: 'QB Site A RFC2455 Test', colB: 'QB Site B RFC2455 Test' },
  { label: 'BER Report', colA: 'QB Site A BER Report', colB: 'QB Site B BER Report' },
  { label: 'Consolidated Asset Form', colA: 'QB Site A Consolidated Asset Form', colB: 'QB Site B Consolidated Asset Form' },
  { label: 'As-Built Final CD', colA: 'QB Site A As-Built Final CD', colB: 'QB Site B As-Built Final CD' },
  { label: 'As Built WP', colA: 'QB Site A As Built WP', colB: 'QB Site B As Built WP' },
  { label: 'As-Built LLD', colA: 'QB Site A As-Built LLD', colB: 'QB Site B As-Built LLD' },
  // sic — "KPI Report3" is mislabeled in the source tracker; it IS Site B's field.
  { label: 'KPI Report', colA: 'QB Site A KPI Report', colB: 'QB Site A KPI Report3' },
  { label: 'POD', colA: 'Site A POD2', colB: 'Site B POD4' },
  { label: 'HW POD', colA: 'HW POD', colB: 'HW POD2' },
]

// Extra checklist columns that gate CXLD (cancelled) exclusion but aren't
// themselves ownership-tracked items.
const CANCELLATION_EXTRA_COLS = [
  'Site A Asset Form', 'Site B Asset Form', 'Site A Packing Slip', 'Site B Packing Slip',
]

// ─────────────────────────────────────────────────────────────────────────────
// 2. NORMALIZATION HELPERS
// ─────────────────────────────────────────────────────────────────────────────

export type ScopRawRow = Record<string, unknown>

function norm(v: unknown): string | null {
  if (v === null || v === undefined) return null
  const s = String(v).trim()
  return s === '' ? null : s.toLowerCase()
}

function isItemDone(v: unknown): boolean {
  const s = norm(v)
  return s !== null && ['yes', 'accepted', 'complete'].includes(s)
}

function startsWithComplete(v: unknown): boolean {
  const s = norm(v)
  return s !== null && s.startsWith('complete')
}

function containsAnyKeyword(v: unknown, keywords: string[]): boolean {
  const s = norm(v)
  if (s === null) return false
  return keywords.some(k => s.includes(k.toLowerCase()))
}

/**
 * Parses Construction Complete Actual into a real Date or null.
 * Blank cells in this column have been observed coming through as a time-only
 * artifact (00:00:00 / a ~1899 Date) rather than a true null. Any value that
 * doesn't resolve to a real modern calendar date is treated as null — a naive
 * not-null check on this column reads every row as "construction complete".
 */
export function parseConstructionCompleteDate(rawValue: unknown): Date | null {
  if (rawValue === null || rawValue === undefined) return null
  if (rawValue instanceof Date) {
    if (isNaN(rawValue.getTime()) || rawValue.getFullYear() < 1990) return null
    return rawValue
  }
  // Excel serial number (xlsx without cellDates) — days since 1899-12-30.
  if (typeof rawValue === 'number') {
    if (rawValue < 40000 || rawValue > 60000) return null // outside ~2009..2064
    const d = new Date(Math.round((rawValue - 25569) * 86400 * 1000))
    return isNaN(d.getTime()) || d.getFullYear() < 1990 ? null : d
  }
  const s = String(rawValue).trim()
  if (!s || s === 'null' || s === 'undefined' || s === 'NaN') return null
  const parsed = new Date(s)
  if (isNaN(parsed.getTime()) || parsed.getFullYear() < 1990) return null
  return parsed
}

/**
 * Alias map layered on GC_CONFIG's roster: every configured GC maps to its
 * own canonical casing, plus any extra raw->canonical entries from settings
 * (seeded with the WaveLink / Viking collisions). Keyed lowercase.
 */
export function buildScopAliasMap(extra: Record<string, string> = {}): Record<string, string> {
  const map: Record<string, string> = {}
  GC_CONFIG.forEach(c => { map[c.gc.trim().toLowerCase()] = c.gc })
  Object.entries(extra).forEach(([k, v]) => { map[k.trim().toLowerCase()] = v })
  return map
}

export function normalizeGCName(rawName: unknown, aliasMap: Record<string, string>): string | null {
  if (rawName === null || rawName === undefined || String(rawName).trim() === '') return null
  const trimmed = String(rawName).trim()
  return aliasMap[trimmed.toLowerCase()] || trimmed
}

function daysSince(date: Date | null, asOfDate: Date): number | null {
  if (!date) return null
  return Math.floor((asOfDate.getTime() - date.getTime()) / (1000 * 60 * 60 * 24))
}

// ─────────────────────────────────────────────────────────────────────────────
// 3. ROW-LEVEL CLASSIFICATION
// ─────────────────────────────────────────────────────────────────────────────

export type PathwaveStatus = 'COMPLETE' | 'IN_PROGRESS' | 'OAD' | 'PENDING_HOP_COMPLETION'
export type QuickbaseStatus = 'COMPLETE' | 'PENDING' | 'PENDING_HOP_COMPLETION'

export interface ScopRow {
  hop: string
  pathId: string
  gc: string | null
  cm: string
  nearSiteA: string
  farSiteB: string
  isConstructionComplete: boolean
  constructionCompleteDate: Date | null
  agingDays: number | null
  oneAndDoneRaw: string
  isOAD: boolean
  pathwaveStatus: PathwaveStatus
  quickbaseStatus: QuickbaseStatus
  fullyComplete: boolean
  _rawRow: ScopRawRow
}

// Resolved subset of ScopSettings the calc functions actually need, plus the
// concrete as-of Date. Build it once per run with resolveScopCalcSettings().
export interface ScopCalcSettings {
  aliasMap: Record<string, string>
  oadKeywords: string[]
  emailPriorityThresholdDays: number
  emailPriorityCap: number
  pathwaveGcOwnedItems: string[]
  asOfDate: Date
}

export function resolveScopCalcSettings(s: ScopSettings = DEFAULT_SCOP): ScopCalcSettings {
  return {
    aliasMap: buildScopAliasMap(s.gcAliasMap),
    oadKeywords: s.oadKeywords?.length ? s.oadKeywords : ['oad'],
    emailPriorityThresholdDays: s.emailPriorityThresholdDays ?? 60,
    emailPriorityCap: s.emailPriorityCap ?? 5,
    pathwaveGcOwnedItems: s.pathwaveGcOwnedItems?.length
      ? s.pathwaveGcOwnedItems
      : DEFAULT_SCOP.pathwaveGcOwnedItems,
    asOfDate: s.asOfDateOverride ? new Date(s.asOfDateOverride + 'T00:00:00') : new Date(),
  }
}

export function pathwaveGcItems(s: ScopCalcSettings): ScopChecklistItem[] {
  const owned = new Set(s.pathwaveGcOwnedItems)
  return PATHWAVE_ALL_ITEMS.filter(i => owned.has(i.label))
}

export function classifyRow(row: ScopRawRow, s: ScopCalcSettings): ScopRow {
  const constructionCompleteDate = parseConstructionCompleteDate(row[SCOP_COLUMNS.CONSTRUCTION_COMPLETE_ACTUAL])
  const isConstructionComplete = constructionCompleteDate !== null
  const agingDays = isConstructionComplete ? daysSince(constructionCompleteDate, s.asOfDate) : null

  const oneAndDone = row[SCOP_COLUMNS.ONE_AND_DONE]
  const isOAD = containsAnyKeyword(oneAndDone, s.oadKeywords)
  const pathwaveComplete = startsWithComplete(oneAndDone)

  // Order matters: OAD is checked before falling through to In Progress (spec §4a).
  let pathwaveStatus: PathwaveStatus
  if (!isConstructionComplete) pathwaveStatus = 'PENDING_HOP_COMPLETION'
  else if (pathwaveComplete) pathwaveStatus = 'COMPLETE'
  else if (isOAD) pathwaveStatus = 'OAD'
  else pathwaveStatus = 'IN_PROGRESS'

  const qbComplete = startsWithComplete(row[SCOP_COLUMNS.QB_DELIVERABLE_STATUS])
  const quickbaseStatus: QuickbaseStatus = !isConstructionComplete
    ? 'PENDING_HOP_COMPLETION'
    : qbComplete ? 'COMPLETE' : 'PENDING'

  return {
    hop: String(row[SCOP_COLUMNS.HOP] ?? '').trim(),
    pathId: String(row[SCOP_COLUMNS.PATH_ID] ?? '').trim(),
    gc: normalizeGCName(row[SCOP_COLUMNS.GENERAL_CONTRACTOR], s.aliasMap),
    cm: String(row[SCOP_COLUMNS.SITE_CM] ?? '').trim(),
    nearSiteA: String(row[SCOP_COLUMNS.NEAR_SITE_A] ?? '').trim(),
    farSiteB: String(row[SCOP_COLUMNS.FAR_SITE_B] ?? '').trim(),
    isConstructionComplete,
    constructionCompleteDate,
    agingDays,
    oneAndDoneRaw: oneAndDone == null ? '' : String(oneAndDone).trim(),
    isOAD,
    pathwaveStatus,
    quickbaseStatus,
    fullyComplete: isConstructionComplete && pathwaveComplete && qbComplete,
    _rawRow: row,
  }
}

/**
 * Filters raw name-keyed rows to Nokia's scope, drops cancelled (CXLD) HOPs,
 * and classifies each. Header-row resolution happens upstream (see the Reports
 * page upload handler) — this takes already-keyed row objects.
 */
export function buildScopDataset(rawRows: ScopRawRow[], s: ScopCalcSettings): ScopRow[] {
  const scoped = rawRows.filter(row => norm(row[SCOP_COLUMNS.SCOPE]) === 'nokia')

  const allChecklistCols = [
    ...PATHWAVE_ALL_ITEMS.slice(0, 6).flatMap(i => [i.colA, i.colB]),
    ...CANCELLATION_EXTRA_COLS,
    ...QUICKBASE_ALL_ITEMS.flatMap(i => [i.colA, i.colB]),
  ]

  const active = scoped.filter(row =>
    !allChecklistCols.some(col => {
      const v = norm(row[col])
      return v !== null && v.includes('cxld')
    })
  )

  return active.map(row => classifyRow(row, s))
}

// ─────────────────────────────────────────────────────────────────────────────
// 4. AGGREGATE VIEWS
// ─────────────────────────────────────────────────────────────────────────────

export interface QuickBaseView {
  complete: number
  pending: number
  pendingHopCompletion: number
  total: number
  trackedTotal: number
  percentComplete: number
}

export function computeQuickBaseView(dataset: ScopRow[]): QuickBaseView {
  const counts = { COMPLETE: 0, PENDING: 0, PENDING_HOP_COMPLETION: 0 }
  dataset.forEach(r => { counts[r.quickbaseStatus]++ })
  const total = dataset.length
  const trackedTotal = counts.COMPLETE + counts.PENDING
  return {
    complete: counts.COMPLETE,
    pending: counts.PENDING,
    pendingHopCompletion: counts.PENDING_HOP_COMPLETION,
    total,
    trackedTotal,
    percentComplete: trackedTotal > 0 ? Math.round((100 * counts.COMPLETE) / trackedTotal) : 0,
  }
}

export interface PathwaveGcBreakdown { inProgress: number; complete: number; oad: number }
export interface OadSite {
  hop: string; pathId: string; gc: string | null; cm: string; note: string; agingDays: number | null
}
export interface PathwaveView {
  complete: number
  inProgress: number
  oad: number
  pendingHopCompletion: number
  total: number
  byGC: Record<string, PathwaveGcBreakdown>
  oadSites: OadSite[]
}

export function computePathwaveView(dataset: ScopRow[]): PathwaveView {
  const counts = { COMPLETE: 0, IN_PROGRESS: 0, OAD: 0, PENDING_HOP_COMPLETION: 0 }
  dataset.forEach(r => { counts[r.pathwaveStatus]++ })

  const byGC: Record<string, PathwaveGcBreakdown> = {}
  dataset
    .filter(r => r.isConstructionComplete && r.gc)
    .forEach(r => {
      const gc = r.gc as string
      if (!byGC[gc]) byGC[gc] = { inProgress: 0, complete: 0, oad: 0 }
      if (r.pathwaveStatus === 'IN_PROGRESS') byGC[gc].inProgress++
      else if (r.pathwaveStatus === 'COMPLETE') byGC[gc].complete++
      else if (r.pathwaveStatus === 'OAD') byGC[gc].oad++
    })

  const oadSites: OadSite[] = dataset
    .filter(r => r.pathwaveStatus === 'OAD')
    .sort((a, b) => (b.agingDays || 0) - (a.agingDays || 0))
    .map(r => ({ hop: r.hop, pathId: r.pathId, gc: r.gc, cm: r.cm, note: r.oneAndDoneRaw, agingDays: r.agingDays }))

  return {
    complete: counts.COMPLETE,
    inProgress: counts.IN_PROGRESS,
    oad: counts.OAD,
    pendingHopCompletion: counts.PENDING_HOP_COMPLETION,
    total: dataset.length,
    byGC,
    oadSites,
  }
}

export interface OverallView {
  trackedTotal: number
  pathwaveComplete: number
  pathwavePending: number
  quickbaseComplete: number
  quickbasePending: number
  fullyComplete: number
  notFullyComplete: number
}

export function computeOverallView(dataset: ScopRow[]): OverallView {
  const tracked = dataset.filter(r => r.isConstructionComplete)
  const pathwaveComplete = tracked.filter(r => r.pathwaveStatus === 'COMPLETE').length
  const quickbaseComplete = tracked.filter(r => r.quickbaseStatus === 'COMPLETE').length
  const fullyComplete = tracked.filter(r => r.fullyComplete).length
  return {
    trackedTotal: tracked.length,
    pathwaveComplete,
    // Intentionally includes OAD — View 3 does not carve OAD out, only View 2 does.
    pathwavePending: tracked.length - pathwaveComplete,
    quickbaseComplete,
    quickbasePending: tracked.length - quickbaseComplete,
    fullyComplete,
    notFullyComplete: tracked.length - fullyComplete,
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// 5. ITEMIZED "WHAT'S MISSING" DETAIL
// ─────────────────────────────────────────────────────────────────────────────

export interface ItemizedMissingRow {
  hop: string
  pathId: string
  gc: string | null
  cm: string
  nearSiteA: string
  farSiteB: string
  site: 'Site A' | 'Site B'
  agingDays: number | null
  missingItems: string
}

/** One row per HOP+Site where at least one item in `itemList` is not done. */
export function buildItemizedMissingItems(rows: ScopRow[], itemList: ScopChecklistItem[]): ItemizedMissingRow[] {
  const out: ItemizedMissingRow[] = []
  rows.forEach(r => {
    ;(['A', 'B'] as const).forEach(site => {
      const missing = itemList
        .filter(item => !isItemDone(r._rawRow[site === 'A' ? item.colA : item.colB]))
        .map(item => item.label)
      if (missing.length === 0) return
      out.push({
        hop: r.hop,
        pathId: r.pathId,
        gc: r.gc,
        cm: r.cm,
        nearSiteA: r.nearSiteA,
        farSiteB: r.farSiteB,
        site: `Site ${site}`,
        agingDays: r.agingDays,
        missingItems: missing.join(', '),
      })
    })
  })
  return out.sort((a, b) => (b.agingDays || 0) - (a.agingDays || 0))
}

// ─────────────────────────────────────────────────────────────────────────────
// 6. PER-GC REPORT PAYLOADS
// ─────────────────────────────────────────────────────────────────────────────

export interface ScopOadDetailRow {
  hop: string; pathId: string; nearSiteA: string; farSiteB: string
  cm: string; agingDays: number | null; note: string
}

export interface ScopGcReport {
  gc: string
  generatedDate: Date
  summary: {
    totalHOPs: number
    completed: number
    outstanding: number // includes OAD, for an honest total
    oldestOutstandingDays: number | null
  }
  sections: {
    pathwave: ItemizedMissingRow[] // excludes OAD rows (spec §4a)
    oad: ScopOadDetailRow[] // only non-empty when the GC has OAD HOPs
    // No quickbase section — all QuickBase items are Nokia-owned.
  }
}

/** Returns null if the GC has nothing outstanding (no report should be made). */
export function buildGCReport(dataset: ScopRow[], gcName: string, s: ScopCalcSettings): ScopGcReport | null {
  const gcRows = dataset.filter(r => r.gc === gcName && r.isConstructionComplete)
  const pathwayPendingRows = gcRows.filter(r => r.pathwaveStatus === 'IN_PROGRESS')
  const oadRows = gcRows.filter(r => r.pathwaveStatus === 'OAD')
  if (pathwayPendingRows.length === 0 && oadRows.length === 0) return null

  const pathwaveMissingItems = buildItemizedMissingItems(pathwayPendingRows, pathwaveGcItems(s))

  const oadDetail: ScopOadDetailRow[] = [...oadRows]
    .sort((a, b) => (b.agingDays || 0) - (a.agingDays || 0))
    .map(r => ({
      hop: r.hop, pathId: r.pathId, nearSiteA: r.nearSiteA, farSiteB: r.farSiteB,
      cm: r.cm, agingDays: r.agingDays, note: r.oneAndDoneRaw,
    }))

  const completeRows = gcRows.filter(r => r.pathwaveStatus === 'COMPLETE')
  const totalOutstanding = pathwayPendingRows.length + oadRows.length

  return {
    gc: gcName,
    generatedDate: s.asOfDate,
    summary: {
      totalHOPs: completeRows.length + totalOutstanding,
      completed: completeRows.length,
      outstanding: totalOutstanding,
      oldestOutstandingDays: totalOutstanding > 0
        ? Math.max(...pathwayPendingRows.concat(oadRows).map(r => r.agingDays || 0))
        : null,
    },
    sections: { pathwave: pathwaveMissingItems, oad: oadDetail },
  }
}

export function buildAllGCReports(dataset: ScopRow[], s: ScopCalcSettings): ScopGcReport[] {
  const gcNames = [...new Set(dataset.map(r => r.gc).filter((g): g is string => !!g))].sort()
  return gcNames
    .map(gc => buildGCReport(dataset, gc, s))
    .filter((r): r is ScopGcReport => r !== null)
}

// ─────────────────────────────────────────────────────────────────────────────
// 7. GC EMAIL (text + mailto) — mirrors buildDecomEmailMailto / buildGrEmailMailto
// ─────────────────────────────────────────────────────────────────────────────

export interface ScopEmailText {
  subject: string
  body: string
  attachmentFilename: string
}

function mmddyyyy(d: Date): string {
  const p = (n: number) => String(n).padStart(2, '0')
  return `${p(d.getMonth() + 1)}/${p(d.getDate())}/${d.getFullYear()}`
}

export function buildGCEmailText(report: ScopGcReport, s: ScopCalcSettings): ScopEmailText {
  const { gc, summary, sections } = report
  const dateStr = mmddyyyy(report.generatedDate)
  const threshold = s.emailPriorityThresholdDays
  const cap = s.emailPriorityCap

  // Dedupe the site-level Pathwave rows down to one entry per HOP (already OAD-free).
  const priorityItems = [...new Map(sections.pathwave.map(r => [r.hop, r])).values()]
    .filter(r => (r.agingDays || 0) >= threshold)
    .sort((a, b) => (b.agingDays || 0) - (a.agingDays || 0))
    .slice(0, cap)

  const subject = `Weekly SCOP Action Items — ${gc} (${dateStr})`

  let body = `Hi ${gc},\n\nSummary for ${gc}:\n`
  body += `• Total HOPs: ${summary.totalHOPs}\n`
  body += `• Completed: ${summary.completed}\n`
  body += `• Outstanding: ${summary.outstanding}\n`
  body += `• Oldest outstanding item: ${summary.oldestOutstandingDays} days since construction complete\n\n`
  body += `Full detail is in the attached report, sorted oldest-first.\n\n`

  if (priorityItems.length > 0) {
    body += `Top priority (aging ${threshold}+ days since construction complete):\n`
    body += priorityItems.map(r => `• ${r.hop} — ${r.agingDays} days`).join('\n')
    body += `\n\nThese are the longest-outstanding items — let's prioritize closing these out first.\n\n`
  } else {
    body += `Nothing is over the ${threshold}-day priority threshold yet, but please keep the attached list moving.\n\n`
  }

  body += `Please review the attached report and let us know a target date for the outstanding items. Happy to hop on a call if anything needs clarifying.\n\nThanks,\nCJ`

  return {
    subject,
    body,
    attachmentFilename: `SCOP_GC_Report_${gc.replace(/[/ ]/g, '_')}_${dateStr.replace(/\//g, '-')}.xlsx`,
  }
}

export function buildScopGcEmailMailto(
  report: ScopGcReport,
  s: ScopCalcSettings,
  emailSettings: { ccList: string[]; gcContactEmails: Record<string, string> },
): string {
  const { subject, body } = buildGCEmailText(report, s)
  const target = report.gc.trim().toLowerCase()
  const toKey = Object.keys(emailSettings.gcContactEmails).find(k => k.trim().toLowerCase() === target)
  const to = toKey ? emailSettings.gcContactEmails[toKey] : ''
  const cc = emailSettings.ccList.join(',')
  return `mailto:${to}?cc=${encodeURIComponent(cc)}&subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`
}

// ─────────────────────────────────────────────────────────────────────────────
// 8. UPLOAD SUPPORT — header-row detection + row keying
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Finds the 0-based index of the header row in a raw `sheet_to_json(header:1)`
 * grid by scanning the first `scanRows` rows for one containing BOTH "HOP" and
 * "General Contractor" as trimmed cell values. Returns -1 if not found.
 */
export function detectScopHeaderRow(rawRows: unknown[][], scanRows = 15): number {
  const limit = Math.min(scanRows, rawRows.length)
  for (let i = 0; i < limit; i++) {
    const cells = (rawRows[i] || []).map(c => String(c ?? '').trim())
    if (cells.includes('HOP') && cells.includes('General Contractor')) return i
  }
  return -1
}

/**
 * Re-keys stored array rows into name-keyed objects. `stored[0]` is the header
 * row (this is how the Decom report stores its data too — header travels with
 * the payload rather than being stripped).
 */
export function keyScopRows(stored: unknown[][]): ScopRawRow[] {
  if (stored.length < 2) return []
  const headers = (stored[0] || []).map(h => String(h ?? '').trim())
  return stored.slice(1)
    .filter(r => Array.isArray(r) && r.some(v => v !== null && v !== undefined && String(v).trim() !== ''))
    .map(r => {
      const obj: ScopRawRow = {}
      headers.forEach((h, i) => { if (h) obj[h] = (r as unknown[])[i] ?? null })
      return obj
    })
}

// ─────────────────────────────────────────────────────────────────────────────
// 9. AGING COLOR
// ─────────────────────────────────────────────────────────────────────────────

export type AgingColor = 'green' | 'amber' | 'red'

export function agingColor(days: number | null, thresholds: { green: number; amber: number }): AgingColor {
  const n = days ?? 0
  if (n < thresholds.green) return 'green'
  if (n < thresholds.amber) return 'amber'
  return 'red'
}
