import { supabase, loadTrackerSnapshot } from './supabase'
import { GC_CONFIG, matches } from './gcConfig'
import { lookupContactEmail, DEFAULT_SCOP, type ScopSettings } from './settings'
import { parseDecomRows } from './decom'
import { buildScopDataset, keyScopRows, resolveScopCalcSettings } from './scop'
import { loadChunkedReport } from './reportChunks'

// Raw SPO report rows are stored as arrays with no header row (see app/reports/page.tsx).
// These are the fixed column positions established by that page's SPO_COL_IDX mapping.
const SPO_COL = {
  customerSiteId: 7,
  name: 8,
  sogName: 33,
  spoNumber: 40,
  spoVendor: 43,
  spoValue: 47,
  grDate: 50,
}

export type GrTier = 'init20' | '60' | '70' | '20' | '30' | 'CR' | null
export type GrTrigger = 'MS15A' | 'MS16A' | 'DECOM_SCOP' | null
export type GrStatus = 'GR Done' | 'Ready to Release' | 'Awaiting Trigger'

// Row classification (foundational — everything else derives from this):
//   Base PO = has SPO Number AND has SOG Name
//   CR      = has SPO Number AND SOG Name is blank — displayed as-is, no "CR" label in UI
//   Error   = has SOG Name but NO SPO Number — excluded entirely, never returned/displayed
export type GrRowType = 'base' | 'cr'

export interface GrRow {
  hop: string            // normalized lowercase key used for matching
  hopDisplay: string      // SPO report Name column, as-is
  pathId: string
  gc: string
  nokiaPm: string
  sogName: string
  spoNumber: string
  spoValue: number
  rowType: GrRowType
  triggerDate: string     // MM/DD/YYYY, blank if unknown
  triggerDateRaw: Date | null
  trigger: GrTrigger
  triggerMet: boolean     // CR rows have no trigger gate — always true until GR'd
  grDate: string          // MM/DD/YYYY, blank if not processed
  grDateRaw: Date | null
  status: GrStatus
  tier: GrTier
  tierLabel: string
  cjActionable: boolean
  isDecomScop: boolean
  pendingReason: string   // blank once GR Done — see pendingReasonFor()
}

export interface TrackerDates {
  ms15a: Date | null
  ms16a: Date | null
  nokiaPm: string
}

// Tracker uses '<>' as the HOP separator, SPO report uses 'to'.
export function normalizeTrackerHop(hop: string): string {
  return hop.replace(' <> ', ' to ').replace('<>', 'to')
}

function matchKey(s: string): string {
  return s.trim().toLowerCase()
}

export function parseDateAny(val: unknown): Date | null {
  if (!val) return null
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val
  if (typeof val === 'number') {
    if (val > 40000 && val < 60000) {
      const d = new Date((val - 25569) * 86400 * 1000)
      return isNaN(d.getTime()) ? null : d
    }
    return null
  }
  const s = String(val).trim()
  if (!s || s === 'null' || s === 'undefined' || s === 'NaN') return null
  const d = new Date(s)
  if (!isNaN(d.getTime())) return d
  const parts = s.split('/')
  if (parts.length === 3) {
    const d2 = new Date(parseInt(parts[2]), parseInt(parts[0]) - 1, parseInt(parts[1]))
    return isNaN(d2.getTime()) ? null : d2
  }
  return null
}

export function fmtDate(d: Date | null): string {
  if (!d) return ''
  return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`
}

export function fmtMoney(n: number): string {
  return `$${n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`
}

export function fmtMoneyShort(n: number): string {
  return `$${Math.round(n / 1000)}k`
}

// SOG Name tier bucketing. Evaluation order matters — "init 20%" must be
// checked before plain "20%" or Init sites get misclassified as Decom/SCOP.
export function classifySog(sogNameRaw: string): {
  tier: GrTier; trigger: GrTrigger; cjActionable: boolean; tierLabel: string; isDecomScop: boolean
} {
  const s = (sogNameRaw || '').trim().toLowerCase()
  if (s.includes('init 20%')) {
    return { tier: 'init20', trigger: 'MS15A', cjActionable: true, tierLabel: 'Init 20% — MS15A (Site Started)', isDecomScop: false }
  }
  if (s.includes('60%')) {
    return { tier: '60', trigger: 'MS16A', cjActionable: true, tierLabel: '60% — MS16A (CX Complete)', isDecomScop: false }
  }
  if (s.includes('70%')) {
    return { tier: '70', trigger: 'MS16A', cjActionable: true, tierLabel: '70% — MS16A (CX Complete, old split)', isDecomScop: false }
  }
  if (s.includes('20%') && !s.includes('init')) {
    return { tier: '20', trigger: 'DECOM_SCOP', cjActionable: true, tierLabel: '20% — Decom + SCOP Complete', isDecomScop: true }
  }
  if (s.includes('30%')) {
    return { tier: '30', trigger: 'DECOM_SCOP', cjActionable: true, tierLabel: '30% — Decom + SCOP Complete (old split)', isDecomScop: true }
  }
  return { tier: null, trigger: null, cjActionable: false, tierLabel: sogNameRaw || 'Unknown', isDecomScop: false }
}

// Parses the raw tracker snapshot rows into a lookup of HOP -> actual dates + Nokia PM.
// Mirrors the header-detection pattern used by every other tracker-driven page.
export function buildTrackerDateMap(rows: unknown[][]): Map<string, TrackerDates> {
  const map = new Map<string, TrackerDates>()
  let headerRow = -1
  for (let i = 0; i < 10; i++) {
    if ((rows[i] as unknown[])?.some(c => String(c).trim() === 'HOP')) { headerRow = i; break }
  }
  if (headerRow === -1) return map

  const headers = rows[headerRow] as string[]
  const col = (name: string) => headers.findIndex(h => String(h).trim() === name)
  const hopCol = col('HOP')
  const don444Col = col('DON 444')
  const ms15aCol = col('MS15 Implementation Start A')
  const ms16aCol = col('MS16 Implementation Ends A')
  const nokiaPmCol = col('Nokia PM')

  for (let i = headerRow + 1; i < rows.length; i++) {
    const row = rows[i] as unknown[]
    const don = String(row[don444Col] || '').trim().toUpperCase()
    if (don !== 'DON 444') continue
    const hop = String(row[hopCol] || '').trim()
    if (!hop || hop === 'undefined') continue
    // Tracker HOP names use '<>' as a separator; SPO report Name uses 'to'.
    // Normalize here so the key matches the SPO row's raw Name column.
    const key = matchKey(normalizeTrackerHop(hop))
    if (map.has(key)) continue
    map.set(key, {
      ms15a: parseDateAny(row[ms15aCol]),
      ms16a: parseDateAny(row[ms16aCol]),
      nokiaPm: String(row[nokiaPmCol] || '').trim(),
    })
  }
  return map
}

// Base PO trigger gate. CR rows never call this — they have no trigger gate (see buildGrRows).
function isTriggerMet(trigger: GrTrigger, dates: TrackerDates | undefined, decomScopStatus: DecomScopStatus | undefined): boolean {
  if (!trigger) return false
  if (trigger === 'DECOM_SCOP') return !!decomScopStatus?.decomComplete && !!decomScopStatus?.scopComplete
  if (!dates) return false
  if (trigger === 'MS15A') return !!dates.ms15a
  if (trigger === 'MS16A') return !!dates.ms16a
  return false
}

export interface DecomScopStatus {
  decomComplete: boolean
  scopComplete: boolean
}

// Per-HOP Decom/SCOP completion status for the 20%/30% GR tiers, keyed the
// same way buildTrackerDateMap keys tracker HOPs (normalized to the SPO
// report's 'to'-separated HOP format). Kept as two separate flags (not one
// merged boolean) so pendingReasonFor() below can say which one is actually
// still blocking. A HOP counts as decom-complete only when EVERY physical
// site row for it (a HOP can span two sites) is 'complete'; scop completeness
// reads ScopRow.fullyComplete directly (one row per HOP already). A HOP
// absent from either tracker defaults to false — "not tracked yet" isn't
// the same as "complete".
export function buildDecomScopCompleteMap(
  decomRawRows: unknown[][],
  scopRawRows: unknown[][],
  scopSettings: ScopSettings = DEFAULT_SCOP,
): Map<string, DecomScopStatus> {
  const decomByHop = new Map<string, boolean>()
  parseDecomRows(decomRawRows).forEach(r => {
    if (!r.hop) return
    const key = matchKey(normalizeTrackerHop(r.hop))
    const isComplete = r.status === 'complete'
    const prev = decomByHop.get(key)
    decomByHop.set(key, prev === undefined ? isComplete : prev && isComplete)
  })

  const scopByHop = new Map<string, boolean>()
  if (scopRawRows.length >= 2) {
    const calc = resolveScopCalcSettings(scopSettings)
    buildScopDataset(keyScopRows(scopRawRows), calc).forEach(r => {
      if (!r.hop) return
      scopByHop.set(matchKey(normalizeTrackerHop(r.hop)), r.fullyComplete)
    })
  }

  const out = new Map<string, DecomScopStatus>()
  new Set([...decomByHop.keys(), ...scopByHop.keys()]).forEach(key => {
    out.set(key, {
      decomComplete: decomByHop.get(key) ?? false,
      scopComplete: scopByHop.get(key) ?? false,
    })
  })
  return out
}

// "Why is this GR pending" — only meaningful while status is "Awaiting
// Trigger" ('Not Yet' in the UI), i.e. something is genuinely blocking it.
// Once the trigger is met, status flips to "Ready to Release" ('Pending' in
// the UI) — that just means it's sitting in CJ's own queue to submit, not
// blocked on anything else, so no reason applies there (or once GR'd).
// CR rows have no trigger gate at all (see buildGrRows) so they never reach
// "Awaiting Trigger" — this is never called for them in practice.
function pendingReasonFor(tier: GrTier, decomScopStatus: DecomScopStatus | undefined): string {
  switch (tier) {
    case 'init20': return 'Pending Start/MSS Install Confirmation'
    case '60':
    case '70': return 'Pending CX Completion'
    case '20':
    case '30': {
      const decomDone = decomScopStatus?.decomComplete ?? false
      const scopDone = decomScopStatus?.scopComplete ?? false
      if (!decomDone && !scopDone) return 'Pending Decom + SCOP Complete'
      return !decomDone ? 'Pending Decom Complete' : 'Pending SCOP Complete'
    }
    default: return 'Pending — Unclassified SOG Tier'
  }
}

// Classifies a raw SPO report row into Base PO / CR / Error+empty (excluded).
// Must run before any other row processing — this is the foundational filter.
function classifyRowType(sogNameRaw: string, spoNumberRaw: string): GrRowType | null {
  const hasSpo = !!spoNumberRaw.trim()
  const hasSog = !!sogNameRaw.trim()
  if (hasSpo && hasSog) return 'base'
  if (hasSpo && !hasSog) return 'cr'
  return null // Error row (SOG present, no SPO) or empty row — excluded entirely
}

function computeStatus(grDate: Date | null, isTriggerMet: boolean): GrStatus {
  if (grDate) return 'GR Done'
  if (isTriggerMet) return 'Ready to Release'
  return 'Awaiting Trigger'
}

// Joins SPO report rows against the tracker date map to build the unified GR row set.
// decomScopCompleteMap gates the 20%/30% (DECOM_SCOP trigger) tiers — see
// buildDecomScopCompleteMap. Defaults to empty so existing callers that don't
// pass one just get those tiers always "not yet met" instead of erroring.
export function buildGrRows(
  spoRows: unknown[][],
  trackerDateMap: Map<string, TrackerDates>,
  decomScopCompleteMap: Map<string, DecomScopStatus> = new Map(),
): GrRow[] {
  const out: GrRow[] = []
  spoRows.forEach(row => {
    const nameRaw = String(row[SPO_COL.name] || '').trim()
    if (!nameRaw) return

    const sogNameRaw = String(row[SPO_COL.sogName] || '').trim()
    const spoNumberRaw = String(row[SPO_COL.spoNumber] || '').trim()
    const rowType = classifyRowType(sogNameRaw, spoNumberRaw)
    if (!rowType) return // Error row (SOG, no SPO) or empty row — excluded entirely

    // CR rows have no SOG tier to classify, so they carry no trigger gate —
    // treated as always eligible for release until GR'd (business decision).
    // "CR" is itself a formal tier so these rows participate in tier filtering.
    const sog = rowType === 'cr'
      ? { tier: 'CR' as GrTier, trigger: null as GrTrigger, cjActionable: true, tierLabel: 'CR', isDecomScop: false }
      : classifySog(sogNameRaw)

    const vendor = row[SPO_COL.spoVendor]
    const gcEntry = GC_CONFIG.find(cfg => matches(vendor, cfg.spo_match))
    const gc = gcEntry ? gcEntry.gc : String(vendor || '').trim()

    // SPO report Name column already uses 'to' — the tracker-side map was
    // normalized when it was built, so we match against the raw SPO name here.
    const key = matchKey(nameRaw)
    const dates = trackerDateMap.get(key)
    const decomScopStatus = decomScopCompleteMap.get(key)
    const triggerMetFlag = rowType === 'cr' ? true : isTriggerMet(sog.trigger, dates, decomScopStatus)

    const grDateRaw = parseDateAny(row[SPO_COL.grDate])
    const status = computeStatus(grDateRaw, triggerMetFlag)
    // trackerDateMap only holds HOPs currently on the DON 444 list (see
    // buildTrackerDateMap) — a HOP with no entry there wasn't dropped, it was
    // pulled from the build plan entirely, which outranks any tier-specific
    // "what's blocking it" reason.
    const pendingReason = status !== 'Awaiting Trigger' ? ''
      : !dates ? 'Removed from Build Plan'
      : pendingReasonFor(sog.tier, decomScopStatus)

    // DECOM_SCOP has no single tracker date to show — it's gated on a status
    // (both trackers complete), not a milestone date — so triggerDate stays
    // blank for that tier same as an unmet MS15A/MS16A trigger would.
    const triggerDateRaw = sog.trigger === 'MS15A' ? (dates?.ms15a ?? null)
      : sog.trigger === 'MS16A' ? (dates?.ms16a ?? null)
      : null

    const rawVal = row[SPO_COL.spoValue]
    const spoValue = typeof rawVal === 'number' ? rawVal : (parseFloat(String(rawVal ?? '0').replace(/[$,]/g, '')) || 0)

    out.push({
      hop: key,
      hopDisplay: nameRaw,
      pathId: String(row[SPO_COL.customerSiteId] || '').trim(),
      gc,
      nokiaPm: dates?.nokiaPm || '',
      sogName: sogNameRaw,
      spoNumber: spoNumberRaw,
      spoValue,
      rowType,
      triggerDate: fmtDate(triggerDateRaw),
      triggerDateRaw,
      trigger: sog.trigger,
      triggerMet: triggerMetFlag,
      grDate: fmtDate(grDateRaw),
      grDateRaw,
      status,
      tier: sog.tier,
      tierLabel: sog.tierLabel,
      cjActionable: sog.cjActionable,
      isDecomScop: sog.isDecomScop,
      pendingReason,
    })
  })
  return out
}

// Sort dropdown shared by the GR Tracker page and Dashboard GR expand panel.
export type GrSortOption = 'trigger' | 'hop' | 'sog' | 'value' | 'status'

export const GR_SORT_OPTIONS: { value: GrSortOption; label: string }[] = [
  { value: 'trigger', label: 'Trigger Date (oldest first)' },
  { value: 'hop', label: 'HOP Name (A→Z)' },
  { value: 'sog', label: 'SOG Name' },
  { value: 'value', label: 'SPO Value (largest first)' },
  { value: 'status', label: 'GR Status' },
]

export function sortGrRowsBy(rows: GrRow[], sortBy: GrSortOption): GrRow[] {
  const statusRank = (r: GrRow) => r.status === 'Ready to Release' ? 0 : r.status === 'Awaiting Trigger' ? 1 : 2
  const arr = [...rows]
  switch (sortBy) {
    case 'hop':
      return arr.sort((a, b) => a.hopDisplay.localeCompare(b.hopDisplay) || a.sogName.localeCompare(b.sogName))
    case 'sog':
      return arr.sort((a, b) => a.sogName.localeCompare(b.sogName) || a.hopDisplay.localeCompare(b.hopDisplay))
    case 'value':
      return arr.sort((a, b) => b.spoValue - a.spoValue)
    case 'status':
      return arr.sort((a, b) => statusRank(a) - statusRank(b))
    case 'trigger':
    default: {
      const t = (r: GrRow) => r.triggerDateRaw ? r.triggerDateRaw.getTime() : Infinity
      return arr.sort((a, b) => t(a) - t(b))
    }
  }
}

// Base PO vs CR volume/value breakdown that drives the KPI tiles.
export interface GrBreakdown {
  totalBasePOs: number
  totalCRs: number
  basePOValueGRd: number
  basePOValuePending: number
  basePOValueNotYet: number
  crValueGRd: number
  crValuePending: number
  crValueNotYet: number
}

export function computeGrBreakdown(rows: GrRow[]): GrBreakdown {
  const base = rows.filter(r => r.rowType === 'base')
  const cr = rows.filter(r => r.rowType === 'cr')
  return {
    totalBasePOs: base.length,
    totalCRs: cr.length,
    basePOValueGRd: base.filter(r => !!r.grDate).reduce((s, r) => s + r.spoValue, 0),
    basePOValuePending: base.filter(r => !r.grDate && r.triggerMet).reduce((s, r) => s + r.spoValue, 0),
    basePOValueNotYet: base.filter(r => !r.grDate && !r.triggerMet).reduce((s, r) => s + r.spoValue, 0),
    crValueGRd: cr.filter(r => !!r.grDate).reduce((s, r) => s + r.spoValue, 0),
    crValuePending: cr.filter(r => !r.grDate && r.triggerMet).reduce((s, r) => s + r.spoValue, 0),
    crValueNotYet: cr.filter(r => !r.grDate && !r.triggerMet).reduce((s, r) => s + r.spoValue, 0),
  }
}

export type GrTileFilter = 'totalBasePOs' | 'totalCRs' | 'basePOGRd' | 'basePOPending' | 'basePONotYet' | 'crGRd' | 'crPending' | 'crNotYet' | null

export function rowsForTileFilter(rows: GrRow[], filter: GrTileFilter): GrRow[] {
  switch (filter) {
    case 'totalBasePOs': return rows.filter(r => r.rowType === 'base')
    case 'totalCRs': return rows.filter(r => r.rowType === 'cr')
    case 'basePOGRd': return rows.filter(r => r.rowType === 'base' && !!r.grDate)
    case 'basePOPending': return rows.filter(r => r.rowType === 'base' && !r.grDate && r.triggerMet)
    case 'basePONotYet': return rows.filter(r => r.rowType === 'base' && !r.grDate && !r.triggerMet)
    case 'crGRd': return rows.filter(r => r.rowType === 'cr' && !!r.grDate)
    case 'crPending': return rows.filter(r => r.rowType === 'cr' && !r.grDate && r.triggerMet)
    case 'crNotYet': return rows.filter(r => r.rowType === 'cr' && !r.grDate && !r.triggerMet)
    default: return rows
  }
}

export interface GrGroups {
  ready: GrRow[]
  awaiting: GrRow[]
  done: GrRow[]
  awareness: GrRow[]
}

export function groupGrRows(rows: GrRow[]): GrGroups {
  return {
    ready: rows.filter(r => r.cjActionable && r.status === 'Ready to Release'),
    awaiting: rows.filter(r => r.cjActionable && r.status === 'Awaiting Trigger'),
    done: rows.filter(r => r.status === 'GR Done'),
    // 20%/30% (Decom/SCOP) tiers are now CJ-actionable (DECOM_SCOP trigger),
    // so "awareness" is no longer about tier — it's whatever tier didn't
    // classify at all (classifySog's null fallback), shown but not acted on.
    awareness: rows.filter(r => !r.cjActionable),
  }
}

// Loads the SPO report + latest tracker snapshot from Supabase and joins them.
export async function loadGrRows(): Promise<GrRow[]> {
  const [spoResult, trackerSnap, decomReport, scopReport] = await Promise.all([
    supabase.from('report_snapshots').select('data').eq('id', 'spo').single(),
    loadTrackerSnapshot(),
    loadChunkedReport('decom'),
    loadChunkedReport('scop'),
  ])
  const spoRows: unknown[][] = spoResult.data?.data ? JSON.parse(spoResult.data.data) : []
  const trackerDateMap = trackerSnap ? buildTrackerDateMap(trackerSnap.data) : new Map<string, TrackerDates>()
  const decomScopCompleteMap = buildDecomScopCompleteMap(decomReport?.rows ?? [], scopReport?.rows ?? [])
  return buildGrRows(spoRows, trackerDateMap, decomScopCompleteMap)
}

// ─────────────────────────────────────────────
// EMAIL GENERATION (Part 5 spec — shared across all 3 surfaces)
// ─────────────────────────────────────────────

export const GR_EMAIL_TO = ['carlos.2.molina.ext@nokia.com', 'deepak.maruti.ext@nokia.com']
export const GR_EMAIL_CC_BASE = [
  'thomas.meinke.ext@nokia.com',
  'steve.jahr.ext@nokia.com',
  'christopher.seebach@nokia.com',
  'george.anson@nokia.com',
  'curtiss.lindsey.ext@nokia.com',
  'emily.rudolph@nokia.com',
  'scott.tomlinson.ext@nokia.com',
  'paul.1.barlow.ext@nokia.com',
]

// GC primary contact emails are not yet known — fill in as they're confirmed.
// Any GC without an entry here simply omits the extra CC line (never blocks the email).
export const GC_PRIMARY_CONTACT: Record<string, string> = {}

export function buildGrEmailMailto(
  gc: string,
  rows: GrRow[],
  emailOverrides?: { financeEmails?: string[]; gcContactEmails?: Record<string, string> }
): string {
  const today = new Date()
  const dateStr = `${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear()}`
  const subject = `GR Release Request — ${gc} — ${dateStr}`
  const div = '═'.repeat(40)

  // Group by HOP first, then list every SOG/CR line underneath its HOP.
  const byHop = new Map<string, { pathId: string; hopDisplay: string; rows: GrRow[] }>()
  rows.forEach(r => {
    if (!byHop.has(r.hop)) byHop.set(r.hop, { pathId: r.pathId, hopDisplay: r.hopDisplay, rows: [] })
    byHop.get(r.hop)!.rows.push(r)
  })
  const hopGroups = Array.from(byHop.values()).sort((a, b) => a.hopDisplay.localeCompare(b.hopDisplay))

  let body = `Please see below for GR release requests ready for processing.\n\n${div}\n\n`
  let total = 0
  hopGroups.forEach(g => {
    body += `★ ${g.hopDisplay} ★  |  Path ID: ${g.pathId}\n`
    g.rows.forEach(r => {
      total += r.spoValue
      body += `  • ${r.sogName || 'CR'} — SPO #: ${r.spoNumber || '—'} — Value: ${fmtMoney(r.spoValue)} — Trigger: ${r.triggerDate || '—'}\n`
    })
    body += `\n`
  })
  body += `${div}\n`
  body += `Total Pending Release: ${fmtMoney(total)}\n\n`
  body += `Thank you,\nCJ`

  const gcContactMap = emailOverrides?.gcContactEmails ?? GC_PRIMARY_CONTACT
  const financeEmails = emailOverrides?.financeEmails ?? GR_EMAIL_CC_BASE
  const gcContact = lookupContactEmail(gcContactMap, gc)
  const cc = [...(gcContact ? [gcContact] : []), ...financeEmails].join(',')
  const to = GR_EMAIL_TO.join(',')

  return `mailto:${to}?cc=${encodeURIComponent(cc)}&subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`
}
