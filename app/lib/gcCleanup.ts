// GC Clean-Up: for GCs who are no longer working the program, which of their
// HOPs still have money sitting unpaid (unreleased GR), and what's blocking
// it — construction never started, construction unfinished, or just Decom
// and/or SCOP left to close out. That unpaid $ is what's available to pay a
// new GC to finish the HOP, so each HOP also carries a reassignment record
// (new GC, their quote, a comment, a status) saved separately from the
// tracker, since a fresh tracker upload would otherwise wipe it.
import { supabase, loadTrackerSnapshot } from './supabase'
import { loadChunkedReport } from './reportChunks'
import {
  type GrRow, type DecomScopStatus, type TrackerDates,
  buildTrackerDateMap, buildDecomScopCompleteMap, buildGrRows, normalizeTrackerHop,
} from './grTracker'
import { type HopCoord, buildHopCoordMap } from './hopCoords'
import { parseDecomRows, STATUS_DISPLAY_LABEL } from './decom'
import {
  type ScopRow, buildScopDataset, keyScopRows, resolveScopCalcSettings,
  buildItemizedMissingItems, PATHWAVE_ALL_ITEMS, QUICKBASE_ALL_ITEMS,
} from './scop'
import { DEFAULT_SCOP } from './settings'

function matchKey(s: string): string {
  return s.trim().toLowerCase()
}

// Same combined list buildGCReport's internal engineering view would use —
// every Pathwave and QuickBase item, not just the GC-owned subset, since
// this page is CJ squaring things away himself, not a GC-facing report.
// QuickBase items are prefixed: a few labels (e.g. "POD") exist in both
// lists for genuinely different columns, and without the prefix they'd
// show as an unreadable duplicate like "POD, POD".
const ALL_SCOP_ITEMS = [
  ...PATHWAVE_ALL_ITEMS,
  ...QUICKBASE_ALL_ITEMS.map(item => ({ ...item, label: `QB: ${item.label}` })),
]

export type CleanupBucket = 'not_started' | 'mid_construction' | 'cleanup_only'

export const BUCKET_LABELS: Record<CleanupBucket, string> = {
  not_started: 'Not started — nothing built yet',
  mid_construction: 'Mid-construction — build left unfinished',
  cleanup_only: 'Construction paid — Decom/SCOP cleanup only',
}

export interface UnpaidTier {
  tier: string
  tierLabel: string
  value: number
  reason: string
  // The original (departed) GC's own SPO for this tier — what Finance
  // cancels once the work is reassigned, so it isn't paid out twice.
  spoNumber: string
  // The vendor on THIS specific SPO line — not always the same as the HOP's
  // tracker GC. A HOP can carry small CR lines issued to an unrelated vendor
  // (a trenching sub, a miscellaneous change order) alongside its real
  // construction SPOs, so Finance needs the exact vendor per line, not just
  // whichever GC the HOP is attributed to overall.
  vendor: string
}

// The tracker's own General Contractor column, per HOP — the authoritative
// answer to "who is this HOP's GC," independent of which vendor happens to
// be on any one SPO line. A HOP can carry SPO rows from several vendors (a
// small CR issued to an unrelated sub, say), so picking "the GC" from an
// arbitrary SPO row was wrong — it could attribute a HOP to a GC who only
// touched one minor line item, while the tracker says someone else (often
// still-active) actually owns it. Same header/DON444-row scanning as
// buildTrackerDateMap in grTracker.ts, first-seen value wins per HOP.
export function buildHopGcMap(rows: unknown[][]): Map<string, string> {
  const out = new Map<string, string>()
  let headerRow = -1
  for (let i = 0; i < 10; i++) {
    if ((rows[i] as unknown[])?.some(c => String(c).trim() === 'HOP')) { headerRow = i; break }
  }
  if (headerRow === -1) return out
  const headers = (rows[headerRow] as unknown[]).map(h => String(h ?? '').trim())
  const col = (name: string) => headers.indexOf(name)
  const hopCol = col('HOP')
  const don444Col = col('DON 444')
  const gcCol = col('General Contractor')
  if (hopCol === -1 || gcCol === -1) return out
  for (let i = headerRow + 1; i < rows.length; i++) {
    const row = rows[i] as unknown[]
    if (!row) continue
    if (String(row[don444Col] ?? '').trim().toUpperCase() !== 'DON 444') continue
    const hop = String(row[hopCol] ?? '').trim()
    if (!hop || hop === 'undefined') continue
    const key = matchKey(normalizeTrackerHop(hop))
    if (out.has(key)) continue
    const gc = String(row[gcCol] ?? '').trim()
    if (gc) out.set(key, gc)
  }
  return out
}

export interface SiteNote {
  siteName: string
  statusLabel: string
}

// Decom is tracked per physical site (a HOP has two), so decomComplete above
// is already the HOP-level AND of both — this is the detail behind that
// boolean: which site(s) are actually still outstanding.
export function buildDecomSiteNotes(decomRawRows: unknown[][]): Map<string, SiteNote[]> {
  const out = new Map<string, SiteNote[]>()
  parseDecomRows(decomRawRows).forEach(r => {
    if (!r.hop || r.status === 'complete') return
    const key = matchKey(normalizeTrackerHop(r.hop))
    out.set(key, [...(out.get(key) ?? []), { siteName: r.siteName, statusLabel: STATUS_DISPLAY_LABEL[r.status] }])
  })
  return out
}

// SCOP is one row per HOP (not per site), but each row still has a Site A
// and Site B half of the checklist — this is which items are actually
// outstanding on each half, the same itemization buildGCReport's email uses.
export function buildScopSiteNotes(scopDataset: ScopRow[]): Map<string, SiteNote[]> {
  const out = new Map<string, SiteNote[]>()
  buildItemizedMissingItems(scopDataset.filter(r => !r.fullyComplete), ALL_SCOP_ITEMS).forEach(row => {
    const key = matchKey(normalizeTrackerHop(row.hop))
    out.set(key, [...(out.get(key) ?? []), { siteName: row.site, statusLabel: row.missingItems }])
  })
  return out
}

export interface CleanupHop {
  hop: string
  hopDisplay: string
  pathId: string
  gc: string
  bucket: CleanupBucket
  totalValue: number
  paidValue: number
  paidPct: number
  unpaidValue: number
  unpaidTiers: UnpaidTier[]
  decomComplete: boolean
  scopComplete: boolean
  // Which site(s) still need Decom, when decomComplete is false — empty if
  // the HOP isn't in the Decom tracker at all (not the same as complete).
  decomPendingSites: SiteNote[]
  // Which SCOP checklist item(s) are still outstanding, per site, when
  // scopComplete is false.
  scopPendingSites: SiteNote[]
  // null when the tracker had no valid Latt./Long. for this HOP's sites —
  // excluded from the map and from clustering, never guessed at.
  coord: HopCoord | null
  // False for fullyPaidButIncomplete (nothing unpaid to fund a new GC with)
  // and readyToRelease (Decom + SCOP already both done — nothing left for a
  // new GC to do, just a GR CJ submits himself). Drives whether a HOP's row
  // is editable (Assign/Quote) or read-only.
  needsReassignment: boolean
}

export interface CleanupGroups {
  not_started: CleanupHop[]
  mid_construction: CleanupHop[]
  cleanup_only: CleanupHop[]
  // Everything GR'd already (nothing left to reassign money for) but Decom
  // and/or SCOP still isn't complete — the original GC was paid in full and
  // nobody is left to finish the paperwork. Not a money problem, but worth
  // surfacing so it doesn't quietly fall through the cracks.
  fullyPaidButIncomplete: CleanupHop[]
  // The opposite case: Decom AND SCOP are both already complete, but the
  // unpaid tier is still sitting there because nobody's submitted the GR
  // yet. There's no clean-up work to reassign here — it only needs CJ to
  // release the payment, so these are kept out of the quoting workflow
  // entirely rather than looking like open work for a new GC.
  readyToRelease: CleanupHop[]
}

function tierSortRank(tier: string): number {
  return ['init20', '60', '70', '20', '30', 'CR'].indexOf(tier)
}

export function buildCleanupHops(
  grRows: GrRow[],
  departedGcs: string[],
  decomScopMap: Map<string, DecomScopStatus>,
  coordMap: Map<string, HopCoord> = new Map(),
  // Current DON 444 HOPs (tracker-keyed, same key space as GrRow.hop). A HOP
  // dropped from DON 444 — a cancelled/removed site — has no business on a
  // quoting list even if old unpaid SPO rows for it still exist.
  don444Hops: Set<string> | null = null,
  decomSiteNotes: Map<string, SiteNote[]> = new Map(),
  scopSiteNotes: Map<string, SiteNote[]> = new Map(),
  // HOPs with an actual MS16 (Construction Complete) date in the tracker.
  // CJ isn't ready to hold a GC accountable for Decom/SCOP on a HOP that
  // isn't physically built yet — construction being unpaid (the mid/not-
  // started buckets below) is a different problem from construction being
  // incomplete, and this gate is specifically about the latter.
  ms16aCompleteHops: Set<string> | null = null,
  // Tracker-authoritative GC per HOP (buildHopGcMap) — this, not any single
  // SPO row's vendor, decides whether a HOP belongs to a departed GC. A HOP
  // whose tracker GC is still active is excluded even if one of its SPO
  // lines happens to be with a departed GC.
  hopGcMap: Map<string, string> | null = null,
): CleanupGroups {
  const departedSet = new Set(departedGcs.map(g => g.trim().toLowerCase()))
  const byHop = new Map<string, GrRow[]>()
  grRows.forEach(r => {
    const trackerGc = hopGcMap?.get(r.hop)
    const isDeparted = trackerGc ? departedSet.has(trackerGc.trim().toLowerCase()) : departedSet.has(r.gc.trim().toLowerCase())
    if (!isDeparted) return
    if (don444Hops && !don444Hops.has(r.hop)) return
    if (ms16aCompleteHops && !ms16aCompleteHops.has(r.hop)) return
    byHop.set(r.hop, [...(byHop.get(r.hop) ?? []), r])
  })

  const groups: CleanupGroups = { not_started: [], mid_construction: [], cleanup_only: [], fullyPaidButIncomplete: [], readyToRelease: [] }

  byHop.forEach((rows, hop) => {
    const totalValue = rows.reduce((s, r) => s + r.spoValue, 0)
    const paidRows = rows.filter(r => !!r.grDate)
    const unpaidRows = rows.filter(r => !r.grDate)
    const paidValue = paidRows.reduce((s, r) => s + r.spoValue, 0)
    const unpaidValue = unpaidRows.reduce((s, r) => s + r.spoValue, 0)
    const unpaidTiers: UnpaidTier[] = unpaidRows
      .slice()
      .sort((a, b) => tierSortRank(a.tier || '') - tierSortRank(b.tier || ''))
      .map(r => ({
        tier: r.tier || '',
        tierLabel: r.tierLabel,
        value: r.spoValue,
        reason: r.pendingReason || (r.status === 'Ready to Release' ? 'Ready to release — not yet submitted' : ''),
        spoNumber: r.spoNumber,
        vendor: r.gc,
      }))
    const status = decomScopMap.get(hop)
    const cleanupHop: CleanupHop = {
      hop,
      hopDisplay: rows[0].hopDisplay,
      pathId: rows[0].pathId,
      gc: hopGcMap?.get(hop) || rows[0].gc,
      bucket: 'cleanup_only', // placeholder, set below
      totalValue,
      paidValue,
      paidPct: totalValue > 0 ? Math.round((paidValue / totalValue) * 100) : 0,
      unpaidValue,
      unpaidTiers,
      decomComplete: status?.decomComplete ?? false,
      scopComplete: status?.scopComplete ?? false,
      decomPendingSites: decomSiteNotes.get(hop) ?? [],
      scopPendingSites: scopSiteNotes.get(hop) ?? [],
      coord: coordMap.get(hop) ?? null,
      needsReassignment: true, // default; the two early-exit branches below override it
    }

    if (unpaidValue <= 0) {
      if (!cleanupHop.decomComplete || !cleanupHop.scopComplete) {
        cleanupHop.needsReassignment = false
        groups.fullyPaidButIncomplete.push(cleanupHop)
      }
      return
    }
    // Every unpaid tier is sitting there only because it hasn't been
    // submitted yet — Decom and SCOP are both already done, so there's no
    // actual clean-up work for a new GC, just a GR CJ needs to release.
    if (cleanupHop.decomComplete && cleanupHop.scopComplete && unpaidTiers.every(t => t.reason.startsWith('Ready to release'))) {
      cleanupHop.needsReassignment = false
      groups.readyToRelease.push(cleanupHop)
      return
    }
    const hasUnpaidTier = (t: string) => unpaidTiers.some(u => u.tier === t)
    cleanupHop.bucket = hasUnpaidTier('init20') ? 'not_started'
      : (hasUnpaidTier('60') || hasUnpaidTier('70')) ? 'mid_construction'
      : 'cleanup_only'
    groups[cleanupHop.bucket].push(cleanupHop)
  })

  const byValueDesc = (a: CleanupHop, b: CleanupHop) => b.unpaidValue - a.unpaidValue
  groups.not_started.sort(byValueDesc)
  groups.mid_construction.sort(byValueDesc)
  groups.cleanup_only.sort(byValueDesc)
  groups.fullyPaidButIncomplete.sort((a, b) => a.hopDisplay.localeCompare(b.hopDisplay))
  groups.readyToRelease.sort(byValueDesc)
  return groups
}

// Same four reports loadGrRows() (app/gr-tracker) joins — tracker, SPO
// report, Decom report, SCOP report — kept here instead of calling
// loadGrRows() directly because this view also needs the per-HOP Decom/SCOP
// map that function builds internally and doesn't return.
export async function loadCleanupGroups(departedGcs: string[]): Promise<CleanupGroups> {
  const [spoResult, trackerSnap, decomReport, scopReport] = await Promise.all([
    supabase.from('report_snapshots').select('data').eq('id', 'spo').single(),
    loadTrackerSnapshot(),
    loadChunkedReport('decom'),
    loadChunkedReport('scop'),
  ])
  const spoRows: unknown[][] = spoResult.data?.data ? JSON.parse(spoResult.data.data) : []
  const trackerDateMap = trackerSnap ? buildTrackerDateMap(trackerSnap.data) : new Map<string, TrackerDates>()
  // Same calc settings fed to both the complete/incomplete boolean and the
  // itemized detail below, so they can never disagree with each other.
  const calc = resolveScopCalcSettings(DEFAULT_SCOP)
  const decomScopMap = buildDecomScopCompleteMap(decomReport?.rows ?? [], scopReport?.rows ?? [], DEFAULT_SCOP)
  const coordMap = trackerSnap ? buildHopCoordMap(trackerSnap.data) : new Map<string, HopCoord>()
  const decomSiteNotes = buildDecomSiteNotes(decomReport?.rows ?? [])
  const scopDataset = scopReport && scopReport.rows.length >= 2 ? buildScopDataset(keyScopRows(scopReport.rows), calc) : []
  const scopSiteNotes = buildScopSiteNotes(scopDataset)
  const grRows = buildGrRows(spoRows, trackerDateMap, decomScopMap)
  // trackerDateMap is already DON 444-only and keyed the same way GrRow.hop
  // is (see buildTrackerDateMap in grTracker.ts), so its key set is exactly
  // "current DON 444 HOPs."
  const don444Hops = new Set(trackerDateMap.keys())
  const ms16aCompleteHops = new Set(
    Array.from(trackerDateMap.entries()).filter(([, d]) => d.ms16a !== null).map(([key]) => key)
  )
  const hopGcMap = trackerSnap ? buildHopGcMap(trackerSnap.data) : new Map<string, string>()
  return buildCleanupHops(grRows, departedGcs, decomScopMap, coordMap, don444Hops, decomSiteNotes, scopSiteNotes, ms16aCompleteHops, hopGcMap)
}

// ─────────────────────────────────────────────
// REASSIGNMENT — saved separately from the tracker (a fresh tracker upload
// would otherwise wipe any GC/notes typed directly into it). One row holds
// every HOP's assignment, the same pm_updates_cache KV pattern used
// elsewhere in this app.
// ─────────────────────────────────────────────

export const CLEANUP_STATUS_OPTIONS = ['Needs Quote', 'Quoted', 'Assigned', 'Complete'] as const
export type CleanupStatus = typeof CLEANUP_STATUS_OPTIONS[number]

export interface CleanupAssignment {
  newGc: string
  quote: number | null
  comment: string
  status: CleanupStatus
  updatedAt: string
  // '' until the quote-confirmation email to the new GC / the SPO-cancellation
  // email to Finance has actually been sent — stamped by gcCleanupEmail.ts
  // right after the mailto opens, so these read as a sent log to follow up
  // against, not just a form field.
  gcEmailSentAt: string
  financeEmailSentAt: string
}

export const EMPTY_ASSIGNMENT: CleanupAssignment = {
  newGc: '', quote: null, comment: '', status: 'Needs Quote', updatedAt: '',
  gcEmailSentAt: '', financeEmailSentAt: '',
}

const ASSIGNMENTS_ID = 'gc-cleanup-assignments'

export async function loadCleanupAssignments(): Promise<Record<string, CleanupAssignment>> {
  const { data } = await supabase.from('pm_updates_cache').select('updates').eq('id', ASSIGNMENTS_ID).single()
  if (!data?.updates) return {}
  try {
    const parsed = JSON.parse(data.updates)
    return parsed && typeof parsed === 'object' ? parsed : {}
  } catch {
    return {}
  }
}

// Read-modify-write under a fresh read, then reads the row back and checks
// the edit landed before resolving — the NTP comments tab lost a comment
// this same way (two overlapping saves on one KV row), so this mirrors the
// fix built there: nothing reports success until it's confirmed on the server.
export async function saveCleanupAssignment(hop: string, patch: Partial<CleanupAssignment>): Promise<CleanupAssignment> {
  const current = await loadCleanupAssignments()
  const next: CleanupAssignment = { ...EMPTY_ASSIGNMENT, ...current[hop], ...patch, updatedAt: new Date().toISOString() }
  const merged = { ...current, [hop]: next }
  const { error } = await supabase.from('pm_updates_cache').upsert({
    id: ASSIGNMENTS_ID,
    updates: JSON.stringify(merged),
    updated_at: new Date().toISOString(),
  })
  if (error) throw new Error(error.message)
  const landed = await loadCleanupAssignments()
  const got = landed[hop]
  if (!got || got.updatedAt !== next.updatedAt) throw new Error('Save did not persist — please try again')
  return got
}
