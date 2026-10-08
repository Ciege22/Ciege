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
  buildTrackerDateMap, buildDecomScopCompleteMap, buildGrRows,
} from './grTracker'

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
}

function tierSortRank(tier: string): number {
  return ['init20', '60', '70', '20', '30', 'CR'].indexOf(tier)
}

export function buildCleanupHops(
  grRows: GrRow[],
  departedGcs: string[],
  decomScopMap: Map<string, DecomScopStatus>,
): CleanupGroups {
  const departedSet = new Set(departedGcs.map(g => g.trim().toLowerCase()))
  const byHop = new Map<string, GrRow[]>()
  grRows.forEach(r => {
    if (!departedSet.has(r.gc.trim().toLowerCase())) return
    byHop.set(r.hop, [...(byHop.get(r.hop) ?? []), r])
  })

  const groups: CleanupGroups = { not_started: [], mid_construction: [], cleanup_only: [], fullyPaidButIncomplete: [] }

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
      }))
    const status = decomScopMap.get(hop)
    const cleanupHop: CleanupHop = {
      hop,
      hopDisplay: rows[0].hopDisplay,
      pathId: rows[0].pathId,
      gc: rows[0].gc,
      bucket: 'cleanup_only', // placeholder, set below
      totalValue,
      paidValue,
      paidPct: totalValue > 0 ? Math.round((paidValue / totalValue) * 100) : 0,
      unpaidValue,
      unpaidTiers,
      decomComplete: status?.decomComplete ?? false,
      scopComplete: status?.scopComplete ?? false,
    }

    if (unpaidValue <= 0) {
      if (!cleanupHop.decomComplete || !cleanupHop.scopComplete) groups.fullyPaidButIncomplete.push(cleanupHop)
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
  const decomScopMap = buildDecomScopCompleteMap(decomReport?.rows ?? [], scopReport?.rows ?? [])
  const grRows = buildGrRows(spoRows, trackerDateMap, decomScopMap)
  return buildCleanupHops(grRows, departedGcs, decomScopMap)
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
}

export const EMPTY_ASSIGNMENT: CleanupAssignment = { newGc: '', quote: null, comment: '', status: 'Needs Quote', updatedAt: '' }

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
