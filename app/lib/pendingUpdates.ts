// Consolidated "PM Updates" list shared by Tracker, CM View, and GC Call
// View. Previously each view staged its own edits in a separate, page-only
// Supabase row (tracker-changes-{date}, cm-updates-{date}, gc-updates-{date})
// — three untracked lists that were easy to lose track of. Now all three
// pages read and write this single row, tagging each entry with which view
// it came from, so any edit made anywhere shows up on the one list.
import { supabase } from './supabase'

export type UpdateSource = 'tracker' | 'cm' | 'gc'

export interface PendingUpdate {
  source: UpdateSource
  // Tracker only — a HOP can span two physical rows (see TrackerRowData.rowKey
  // in app/tracker/page.tsx), so tracker edits dedup on rowKey+field rather
  // than hop+field. CM/GC entries have no rowKey and dedup on hop+field.
  rowKey?: string
  hop: string
  field: string
  oldValue: string
  newValue: string
  timestamp: string
  user?: string
  completed?: boolean
}

export const SOURCE_LABELS: Record<UpdateSource, string> = {
  tracker: 'Tracker',
  cm: 'CM View',
  gc: 'GC View',
}

export const SOURCE_BADGE_CLASSES: Record<UpdateSource, string> = {
  tracker: 'bg-blue-200 text-blue-900',
  cm: 'bg-purple-200 text-purple-900',
  gc: 'bg-orange-200 text-orange-900',
}

// One permanent row, not date-keyed — earlier versions reset to empty every
// calendar day, which silently dropped anything still open at midnight.
// Open items now persist until done; only completed ones age out (see
// HISTORY_DAYS below), so the list still reads as "today's stuff" without
// erasing work that spans a day boundary.
const STORAGE_ID = 'pending-updates'

// How long a completed entry stays visible (as history/audit trail) before
// it's pruned on load. Open (incomplete) entries are never pruned this way —
// they only leave the list via Clear Completed/Clear All or getting undone.
const HISTORY_DAYS = 7

function pruneOldCompleted(updates: PendingUpdate[]): PendingUpdate[] {
  const cutoff = Date.now() - HISTORY_DAYS * 24 * 60 * 60 * 1000
  return updates.filter(u => !u.completed || new Date(u.timestamp).getTime() >= cutoff)
}

export async function loadPendingUpdates(): Promise<PendingUpdate[]> {
  const { data } = await supabase
    .from('pm_updates_cache')
    .select('updates')
    .eq('id', STORAGE_ID)
    .single()
  if (!data?.updates) return []
  let parsed: PendingUpdate[]
  try {
    parsed = JSON.parse(data.updates)
  } catch {
    return []
  }
  const pruned = pruneOldCompleted(parsed)
  if (pruned.length !== parsed.length) await persistPendingUpdates(pruned)
  return pruned
}

export async function persistPendingUpdates(updates: PendingUpdate[]): Promise<void> {
  await supabase.from('pm_updates_cache').upsert({
    id: STORAGE_ID,
    updates: JSON.stringify(updates),
    updated_at: new Date().toISOString(),
  })
}

function updateKey(u: Pick<PendingUpdate, 'rowKey' | 'hop' | 'field'>): string {
  return `${u.rowKey ?? u.hop}|${u.field}`
}

// A second edit to the same field (same HOP, or same physical row for
// Tracker) replaces the pending entry instead of stacking a duplicate —
// same collapse-to-latest behavior each view had on its own list.
export function upsertPendingUpdate(prev: PendingUpdate[], next: Omit<PendingUpdate, 'completed'>): PendingUpdate[] {
  const key = updateKey(next)
  return [...prev.filter(u => updateKey(u) !== key), { ...next, completed: false }]
}
