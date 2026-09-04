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

// Still date-keyed so the list resets to a fresh slate each calendar day —
// matches the reset behavior every view's own list had individually.
function storageId(): string {
  return `pending-updates-${new Date().toISOString().slice(0, 10)}`
}

export async function loadPendingUpdates(): Promise<PendingUpdate[]> {
  const { data } = await supabase
    .from('pm_updates_cache')
    .select('updates')
    .eq('id', storageId())
    .single()
  if (!data?.updates) return []
  try {
    return JSON.parse(data.updates)
  } catch {
    return []
  }
}

export async function persistPendingUpdates(updates: PendingUpdate[]): Promise<void> {
  await supabase.from('pm_updates_cache').upsert({
    id: storageId(),
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
