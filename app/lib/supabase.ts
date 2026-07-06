import { createClient } from '@supabase/supabase-js'

const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL!
const supabaseKey = process.env.NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY!

export const supabase = createClient(supabaseUrl, supabaseKey)

export async function saveTrackerSnapshot(filename: string, hopCount: number, data: unknown[][]) {
  // Delete snapshots older than 4 weeks
  const fourWeeksAgo = new Date()
  fourWeeksAgo.setDate(fourWeeksAgo.getDate() - 28)
  await supabase
    .from('tracker_snapshot')
    .delete()
    .lt('uploaded_at', fourWeeksAgo.toISOString())

  // Save new snapshot
  const { data: newSnap, error } = await supabase
    .from('tracker_snapshot')
    .insert({ filename, hop_count: hopCount, data: JSON.stringify(data) })
    .select()
    .single()

  if (error) { console.error('Error saving snapshot:', error); return null }
  return newSnap
}

export async function getPreviousSnapshot(): Promise<{ id: string; filename: string; uploaded_at: string; data: unknown[][] } | null> {
  const { data, error } = await supabase
    .from('tracker_snapshot')
    .select('*')
    .order('uploaded_at', { ascending: false })
    .range(1, 1)
    .single()
  if (error || !data) return null
  return { ...data, data: JSON.parse(data.data) }
}

export async function getAllSnapshots(): Promise<{ id: string; filename: string; uploaded_at: string; hop_count: number }[]> {
  const { data, error } = await supabase
    .from('tracker_snapshot')
    .select('id, filename, uploaded_at, hop_count')
    .order('uploaded_at', { ascending: false })
  if (error || !data) return []
  return data
}

export async function saveTrackerChanges(uploadId: string, changes: unknown[]) {
  if (changes.length === 0) return
  const { error } = await supabase.from('tracker_changes').insert(changes)
  if (error) console.error('Error saving changes:', error)
}

export async function saveSchemaChanges(uploadId: string, changes: unknown[]) {
  if (changes.length === 0) return
  const { error } = await supabase.from('tracker_schema_changes').insert(changes)
  if (error) console.error('Error saving schema changes:', error)
}

export async function getChangeLog(): Promise<unknown[]> {
  const { data, error } = await supabase
    .from('tracker_changes')
    .select('*')
    .order('uploaded_at', { ascending: false })
    .limit(500)
  if (error || !data) return []
  return data
}

export async function getSchemaChanges(): Promise<unknown[]> {
  const { data, error } = await supabase
    .from('tracker_schema_changes')
    .select('*')
    .order('uploaded_at', { ascending: false })
    .limit(100)
  if (error || !data) return []
  return data
}

export async function loadTrackerSnapshot(): Promise<{ filename: string; uploaded_at: string; hop_count: number; data: unknown[][] } | null> {
  const { data, error } = await supabase
    .from('tracker_snapshot')
    .select('*')
    .order('uploaded_at', { ascending: false })
    .limit(1)
    .single()
  if (error || !data) return null
  return { ...data, data: JSON.parse(data.data) }
}
