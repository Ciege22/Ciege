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
