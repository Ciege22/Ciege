import { createClient } from '@supabase/supabase-js'

const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL!
const supabaseKey = process.env.NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY!

export const supabase = createClient(supabaseUrl, supabaseKey)

export async function saveTrackerSnapshot(filename: string, hopCount: number, data: unknown[][]) {
  await supabase.from('tracker_snapshot').delete().neq('id', '00000000-0000-0000-0000-000000000000')
  const { error } = await supabase
    .from('tracker_snapshot')
    .insert({ filename, hop_count: hopCount, data: JSON.stringify(data) })
  if (error) console.error('Error saving snapshot:', error)
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
