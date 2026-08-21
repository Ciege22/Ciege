import { supabase } from './supabase'

const CHUNK_SIZE = 500

export interface ChunkedReportMeta {
  filename: string
  uploaded_at: string
  chunkCount: number
}

export interface ChunkedReportResult {
  filename: string
  uploaded_at: string
  rows: unknown[][]
}

// Splits `rows` into CHUNK_SIZE-row records (id = `${prefix}_chunk_0`, `_chunk_1`, ...)
// plus a `${prefix}_meta` record pointing at how many chunks exist, so a single
// large report_snapshots row never has to hold a payload big enough to trip a
// Supabase/Cloudflare size or timeout limit. Any chunk left over from a previous,
// larger upload (fewer chunks this time) is deleted so stale rows don't linger.
export async function saveChunkedReport(
  prefix: string,
  filename: string,
  rows: unknown[][]
): Promise<{ error: unknown }> {
  const uploaded_at = new Date().toISOString()
  const chunks: unknown[][][] = []
  for (let i = 0; i < rows.length; i += CHUNK_SIZE) {
    chunks.push(rows.slice(i, i + CHUNK_SIZE))
  }
  // Always write at least one (possibly empty) chunk so a zero-row upload still
  // produces a consistent meta + chunk_0 pair rather than a dangling meta record.
  if (chunks.length === 0) chunks.push([])

  const { data: prevMetaRow } = await supabase
    .from('report_snapshots')
    .select('data')
    .eq('id', `${prefix}_meta`)
    .single()
  const prevChunkCount: number = prevMetaRow?.data ? (JSON.parse(prevMetaRow.data) as ChunkedReportMeta).chunkCount : 0

  for (let i = 0; i < chunks.length; i++) {
    const { error } = await supabase.from('report_snapshots').upsert({
      id: `${prefix}_chunk_${i}`,
      filename,
      uploaded_at,
      data: JSON.stringify(chunks[i]),
    })
    if (error) return { error }
  }

  if (prevChunkCount > chunks.length) {
    const staleIds = Array.from({ length: prevChunkCount - chunks.length }, (_, i) => `${prefix}_chunk_${chunks.length + i}`)
    await supabase.from('report_snapshots').delete().in('id', staleIds)
  }

  const meta: ChunkedReportMeta = { filename, uploaded_at, chunkCount: chunks.length }
  const { error: metaError } = await supabase.from('report_snapshots').upsert({
    id: `${prefix}_meta`,
    filename,
    uploaded_at,
    data: JSON.stringify(meta),
  })
  if (metaError) return { error: metaError }

  return { error: null }
}

export async function loadChunkedReport(prefix: string): Promise<ChunkedReportResult | null> {
  const { data: metaRow } = await supabase
    .from('report_snapshots')
    .select('data')
    .eq('id', `${prefix}_meta`)
    .single()
  if (!metaRow?.data) return null

  const meta = JSON.parse(metaRow.data) as ChunkedReportMeta

  const chunkResults = await Promise.all(
    Array.from({ length: meta.chunkCount }, (_, i) =>
      supabase.from('report_snapshots').select('data').eq('id', `${prefix}_chunk_${i}`).single()
    )
  )

  const rows: unknown[][] = []
  for (const r of chunkResults) {
    if (r.data?.data) rows.push(...(JSON.parse(r.data.data) as unknown[][]))
  }

  return { filename: meta.filename, uploaded_at: meta.uploaded_at, rows }
}
