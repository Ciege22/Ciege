'use client'

import { useEffect, useState } from 'react'
import { loadTrackerSnapshot } from '../lib/supabase'
import {
  buildSpoBucket, HopSpo, SpoDraft, loadSpoDraft, saveSpoDraft, rowsFor, openSpoEmail,
  downloadSpoExcel, stageSpoRequested,
} from '../lib/spoRequest'

// Email / Excel / stamp actions for an SPO request batch. Presentational — the
// caller supplies the bucket and draft, so the same bar serves the SPO page
// and the dashboard's Needs SPO Requested window.
export default function SpoBatchBar({ bucket, draft, onDraftChange }: {
  bucket: HopSpo[]
  draft: SpoDraft
  onDraftChange: (d: SpoDraft) => void
}) {
  const [status, setStatus] = useState<string | null>(null)
  const missingValues = bucket.reduce((n, h) => n + rowsFor(draft, h.hop).filter(r => !r.value.trim()).length, 0)

  const onStamp = async () => {
    setStatus('Staging CX SPO Request dates…')
    try {
      const n = await stageSpoRequested(bucket)
      setStatus(`Staged CX SPO Request = today on ${n} HOP${n === 1 ? '' : 's'}. Review on the Pending Updates list.`)
    } catch (err) {
      setStatus(err instanceof Error ? `Failed: ${err.message}` : 'Failed to stage updates.')
    }
  }

  return (
    <div className="bg-zinc-900 border border-white/10 rounded-xl p-3">
      <div className="flex flex-wrap items-center gap-3">
        <input
          type="email"
          placeholder="Team email address"
          value={draft.to}
          onChange={e => { const next = { ...draft, to: e.target.value }; onDraftChange(next) }}
          className="flex-1 min-w-[200px] bg-zinc-800 text-sm rounded px-3 py-2 border border-white/10"
        />
        <button onClick={() => openSpoEmail(bucket, draft)} disabled={bucket.length === 0}
          className="bg-blue-600 hover:bg-blue-500 disabled:opacity-40 text-white text-sm font-semibold rounded px-4 py-2">
          ✉️ Email SPO requests
        </button>
        <button onClick={() => downloadSpoExcel(bucket, draft)} disabled={bucket.length === 0}
          className="bg-emerald-700 hover:bg-emerald-600 disabled:opacity-40 text-white text-sm font-semibold rounded px-4 py-2">
          ⬇️ Download Excel
        </button>
        <button onClick={onStamp} disabled={bucket.length === 0}
          className="bg-amber-700 hover:bg-amber-600 disabled:opacity-40 text-white text-sm font-semibold rounded px-4 py-2">
          📌 Stamp CX SPO Request = today
        </button>
      </div>
      <div className="mt-2 text-xs text-zinc-400">
        {bucket.length} HOP{bucket.length === 1 ? '' : 's'} · {bucket.length * 3} payment rows
        {missingValues > 0 && <span className="text-amber-400"> · {missingValues} SPO value{missingValues === 1 ? '' : 's'} still blank (enter them on the SPO Requests page)</span>}
      </div>
      {status && <div className="mt-2 text-xs text-zinc-300">{status}</div>}
    </div>
  )
}

// Dashboard wrapper: loads the bucket and saved draft, and limits the batch to
// the HOPs the window is currently showing.
export function SpoBatchForHops({ hopNames }: { hopNames: string[] }) {
  const [bucket, setBucket] = useState<HopSpo[] | null>(null)
  const [draft, setDraft] = useState<SpoDraft>({ to: '', hops: {} })

  useEffect(() => {
    let cancelled = false
    Promise.all([loadTrackerSnapshot(), loadSpoDraft()]).then(([snap, d]) => {
      if (cancelled) return
      setDraft(d)
      const wanted = new Set(hopNames)
      setBucket(snap ? buildSpoBucket(snap.data).filter(h => wanted.has(h.hop)) : [])
    })
    return () => { cancelled = true }
    // hopNames identity changes each render of the caller; key on its contents
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [hopNames.join('|')])

  if (bucket === null) return <div className="text-xs text-zinc-500">Loading SPO batch…</div>
  return (
    <SpoBatchBar bucket={bucket} draft={draft} onDraftChange={d => { setDraft(d); saveSpoDraft(d) }} />
  )
}
