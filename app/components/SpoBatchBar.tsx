'use client'

import { ReactNode, useEffect, useMemo, useState } from 'react'
import { loadTrackerSnapshot } from '../lib/supabase'
import { loadEmailSettings, EmailRouting } from '../lib/settings'
import {
  buildSpoBucket, HopSpo, SpoDraft, loadSpoDraft, rowsFor, openSpoEmail, spoRecipients,
  downloadSpoExcel, stageSpoRequested,
} from '../lib/spoRequest'

// Email / Excel / stamp actions for an SPO request batch, with a contractor
// filter so one GC's pricing can be worked at a time. The caller supplies the
// bucket and draft; the children render prop receives the GC-filtered list so
// the page's cards and the actions always cover the same HOPs.
export default function SpoBatchBar({ bucket, draft, children }: {
  bucket: HopSpo[]
  draft: SpoDraft
  children?: (scoped: HopSpo[]) => ReactNode
}) {
  const [gc, setGc] = useState<string>('ALL')
  const [status, setStatus] = useState<string | null>(null)
  const [routing, setRouting] = useState<Record<string, EmailRouting> | undefined>(undefined)

  useEffect(() => {
    loadEmailSettings().then(e => setRouting(e.routing)).catch(() => setRouting(undefined))
  }, [])
  const recipients = spoRecipients(routing)

  const gcs = useMemo(() => {
    const counts = new Map<string, number>()
    bucket.forEach(h => { const k = h.gc || '(no GC)'; counts.set(k, (counts.get(k) ?? 0) + 1) })
    return Array.from(counts.entries()).sort((a, b) => a[0].localeCompare(b[0]))
  }, [bucket])

  const scoped = useMemo(
    () => (gc === 'ALL' ? bucket : bucket.filter(h => (h.gc || '(no GC)') === gc)),
    [bucket, gc],
  )
  const missingValues = scoped.reduce((n, h) => n + rowsFor(draft, h.hop).filter(r => !r.value.trim()).length, 0)

  const onStamp = async () => {
    setStatus('Staging CX SPO Request dates…')
    try {
      const n = await stageSpoRequested(scoped)
      setStatus(`Staged CX SPO Request = today on ${n} HOP${n === 1 ? '' : 's'}. Review on the Pending Updates list.`)
    } catch (err) {
      setStatus(err instanceof Error ? `Failed: ${err.message}` : 'Failed to stage updates.')
    }
  }

  return (
    <div>
      <div className="bg-zinc-900 border border-white/10 rounded-xl p-3">
        <div className="flex flex-wrap items-center gap-3">
          <select value={gc} onChange={e => setGc(e.target.value)}
            className="bg-zinc-800 text-sm rounded px-3 py-2 border border-white/10">
            <option value="ALL">All contractors ({bucket.length})</option>
            {gcs.map(([name, n]) => <option key={name} value={name}>{name} ({n})</option>)}
          </select>
          <div className="flex-1 min-w-[200px] text-xs text-zinc-400">
            <div>To: <span className="text-zinc-200">{recipients.to || 'not set'}</span></div>
            <div>CC: <span className="text-zinc-200">{recipients.cc || 'none'}</span></div>
            <div className="text-zinc-500">Set in Settings → Email Routing → SPO Requests</div>
          </div>
          <button onClick={() => openSpoEmail(scoped, draft, routing)} disabled={scoped.length === 0}
            className="bg-blue-600 hover:bg-blue-500 disabled:opacity-40 text-white text-sm font-semibold rounded px-4 py-2">
            ✉️ Email SPO requests
          </button>
          <button onClick={() => downloadSpoExcel(scoped, draft)} disabled={scoped.length === 0}
            className="bg-emerald-700 hover:bg-emerald-600 disabled:opacity-40 text-white text-sm font-semibold rounded px-4 py-2">
            ⬇️ Download Excel
          </button>
          <button onClick={onStamp} disabled={scoped.length === 0}
            className="bg-amber-700 hover:bg-amber-600 disabled:opacity-40 text-white text-sm font-semibold rounded px-4 py-2">
            📌 Stamp CX SPO Request = today
          </button>
        </div>
        <div className="mt-2 text-xs text-zinc-400">
          {scoped.length} HOP{scoped.length === 1 ? '' : 's'} · {scoped.length * 3} payment rows
          {missingValues > 0 && <span className="text-amber-400"> · {missingValues} SPO value{missingValues === 1 ? '' : 's'} still blank</span>}
        </div>
        {status && <div className="mt-2 text-xs text-zinc-300">{status}</div>}
      </div>
      {children && <div className="mt-4">{children(scoped)}</div>}
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
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [hopNames.join('|')])

  if (bucket === null) return <div className="text-xs text-zinc-500">Loading SPO batch…</div>
  return <SpoBatchBar bucket={bucket} draft={draft} />
}
