'use client'

export const dynamic = 'force-dynamic'

import { useEffect, useState } from 'react'
import { useIsAuthenticated } from '@azure/msal-react'
import { useSharePointFile } from '../components/SharePointDataProvider'
import MicrosoftLogin from '../components/MicrosoftLogin'
import SharePointSourceLine from '../components/SharePointSourceLine'
import BackToDashboard from '../components/BackToDashboard'
import SpoBatchBar from '../components/SpoBatchBar'
import {
  buildSpoBucket, SpoDraft, PaymentRow, loadSpoDraft, saveSpoDraft,
  rowsFor, cpoText, todayStr,
} from '../lib/spoRequest'

// Full SPO Request view — the same bucket and batch actions as the dashboard's
// Needs SPO Requested window, plus per-row entry of each HOP's payment split.
export default function SpoRequestPage() {
  const isAuthenticated = useIsAuthenticated()
  const trackerFile = useSharePointFile('tracker')

  const [draftLoading, setDraftLoading] = useState(true)
  const [draft, setDraft] = useState<SpoDraft>({ to: '', hops: {} })

  useEffect(() => {
    let cancelled = false
    loadSpoDraft().then(d => {
      if (!cancelled) { setDraft(d); setDraftLoading(false) }
    })
    return () => { cancelled = true }
  }, [])

  const bucket = trackerFile.rows ? buildSpoBucket(trackerFile.rows) : []
  const loading = draftLoading || (isAuthenticated && !trackerFile.rows && !trackerFile.error)
  const error = trackerFile.error

  const updateDraft = (next: SpoDraft) => { setDraft(next); saveSpoDraft(next) }

  const updateRow = (hop: string, idx: number, field: keyof PaymentRow, val: string) => {
    const current = rowsFor(draft, hop).map(r => ({ ...r }))
    current[idx][field] = val
    updateDraft({ ...draft, hops: { ...draft.hops, [hop]: current } })
  }

  if (!isAuthenticated) {
    return (
      <div className="min-h-screen bg-zinc-950 text-zinc-100 p-4">
        <BackToDashboard />
        <div className="rounded-2xl border border-white/10 bg-white/5 p-6 text-center">
          <p className="text-sm text-zinc-400 mb-3">Sign in with your Nokia account to load the tracker from SharePoint.</p>
          <div className="flex justify-center"><MicrosoftLogin /></div>
        </div>
      </div>
    )
  }
  if (loading) return <div className="p-6 text-zinc-400 text-sm">Loading SPO bucket…</div>
  if (error) return <div className="p-6 text-red-300 text-sm">{error}</div>

  return (
    <div className="min-h-screen bg-zinc-950 text-zinc-100 p-4">
      <BackToDashboard />
      <h1 className="text-2xl font-semibold text-white mb-1">SPO Requests</h1>
      <p className="text-sm text-zinc-400 mb-1">
        {bucket.length} HOP{bucket.length === 1 ? '' : 's'} with a CPO received and no SPO requested yet. Enter each payment split&apos;s SPO value, then email the batch.
      </p>
      <div className="mb-4">
        <SharePointSourceLine fileKey="tracker" />
      </div>

      <SpoBatchBar bucket={bucket} draft={draft}>
        {scoped => scoped.length === 0 ? (
          <div className="text-zinc-400 text-sm">Nothing waiting — every CPO-received HOP already has an SPO requested.</div>
        ) : (
          <div className="space-y-3">
            {scoped.map((h, idx) => (
              <div key={h.hop}>
                {(idx === 0 || (scoped[idx - 1].gc || '') !== (h.gc || '')) && (
                  <div className="text-sm font-bold text-amber-300 mt-4 mb-2">{h.gc || '(no GC)'}</div>
                )}
                <div className="rounded-xl border border-white/10 bg-zinc-900 p-4">
                  <div className="flex flex-wrap items-baseline gap-x-4 gap-y-1 mb-3">
                    <span className="text-sm font-bold text-white">HOP {h.appPathId}</span>
                    <span className="text-sm text-zinc-300">{h.hop}</span>
                    <span className="text-xs text-zinc-500">Customer Site ID {h.pathId}</span>
                    <span className="text-xs text-zinc-500">CPO# {cpoText(h) || '—'}</span>
                    <span className="text-xs text-zinc-500">{h.vendor}</span>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full text-xs">
                      <thead>
                        <tr className="text-zinc-500 text-left">
                          <th className="py-1 pr-3">SPO type</th>
                          <th className="py-1 pr-3">% payment</th>
                          <th className="py-1 pr-3">SPO value</th>
                          <th className="py-1 pr-3">SPO Request</th>
                        </tr>
                      </thead>
                      <tbody>
                        {rowsFor(draft, h.hop).map((r, i) => (
                          <tr key={i} className="border-t border-white/5">
                            <td className="py-1.5 pr-3 text-zinc-300">Base</td>
                            <td className="py-1.5 pr-3">
                              <input value={r.pct} onChange={e => updateRow(h.hop, i, 'pct', e.target.value)}
                                className="w-20 bg-zinc-800 rounded px-2 py-1 border border-white/10" />
                            </td>
                            <td className="py-1.5 pr-3">
                              <input value={r.value} placeholder="$0.00" onChange={e => updateRow(h.hop, i, 'value', e.target.value)}
                                className="w-32 bg-zinc-800 rounded px-2 py-1 border border-white/10" />
                            </td>
                            <td className="py-1.5 pr-3 text-zinc-300">{todayStr()}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}
      </SpoBatchBar>
    </div>
  )
}
