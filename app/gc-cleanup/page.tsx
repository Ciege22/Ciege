'use client'

export const dynamic = 'force-dynamic'

import { useEffect, useMemo, useState } from 'react'
import BackToDashboard from '../components/BackToDashboard'
import { GC_CONFIG } from '../lib/gcConfig'
import { fmtMoney } from '../lib/grTracker'
import { loadDepartedGcs, saveDepartedGcs } from '../lib/settings'
import {
  type CleanupBucket, type CleanupHop, type CleanupGroups, type CleanupAssignment,
  BUCKET_LABELS, CLEANUP_STATUS_OPTIONS,
  loadCleanupGroups, loadCleanupAssignments, saveCleanupAssignment,
} from '../lib/gcCleanup'

const BUCKET_ORDER: CleanupBucket[] = ['cleanup_only', 'mid_construction', 'not_started']

export default function GcCleanupPage() {
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)
  const [departedGcs, setDepartedGcs] = useState<string[]>([])
  const [savingGcList, setSavingGcList] = useState(false)
  const [groups, setGroups] = useState<CleanupGroups | null>(null)
  const [assignments, setAssignments] = useState<Record<string, CleanupAssignment>>({})

  async function loadAll(gcs: string[]) {
    setLoading(true)
    setError(null)
    try {
      const [g, a] = await Promise.all([loadCleanupGroups(gcs), loadCleanupAssignments()])
      setGroups(g)
      setAssignments(a)
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Could not load clean-up data.')
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => {
    loadDepartedGcs().then(gcs => { setDepartedGcs(gcs); loadAll(gcs) })
  }, [])

  async function toggleGc(gc: string) {
    const next = departedGcs.includes(gc) ? departedGcs.filter(g => g !== gc) : [...departedGcs, gc]
    setDepartedGcs(next)
    setSavingGcList(true)
    try {
      await saveDepartedGcs(next)
      await loadAll(next)
    } finally {
      setSavingGcList(false)
    }
  }

  const totals = useMemo(() => {
    if (!groups) return null
    const all = [...groups.not_started, ...groups.mid_construction, ...groups.cleanup_only]
    return {
      hops: all.length,
      unpaid: all.reduce((s, h) => s + h.unpaidValue, 0),
      incomplete: groups.fullyPaidButIncomplete.length,
    }
  }, [groups])

  return (
    <div className="min-h-screen bg-gray-950 text-white p-4 md:p-6">
      <datalist id="gc-cleanup-roster">
        {GC_CONFIG.map(cfg => <option key={cfg.gc} value={cfg.gc} />)}
      </datalist>
      <BackToDashboard />
      <h1 className="text-2xl font-bold mb-1">🧹 GC Clean-Up</h1>
      <p className="text-gray-400 text-sm mb-6">
        HOPs that belonged to a GC no longer on the program, with whatever SPO money is still unpaid on them —
        that unpaid $ is what&apos;s available to pay someone else to finish the work.
      </p>

      <div className="bg-gray-900 rounded-xl border border-gray-700 p-4 mb-6">
        <h2 className="text-sm font-semibold text-gray-300 mb-3">GCs no longer with us {savingGcList && <span className="text-gray-500">(saving…)</span>}</h2>
        <div className="flex flex-wrap gap-2">
          {GC_CONFIG.map(cfg => (
            <label key={cfg.gc} className={`flex items-center gap-2 rounded-lg px-3 py-1.5 text-sm cursor-pointer border ${departedGcs.includes(cfg.gc) ? 'bg-red-950 border-red-700 text-red-200' : 'bg-gray-800 border-gray-700 text-gray-300'}`}>
              <input type="checkbox" checked={departedGcs.includes(cfg.gc)} onChange={() => toggleGc(cfg.gc)} />
              {cfg.gc}
            </label>
          ))}
        </div>
      </div>

      {error && <p className="bg-red-950 border border-red-700 text-red-200 text-sm rounded-lg p-3 mb-4">{error}</p>}
      {loading && <p className="text-gray-400 text-sm">Loading…</p>}

      {!loading && groups && departedGcs.length === 0 && (
        <p className="text-gray-500 text-sm">Check off a departed GC above to see their outstanding HOPs.</p>
      )}

      {!loading && groups && departedGcs.length > 0 && totals && (
        <>
          <div className="grid grid-cols-2 md:grid-cols-3 gap-3 mb-6">
            <Tile label="HOPs needing cleanup" value={String(totals.hops)} />
            <Tile label="Unpaid $ available to reassign" value={fmtMoney(totals.unpaid)} />
            <Tile label="Fully paid but Decom/SCOP incomplete" value={String(totals.incomplete)} warn={totals.incomplete > 0} />
          </div>

          {BUCKET_ORDER.map(bucket => (
            <BucketSection
              key={bucket}
              bucket={bucket}
              rows={groups[bucket]}
              assignments={assignments}
              onSaved={(hop, a) => setAssignments(prev => ({ ...prev, [hop]: a }))}
            />
          ))}

          {groups.fullyPaidButIncomplete.length > 0 && (
            <div className="mb-8">
              <h2 className="text-lg font-semibold text-amber-300 mb-2">⚠️ Fully paid, but Decom/SCOP still incomplete</h2>
              <p className="text-gray-500 text-xs mb-3">The original GC was paid in full — there&apos;s no unpaid tier to fund a replacement, so this needs a different fix (follow up with the GC, or absorb the cost of having someone else close it out).</p>
              <SimpleHopList rows={groups.fullyPaidButIncomplete} />
            </div>
          )}
        </>
      )}
    </div>
  )
}

function Tile({ label, value, warn }: { label: string; value: string; warn?: boolean }) {
  return (
    <div className={`rounded-xl border p-4 text-center ${warn ? 'bg-amber-950 border-amber-700' : 'bg-gray-900 border-gray-700'}`}>
      <p className={`text-2xl font-bold ${warn ? 'text-amber-300' : 'text-white'}`}>{value}</p>
      <p className="text-xs text-gray-400 mt-1">{label}</p>
    </div>
  )
}

function SimpleHopList({ rows }: { rows: CleanupHop[] }) {
  return (
    <div className="overflow-x-auto rounded-xl border border-gray-700">
      <table className="w-full text-xs">
        <thead>
          <tr className="bg-gray-800 text-gray-400">
            <th className="text-left p-2">HOP</th>
            <th className="text-left p-2">Original GC</th>
            <th className="text-left p-2">Decom</th>
            <th className="text-left p-2">SCOP</th>
          </tr>
        </thead>
        <tbody>
          {rows.map(h => (
            <tr key={h.hop} className="border-t border-gray-800 bg-gray-900">
              <td className="p-2 font-semibold text-white whitespace-nowrap">{h.hopDisplay}</td>
              <td className="p-2 text-gray-300">{h.gc}</td>
              <td className="p-2">{h.decomComplete ? <span className="text-green-400">✓</span> : <span className="text-red-400">✗</span>}</td>
              <td className="p-2">{h.scopComplete ? <span className="text-green-400">✓</span> : <span className="text-red-400">✗</span>}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

function BucketSection({ bucket, rows, assignments, onSaved }: {
  bucket: CleanupBucket
  rows: CleanupHop[]
  assignments: Record<string, CleanupAssignment>
  onSaved: (hop: string, a: CleanupAssignment) => void
}) {
  if (rows.length === 0) return null
  return (
    <div className="mb-8">
      <h2 className="text-lg font-semibold text-white mb-2">{BUCKET_LABELS[bucket]} ({rows.length})</h2>
      <div className="overflow-x-auto rounded-xl border border-gray-700">
        <table className="w-full text-xs">
          <thead>
            <tr className="bg-gray-800 text-gray-400">
              <th className="text-left p-2">HOP</th>
              <th className="text-left p-2">Path ID</th>
              <th className="text-left p-2">Original GC</th>
              <th className="text-left p-2">Paid / Total</th>
              <th className="text-left p-2">Unpaid $ available</th>
              <th className="text-left p-2">Unpaid tiers</th>
              <th className="text-left p-2">Decom</th>
              <th className="text-left p-2">SCOP</th>
              <th className="text-left p-2">Assign new GC</th>
              <th className="text-left p-2">Their quote</th>
              <th className="text-left p-2">Gap</th>
              <th className="text-left p-2">Status</th>
              <th className="text-left p-2">Comment</th>
              <th className="text-left p-2">Save</th>
            </tr>
          </thead>
          <tbody>
            {rows.map(h => (
              <CleanupRow key={h.hop} hop={h} assignment={assignments[h.hop]} onSaved={a => onSaved(h.hop, a)} />
            ))}
          </tbody>
        </table>
      </div>
    </div>
  )
}

function CleanupRow({ hop, assignment, onSaved }: {
  hop: CleanupHop
  assignment: CleanupAssignment | undefined
  onSaved: (a: CleanupAssignment) => void
}) {
  const [newGc, setNewGc] = useState(assignment?.newGc ?? '')
  const [quote, setQuote] = useState(assignment?.quote != null ? String(assignment.quote) : '')
  const [status, setStatus] = useState(assignment?.status ?? 'Needs Quote')
  const [comment, setComment] = useState(assignment?.comment ?? '')
  const [saving, setSaving] = useState(false)
  const [saveError, setSaveError] = useState<string | null>(null)
  const [savedAt, setSavedAt] = useState<string | null>(null)

  const quoteNum = quote.trim() === '' ? null : Number(quote)
  const gap = quoteNum != null && !Number.isNaN(quoteNum) ? quoteNum - hop.unpaidValue : null
  const isDirty = newGc !== (assignment?.newGc ?? '') || quote !== (assignment?.quote != null ? String(assignment.quote) : '')
    || status !== (assignment?.status ?? 'Needs Quote') || comment !== (assignment?.comment ?? '')

  async function save() {
    setSaving(true)
    setSaveError(null)
    try {
      const saved = await saveCleanupAssignment(hop.hop, {
        newGc, status,
        quote: quoteNum != null && !Number.isNaN(quoteNum) ? quoteNum : null,
        comment,
      })
      onSaved(saved)
      setSavedAt(new Date().toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }))
    } catch (err) {
      setSaveError(err instanceof Error ? err.message : 'Save failed — try again.')
    } finally {
      setSaving(false)
    }
  }

  return (
    <tr className={`border-t border-gray-800 align-top ${isDirty ? 'bg-amber-950/40' : 'bg-gray-900'}`}>
      <td className="p-2 font-semibold text-white whitespace-nowrap">{hop.hopDisplay}</td>
      <td className="p-2 text-gray-400 whitespace-nowrap">{hop.pathId || '—'}</td>
      <td className="p-2 text-gray-300 whitespace-nowrap">{hop.gc}</td>
      <td className="p-2 text-gray-300 whitespace-nowrap">{fmtMoney(hop.paidValue)} / {fmtMoney(hop.totalValue)} ({hop.paidPct}%)</td>
      <td className="p-2 text-emerald-300 font-semibold whitespace-nowrap">{fmtMoney(hop.unpaidValue)}</td>
      <td className="p-2 max-w-[14rem]">
        {hop.unpaidTiers.map(t => {
          // "Ready to release" means the work is already done — this is just
          // sitting in CJ's own GR queue, not something a new GC needs to fix.
          const readyToRelease = t.reason.startsWith('Ready to release')
          return (
            <div key={t.tier} className="whitespace-nowrap text-gray-400">
              {t.tierLabel}: {fmtMoney(t.value)}
              {t.reason && <span className={readyToRelease ? ' text-sky-400' : ' text-amber-400'}> — {t.reason}</span>}
            </div>
          )
        })}
      </td>
      <td className="p-2">{hop.decomComplete ? <span className="text-green-400">✓</span> : <span className="text-red-400">✗</span>}</td>
      <td className="p-2">{hop.scopComplete ? <span className="text-green-400">✓</span> : <span className="text-red-400">✗</span>}</td>
      <td className="p-2">
        <input list="gc-cleanup-roster" value={newGc} onChange={e => setNewGc(e.target.value)} placeholder="New GC…"
          className="w-36 bg-gray-800 text-white text-xs rounded px-2 py-1 border border-gray-600 focus:outline-none focus:border-blue-500" />
      </td>
      <td className="p-2">
        <input type="number" value={quote} onChange={e => setQuote(e.target.value)} placeholder="$"
          className="w-24 bg-gray-800 text-white text-xs rounded px-2 py-1 border border-gray-600 focus:outline-none focus:border-blue-500" />
      </td>
      <td className={`p-2 font-semibold whitespace-nowrap ${gap == null ? 'text-gray-600' : gap > 0 ? 'text-red-400' : 'text-green-400'}`}>
        {gap == null ? '—' : gap > 0 ? `Short ${fmtMoney(gap)}` : `Covers +${fmtMoney(-gap)}`}
      </td>
      <td className="p-2">
        <select value={status} onChange={e => setStatus(e.target.value as CleanupAssignment['status'])}
          className="bg-gray-800 text-gray-300 text-xs rounded px-1 py-1 border border-gray-600 focus:outline-none focus:border-blue-500">
          {CLEANUP_STATUS_OPTIONS.map(s => <option key={s} value={s}>{s}</option>)}
        </select>
      </td>
      <td className="p-2">
        <textarea rows={2} value={comment} onChange={e => setComment(e.target.value)} placeholder="Notes…"
          className="w-48 bg-gray-800 text-white text-xs rounded px-2 py-1 border border-gray-600 focus:outline-none focus:border-blue-500" />
      </td>
      <td className="p-2 whitespace-nowrap">
        <button onClick={save} disabled={!isDirty || saving}
          className="bg-blue-700 hover:bg-blue-600 disabled:opacity-30 disabled:cursor-not-allowed text-white text-xs px-2 py-1 rounded">
          {saving ? 'Saving…' : 'Save'}
        </button>
        {isDirty && !saving && <div className="text-amber-400 text-[10px] mt-1">unsaved</div>}
        {!isDirty && savedAt && <div className="text-emerald-400 text-[10px] mt-1">saved {savedAt}</div>}
        {saveError && <div className="text-red-400 text-[10px] mt-1">{saveError}</div>}
      </td>
    </tr>
  )
}
