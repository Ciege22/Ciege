'use client'

export const dynamic = 'force-dynamic'

import { useEffect, useMemo, useState } from 'react'
import dynamicImport from 'next/dynamic'
import BackToDashboard from '../components/BackToDashboard'
import ColumnFilterMenu from '../components/ColumnFilterMenu'
import { GC_CONFIG } from '../lib/gcConfig'
import { fmtMoney } from '../lib/grTracker'
import { loadDepartedGcs, saveDepartedGcs } from '../lib/settings'
import {
  type CleanupBucket, type CleanupHop, type CleanupGroups, type CleanupAssignment,
  BUCKET_LABELS, CLEANUP_STATUS_OPTIONS,
  loadCleanupGroups, loadCleanupAssignments, saveCleanupAssignment,
} from '../lib/gcCleanup'
import { clusterByDistance, type GeoCluster } from '../lib/hopCoords'

// Leaflet touches `window` at import time — see app/map/page.tsx for the
// same ssr:false pattern this mirrors.
const CleanupMap = dynamicImport(() => import('./CleanupMap'), {
  ssr: false,
  loading: () => <div className="flex items-center justify-center h-[60vh] text-gray-400 text-sm">Loading map…</div>,
})

const BUCKET_ORDER: CleanupBucket[] = ['cleanup_only', 'mid_construction', 'not_started']
const BUCKET_EMOJI: Record<CleanupBucket, string> = { cleanup_only: '🧾', mid_construction: '🏗️', not_started: '🚧' }
const DEFAULT_RADIUS_MILES = 100

type FilterCol = 'gc' | 'decom' | 'scop' | 'status'
const FILTER_COLS: { key: FilterCol; label: string }[] = [
  { key: 'gc', label: 'Original GC' },
  { key: 'decom', label: 'Decom' },
  { key: 'scop', label: 'SCOP' },
  { key: 'status', label: 'Status' },
]

export default function GcCleanupPage() {
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)
  const [departedGcs, setDepartedGcs] = useState<string[]>([])
  const [savingGcList, setSavingGcList] = useState(false)
  const [groups, setGroups] = useState<CleanupGroups | null>(null)
  const [assignments, setAssignments] = useState<Record<string, CleanupAssignment>>({})
  const [openBucket, setOpenBucket] = useState<CleanupBucket | 'fullyPaidButIncomplete' | null>(null)
  const [mapOpen, setMapOpen] = useState(false)
  const [radiusMiles, setRadiusMiles] = useState(DEFAULT_RADIUS_MILES)

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

  const allHops = useMemo(() => groups ? [...groups.not_started, ...groups.mid_construction, ...groups.cleanup_only] : [], [groups])
  const totalUnpaid = useMemo(() => allHops.reduce((s, h) => s + h.unpaidValue, 0), [allHops])

  const clusters = useMemo<GeoCluster<CleanupHop>[]>(() => {
    const points = allHops.filter(h => h.coord).map(h => ({ item: h, coord: h.coord! }))
    return clusterByDistance(points, radiusMiles)
  }, [allHops, radiusMiles])
  const unmapped = allHops.length - clusters.reduce((s, c) => s + c.members.length, 0)

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

      {!loading && groups && departedGcs.length > 0 && (
        <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-6">
          {BUCKET_ORDER.map(bucket => (
            <Tile key={bucket} emoji={BUCKET_EMOJI[bucket]} label={BUCKET_LABELS[bucket]} count={groups[bucket].length}
              sub={fmtMoney(groups[bucket].reduce((s, h) => s + h.unpaidValue, 0))}
              onClick={() => setOpenBucket(bucket)} />
          ))}
          <Tile emoji="⚠️" label="Fully paid, Decom/SCOP incomplete" count={groups.fullyPaidButIncomplete.length}
            sub="no money left to reassign" warn={groups.fullyPaidButIncomplete.length > 0}
            onClick={() => setOpenBucket('fullyPaidButIncomplete')} />
          <Tile emoji="💰" label="Total unpaid $ available" count={null} sub={fmtMoney(totalUnpaid)}
            onClick={() => setOpenBucket('cleanup_only')} />
          <Tile emoji="🗺️" label="Map & distance clusters" count={clusters.length || null}
            sub={unmapped > 0 ? `${unmapped} HOP(s) have no coordinates` : `within ${radiusMiles}mi`}
            onClick={() => setMapOpen(true)} />
        </div>
      )}

      {!loading && groups && openBucket && (
        <BucketModal
          title={openBucket === 'fullyPaidButIncomplete' ? '⚠️ Fully paid, Decom/SCOP still incomplete' : BUCKET_LABELS[openBucket]}
          rows={groups[openBucket]}
          editable={openBucket !== 'fullyPaidButIncomplete'}
          assignments={assignments}
          onSaved={(hop, a) => setAssignments(prev => ({ ...prev, [hop]: a }))}
          onClose={() => setOpenBucket(null)}
        />
      )}

      {mapOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 p-4" onClick={() => setMapOpen(false)}>
          <div className="bg-gray-900 border border-gray-700 rounded-xl shadow-2xl w-full max-w-6xl max-h-[90vh] overflow-y-auto" onClick={e => e.stopPropagation()}>
            <div className="flex items-center justify-between px-6 py-4 border-b border-gray-700 sticky top-0 bg-gray-900">
              <h2 className="text-white font-semibold text-lg">🗺️ Map & distance clusters</h2>
              <button onClick={() => setMapOpen(false)} className="text-gray-400 hover:text-white text-xl leading-none">&times;</button>
            </div>
            <div className="p-6">
              <div className="flex items-center gap-3 mb-4">
                <label className="text-sm text-gray-300">Cluster radius</label>
                <input type="number" min={5} step={5} value={radiusMiles}
                  onChange={e => setRadiusMiles(Math.max(5, Number(e.target.value) || DEFAULT_RADIUS_MILES))}
                  className="w-24 bg-gray-800 text-white text-sm rounded px-2 py-1 border border-gray-600 focus:outline-none focus:border-blue-500" />
                <span className="text-sm text-gray-400">miles</span>
                <span className="text-xs text-gray-500 ml-2">
                  Greedy grouping by distance — HOPs within {radiusMiles}mi of a cluster&apos;s center get pulled in. Not an optimal grouping, just a practical way to bundle nearby work for one GC.
                </span>
              </div>
              {unmapped > 0 && (
                <p className="text-amber-400 text-xs mb-3">{unmapped} HOP(s) have no valid Latt./Long. in the tracker and are left off the map.</p>
              )}
              <CleanupMap clusters={clusters} radiusMiles={radiusMiles} />
              <div className="mt-4 grid md:grid-cols-2 gap-3">
                {clusters.map(c => (
                  <div key={c.id} className="bg-gray-800 rounded-lg border border-gray-700 p-3">
                    <p className="text-sm font-semibold text-white mb-1">
                      Cluster {c.id} — {c.members.length} HOP{c.members.length === 1 ? '' : 's'} · {fmtMoney(c.members.reduce((s, h) => s + h.unpaidValue, 0))}
                    </p>
                    <ul className="text-xs text-gray-400 space-y-0.5 max-h-32 overflow-y-auto">
                      {c.members.map(h => <li key={h.hop}>{h.hopDisplay} <span className="text-gray-600">({h.gc})</span></li>)}
                    </ul>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

function Tile({ emoji, label, count, sub, warn, onClick }: { emoji: string; label: string; count: number | null; sub: string; warn?: boolean; onClick: () => void }) {
  return (
    <button onClick={onClick}
      className={`text-left rounded-xl border p-4 hover:border-blue-500 transition-colors ${warn ? 'bg-amber-950 border-amber-700' : 'bg-gray-900 border-gray-700'}`}>
      <p className="text-2xl mb-1">{emoji}</p>
      {count != null && <p className={`text-2xl font-bold ${warn ? 'text-amber-300' : 'text-white'}`}>{count}</p>}
      <p className={`text-sm font-semibold ${warn ? 'text-amber-200' : 'text-gray-200'}`}>{sub}</p>
      <p className="text-xs text-gray-500 mt-1">{label}</p>
    </button>
  )
}

function cellValue(h: CleanupHop, col: FilterCol, assignments: Record<string, CleanupAssignment>): string {
  if (col === 'gc') return h.gc
  if (col === 'decom') return h.decomComplete ? 'Complete' : 'Incomplete'
  if (col === 'scop') return h.scopComplete ? 'Complete' : 'Incomplete'
  return assignments[h.hop]?.status ?? 'Needs Quote'
}

function BucketModal({ title, rows, editable, assignments, onSaved, onClose }: {
  title: string
  rows: CleanupHop[]
  editable: boolean
  assignments: Record<string, CleanupAssignment>
  onSaved: (hop: string, a: CleanupAssignment) => void
  onClose: () => void
}) {
  const [search, setSearch] = useState('')
  const [filters, setFilters] = useState<Partial<Record<FilterCol, Set<string>>>>({})
  const [openCol, setOpenCol] = useState<FilterCol | null>(null)

  const valuesByCol = useMemo(() => {
    const out = {} as Record<FilterCol, string[]>
    FILTER_COLS.forEach(c => {
      out[c.key] = Array.from(new Set(rows.map(h => cellValue(h, c.key, assignments)))).sort()
    })
    return out
    // Intentionally keyed on rows only — recomputing live off `assignments` on every
    // keystroke would reshuffle the filter's own value list out from under an open menu.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rows])

  const visible = useMemo(() => {
    const q = search.trim().toLowerCase()
    return rows.filter(h => {
      for (const key of Object.keys(filters) as FilterCol[]) {
        const allowed = filters[key]
        if (allowed && !allowed.has(cellValue(h, key, assignments))) return false
      }
      return !q || [h.hopDisplay, h.pathId, h.gc].some(v => v.toLowerCase().includes(q))
    })
  }, [rows, filters, search, assignments])

  function setColumnFilter(key: FilterCol, selected: Set<string> | null) {
    setFilters(prev => {
      const next = { ...prev }
      if (selected === null || selected.size === valuesByCol[key].length) delete next[key]
      else next[key] = selected
      return next
    })
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 p-4" onClick={onClose}>
      <div className="bg-gray-900 border border-gray-700 rounded-xl shadow-2xl w-full max-w-[96vw] max-h-[90vh] flex flex-col" onClick={e => e.stopPropagation()}>
        <div className="flex items-center justify-between px-6 py-4 border-b border-gray-700">
          <h2 className="text-white font-semibold text-lg">
            {title} <span className="text-gray-400 text-sm ml-2">({visible.length === rows.length ? `${rows.length} HOPs` : `${visible.length} of ${rows.length} HOPs`})</span>
          </h2>
          <div className="flex items-center gap-3">
            <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search HOP, Path ID, GC"
              className="w-56 bg-gray-800 text-white text-sm rounded px-2 py-1 border border-gray-600 focus:outline-none focus:border-blue-500" />
            {Object.keys(filters).length > 0 && (
              <button onClick={() => setFilters({})} className="text-gray-400 hover:text-white text-xs underline">✕ Clear filters</button>
            )}
            <button onClick={onClose} className="text-gray-400 hover:text-white text-xl leading-none">&times;</button>
          </div>
        </div>
        <div className="overflow-auto p-4">
          <table className="w-full text-xs">
            <thead className="sticky top-0 z-10 bg-gray-800">
              <tr className="bg-gray-800 text-gray-400">
                <th className="text-left p-2 bg-gray-800">HOP</th>
                <th className="text-left p-2 bg-gray-800">Path ID</th>
                <FilterableHeader label="Original GC" col="gc" openCol={openCol} setOpenCol={setOpenCol} values={valuesByCol.gc} filters={filters} onApply={setColumnFilter} />
                <th className="text-left p-2 bg-gray-800">Paid / Total</th>
                <th className="text-left p-2 bg-gray-800">Unpaid $ available</th>
                <th className="text-left p-2 bg-gray-800">Unpaid tiers</th>
                <FilterableHeader label="Decom" col="decom" openCol={openCol} setOpenCol={setOpenCol} values={valuesByCol.decom} filters={filters} onApply={setColumnFilter} />
                <FilterableHeader label="SCOP" col="scop" openCol={openCol} setOpenCol={setOpenCol} values={valuesByCol.scop} filters={filters} onApply={setColumnFilter} />
                {editable && <th className="text-left p-2 bg-gray-800">Assign new GC</th>}
                {editable && <th className="text-left p-2 bg-gray-800">Their quote</th>}
                {editable && <th className="text-left p-2 bg-gray-800">Gap</th>}
                {editable && <FilterableHeader label="Status" col="status" openCol={openCol} setOpenCol={setOpenCol} values={valuesByCol.status} filters={filters} onApply={setColumnFilter} />}
                {editable && <th className="text-left p-2 bg-gray-800">Comment</th>}
                {editable && <th className="text-left p-2 bg-gray-800">Save</th>}
              </tr>
            </thead>
            <tbody>
              {visible.map(h => (
                editable
                  ? <CleanupRow key={h.hop} hop={h} assignment={assignments[h.hop]} onSaved={a => onSaved(h.hop, a)} />
                  : <ReadOnlyRow key={h.hop} hop={h} />
              ))}
              {visible.length === 0 && (
                <tr><td colSpan={13} className="p-4 text-center text-gray-500">No HOPs match these filters.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  )
}

function FilterableHeader({ label, col, openCol, setOpenCol, values, filters, onApply }: {
  label: string
  col: FilterCol
  openCol: FilterCol | null
  setOpenCol: (c: FilterCol | null) => void
  values: string[]
  filters: Partial<Record<FilterCol, Set<string>>>
  onApply: (col: FilterCol, sel: Set<string> | null) => void
}) {
  const active = !!filters[col]
  return (
    <th className="relative text-left p-2 whitespace-nowrap bg-gray-800">
      <button type="button" onClick={() => setOpenCol(openCol === col ? null : col)} className="inline-flex items-center gap-1 hover:text-white">
        {label}<span className={active ? 'text-emerald-400' : 'text-gray-600'}>{active ? ' ⏷' : ' ▾'}</span>
      </button>
      {openCol === col && (
        <ColumnFilterMenu values={values} selected={filters[col] ?? null} sort={null} onSort={() => {}}
          onApply={sel => onApply(col, sel)} onClose={() => setOpenCol(null)} />
      )}
    </th>
  )
}

// Decom is tracked per physical site (a HOP has two) — the checkmark stays
// HOP-level since that's what's quoted, but the pending site name(s) show
// underneath so it's clear which end of the HOP still needs the work.
function DecomCell({ hop }: { hop: CleanupHop }) {
  if (hop.decomComplete) return <td className="p-2"><span className="text-green-400">✓</span></td>
  return (
    <td className="p-2 max-w-[10rem]">
      <span className="text-red-400">✗</span>
      {hop.decomPendingSites.length > 0 && (
        <div className="mt-0.5 text-[10px] text-gray-500 leading-tight">
          {hop.decomPendingSites.map(s => (
            <div key={s.siteName} title={`${s.siteName}: ${s.statusLabel}`} className="max-w-[10rem] overflow-hidden text-ellipsis whitespace-nowrap">
              {s.siteName}: {s.statusLabel}
            </div>
          ))}
        </div>
      )}
    </td>
  )
}

function ReadOnlyRow({ hop }: { hop: CleanupHop }) {
  return (
    <tr className="border-t border-gray-800 bg-gray-900">
      <td className="p-2 font-semibold text-white whitespace-nowrap">{hop.hopDisplay}</td>
      <td className="p-2 text-gray-400 whitespace-nowrap">{hop.pathId || '—'}</td>
      <td className="p-2 text-gray-300 whitespace-nowrap">{hop.gc}</td>
      <td className="p-2 text-gray-300 whitespace-nowrap">{fmtMoney(hop.paidValue)} / {fmtMoney(hop.totalValue)} ({hop.paidPct}%)</td>
      <td className="p-2 text-gray-500">—</td>
      <td className="p-2 text-gray-600">—</td>
      <DecomCell hop={hop} />
      <td className="p-2">{hop.scopComplete ? <span className="text-green-400">✓</span> : <span className="text-red-400">✗</span>}</td>
    </tr>
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
      <td className="p-2 w-56 max-w-[14rem] overflow-hidden">
        {hop.unpaidTiers.map(t => {
          const readyToRelease = t.reason.startsWith('Ready to release')
          const full = `${t.tierLabel}: ${fmtMoney(t.value)}${t.reason ? ` — ${t.reason}` : ''}`
          return (
            // title gives the full text on hover; nothing is cut from the DOM
            // so a drag-select still copies the whole line, only the ellipsis is visual.
            <div key={t.tier} title={full} className="max-w-[14rem] overflow-hidden text-ellipsis whitespace-nowrap text-gray-400">
              {t.tierLabel}: {fmtMoney(t.value)}
              {t.reason && <span className={readyToRelease ? ' text-sky-400' : ' text-amber-400'}> — {t.reason}</span>}
            </div>
          )
        })}
      </td>
      <DecomCell hop={hop} />
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
