'use client'

export const dynamic = 'force-dynamic'

import { useEffect, useMemo, useState } from 'react'
import dynamicImport from 'next/dynamic'
import BackToDashboard from '../components/BackToDashboard'
import ColumnFilterMenu from '../components/ColumnFilterMenu'
import { GC_CONFIG } from '../lib/gcConfig'
import { fmtMoney } from '../lib/grTracker'
import { loadDepartedGcs, saveDepartedGcs, loadEmailSettings, type EmailSettings } from '../lib/settings'
import {
  type CleanupHop, type CleanupGroups, type CleanupAssignment, type SiteNote,
  CLEANUP_STATUS_OPTIONS,
  loadCleanupGroups, loadCleanupAssignments, saveCleanupAssignment,
} from '../lib/gcCleanup'
import {
  type QuotedHop, quotedHops, groupByNewGc, cancellableHops, openGcQuoteEmail, openCancellationEmail, tierPercent,
} from '../lib/gcCleanupEmail'
import { clusterByDistance, type GeoCluster } from '../lib/hopCoords'

// Leaflet touches `window` at import time — see app/map/page.tsx for the
// same ssr:false pattern this mirrors.
const CleanupMap = dynamicImport(() => import('./CleanupMap'), {
  ssr: false,
  loading: () => <div className="flex items-center justify-center h-[60vh] text-gray-400 text-sm">Loading map…</div>,
})

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
  const [openBucket, setOpenBucket] = useState<'fullyPaidButIncomplete' | 'all' | 'unpaid' | 'quoted' | 'pendingQuoting' | null>(null)
  const [mapOpen, setMapOpen] = useState(false)
  const [emailCenterOpen, setEmailCenterOpen] = useState(false)
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
  // Every HOP the page shows anywhere, across all four tiles — the number
  // CJ would otherwise have to add up from the tiles himself.
  const everyHop = useMemo(() => groups ? [...allHops, ...groups.fullyPaidButIncomplete] : [], [groups, allHops])
  const quoted = useMemo(() => quotedHops(everyHop, assignments), [everyHop, assignments])
  // The tile only counts what still needs a GC email — already-sent ones stay
  // visible inside the Email Center (so nothing's lost) but shouldn't inflate
  // the "needs action" badge.
  const notYetGcEmailed = useMemo(() => quoted.filter(q => !q.assignment.gcEmailSentAt), [quoted])
  // "Quoted" / "Pending Quoting" is a strict partition of every HOP on the
  // page — same quote-entered definition the Email Center uses — so the two
  // always add up to the Total HOPs tile.
  const quotedCleanupHops = useMemo(() => quoted.map(q => q.hop), [quoted])
  const pendingQuotingHops = useMemo(() => {
    const quotedKeys = new Set(quotedCleanupHops.map(h => h.hop))
    return everyHop.filter(h => !quotedKeys.has(h.hop))
  }, [everyHop, quotedCleanupHops])

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
        that unpaid $ is what&apos;s available to pay someone else to finish the work. Only HOPs with an actual
        MS16 Construction Complete date show up here — until construction is physically done, there&apos;s
        nothing to hold a GC accountable for on Decom/SCOP yet.
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
          <Tile emoji="📋" label="Total HOPs across every tile below" count={everyHop.length}
            sub={`${departedGcs.length} GC${departedGcs.length === 1 ? '' : 's'} departed`}
            onClick={() => setOpenBucket('all')} />
          <Tile emoji="✅" label="Quoted" count={quotedCleanupHops.length}
            sub={fmtMoney(quoted.reduce((s, q) => s + (q.assignment.quote ?? 0), 0))}
            onClick={() => setOpenBucket('quoted')} />
          <Tile emoji="📝" label="Pending Quoting" count={pendingQuotingHops.length}
            sub={fmtMoney(pendingQuotingHops.reduce((s, h) => s + h.unpaidValue, 0))}
            onClick={() => setOpenBucket('pendingQuoting')} />
          <Tile emoji="⚠️" label="Fully paid, Decom/SCOP incomplete" count={groups.fullyPaidButIncomplete.length}
            sub="no money left to reassign" warn={groups.fullyPaidButIncomplete.length > 0}
            onClick={() => setOpenBucket('fullyPaidButIncomplete')} />
          <Tile emoji="💰" label="Total unpaid $ available" count={allHops.length} sub={fmtMoney(totalUnpaid)}
            onClick={() => setOpenBucket('unpaid')} />
          <Tile emoji="🗺️" label="Map & distance clusters" count={clusters.length || null}
            sub={unmapped > 0 ? `${unmapped} HOP(s) have no coordinates` : `within ${radiusMiles}mi`}
            onClick={() => setMapOpen(true)} />
          <Tile emoji="📧" label="Ready to email — quoted HOPs" count={notYetGcEmailed.length}
            sub={notYetGcEmailed.length === 0 ? (quoted.length === 0 ? 'assign a GC + quote first' : 'all caught up — nothing new to email') : `${groupByNewGc(notYetGcEmailed).length} GC(s) to email`}
            onClick={() => setEmailCenterOpen(true)} />
        </div>
      )}

      {!loading && groups && openBucket && (
        <BucketModal
          title={
            openBucket === 'all' ? `📋 All ${everyHop.length} HOPs`
            : openBucket === 'fullyPaidButIncomplete' ? '⚠️ Fully paid, Decom/SCOP still incomplete'
            : openBucket === 'unpaid' ? '💰 Total unpaid $ available'
            : openBucket === 'quoted' ? '✅ Quoted'
            : '📝 Pending Quoting'
          }
          rows={
            openBucket === 'all' ? everyHop
            : openBucket === 'fullyPaidButIncomplete' ? groups.fullyPaidButIncomplete
            : openBucket === 'unpaid' ? allHops
            : openBucket === 'quoted' ? quotedCleanupHops
            : pendingQuotingHops
          }
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

      {emailCenterOpen && (
        <EmailCenterModal
          quoted={quoted}
          onSaved={(hop, a) => setAssignments(prev => ({ ...prev, [hop]: a }))}
          onClose={() => setEmailCenterOpen(false)}
        />
      )}
    </div>
  )
}

function fmtSent(iso: string): string {
  if (!iso) return 'not sent'
  return new Date(iso).toLocaleString([], { month: 'numeric', day: 'numeric', hour: 'numeric', minute: '2-digit' })
}

function EmailCenterModal({ quoted, onSaved, onClose }: {
  quoted: QuotedHop[]
  onSaved: (hop: string, a: CleanupAssignment) => void
  onClose: () => void
}) {
  const [emailSettings, setEmailSettings] = useState<EmailSettings | null>(null)
  const [sendingGc, setSendingGc] = useState<string | null>(null)
  const [sendingFinance, setSendingFinance] = useState(false)
  const [error, setError] = useState<string | null>(null)
  // Off by default: once a HOP's been emailed it drops out of view instead of
  // piling up here, so this stays "what's new" rather than a growing log.
  const [showSent, setShowSent] = useState(false)

  useEffect(() => {
    loadEmailSettings().then(setEmailSettings).catch(() => setEmailSettings(null))
  }, [])

  const gcGroups = groupByNewGc(quoted)
  const cancellableAll = cancellableHops(quoted)

  // Sending an email is also CJ's own signal that this HOP moved past "just
  // quoted" — advances status unless he's already manually marked it Complete.
  async function stampSent(items: QuotedHop[], field: 'gcEmailSentAt' | 'financeEmailSentAt') {
    const stampedAt = new Date().toISOString()
    for (const { hop, assignment } of items) {
      const status = assignment.status === 'Complete' ? 'Complete' : 'Assigned'
      const saved = await saveCleanupAssignment(hop.hop, { [field]: stampedAt, status })
      onSaved(hop.hop, saved)
    }
  }

  async function handleEmailGc(newGc: string, items: QuotedHop[]) {
    if (!emailSettings || items.length === 0) return
    setSendingGc(newGc)
    setError(null)
    try {
      openGcQuoteEmail(newGc, items, emailSettings.gcContactEmails, emailSettings.ccList, emailSettings.routing)
      await stampSent(items, 'gcEmailSentAt')
    } catch (err) {
      setError(err instanceof Error ? err.message : 'The email opened, but stamping it as sent failed — try again so it stays tracked.')
    } finally {
      setSendingGc(null)
    }
  }

  async function handleEmailFinance(items: QuotedHop[]) {
    if (!emailSettings || items.length === 0) return
    setSendingFinance(true)
    setError(null)
    try {
      openCancellationEmail(items, emailSettings.financeEmails, emailSettings.routing)
      await stampSent(items, 'financeEmailSentAt')
    } catch (err) {
      setError(err instanceof Error ? err.message : 'The email opened, but stamping it as sent failed — try again so it stays tracked.')
    } finally {
      setSendingFinance(false)
    }
  }

  const sentGcCount = quoted.filter(q => q.assignment.gcEmailSentAt).length
  const sentFinanceCount = cancellableAll.filter(q => q.assignment.financeEmailSentAt).length

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 p-4" onClick={onClose}>
      <div className="bg-gray-900 border border-gray-700 rounded-xl shadow-2xl w-full max-w-4xl max-h-[90vh] overflow-y-auto" onClick={e => e.stopPropagation()}>
        <div className="flex items-center justify-between px-6 py-4 border-b border-gray-700 sticky top-0 bg-gray-900">
          <h2 className="text-white font-semibold text-lg">📧 Email Center — {quoted.length} quoted HOP{quoted.length === 1 ? '' : 's'}</h2>
          <div className="flex items-center gap-4">
            <label className="flex items-center gap-1.5 text-xs text-gray-400 cursor-pointer">
              <input type="checkbox" checked={showSent} onChange={e => setShowSent(e.target.checked)} />
              Show already-emailed ({sentGcCount} GC, {sentFinanceCount} finance)
            </label>
            <button onClick={onClose} className="text-gray-400 hover:text-white text-xl leading-none">&times;</button>
          </div>
        </div>
        <div className="p-6 space-y-6">
          {error && <p className="bg-red-950 border border-red-700 text-red-200 text-sm rounded-lg p-3">{error}</p>}
          {quoted.length === 0 && (
            <p className="text-gray-500 text-sm">No HOPs have both a new GC and a quote entered yet — fill those in from a tile&apos;s table first.</p>
          )}

          {gcGroups.length > 0 && (
            <div>
              <h3 className="text-sm font-semibold text-gray-300 mb-2">Quote confirmation — one email per GC</h3>
              <div className="space-y-3">
                {gcGroups.map(({ newGc, items }) => {
                  const pending = items.filter(i => !i.assignment.gcEmailSentAt)
                  const sent = items.filter(i => i.assignment.gcEmailSentAt)
                  if (pending.length === 0 && !showSent) return null
                  return (
                    <div key={newGc} className="bg-gray-800 rounded-lg border border-gray-700 p-3">
                      <div className="flex items-center justify-between mb-2 gap-3">
                        <p className="text-sm font-semibold text-white">
                          {newGc} — {pending.length} new{sent.length > 0 ? `, ${sent.length} already sent` : ''}
                          {pending.length > 0 && <> · {fmtMoney(pending.reduce((s, i) => s + (i.assignment.quote ?? 0), 0))}</>}
                        </p>
                        {pending.length > 0 ? (
                          <button onClick={() => handleEmailGc(newGc, pending)} disabled={sendingGc === newGc || !emailSettings}
                            className="bg-blue-700 hover:bg-blue-600 disabled:opacity-40 text-white text-xs font-semibold px-3 py-1.5 rounded whitespace-nowrap">
                            {sendingGc === newGc ? 'Opening…' : `✉️ Email ${newGc} (${pending.length})`}
                          </button>
                        ) : (
                          <button onClick={() => handleEmailGc(newGc, sent)} disabled={sendingGc === newGc || !emailSettings}
                            className="bg-gray-700 hover:bg-gray-600 disabled:opacity-40 text-gray-200 text-xs font-semibold px-3 py-1.5 rounded whitespace-nowrap">
                            {sendingGc === newGc ? 'Opening…' : 'Re-email anyway'}
                          </button>
                        )}
                      </div>
                      {pending.length > 0 && (
                        <ul className="text-xs text-gray-400 space-y-0.5">
                          {pending.map(({ hop, assignment }) => (
                            <li key={hop.hop} className="flex justify-between gap-2">
                              <span>{hop.hopDisplay} <span className="text-gray-600">({hop.pathId || '—'})</span> — {fmtMoney(assignment.quote ?? 0)}</span>
                              <span className="text-amber-400">not sent</span>
                            </li>
                          ))}
                        </ul>
                      )}
                      {showSent && sent.length > 0 && (
                        <ul className={`text-xs text-gray-500 space-y-0.5 ${pending.length > 0 ? 'mt-2 pt-2 border-t border-gray-700' : ''}`}>
                          {sent.map(({ hop, assignment }) => (
                            <li key={hop.hop} className="flex justify-between gap-2">
                              <span>{hop.hopDisplay} ({hop.pathId || '—'}) — {fmtMoney(assignment.quote ?? 0)}</span>
                              <span className="text-emerald-400">{fmtSent(assignment.gcEmailSentAt)}</span>
                            </li>
                          ))}
                        </ul>
                      )}
                    </div>
                  )
                })}
              </div>
            </div>
          )}

          {quoted.length > 0 && (
            <div>
              <h3 className="text-sm font-semibold text-gray-300 mb-2">SPO cancellation — one email to Finance</h3>
              {cancellableAll.length === 0 ? (
                <p className="text-gray-500 text-xs">None of the quoted HOPs have an unpaid SPO left — nothing for Finance to cancel.</p>
              ) : (
                (() => {
                  const pending = cancellableAll.filter(i => !i.assignment.financeEmailSentAt)
                  const sent = cancellableAll.filter(i => i.assignment.financeEmailSentAt)
                  if (pending.length === 0 && !showSent) {
                    return <p className="text-gray-500 text-xs">All {sent.length} cancellable HOP{sent.length === 1 ? '' : 's'} already emailed to Finance — check &quot;Show already-emailed&quot; above to review or resend.</p>
                  }
                  return (
                    <div className="bg-gray-800 rounded-lg border border-gray-700 p-3">
                      <div className="flex items-center justify-between mb-2 gap-3">
                        <p className="text-sm font-semibold text-white">
                          {pending.length} new{sent.length > 0 ? `, ${sent.length} already sent` : ''}
                          {pending.length > 0 && <> · {pending.reduce((s, i) => s + i.hop.unpaidTiers.length, 0)} SPO(s) to cancel</>}
                        </p>
                        {pending.length > 0 ? (
                          <button onClick={() => handleEmailFinance(pending)} disabled={sendingFinance || !emailSettings}
                            className="bg-red-800 hover:bg-red-700 disabled:opacity-40 text-white text-xs font-semibold px-3 py-1.5 rounded whitespace-nowrap">
                            {sendingFinance ? 'Opening…' : `✉️ Email Finance (${pending.length})`}
                          </button>
                        ) : (
                          <button onClick={() => handleEmailFinance(sent)} disabled={sendingFinance || !emailSettings}
                            className="bg-gray-700 hover:bg-gray-600 disabled:opacity-40 text-gray-200 text-xs font-semibold px-3 py-1.5 rounded whitespace-nowrap">
                            {sendingFinance ? 'Opening…' : 'Re-email anyway'}
                          </button>
                        )}
                      </div>
                      {pending.length > 0 && (
                        <ul className="text-xs text-gray-400 space-y-0.5">
                          {pending.map(({ hop }) => (
                            <li key={hop.hop} className="flex justify-between gap-2">
                              <span>{hop.hopDisplay} (GC: {hop.gc}) — {hop.unpaidTiers.map(t => `${t.spoNumber || '—'} (${tierPercent(t.tier)}, ${t.vendor})`).join(', ')}</span>
                              <span className="text-amber-400">not sent</span>
                            </li>
                          ))}
                        </ul>
                      )}
                      {showSent && sent.length > 0 && (
                        <ul className={`text-xs text-gray-500 space-y-0.5 ${pending.length > 0 ? 'mt-2 pt-2 border-t border-gray-700' : ''}`}>
                          {sent.map(({ hop, assignment }) => (
                            <li key={hop.hop} className="flex justify-between gap-2">
                              <span>{hop.hopDisplay} (GC: {hop.gc}) — {hop.unpaidTiers.map(t => `${t.spoNumber || '—'} (${tierPercent(t.tier)}, ${t.vendor})`).join(', ')}</span>
                              <span className="text-emerald-400">{fmtSent(assignment.financeEmailSentAt)}</span>
                            </li>
                          ))}
                        </ul>
                      )}
                    </div>
                  )
                })()
              )}
            </div>
          )}
        </div>
      </div>
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

function BucketModal({ title, rows, assignments, onSaved, onClose }: {
  title: string
  rows: CleanupHop[]
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
                <th className="text-left p-2 bg-gray-800">Assign new GC</th>
                <th className="text-left p-2 bg-gray-800">Their quote</th>
                <th className="text-left p-2 bg-gray-800">Gap</th>
                <FilterableHeader label="Status" col="status" openCol={openCol} setOpenCol={setOpenCol} values={valuesByCol.status} filters={filters} onApply={setColumnFilter} />
                <th className="text-left p-2 bg-gray-800">Comment</th>
                <th className="text-left p-2 bg-gray-800">Save</th>
              </tr>
            </thead>
            <tbody>
              {visible.map(h => (
                h.needsReassignment
                  ? <CleanupRow key={h.hop} hop={h} assignment={assignments[h.hop]} onSaved={a => onSaved(h.hop, a)} />
                  : <ReadOnlyRow key={h.hop} hop={h} />
              ))}
              {visible.length === 0 && (
                <tr><td colSpan={14} className="p-4 text-center text-gray-500">No HOPs match these filters.</td></tr>
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
function SiteStatusCell({ complete, pendingSites }: { complete: boolean; pendingSites: SiteNote[] }) {
  if (complete) return <td className="p-2"><span className="text-green-400">✓</span></td>
  return (
    <td className="p-2 max-w-[12rem]">
      <span className="text-red-400">✗</span>
      {pendingSites.length > 0 && (
        <div className="mt-0.5 text-[10px] text-gray-500 leading-tight">
          {pendingSites.map((s, i) => (
            <div key={`${s.siteName}-${i}`} title={`${s.siteName}: ${s.statusLabel}`} className="max-w-[12rem] overflow-hidden text-ellipsis whitespace-nowrap">
              {s.siteName}: {s.statusLabel}
            </div>
          ))}
        </div>
      )}
    </td>
  )
}

function DecomCell({ hop }: { hop: CleanupHop }) {
  return <SiteStatusCell complete={hop.decomComplete} pendingSites={hop.decomPendingSites} />
}

function ScopCell({ hop }: { hop: CleanupHop }) {
  return <SiteStatusCell complete={hop.scopComplete} pendingSites={hop.scopPendingSites} />
}

// Unpaid $ + the itemized tier list — shared between the editable row and
// the read-only one, since a "ready to release" HOP still has real unpaid $
// and tiers to show, just nothing for a new GC to do about them.
function UnpaidCells({ hop }: { hop: CleanupHop }) {
  if (hop.unpaidTiers.length === 0) {
    return (
      <>
        <td className="p-2 text-gray-500">—</td>
        <td className="p-2 text-gray-600">—</td>
      </>
    )
  }
  return (
    <>
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
    </>
  )
}

function ReadOnlyRow({ hop }: { hop: CleanupHop }) {
  const note = hop.unpaidValue > 0
    ? 'Decom + SCOP already complete — just needs the GR submitted, no new GC needed'
    : 'Fully paid — no unpaid tier to reassign money against'
  return (
    <tr className="border-t border-gray-800 bg-gray-900">
      <td className="p-2 font-semibold text-white whitespace-nowrap">{hop.hopDisplay}</td>
      <td className="p-2 text-gray-400 whitespace-nowrap">{hop.pathId || '—'}</td>
      <td className="p-2 text-gray-300 whitespace-nowrap">{hop.gc}</td>
      <td className="p-2 text-gray-300 whitespace-nowrap">{fmtMoney(hop.paidValue)} / {fmtMoney(hop.totalValue)} ({hop.paidPct}%)</td>
      <UnpaidCells hop={hop} />
      <DecomCell hop={hop} />
      <ScopCell hop={hop} />
      <td colSpan={6} className="p-2 text-gray-600 text-center italic">{note}</td>
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
      <UnpaidCells hop={hop} />
      <DecomCell hop={hop} />
      <ScopCell hop={hop} />
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
