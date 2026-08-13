'use client'

export const dynamic = 'force-dynamic'

import { useState, useEffect } from 'react'
import { GC_CONFIG } from '../lib/gcConfig'
import {
  GrRow, loadGrRows, groupGrRows, sortGrRows, buildGrEmailMailto, fmtMoney, fmtMoneyShort,
} from '../lib/grTracker'

const TIER_OPTIONS: { value: string; label: string }[] = [
  { value: 'all', label: 'All Tiers' },
  { value: 'init20', label: 'Init 20% (MS15A)' },
  { value: '60', label: '60% (MS16A)' },
  { value: '70', label: '70% (MS16A)' },
  { value: '20', label: '20% Decom/SCOP' },
  { value: '30', label: '30% Decom/SCOP' },
]

const STATUS_OPTIONS = ['All', 'Ready to Release', 'Awaiting Trigger', 'GR Done'] as const
type StatusOption = typeof STATUS_OPTIONS[number]

type SpecialFilter = null | 'ready' | 'awaiting' | 'doneThisMonth' | 'decomScop'

function KpiTile({ emoji, label, value, sub, color, onClick }: {
  emoji: string; label: string; value: string; sub: string; color: string; onClick: () => void
}) {
  return (
    <div onClick={onClick}
      className={`bg-gray-900 rounded-xl border p-4 text-center cursor-pointer hover:border-blue-500 hover:bg-gray-800 transition-all ${color}`}>
      <p className="text-2xl mb-1">{emoji}</p>
      <p className="text-white text-2xl font-bold">{value}</p>
      <p className="text-gray-400 text-xs mt-1 font-semibold">{label}</p>
      <p className="text-gray-600 text-xs mt-0.5">{sub}</p>
    </div>
  )
}

function StatusChip({ status }: { status: GrRow['status'] }) {
  if (status === 'GR Done') return <span className="bg-green-900 text-green-200 text-xs px-2 py-0.5 rounded-full whitespace-nowrap">✓ GR Done</span>
  if (status === 'Ready to Release') return <span className="bg-orange-900 text-orange-200 text-xs px-2 py-0.5 rounded-full whitespace-nowrap">⏳ Pending</span>
  return <span className="bg-gray-700 text-gray-300 text-xs px-2 py-0.5 rounded-full whitespace-nowrap">Not Yet</span>
}

export default function GrTrackerPage() {
  const [rows, setRows] = useState<GrRow[]>([])
  const [loaded, setLoaded] = useState(false)
  const [gcFilter, setGcFilter] = useState('ALL')
  const [tierFilter, setTierFilter] = useState('all')
  const [statusFilter, setStatusFilter] = useState<StatusOption>('All')
  const [specialFilter, setSpecialFilter] = useState<SpecialFilter>(null)
  const [emailGroups, setEmailGroups] = useState<{ gc: string; count: number; mailto: string }[] | null>(null)

  useEffect(() => {
    const load = async () => {
      const grRows = await loadGrRows()
      setRows(grRows)
      setLoaded(true)
    }
    load()
  }, [])

  const groups = groupGrRows(rows)
  const now = new Date()
  const doneThisMonth = groups.done.filter(r => r.grDateRaw && r.grDateRaw.getMonth() === now.getMonth() && r.grDateRaw.getFullYear() === now.getFullYear())
  const totalPendingValue = groups.ready.reduce((sum, r) => sum + r.spoValue, 0)
  const awarenessTotalValue = groups.awareness.reduce((sum, r) => sum + r.spoValue, 0)

  const selectTile = (filter: SpecialFilter) => {
    setSpecialFilter(filter)
    setGcFilter('ALL')
    setTierFilter('all')
    setStatusFilter('All')
  }

  let displayRows: GrRow[]
  if (specialFilter === 'ready') displayRows = groups.ready
  else if (specialFilter === 'awaiting') displayRows = groups.awaiting
  else if (specialFilter === 'doneThisMonth') displayRows = doneThisMonth
  else if (specialFilter === 'decomScop') displayRows = groups.awareness
  else {
    displayRows = rows.filter(r => {
      if (gcFilter !== 'ALL' && r.gc !== gcFilter) return false
      if (tierFilter !== 'all' && r.tier !== tierFilter) return false
      if (statusFilter !== 'All' && r.status !== statusFilter) return false
      return true
    })
  }
  displayRows = sortGrRows(displayRows)

  const generateEmails = () => {
    const byGc = new Map<string, GrRow[]>()
    groups.ready.forEach(r => {
      if (!byGc.has(r.gc)) byGc.set(r.gc, [])
      byGc.get(r.gc)!.push(r)
    })
    const out: { gc: string; count: number; mailto: string }[] = []
    byGc.forEach((gcRows, gc) => {
      out.push({ gc, count: gcRows.length, mailto: buildGrEmailMailto(gc, sortGrRows(gcRows)) })
    })
    out.sort((a, b) => a.gc.localeCompare(b.gc))
    setEmailGroups(out)
  }

  return (
    <div className="min-h-screen bg-gray-950 text-white p-6">
      <div className="max-w-7xl mx-auto">

        <div className="mb-6">
          <h1 className="text-3xl font-bold">💰 GR / Invoicing Tracker</h1>
          <p className="text-gray-400 mt-1">Goods Receipt release tracking — DON 444 · Viaero MW Program</p>
        </div>

        {!loaded && <p className="text-gray-400">Loading SPO report and tracker data...</p>}

        {loaded && rows.length === 0 && (
          <div className="bg-gray-900 rounded-xl border border-gray-700 p-12 text-center">
            <p className="text-gray-400 text-xl">📂 No SPO report data found — upload the SPO Master Report on the Reports page</p>
          </div>
        )}

        {loaded && rows.length > 0 && (
          <>
            {/* KPI Tiles */}
            <div className="grid grid-cols-2 md:grid-cols-5 gap-3 mb-6">
              <KpiTile emoji="🟠" label="Ready to Release" value={String(groups.ready.length)} sub="CJ action needed"
                color={groups.ready.length > 0 ? 'border-orange-600' : 'border-gray-700'}
                onClick={() => selectTile('ready')} />
              <KpiTile emoji="💵" label="Total Value Pending" value={fmtMoneyShort(totalPendingValue)} sub="sum of Ready SPOs"
                color={groups.ready.length > 0 ? 'border-orange-600' : 'border-gray-700'}
                onClick={() => selectTile('ready')} />
              <KpiTile emoji="⏳" label="Awaiting Trigger" value={String(groups.awaiting.length)} sub="not yet eligible"
                color="border-gray-700"
                onClick={() => selectTile('awaiting')} />
              <KpiTile emoji="✅" label="GR Done This Month" value={String(doneThisMonth.length)} sub={now.toLocaleString('default', { month: 'long' })}
                color="border-green-700"
                onClick={() => selectTile('doneThisMonth')} />
              <KpiTile emoji="👁" label="Awaiting Other PM" value={String(groups.awareness.length)} sub="Decom/SCOP — awareness"
                color="border-gray-700"
                onClick={() => selectTile('decomScop')} />
            </div>

            {/* Filter Bar */}
            <div className="flex flex-wrap gap-3 mb-4 items-center">
              <select value={gcFilter} onChange={(e) => { setGcFilter(e.target.value); setSpecialFilter(null) }}
                className="bg-gray-800 border border-gray-600 text-white text-sm rounded px-3 py-2">
                <option value="ALL">All GCs</option>
                {GC_CONFIG.map(cfg => <option key={cfg.gc} value={cfg.gc}>{cfg.gc}</option>)}
              </select>
              <select value={tierFilter} onChange={(e) => { setTierFilter(e.target.value); setSpecialFilter(null) }}
                className="bg-gray-800 border border-gray-600 text-white text-sm rounded px-3 py-2">
                {TIER_OPTIONS.map(t => <option key={t.value} value={t.value}>{t.label}</option>)}
              </select>
              <select value={statusFilter} onChange={(e) => { setStatusFilter(e.target.value as StatusOption); setSpecialFilter(null) }}
                className="bg-gray-800 border border-gray-600 text-white text-sm rounded px-3 py-2">
                {STATUS_OPTIONS.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
              {specialFilter && (
                <button onClick={() => setSpecialFilter(null)}
                  className="text-gray-400 hover:text-white text-xs underline">✕ Clear tile filter</button>
              )}
              <span className="text-gray-500 text-xs ml-auto">{displayRows.length} rows</span>
            </div>

            {/* Batch Email */}
            <div className="mb-4">
              <button onClick={generateEmails} disabled={groups.ready.length === 0}
                className="bg-blue-600 hover:bg-blue-700 disabled:opacity-40 disabled:cursor-not-allowed text-white px-4 py-2 rounded-lg text-sm font-semibold">
                ✉️ Generate GR Emails — All GCs
              </button>
              {emailGroups && (
                <div className="flex flex-wrap gap-2 mt-3">
                  {emailGroups.map(eg => (
                    <a key={eg.gc} href={eg.mailto}
                      className="bg-gray-800 hover:bg-gray-700 border border-gray-600 text-white text-xs px-3 py-2 rounded-lg font-semibold">
                      ✉️ {eg.gc} ({eg.count})
                    </a>
                  ))}
                </div>
              )}
            </div>

            {/* Main Table */}
            <div className="overflow-x-auto bg-gray-900 rounded-xl border border-gray-700">
              <table className="w-full text-xs">
                <thead>
                  <tr className="bg-gray-800 text-gray-400">
                    <th className="text-left p-2">HOP</th>
                    <th className="text-left p-2">Path ID</th>
                    <th className="text-left p-2">GC</th>
                    <th className="text-left p-2">SOG Name</th>
                    <th className="text-left p-2">SPO #</th>
                    <th className="text-left p-2">SPO Value</th>
                    <th className="text-left p-2">Trigger Date</th>
                    <th className="text-left p-2">GR Status</th>
                    <th className="text-left p-2">GR Date</th>
                  </tr>
                </thead>
                <tbody>
                  {displayRows.map((r, i) => (
                    <tr key={`${r.hop}-${r.sogName}-${i}`}
                      className={`border-t border-gray-800 ${r.isDecomScop ? 'bg-gray-950' : 'bg-gray-900'}`}>
                      <td className="p-2 font-semibold whitespace-nowrap">
                        {r.isDecomScop && <span className="mr-1">👁</span>}
                        <span className={r.isDecomScop ? 'text-gray-500' : 'text-white'}>{r.hopDisplay}</span>
                      </td>
                      <td className={`p-2 whitespace-nowrap ${r.isDecomScop ? 'text-gray-600' : 'text-gray-400'}`}>{r.pathId || '—'}</td>
                      <td className={`p-2 whitespace-nowrap ${r.isDecomScop ? 'text-gray-500' : 'text-gray-300'}`}>{r.gc}</td>
                      <td className={`p-2 whitespace-nowrap ${r.isDecomScop ? 'text-gray-500' : 'text-gray-300'}`}>{r.sogName}</td>
                      <td className={`p-2 whitespace-nowrap ${r.isDecomScop ? 'text-gray-500' : 'text-gray-300'}`}>{r.spoNumber || '—'}</td>
                      <td className={`p-2 whitespace-nowrap ${r.isDecomScop ? 'text-gray-500' : 'text-gray-300'}`}>{fmtMoney(r.spoValue)}</td>
                      <td className={`p-2 whitespace-nowrap ${r.isDecomScop ? 'text-gray-500' : 'text-gray-300'}`}>{r.triggerDate || '—'}</td>
                      <td className="p-2"><StatusChip status={r.status} /></td>
                      <td className={`p-2 whitespace-nowrap ${r.isDecomScop ? 'text-gray-500' : 'text-gray-300'}`}>{r.grDate || '—'}</td>
                    </tr>
                  ))}
                  {displayRows.length === 0 && (
                    <tr><td colSpan={9} className="p-6 text-center text-gray-500">No rows match the current filters</td></tr>
                  )}
                </tbody>
              </table>
            </div>

            {groups.awareness.length > 0 && (
              <p className="text-gray-600 text-xs mt-3">👁 {groups.awareness.length} Decom/SCOP rows totaling {fmtMoneyShort(awarenessTotalValue)} — awareness only, not CJ's action</p>
            )}
          </>
        )}
      </div>
    </div>
  )
}
