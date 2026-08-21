'use client'

export const dynamic = 'force-dynamic'

import { useState, useEffect, useMemo } from 'react'
import { GC_CONFIG } from '../lib/gcConfig'
import {
  GrRow, GrTileFilter, GrSortOption, GR_SORT_OPTIONS,
  loadGrRows, groupGrRows, sortGrRowsBy, computeGrBreakdown, rowsForTileFilter,
  buildGrEmailMailto, fmtMoney, fmtMoneyShort,
} from '../lib/grTracker'
import BackToDashboard from '../components/BackToDashboard'
import { EmailSettings, DEFAULT_EMAIL, loadEmailSettings } from '../lib/settings'

const TIER_CHECKBOX_OPTIONS: { value: string; label: string }[] = [
  { value: 'init20', label: 'Init 20% (MS15A)' },
  { value: '60', label: '60% (MS16A)' },
  { value: '70', label: '70% (MS16A)' },
  { value: '20', label: '20% Decom/SCOP' },
  { value: '30', label: '30% Decom/SCOP' },
  { value: 'CR', label: 'CR (SOG Name blank)' },
]
const ALL_TIER_VALUES = TIER_CHECKBOX_OPTIONS.map(t => t.value)

const ROW_TYPE_OPTIONS: { value: 'all' | 'base' | 'cr'; label: string }[] = [
  { value: 'all', label: 'Show All' },
  { value: 'base', label: 'Base POs Only' },
  { value: 'cr', label: 'CRs Only' },
]

const STATUS_OPTIONS = ['All', 'Ready to Release', 'Awaiting Trigger', 'GR Done'] as const
type StatusOption = typeof STATUS_OPTIONS[number]

function KpiTile({ emoji, label, value, sub, color, valueColor, active, onClick }: {
  emoji: string; label: string; value: string; sub: string; color: string; valueColor?: string; active?: boolean; onClick: () => void
}) {
  return (
    <div onClick={onClick}
      className={`bg-gray-900 rounded-xl border p-4 text-center cursor-pointer hover:border-blue-500 hover:bg-gray-800 transition-all ${active ? 'ring-2 ring-blue-500' : ''} ${color}`}>
      <p className="text-2xl mb-1">{emoji}</p>
      <p className={`text-2xl font-bold ${valueColor || 'text-white'}`}>{value}</p>
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
  const [pmFilter, setPmFilter] = useState('CJ')
  const [gcFilter, setGcFilter] = useState('ALL')
  const [tierFilters, setTierFilters] = useState<Set<string>>(new Set(ALL_TIER_VALUES))
  const [rowTypeFilter, setRowTypeFilter] = useState<'all' | 'base' | 'cr'>('all')
  const [statusFilter, setStatusFilter] = useState<StatusOption>('All')
  const [sortBy, setSortBy] = useState<GrSortOption>('trigger')
  const [specialFilter, setSpecialFilter] = useState<GrTileFilter>(null)
  const [emailGroups, setEmailGroups] = useState<{ gc: string; count: number; mailto: string }[] | null>(null)
  const [emailSettings, setEmailSettings] = useState<EmailSettings>(DEFAULT_EMAIL)

  useEffect(() => {
    loadEmailSettings().then(setEmailSettings)
  }, [])

  useEffect(() => {
    const load = async () => {
      const grRows = await loadGrRows()
      setRows(grRows)
      setLoaded(true)
    }
    load()
  }, [])

  const pmOptions = useMemo(() => {
    const set = new Set<string>()
    rows.forEach(r => { if (r.nokiaPm) set.add(r.nokiaPm) })
    return ['ALL', ...Array.from(set).sort()]
  }, [rows])

  const pmFiltered = pmFilter === 'ALL' ? rows : rows.filter(r => r.nokiaPm === pmFilter)
  const breakdown = computeGrBreakdown(pmFiltered)

  const selectTile = (filter: GrTileFilter) => {
    setSpecialFilter(prev => prev === filter ? null : filter)
    setGcFilter('ALL')
    setTierFilters(new Set(ALL_TIER_VALUES))
    setRowTypeFilter('all')
    setStatusFilter('All')
  }

  const toggleTier = (value: string) => {
    setSpecialFilter(null)
    setTierFilters(prev => {
      const next = new Set(prev)
      if (next.has(value)) next.delete(value)
      else next.add(value)
      return next
    })
  }

  let displayRows: GrRow[]
  if (specialFilter) {
    displayRows = rowsForTileFilter(pmFiltered, specialFilter)
  } else {
    displayRows = pmFiltered.filter(r => {
      if (gcFilter !== 'ALL' && r.gc !== gcFilter) return false
      if (statusFilter !== 'All' && r.status !== statusFilter) return false
      return true
    })
  }
  // Row type toggle + tier checkboxes filter instantly on top of whatever's displayed,
  // regardless of how the row set got there (dropdowns or a tile click).
  if (rowTypeFilter !== 'all') {
    displayRows = displayRows.filter(r => r.rowType === rowTypeFilter)
  }
  if (tierFilters.size < ALL_TIER_VALUES.length) {
    displayRows = displayRows.filter(r => r.tier != null && tierFilters.has(r.tier))
  }
  displayRows = sortGrRowsBy(displayRows, sortBy)

  const generateEmails = () => {
    const groups = groupGrRows(rows)
    const byGc = new Map<string, GrRow[]>()
    groups.ready.forEach(r => {
      if (!byGc.has(r.gc)) byGc.set(r.gc, [])
      byGc.get(r.gc)!.push(r)
    })
    const out: { gc: string; count: number; mailto: string }[] = []
    byGc.forEach((gcRows, gc) => {
      out.push({ gc, count: gcRows.length, mailto: buildGrEmailMailto(gc, gcRows, { financeEmails: emailSettings.financeEmails, gcContactEmails: emailSettings.gcContactEmails }) })
    })
    out.sort((a, b) => a.gc.localeCompare(b.gc))
    setEmailGroups(out)
  }

  const readyCount = groupGrRows(rows).ready.length

  return (
    <div className="min-h-screen bg-gray-950 text-white p-6">
      <div className="max-w-7xl mx-auto">

        <BackToDashboard />

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
            {/* PM Filter */}
            {pmOptions.length > 1 && (
              <div className="flex gap-2 mb-4 flex-wrap items-center">
                <span className="text-gray-500 text-xs font-semibold">Nokia PM:</span>
                {pmOptions.map(pm => (
                  <button key={pm} onClick={() => setPmFilter(pm)}
                    className={`px-4 py-1.5 rounded-lg text-sm font-semibold transition-all ${pmFilter === pm ? 'bg-blue-600 text-white' : 'bg-gray-800 text-gray-300 hover:bg-gray-700'}`}>
                    {pm}
                  </button>
                ))}
              </div>
            )}

            {/* KPI Tiles — Row 1: Volume */}
            <div className="grid grid-cols-2 gap-3 mb-3">
              <KpiTile emoji="📄" label="Total Base POs" value={String(breakdown.totalBasePOs)} sub="valid SPO + SOG rows"
                color="border-gray-700" active={specialFilter === 'totalBasePOs'}
                onClick={() => selectTile('totalBasePOs')} />
              <KpiTile emoji="🔧" label="Total CRs" value={String(breakdown.totalCRs)} sub="SPO issued, no SOG tier"
                color="border-gray-700" active={specialFilter === 'totalCRs'}
                onClick={() => selectTile('totalCRs')} />
            </div>

            {/* KPI Tiles — Row 2: Value */}
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-6">
              <KpiTile emoji="✅" label="Base PO Value GR'd" value={fmtMoneyShort(breakdown.basePOValueGRd)} sub="GR Date populated"
                color="border-green-700" valueColor="text-green-400" active={specialFilter === 'basePOGRd'}
                onClick={() => selectTile('basePOGRd')} />
              <KpiTile emoji="⏳" label="Base PO Value Pending" value={fmtMoneyShort(breakdown.basePOValuePending)} sub="trigger met, GR Date blank"
                color="border-orange-700" valueColor="text-orange-400" active={specialFilter === 'basePOPending'}
                onClick={() => selectTile('basePOPending')} />
              <KpiTile emoji="✅" label="CR Value GR'd" value={fmtMoneyShort(breakdown.crValueGRd)} sub="GR Date populated"
                color="border-green-700" valueColor="text-green-400" active={specialFilter === 'crGRd'}
                onClick={() => selectTile('crGRd')} />
              <KpiTile emoji="⏳" label="CR Value Pending" value={fmtMoneyShort(breakdown.crValuePending)} sub="GR Date blank"
                color="border-orange-700" valueColor="text-orange-400" active={specialFilter === 'crPending'}
                onClick={() => selectTile('crPending')} />
            </div>

            {/* Filter Bar */}
            <div className="flex flex-wrap gap-3 mb-4 items-center">
              <select value={gcFilter} onChange={(e) => { setGcFilter(e.target.value); setSpecialFilter(null) }}
                className="bg-gray-800 border border-gray-600 text-white text-sm rounded px-3 py-2">
                <option value="ALL">All GCs</option>
                {GC_CONFIG.map(cfg => <option key={cfg.gc} value={cfg.gc}>{cfg.gc}</option>)}
              </select>
              <select value={statusFilter} onChange={(e) => { setStatusFilter(e.target.value as StatusOption); setSpecialFilter(null) }}
                className="bg-gray-800 border border-gray-600 text-white text-sm rounded px-3 py-2">
                {STATUS_OPTIONS.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
              <select value={sortBy} onChange={(e) => setSortBy(e.target.value as GrSortOption)}
                className="bg-gray-800 border border-gray-600 text-white text-sm rounded px-3 py-2">
                {GR_SORT_OPTIONS.map(o => <option key={o.value} value={o.value}>Sort: {o.label}</option>)}
              </select>
              {specialFilter && (
                <button onClick={() => setSpecialFilter(null)}
                  className="text-gray-400 hover:text-white text-xs underline">✕ Clear tile filter</button>
              )}
              <span className="text-gray-500 text-xs ml-auto">{displayRows.length} rows</span>
            </div>

            {/* Tier Filter — multi-select */}
            <div className="flex flex-wrap gap-3 items-center bg-gray-900 border border-gray-700 rounded-lg px-3 py-2 mb-4">
              <span className="text-gray-500 text-xs font-semibold">Tier:</span>
              {TIER_CHECKBOX_OPTIONS.map(t => (
                <label key={t.value} className="flex items-center gap-1.5 text-xs text-gray-300 cursor-pointer">
                  <input type="checkbox" checked={tierFilters.has(t.value)}
                    onChange={() => toggleTier(t.value)}
                    className="w-3.5 h-3.5 cursor-pointer accent-blue-500" />
                  {t.label}
                </label>
              ))}
            </div>

            {/* Batch Email */}
            <div className="mb-4">
              <button onClick={generateEmails} disabled={readyCount === 0}
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

            {/* Row Type Toggle */}
            <div className="flex gap-2 mb-3">
              {ROW_TYPE_OPTIONS.map(o => (
                <button key={o.value} onClick={() => setRowTypeFilter(o.value)}
                  className={`px-4 py-1.5 rounded-lg text-sm font-semibold transition-all ${rowTypeFilter === o.value ? 'bg-blue-600 text-white' : 'bg-gray-800 text-gray-300 hover:bg-gray-700'}`}>
                  {o.label}
                </button>
              ))}
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
                    <tr key={`${r.hop}-${r.sogName}-${r.spoNumber}-${i}`}
                      className={`border-t border-gray-800 ${r.isDecomScop ? 'bg-gray-950' : 'bg-gray-900'}`}>
                      <td className="p-2 font-semibold whitespace-nowrap">
                        {r.isDecomScop && <span className="mr-1">👁</span>}
                        <span className={r.isDecomScop ? 'text-gray-500' : 'text-white'}>{r.hopDisplay}</span>
                      </td>
                      <td className={`p-2 whitespace-nowrap ${r.isDecomScop ? 'text-gray-600' : 'text-gray-400'}`}>{r.pathId || '—'}</td>
                      <td className={`p-2 whitespace-nowrap ${r.isDecomScop ? 'text-gray-500' : 'text-gray-300'}`}>{r.gc}</td>
                      <td className={`p-2 whitespace-nowrap ${r.isDecomScop ? 'text-gray-500' : 'text-gray-300'}`}>{r.sogName || 'CR'}</td>
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
          </>
        )}
      </div>
    </div>
  )
}
