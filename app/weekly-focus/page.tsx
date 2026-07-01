'use client'

export const dynamic = 'force-dynamic'

import { useState, useEffect, useCallback } from 'react'
import { supabase, loadTrackerSnapshot } from '../lib/supabase'

interface HOP {
  hop: string
  gc: string
  cm: string
  nokiaPm: string
  regionPm: string
  ms15f: string
  ms15a: string
  ms16f: string
  ms16a: string
  mss: string
  powerUp: string
  hasNtp: boolean
  hasMat: boolean
  ntpOwner: string
  ntpWaitingOn: string
  matForecast: string
  matReceived: string
  gcPickup: boolean
  wpApproved: boolean
  vendorWindow: string
  daysOut: number | null
  daysElapsed: number | null
  inProgress: boolean
  complete: boolean
  over18d: boolean
}

interface Action {
  id: string
  hop_name: string
  action_text: string
  action_type: string
  created_at: string
  completed: boolean
  completed_at: string | null
  auto_completed: boolean
  source_field: string | null
}

function parseDateAny(val: unknown): Date | null {
  if (!val) return null
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val
  if (typeof val === 'number') {
    const d = new Date((val - 25569) * 86400 * 1000)
    return isNaN(d.getTime()) ? null : d
  }
  const d = new Date(String(val))
  return isNaN(d.getTime()) ? null : d
}

function parseDate(val: unknown): Date | null {
  const d = parseDateAny(val)
  if (!d) return null
  return d.getFullYear() >= 2025 ? d : null
}

function fmtDate(d: Date | null): string {
  if (!d) return ''
  return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`
}

function fmtDM(d: Date | null): string {
  if (!d) return ''
  return `${d.getMonth() + 1}/${d.getDate()}`
}

function daysBetween(a: Date, b: Date): number {
  return Math.round((b.getTime() - a.getTime()) / (1000 * 60 * 60 * 24))
}

const ACTION_TYPES = [
  'Chase NTP',
  'Chase Material',
  'Chase GC',
  'Chase CM',
  'Chase Viaero',
  'Chase ITW/Samsung',
  'Update Tracker',
  'Review/Monitor',
  'Other'
]

const AUTO_COMPLETE_MAP: Record<string, string[]> = {
  'NTP A': ['Chase NTP'],
  'Material Received A ': ['Chase Material'],
  'MS16 Implementation Ends A': ['Chase GC', 'Update Tracker'],
  'GC Material Pick-up (A)': ['Chase GC'],
  'MSS Completed NMS Ready ': ['Update Tracker'],
}

const newActionTextRef = { current: {} as Record<string, string> }
const newActionTypeRef = { current: {} as Record<string, string> }

interface HopRowProps {
  h: HOP
  showElapsed: boolean
  isExpanded: boolean
  hopAllActions: Action[]
  hopOpenActions: Action[]
  mode: string
  onToggle: (hop: string) => void
  onToggleAction: (action: Action) => void
  onAddAction: (hop: string) => void
  actionText: string
  actionType: string
  onActionTextChange: (hop: string, val: string) => void
  onActionTypeChange: (hop: string, val: string) => void
}

function HopRow({ h, showElapsed, isExpanded, hopAllActions, hopOpenActions, mode,
  onToggle, onToggleAction, onAddAction, actionText, actionType,
  onActionTextChange, onActionTypeChange }: HopRowProps) {
  const hasBlocker = !h.hasNtp || !h.hasMat || h.vendorWindow.includes('🔴') || h.over18d

  return (
    <div className={`rounded-lg border mb-2 overflow-hidden ${h.over18d ? 'border-red-700' : hasBlocker ? 'border-yellow-800' : 'border-gray-700'}`}>
      <div
        className={`flex items-center justify-between p-3 cursor-pointer ${h.over18d ? 'bg-red-950' : hasBlocker ? 'bg-yellow-950' : 'bg-gray-900'} hover:bg-opacity-80`}
        onClick={() => onToggle(h.hop)}>
        <div className="flex items-center gap-3 flex-wrap">
          <span className="font-bold text-white text-sm">{h.hop}</span>
          <span className="text-gray-400 text-xs">{mode === 'gc' ? h.gc : h.cm}</span>
          {showElapsed
            ? <span className={`text-xs font-bold ${h.over18d ? 'text-red-400' : 'text-green-400'}`}>{h.daysElapsed}d elapsed</span>
            : h.daysOut !== null && <span className={`text-xs font-bold ${h.daysOut <= 7 ? 'text-red-400' : 'text-yellow-400'}`}>{h.daysOut}d out</span>
          }
          {!h.hasNtp && <span className="bg-red-900 text-red-200 text-xs px-2 py-0.5 rounded-full">NTP ✗</span>}
          {!h.hasMat && <span className="bg-orange-900 text-orange-200 text-xs px-2 py-0.5 rounded-full">Mat ✗</span>}
          {h.vendorWindow.includes('🔴') && <span className="bg-red-900 text-red-200 text-xs px-2 py-0.5 rounded-full">Vendor ⚠️</span>}
          {h.over18d && <span className="bg-red-800 text-red-200 text-xs px-2 py-0.5 rounded-full">⚠️ Over 18d</span>}
          {hopOpenActions.length > 0 && (
            <span className="bg-blue-900 text-blue-200 text-xs px-2 py-0.5 rounded-full">
              {hopOpenActions.length} open action{hopOpenActions.length > 1 ? 's' : ''}
            </span>
          )}
        </div>
        <span className="text-gray-500 text-xs whitespace-nowrap">{isExpanded ? '▲' : '▼'}</span>
      </div>

      {isExpanded && (
        <div className="bg-gray-950 border-t border-gray-800 p-4">
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-4 text-xs">
            <div><span className="text-gray-500">GC:</span> <span className="text-white">{h.gc || '—'}</span></div>
            <div><span className="text-gray-500">CM:</span> <span className="text-white">{h.cm || '—'}</span></div>
            <div><span className="text-gray-500">FC Start:</span> <span className="text-white">{h.ms15f || '—'}</span></div>
            <div><span className="text-gray-500">FC End:</span> <span className="text-white">{h.ms16f || '—'}</span></div>
            <div><span className="text-gray-500">Act Start:</span> <span className="text-white">{h.ms15a || '—'}</span></div>
            <div><span className="text-gray-500">Act End:</span> <span className="text-white">{h.ms16a || '—'}</span></div>
            <div><span className="text-gray-500">NTP:</span> <span className={h.hasNtp ? 'text-green-400' : 'text-red-400'}>{h.hasNtp ? '✓ Confirmed' : '✗ Pending'}</span></div>
            <div><span className="text-gray-500">Material:</span> <span className={h.hasMat ? 'text-green-400' : 'text-red-400'}>{h.hasMat ? '✓ Received' : '✗ Pending'}</span></div>
            {!h.hasNtp && h.ntpWaitingOn && <div className="col-span-2"><span className="text-gray-500">NTP Waiting On:</span> <span className="text-yellow-300">{h.ntpWaitingOn}</span></div>}
            {!h.hasMat && h.matForecast && <div><span className="text-gray-500">Mat Forecast:</span> <span className="text-yellow-300">{h.matForecast}</span></div>}
            {h.vendorWindow !== '✅ Clear' && <div className="col-span-2"><span className="text-gray-500">Vendor:</span> <span className="text-red-400">{h.vendorWindow}</span></div>}
            {h.mss && <div><span className="text-gray-500">MSS:</span> <span className="text-green-400">{h.mss}</span></div>}
            {h.powerUp && <div><span className="text-gray-500">Power-Up:</span> <span className="text-green-400">{h.powerUp}</span></div>}
          </div>

          <div className="border-t border-gray-800 pt-3">
            <h4 className="text-xs font-bold text-gray-400 mb-2">ACTIONS</h4>
            {hopAllActions.length > 0 && (
              <div className="space-y-1 mb-3">
                {hopAllActions.map(action => (
                  <div key={action.id} className={`flex items-center gap-2 text-xs p-2 rounded ${action.completed ? 'bg-gray-900 opacity-50' : 'bg-gray-800'}`}>
                    <input type="checkbox" checked={action.completed}
                      onChange={() => onToggleAction(action)}
                      className="w-4 h-4 cursor-pointer accent-green-500" />
                    <span className={`flex-1 ${action.completed ? 'line-through text-gray-500' : 'text-white'}`}>{action.action_text}</span>
                    <span className="text-gray-600 whitespace-nowrap">{action.action_type}</span>
                    {action.auto_completed && <span className="text-green-600 text-xs">auto ✓</span>}
                    {action.completed && action.completed_at && (
                      <span className="text-gray-600 text-xs whitespace-nowrap">
                        {new Date(action.completed_at).toLocaleDateString('en-US', { month: 'numeric', day: 'numeric' })}
                      </span>
                    )}
                  </div>
                ))}
              </div>
            )}
            <div className="flex gap-2 flex-wrap">
              <input
                type="text"
                placeholder="Add action..."
                value={actionText}
                onChange={(e) => onActionTextChange(h.hop, e.target.value)}
                onKeyDown={(e) => { if (e.key === 'Enter') onAddAction(h.hop) }}
                className="flex-1 min-w-48 bg-gray-800 text-white text-xs rounded px-2 py-1 border border-gray-600 focus:outline-none focus:border-blue-500"
              />
              <select
                value={actionType}
                onChange={(e) => onActionTypeChange(h.hop, e.target.value)}
                className="bg-gray-800 text-gray-300 text-xs rounded px-2 py-1 border border-gray-600 focus:outline-none">
                {ACTION_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
              </select>
              <button onClick={() => onAddAction(h.hop)}
                className="bg-blue-700 hover:bg-blue-600 text-white text-xs px-3 py-1 rounded font-semibold">
                + Add
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

export default function WeeklyFocusPage() {
  const [hops, setHops] = useState<HOP[]>([])
  const [loaded, setLoaded] = useState(false)
  const [fileName, setFileName] = useState('')
  const [snapshotTime, setSnapshotTime] = useState('')
  const [mode, setMode] = useState<'gc' | 'cm'>('gc')
  const [filterCM, setFilterCM] = useState('All')
  const [filterGC, setFilterGC] = useState('All')
  const [actions, setActions] = useState<Action[]>([])
  const [newActionText, setNewActionText] = useState<Record<string, string>>({})
  const [newActionType, setNewActionType] = useState<Record<string, string>>({})
  const [expandedHops, setExpandedHops] = useState<Set<string>>(new Set())
  const today = new Date()

  // Load actions from Supabase
  const loadActions = async () => {
    const { data } = await supabase
      .from('hop_actions')
      .select('*')
      .order('created_at', { ascending: false })
    if (data) setActions(data)
  }

  // Auto-complete actions based on tracker data
  const autoCompleteActions = useCallback(async (parsedHops: HOP[]) => {
    const { data: openActions } = await supabase
      .from('hop_actions')
      .select('*')
      .eq('completed', false)
    if (!openActions) return

    for (const action of openActions) {
      const hop = parsedHops.find(h => h.hop === action.hop_name)
      if (!hop) continue

      let shouldComplete = false
      let sourceField = ''

      if (action.action_type === 'Chase NTP' && hop.hasNtp) {
        shouldComplete = true; sourceField = 'NTP A'
      }
      if (action.action_type === 'Chase Material' && hop.hasMat) {
        shouldComplete = true; sourceField = 'Material Received A'
      }
      if (action.action_type === 'Chase GC' && hop.gcPickup) {
        shouldComplete = true; sourceField = 'GC Material Pick-up (A)'
      }
      if (action.action_type === 'Chase CM' && hop.complete) {
        shouldComplete = true; sourceField = 'MS16 Implementation Ends A'
      }

      if (shouldComplete) {
        await supabase.from('hop_actions').update({
          completed: true,
          completed_at: new Date().toISOString(),
          auto_completed: true,
          source_field: sourceField
        }).eq('id', action.id)
      }
    }
    await loadActions()
  }, [])

  useEffect(() => {
    const loadFromSnapshot = async () => {
      const snap = await loadTrackerSnapshot()
      if (!snap) return
      setSnapshotTime(
        new Date(snap.uploaded_at).toLocaleDateString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric' }) +
        ' at ' + new Date(snap.uploaded_at).toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })
      )
      setFileName(snap.filename)
      processRows(snap.data)
    }
    loadFromSnapshot()
    loadActions()
  }, [])

  const processRows = useCallback((rows: unknown[][]) => {
    let headerRow = -1
    for (let i = 0; i < 10; i++) {
      if ((rows[i] as unknown[])?.some(c => String(c).trim() === 'HOP')) { headerRow = i; break }
    }
    if (headerRow === -1) return

    const headers = rows[headerRow] as string[]
    const col = (name: string) => headers.findIndex(h => String(h).trim() === name)

    const hopCol      = col('HOP')
    const gcCol       = col('General Contractor')
    const newCmCol    = col('New CM')
    const nokiaPmCol  = col('Nokia PM')
    const regionPmCol = col('Region PM')
    const don444Col   = col('DON 444')
    const ms15fCol    = col('MS15 Implementation Start F')
    const ms15aCol    = col('MS15 Implementation Start A')
    const ms16fCol    = col('MS16 Implementation Ends F')
    const ms16aCol    = col('MS16 Implementation Ends A')
    const mssCol      = col('MSS Completed NMS Ready ')
    const powerCol    = col('Power-Up Completion')
    const ntpCol      = col('NTP A')
    const matCol      = headers.findIndex(h => String(h).trim().replace(/\s+$/, '') === 'Material Received A'.trim())
    const matFcCol    = col('Material Forecast +4ish')
    const wpCol       = col('Work Package Approved in QB')
    const pickupCol   = col('GC Material Pick-up (A)')
    const ntpOwnCol   = col('NTP Action Owner')
    const ntpWaitCol  = col('NTP is waiting on')
    const itwSCol     = col('ITW Schedule Start')
    const itwECol     = col('ITW Schedule Complete')
    const ssSCol      = col('Samsung Schedule Start')
    const ssECol      = col('Samsung Schedule Complete')

    const hopRows = new Map<string, unknown[][]>()
    for (let i = headerRow + 1; i < rows.length; i++) {
      const row = rows[i] as unknown[]
      const don = String(row[don444Col] || '').trim().toUpperCase()
      if (don !== 'DON 444') continue
      const hop = String(row[hopCol] || '').trim()
      if (!hop || hop === 'undefined') continue
      if (!hopRows.has(hop)) hopRows.set(hop, [])
      hopRows.get(hop)!.push(row)
    }

    const parsed: HOP[] = []
    hopRows.forEach((rows2, hop) => {
      const row    = rows2[0]
      const ms15f  = parseDateAny(row[ms15fCol])
      const ms15a  = parseDate(row[ms15aCol])
      const ms16f  = parseDateAny(row[ms16fCol])
      const ms16a  = parseDate(row[ms16aCol])
      const ntpDate = parseDate(row[ntpCol])
      const matDate = parseDateAny(row[matCol])
      const wpDate  = parseDateAny(row[wpCol])
      const pickupD = parseDateAny(row[pickupCol])
      const mssDate = parseDateAny(row[mssCol])
      const powerDate = parseDateAny(row[powerCol])

      const hasNtp     = !!(ntpDate && ntpDate.getFullYear() >= 2025)
      const hasMat     = !!(matDate && matDate.getFullYear() >= 2020)
      const wpApproved = !!wpDate
      const gcPickup   = !!pickupD
      const started    = !!ms15a
      const complete   = !!ms16a
      const inProgress = started && !complete
      const daysOut    = ms15f ? daysBetween(today, ms15f) : null
      const daysElapsed = inProgress && ms15a ? daysBetween(ms15a, today) : null
      const over18d    = inProgress && (daysElapsed ?? 0) > 18

      // Vendor window
      const parts: string[] = []
      rows2.forEach(r => {
        const itwS = parseDateAny(r[itwSCol]); const itwE = parseDateAny(r[itwECol])
        const ssS  = parseDateAny(r[ssSCol]);  const ssE  = parseDateAny(r[ssECol])
        if (ms15f) {
          if (itwS && itwE && itwS <= ms15f && ms15f <= itwE) parts.push(`🔴 ITW thru ${fmtDM(itwE)}`)
          if (ssS && ssE && ssS <= ms15f && ms15f <= ssE) parts.push(`🔴 Samsung thru ${fmtDM(ssE)}`)
        }
      })
      const vendorWindow = parts.length > 0 ? parts.join(' | ') : '✅ Clear'

      parsed.push({
        hop,
        gc:          String(row[gcCol] || '').trim(),
        cm:          String(row[newCmCol] || '').trim(),
        nokiaPm:     String(row[nokiaPmCol] || '').trim(),
        regionPm:    String(row[regionPmCol] || '').trim(),
        ms15f:       fmtDate(ms15f),
        ms15a:       fmtDate(ms15a),
        ms16f:       fmtDate(ms16f),
        ms16a:       fmtDate(ms16a),
        mss:         fmtDate(mssDate),
        powerUp:     fmtDate(powerDate),
        hasNtp, hasMat, wpApproved, gcPickup,
        ntpOwner:    String(row[ntpOwnCol] || '').trim(),
        ntpWaitingOn: String(row[ntpWaitCol] || '').trim(),
        matForecast: fmtDate(parseDateAny(row[matFcCol])),
        matReceived: hasMat ? fmtDate(matDate) : '',
        vendorWindow, daysOut, daysElapsed, inProgress, complete, over18d
      })
    })

    setHops(parsed)
    setLoaded(true)
    autoCompleteActions(parsed)
  }, [today, autoCompleteActions])

  // Filter hops based on mode
  const filteredHops = hops.filter(h => {
    const pmMatch = mode === 'gc'
      ? h.nokiaPm.toUpperCase() === 'CJ'
      : h.regionPm.toUpperCase() === 'CJ'
    const cmMatch = filterCM === 'All' || h.cm === filterCM
    const gcMatch = filterGC === 'All' || h.gc === filterGC
    return pmMatch && cmMatch && gcMatch
  })

  // Sections
  const needsAttention = filteredHops.filter(h =>
    !h.inProgress && !h.complete &&
    h.daysOut !== null && h.daysOut >= 0 && h.daysOut <= 7 &&
    ((!h.hasNtp || !h.hasMat) || h.vendorWindow.includes('🔴'))
  ).sort((a, b) => (a.daysOut ?? 0) - (b.daysOut ?? 0))

  const active = filteredHops.filter(h => h.inProgress)
    .sort((a, b) => (b.daysElapsed ?? 0) - (a.daysElapsed ?? 0))

  const thisWeekReady = filteredHops.filter(h =>
    !h.inProgress && !h.complete &&
    h.daysOut !== null && h.daysOut >= 0 && h.daysOut <= 7 &&
    h.hasNtp && h.hasMat && !h.vendorWindow.includes('🔴')
  ).sort((a, b) => (a.daysOut ?? 0) - (b.daysOut ?? 0))

  const next2Weeks = filteredHops.filter(h =>
    !h.inProgress && !h.complete &&
    h.daysOut !== null && h.daysOut > 7 && h.daysOut <= 14
  ).sort((a, b) => (a.daysOut ?? 0) - (b.daysOut ?? 0))

  const ntpUrgent = filteredHops.filter(h =>
    !h.hasNtp && !h.complete &&
    h.daysOut !== null && h.daysOut <= 14
  ).sort((a, b) => (a.daysOut ?? 0) - (b.daysOut ?? 0))

  // Unique CMs and GCs for filters
  const uniqueCMs = ['All', ...Array.from(new Set(filteredHops.map(h => h.cm).filter(Boolean))).sort()]
  const uniqueGCs = ['All', ...Array.from(new Set(filteredHops.map(h => h.gc).filter(Boolean))).sort()]

  // Actions helpers
  const hopActions = (hop: string) => actions.filter(a => a.hop_name === hop)
  const openActions = (hop: string) => hopActions(hop).filter(a => !a.completed)
  const totalOpenActions = actions.filter(a => !a.completed).length

  const addAction = async (hop: string) => {
    const text = newActionText[hop]?.trim()
    const type = newActionType[hop] || 'Other'
    if (!text) return
    const { data } = await supabase.from('hop_actions').insert({
      hop_name: hop, action_text: text, action_type: type
    }).select()
    if (data) {
      setActions(prev => [data[0], ...prev])
      setNewActionText(prev => ({ ...prev, [hop]: '' }))
    }
  }

  const toggleAction = async (action: Action) => {
    const { data } = await supabase.from('hop_actions').update({
      completed: !action.completed,
      completed_at: !action.completed ? new Date().toISOString() : null,
      auto_completed: false
    }).eq('id', action.id).select()
    if (data) {
      setActions(prev => prev.map(a => a.id === action.id ? data[0] : a))
    }
  }

  const toggleHop = (hop: string) => {
    setExpandedHops(prev => {
      const next = new Set(prev)
      if (next.has(hop)) next.delete(hop)
      else next.add(hop)
      return next
    })
  }

  const Section = ({ title, rows, showElapsed = false, color = 'gray' }: {
    title: string, rows: HOP[], showElapsed?: boolean, color?: string
  }) => {
    const bgMap: Record<string, string> = {
      red: 'bg-red-900 border-red-700',
      yellow: 'bg-yellow-900 border-yellow-700',
      blue: 'bg-blue-900 border-blue-700',
      orange: 'bg-orange-900 border-orange-700',
      gray: 'bg-gray-800 border-gray-600'
    }
    return (
      <div className="mb-6">
        <div className={`flex items-center justify-between px-4 py-2 rounded-t-lg border ${bgMap[color]}`}>
          <h3 className="text-sm font-bold text-white">{title}</h3>
          <span className="text-white text-xs font-bold">{rows.length} HOPs</span>
        </div>
        {rows.length === 0
          ? <div className="bg-gray-900 rounded-b-lg border border-gray-700 border-t-0 p-4 text-center">
              <p className="text-green-400 text-sm">✅ Nothing here — you're ahead</p>
            </div>
          : <div className="bg-gray-950 rounded-b-lg border border-gray-700 border-t-0 p-3">
              {rows.map(h => (
                <HopRow
                  key={h.hop}
                  h={h}
                  showElapsed={showElapsed}
                  isExpanded={expandedHops.has(h.hop)}
                  hopAllActions={actions.filter(a => a.hop_name === h.hop)}
                  hopOpenActions={actions.filter(a => a.hop_name === h.hop && !a.completed)}
                  mode={mode}
                  onToggle={toggleHop}
                  onToggleAction={toggleAction}
                  onAddAction={addAction}
                  actionText={newActionText[h.hop] || ''}
                  actionType={newActionType[h.hop] || 'Other'}
                  onActionTextChange={(hop, val) => setNewActionText(prev => ({ ...prev, [hop]: val }))}
                  onActionTypeChange={(hop, val) => setNewActionType(prev => ({ ...prev, [hop]: val }))}
                />
              ))}
            </div>
        }
      </div>
    )
  }

  return (
    <div className="min-h-screen bg-gray-950 text-white p-6">
      <div className="max-w-5xl mx-auto">

        {/* Header */}
        <div className="flex items-center justify-between mb-6">
          <div>
            <h1 className="text-3xl font-bold">Weekly Focus</h1>
            <p className="text-gray-400 mt-1">Your work area — actions, reminders, and what needs attention this week</p>
          </div>
          {totalOpenActions > 0 && (
            <div className="bg-blue-900 border border-blue-600 rounded-lg px-4 py-2 text-center">
              <p className="text-blue-200 text-xs">Open Actions</p>
              <p className="text-white text-2xl font-bold">{totalOpenActions}</p>
            </div>
          )}
        </div>

        {/* Snapshot info */}
        {snapshotTime && (
          <div className="mb-4 bg-gray-900 border border-gray-700 rounded-lg px-4 py-2 flex items-center justify-between">
            <p className="text-green-400 text-sm font-semibold">📡 Live data from {snapshotTime}</p>
            <p className="text-gray-500 text-xs">{fileName} — {hops.length} HOPs total</p>
          </div>
        )}
        {!snapshotTime && (
          <div className="mb-4 bg-gray-900 border border-gray-700 rounded-lg px-4 py-8 text-center">
            <p className="text-gray-400">No tracker data — go to Dashboard to upload your tracker</p>
          </div>
        )}

        {loaded && (
          <>
            {/* Mode Toggle */}
            <div className="flex gap-3 mb-4">
              <button onClick={() => setMode('gc')}
                className={`px-5 py-2 rounded-lg text-sm font-semibold transition-all ${mode === 'gc' ? 'bg-blue-600 text-white' : 'bg-gray-800 text-gray-300 hover:bg-gray-700'}`}>
                🏗️ GC / PM Mode
              </button>
              <button onClick={() => setMode('cm')}
                className={`px-5 py-2 rounded-lg text-sm font-semibold transition-all ${mode === 'cm' ? 'bg-teal-600 text-white' : 'bg-gray-800 text-gray-300 hover:bg-gray-700'}`}>
                👤 CM Mode
              </button>
              <div className="text-gray-600 text-xs self-center">
                {mode === 'gc' ? 'Nokia PM = CJ — GC schedule & finances' : 'Region PM = CJ — CM field management'}
              </div>
            </div>

            {/* Filters */}
            <div className="flex gap-3 mb-6 flex-wrap">
              <div className="flex items-center gap-2">
                <span className="text-gray-500 text-xs">CM:</span>
                <select value={filterCM} onChange={(e) => setFilterCM(e.target.value)}
                  className="bg-gray-800 text-gray-300 text-xs rounded px-2 py-1 border border-gray-600">
                  {uniqueCMs.map(cm => <option key={cm} value={cm}>{cm}</option>)}
                </select>
              </div>
              <div className="flex items-center gap-2">
                <span className="text-gray-500 text-xs">GC:</span>
                <select value={filterGC} onChange={(e) => setFilterGC(e.target.value)}
                  className="bg-gray-800 text-gray-300 text-xs rounded px-2 py-1 border border-gray-600">
                  {uniqueGCs.map(gc => <option key={gc} value={gc}>{gc}</option>)}
                </select>
              </div>
              <div className="text-gray-500 text-xs self-center">
                {filteredHops.length} HOPs in view
              </div>
            </div>

            {/* KPI Row */}
            <div className="grid grid-cols-2 md:grid-cols-5 gap-3 mb-6">
              {[
                { label: 'Active', value: active.length, color: 'text-blue-400' },
                { label: 'Needs Attention', value: needsAttention.length, color: 'text-red-400' },
                { label: 'This Week Ready', value: thisWeekReady.length, color: 'text-green-400' },
                { label: 'Next 2 Weeks', value: next2Weeks.length, color: 'text-yellow-400' },
                { label: 'NTP Urgent', value: ntpUrgent.length, color: 'text-orange-400' },
              ].map(({ label, value, color }) => (
                <div key={label} className="bg-gray-900 rounded-xl border border-gray-700 p-3 text-center">
                  <p className="text-gray-500 text-xs">{label}</p>
                  <p className={`text-2xl font-bold ${color}`}>{value}</p>
                </div>
              ))}
            </div>

            {/* Sections */}
            <Section title="🔴 Needs Attention Now — Starts This Week With Blockers" rows={needsAttention} color="red" />
            <Section title="🔨 Active — Drive to Completion" rows={active} showElapsed color="blue" />
            <Section title="✅ Starting This Week — Confirm Ready" rows={thisWeekReady} color="gray" />
            <Section title="🟠 Starting Next 2 Weeks — Get Ahead Now" rows={next2Weeks} color="orange" />
            <Section title="🚦 NTP Urgent — Missing NTP ≤14 Days" rows={ntpUrgent} color="yellow" />
          </>
        )}

      </div>
    </div>
  )
}
