'use client'

export const dynamic = 'force-dynamic'

import { useState, useCallback, useEffect } from 'react'
import { loadTrackerSnapshot, supabase } from '../lib/supabase'
import { GC_CONFIG } from '../lib/gcConfig'
import BackToDashboard from '../components/BackToDashboard'
import {
  ThresholdSettings, DEFAULT_THRESHOLDS, loadThresholdSettings,
  ProgramSettings, DEFAULT_PROGRAM, loadProgramSettings, crewCountForGc,
} from '../lib/settings'

// A pull-in candidate more than this many days out goes in the Pull-In Queue
// (Scenario B) instead of being suggested as an immediate Scenario A pull-in.
// Not one of the 5 settings fields — kept as a separate local constant so
// tuning the ramp-up window doesn't silently change pull-in behavior too.
const PULL_IN_WINDOW = 30

// How much earlier another GC's crew has to be available before we bother
// surfacing an informational reassign note. Not specified by name in the
// spec ("significantly earlier") — chosen as a reasonable default.
const REASSIGN_THRESHOLD_DAYS = 14

interface HOP {
  hop: string
  gc: string
  pathId: string
  ms15f: Date | null
  ms15a: Date | null
  ms16f: Date | null
  ms16a: Date | null
  hasNtp: boolean
  hasMat: boolean
  vendorConflicts: { vendor: string; start: Date; end: Date }[]
}

type Recommendation = 'green' | 'pullIn' | 'pushOut' | 'vendorConflict' | 'pullInQueue' | 'onHold' | 'notReady'

interface Row {
  hop: string
  gc: string
  crew: number
  ms15f: Date | null
  ms15a: Date | null
  ms16f: Date | null
  daysElapsed: number | null
  hasNtp: boolean
  hasMat: boolean
  vendorConflict: string
  recommendation: Recommendation
  proposedStart: Date | null
  reassignNote: string
}

function parseDate(val: unknown): Date | null {
  if (!val) return null
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val
  if (typeof val === 'number') {
    const d = new Date((val - 25569) * 86400 * 1000)
    return isNaN(d.getTime()) ? null : d
  }
  const d = new Date(String(val))
  return isNaN(d.getTime()) ? null : d
}

function addDays(date: Date, days: number): Date {
  const d = new Date(date)
  d.setDate(d.getDate() + days)
  return d
}

function fmtDate(d: Date | null): string {
  if (!d) return ''
  return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`
}

function fmtShort(d: Date | null): string {
  if (!d) return ''
  return `${d.getMonth() + 1}/${d.getDate()}`
}

function daysBetween(a: Date, b: Date): number {
  return Math.round((b.getTime() - a.getTime()) / (1000 * 60 * 60 * 24))
}

// Vendor conflict is checked against MS15F only, and only ever called for
// pipeline HOPs — never in-progress ones.
function vendorConflictFor(h: HOP): string {
  if (!h.ms15f) return ''
  for (const c of h.vendorConflicts) {
    if (c.start <= h.ms15f && h.ms15f <= c.end) return `${c.vendor} on site thru ${fmtShort(c.end)}`
  }
  return ''
}

function isGreen(h: HOP): boolean {
  return h.hasNtp && h.hasMat && !vendorConflictFor(h)
}

interface GcResult {
  gc: string
  numCrews: number
  inProgress: Row[]
  pipeline: Row[]
  pullInQueue: Row[]
  held: Row[]
  rampUp: { triggered: boolean; greenCount: number; capacity: number; recommendedCrews: number }
}

function runSimulation(
  hops: HOP[],
  program: ProgramSettings,
  t: ThresholdSettings,
  holds: Set<string>,
  today: Date
): GcResult[] {
  const gcNames = GC_CONFIG.map(c => c.gc)

  // Earliest each GC's least-busy crew could take on a brand new HOP right
  // now — used only for the informational "FYI: other GC available" note.
  const gcEarliestFree: Record<string, Date> = {}

  const results: GcResult[] = gcNames.map(gc => {
    const numCrews = crewCountForGc(program, gc)
    const gcHops = hops.filter(h => h.gc?.trim().toLowerCase() === gc.trim().toLowerCase())

    const complete = (h: HOP) => !!h.ms16a
    const onHold = (h: HOP) => holds.has(h.hop)
    const inProgressHops = gcHops.filter(h => !complete(h) && !onHold(h) && !!h.ms15a)
    const pipelineHops = gcHops.filter(h => !complete(h) && !onHold(h) && !h.ms15a && !!h.ms15f)
    const heldHops = gcHops.filter(h => !complete(h) && onHold(h))

    // Pipeline HOPs sorted by MS15F ascending, then assigned to crews by
    // straight alternation (Crew 1, Crew 2, Crew 1, Crew 2...).
    const sortedPipeline = [...pipelineHops].sort((a, b) => (a.ms15f as Date).getTime() - (b.ms15f as Date).getTime())
    const crewOf = new Map<string, number>()
    sortedPipeline.forEach((h, i) => crewOf.set(h.hop, i % numCrews))

    // In-progress HOPs are always Crew 1 (index 0).
    const inProgressRows: Row[] = inProgressHops.map(h => {
      const elapsed = h.ms15a ? daysBetween(h.ms15a, today) : null
      return {
        hop: h.hop, gc, crew: 1,
        ms15f: h.ms15f, ms15a: h.ms15a, ms16f: h.ms16f,
        daysElapsed: elapsed,
        hasNtp: h.hasNtp, hasMat: h.hasMat,
        vendorConflict: '', // never flagged on in-progress HOPs
        recommendation: 'green',
        proposedStart: null,
        reassignNote: '',
      }
    })

    // Crew 1's starting availability accounts for whatever's already
    // in-progress; other crews start free as of today.
    const crewFree: Date[] = Array.from({ length: numCrews }, () => new Date(today))
    if (inProgressHops.length > 0) {
      const occupiedUntil = inProgressHops.reduce((latest, h) => {
        const end = h.ms16f || (h.ms15a ? addDays(h.ms15a, t.hopDuration) : today)
        return end > latest ? end : latest
      }, today)
      crewFree[0] = addDays(occupiedUntil, 1)
    }
    gcEarliestFree[gc] = crewFree.reduce((min, d) => (d < min ? d : min), crewFree[0])

    // Build each crew's queue: green HOPs first (they jump ahead), then
    // non-green HOPs — both groups keeping their relative MS15F order. This
    // is what implements both Scenario A (crew free earlier than MS15F) and
    // Scenario B (green HOP jumps ahead of non-green ones) with one pass.
    const pipelineRows: Row[] = []
    for (let crewIdx = 0; crewIdx < numCrews; crewIdx++) {
      const queueHops = sortedPipeline.filter(h => crewOf.get(h.hop) === crewIdx)
      const greenQueue = queueHops.filter(isGreen)
      const nonGreenQueue = queueHops.filter(h => !isGreen(h))
      const orderedQueue = [...greenQueue, ...nonGreenQueue]

      orderedQueue.forEach(h => {
        const originalStart = h.ms15f as Date
        const daysOut = daysBetween(today, originalStart)
        const green = isGreen(h)
        const vc = vendorConflictFor(h)
        const pushTriggered = !green && daysOut <= t.pushWindow && (!h.hasNtp || !h.hasMat)

        let proposedStart: Date
        if (green) {
          // Pulled in as early as the crew queue allows.
          proposedStart = crewFree[crewIdx]
        } else {
          const desired = pushTriggered ? addDays(originalStart, t.pushAmount) : originalStart
          proposedStart = desired > crewFree[crewIdx] ? desired : crewFree[crewIdx]
        }

        const proposedEnd = h.ms16f || addDays(proposedStart, t.hopDuration)
        crewFree[crewIdx] = addDays(proposedEnd, 1)

        let recommendation: Recommendation
        const pulledIn = proposedStart.getTime() < originalStart.getTime()
        if (pushTriggered) recommendation = 'pushOut'
        else if (vc) recommendation = 'vendorConflict'
        else if (pulledIn && daysOut <= PULL_IN_WINDOW) recommendation = 'pullIn'
        else if (green && daysOut > PULL_IN_WINDOW && pulledIn) recommendation = 'pullInQueue'
        else if (green) recommendation = 'green'
        else recommendation = 'notReady'

        // Informational reassign note — only for HOPs that aren't already
        // being pulled in / pushed, so it doesn't compete with an active
        // recommendation.
        let reassignNote = ''
        if (recommendation === 'notReady' || recommendation === 'green') {
          let bestGc = ''; let bestDate: Date | null = null
          gcNames.forEach(otherGc => {
            if (otherGc === gc) return
            const free = gcEarliestFree[otherGc]
            if (free && (!bestDate || free < bestDate)) { bestDate = free; bestGc = otherGc }
          })
          if (bestGc && bestDate && daysBetween(bestDate, proposedStart) >= REASSIGN_THRESHOLD_DAYS) {
            reassignNote = `FYI: ${bestGc} crew available ${fmtDate(bestDate)}`
          }
        }

        pipelineRows.push({
          hop: h.hop, gc, crew: crewIdx + 1,
          ms15f: h.ms15f, ms15a: h.ms15a, ms16f: h.ms16f,
          daysElapsed: null,
          hasNtp: h.hasNtp, hasMat: h.hasMat,
          vendorConflict: vc,
          recommendation,
          proposedStart,
          reassignNote,
        })
      })
    }

    const heldRows: Row[] = heldHops.map(h => ({
      hop: h.hop, gc, crew: crewOf.get(h.hop) != null ? (crewOf.get(h.hop) as number) + 1 : 0,
      ms15f: h.ms15f, ms15a: h.ms15a, ms16f: h.ms16f,
      daysElapsed: null,
      hasNtp: h.hasNtp, hasMat: h.hasMat,
      vendorConflict: '',
      recommendation: 'onHold',
      proposedStart: null,
      reassignNote: '',
    }))

    const pullInQueue = pipelineRows.filter(r => r.recommendation === 'pullInQueue')

    // Ramp-up: fully green HOPs starting within rampUpWindow days, compared
    // to current crew capacity.
    const greenSoon = pipelineHops.filter(h => {
      if (!isGreen(h) || !h.ms15f) return false
      const daysOut = daysBetween(today, h.ms15f)
      return daysOut >= 0 && daysOut <= t.rampUpWindow
    })
    const capacity = numCrews * t.rampUpThreshold
    const triggered = greenSoon.length > capacity
    const recommendedCrews = triggered ? Math.ceil((greenSoon.length - capacity) / t.rampUpThreshold) : 0

    return {
      gc, numCrews,
      inProgress: inProgressRows,
      pipeline: pipelineRows,
      pullInQueue,
      held: heldRows,
      rampUp: { triggered, greenCount: greenSoon.length, capacity, recommendedCrews },
    }
  })

  return results
}

const RECOMMENDATION_LABEL: Record<Recommendation, string> = {
  green: '✅ Green',
  pullIn: '⬆️ Pull In',
  pushOut: '⬇️ Push Out',
  vendorConflict: '⚠️ Vendor Conflict',
  pullInQueue: '🚀 Pull-In Queue',
  onHold: '🔴 On Hold',
  notReady: '🟡 Not Ready',
}

const RECOMMENDATION_COLOR: Record<Recommendation, string> = {
  green: 'text-green-400',
  pullIn: 'text-green-400',
  pushOut: 'text-red-400',
  vendorConflict: 'text-yellow-400',
  pullInQueue: 'text-purple-400',
  onHold: 'text-red-400',
  notReady: 'text-gray-400',
}

export default function SchedulePage() {
  const [hops, setHops] = useState<HOP[]>([])
  const [loaded, setLoaded] = useState(false)
  const [snapshotTime, setSnapshotTime] = useState('')
  const [fileName, setFileName] = useState('')
  const [program, setProgram] = useState<ProgramSettings>(DEFAULT_PROGRAM)
  const [thresholds, setThresholds] = useState<ThresholdSettings>(DEFAULT_THRESHOLDS)
  const [holds, setHolds] = useState<Set<string>>(new Set())
  const [collapsed, setCollapsed] = useState<Record<string, boolean>>({})
  const [rampUpOpen, setRampUpOpen] = useState<Record<string, boolean>>({})
  const [hasOptimized, setHasOptimized] = useState(false)
  const today = new Date()
  today.setHours(0, 0, 0, 0)

  useEffect(() => {
    const loadSettings = async () => {
      const [p, t] = await Promise.all([loadProgramSettings(), loadThresholdSettings()])
      setProgram(p)
      setThresholds(t)
    }
    loadSettings()
  }, [])

  useEffect(() => {
    const loadHolds = async () => {
      const { data } = await supabase.from('pm_updates_cache').select('id').like('id', 'schedule-hold-%')
      const set = new Set((data || []).map(row => row.id.slice('schedule-hold-'.length)))
      setHolds(set)
    }
    loadHolds()
  }, [])

  const processRows = useCallback((rows: unknown[][], _filename: string) => {
    let headerRow = -1
    for (let i = 0; i < 10; i++) {
      if ((rows[i] as unknown[])?.some(c => String(c).trim() === 'HOP')) { headerRow = i; break }
    }
    if (headerRow === -1) { alert('Could not find header row'); return }

    const headers = rows[headerRow] as string[]
    const col = (name: string) => headers.findIndex(h => String(h).trim() === name)

    const hopCol     = col('HOP')
    const gcCol      = col('General Contractor')
    const nokiaPmCol = col('Nokia PM')
    const don444Col  = col('DON 444')
    const ms15fCol   = col('MS15 Implementation Start F')
    const ms15aCol   = col('MS15 Implementation Start A')
    const ms16fCol   = col('MS16 Implementation Ends F')
    const ms16aCol   = col('MS16 Implementation Ends A')
    const ntpCol     = col('NTP A')
    const matCol     = headers.findIndex(h => String(h).trim() === 'Material Received A ')
    const itwSCol    = col('ITW Schedule Start')
    const itwECol    = col('ITW Schedule Complete')
    const ssSCol     = col('Samsung Schedule Start')
    const ssECol     = col('Samsung Schedule Complete')
    const pathIdCol  = headers.findIndex(h => String(h).trim().replace(/^'+|'+$/g, '') === 'Path ID')

    const hopRows = new Map<string, unknown[][]>()
    for (let i = headerRow + 1; i < rows.length; i++) {
      const row = rows[i] as unknown[]
      const don = String(row[don444Col] || '').trim().toUpperCase()
      if (don !== 'DON 444') continue
      const nokiaPm = String(row[nokiaPmCol] || '').trim().toUpperCase()
      if (nokiaPm !== 'CJ') continue
      const hop = String(row[hopCol] || '').trim()
      if (!hop || hop === 'undefined') continue
      if (!hopRows.has(hop)) hopRows.set(hop, [])
      hopRows.get(hop)!.push(row)
    }

    const parsed: HOP[] = []
    hopRows.forEach((rows2, hop) => {
      const row = rows2.find(r => String(r[gcCol] || '').trim()) || rows2[0]
      const ms15f = parseDate(row[ms15fCol])
      const ms15a = parseDate(row[ms15aCol])
      const ms16f = parseDate(row[ms16fCol])
      const ms16a = parseDate(row[ms16aCol])
      const ntpDate = parseDate(row[ntpCol])
      const matDate = parseDate(row[matCol])
      const hasNtp = !!(ntpDate && ntpDate.getFullYear() >= 2025)
      const hasMat = !!(matDate && matDate.getFullYear() >= 2020)

      const vendorConflicts: { vendor: string; start: Date; end: Date }[] = []
      rows2.forEach(r => {
        const itwS = parseDate(r[itwSCol]); const itwE = parseDate(r[itwECol])
        const ssS  = parseDate(r[ssSCol]);  const ssE  = parseDate(r[ssECol])
        if (itwS && itwE) vendorConflicts.push({ vendor: 'ITW', start: itwS, end: itwE })
        if (ssS && ssE)   vendorConflicts.push({ vendor: 'Samsung', start: ssS, end: ssE })
      })

      parsed.push({
        hop,
        gc: String(row[gcCol] || '').trim(),
        pathId: String(row[pathIdCol] || '').trim().replace(/^'+|'+$/g, ''),
        ms15f, ms15a, ms16f, ms16a,
        hasNtp, hasMat,
        vendorConflicts,
      })
    })

    setHops(parsed)
    setLoaded(true)
  }, [])

  useEffect(() => {
    const loadFromSnapshot = async () => {
      const snap = await loadTrackerSnapshot()
      if (!snap) return
      setSnapshotTime(new Date(snap.uploaded_at).toLocaleDateString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric' }) + ' at ' + new Date(snap.uploaded_at).toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' }))
      setFileName(snap.filename)
      processRows(snap.data, snap.filename)
    }
    loadFromSnapshot()
  }, [processRows])

  const toggleHold = async (hop: string) => {
    const id = `schedule-hold-${hop}`
    if (holds.has(hop)) {
      setHolds(prev => {
        const next = new Set(prev)
        next.delete(hop)
        return next
      })
      await supabase.from('pm_updates_cache').delete().eq('id', id)
    } else {
      setHolds(prev => new Set(prev).add(hop))
      await supabase.from('pm_updates_cache').upsert({
        id,
        updates: JSON.stringify({ heldAt: new Date().toISOString() }),
        updated_at: new Date().toISOString(),
      })
    }
  }

  const results = runSimulation(hops, program, thresholds, holds, today)

  const copyProposedDates = () => {
    const lines = ['HOP | Current MS15F | Proposed MS15F | Crew | GC']
    results.forEach(r => {
      r.pipeline.forEach(row => {
        lines.push(`${row.hop} | ${fmtDate(row.ms15f)} | ${fmtDate(row.proposedStart)} | Crew ${row.crew} | ${row.gc}`)
      })
    })
    navigator.clipboard.writeText(lines.join('\n'))
      .then(() => alert('✅ Copied proposed dates to clipboard!'))
      .catch(() => alert('Copy failed — please try manually'))
  }

  return (
    <div className="min-h-screen bg-gray-950 text-white p-6">
      <div className="max-w-full mx-auto">

        <BackToDashboard />

        <div className="mb-6 flex items-center justify-between flex-wrap gap-3">
          <div>
            <h1 className="text-3xl font-bold">Schedule Optimizer</h1>
            <p className="text-gray-400 mt-1">Crew queues, push/pull recommendations, and ramp-up readiness — DON 444 · CJ program.</p>
          </div>
          <div className="flex gap-2">
            <button onClick={() => setHasOptimized(true)}
              className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-semibold">
              ⚡ Reoptimize
            </button>
            {hasOptimized && (
              <button onClick={copyProposedDates}
                className="bg-teal-600 hover:bg-teal-700 text-white px-4 py-2 rounded-lg text-sm font-semibold">
                📋 Copy Proposed Dates
              </button>
            )}
          </div>
        </div>

        {snapshotTime && (
          <div className="mb-4 bg-gray-900 border border-gray-700 rounded-lg px-4 py-2 flex items-center justify-between">
            <p className="text-green-400 text-sm font-semibold">📡 Live data from {snapshotTime}</p>
            <p className="text-gray-500 text-xs">{fileName} — {hops.length} HOPs · Upload new tracker on Dashboard to refresh</p>
          </div>
        )}
        {!snapshotTime && (
          <div className="mb-4 bg-gray-900 border border-gray-700 rounded-lg px-4 py-8 text-center">
            <p className="text-gray-400">No tracker data found — go to Dashboard to upload your tracker</p>
          </div>
        )}

        {loaded && results.map(r => {
          const isCollapsed = !!collapsed[r.gc]
          return (
            <div key={r.gc} className="mb-6">
              {r.rampUp.triggered && (
                <div className="mb-2 bg-purple-950 border border-purple-700 rounded-xl p-3 cursor-pointer"
                  onClick={() => setRampUpOpen(prev => ({ ...prev, [r.gc]: !prev[r.gc] }))}>
                  <p className="text-purple-200 text-sm font-semibold">🚀 Engage {r.gc} — Ramp Up Crew</p>
                  {rampUpOpen[r.gc] && (
                    <p className="text-purple-300 text-xs mt-1">
                      {r.rampUp.greenCount} HOPs ready to start, current capacity covers {r.rampUp.capacity} — recommend adding {r.rampUp.recommendedCrews} crew{r.rampUp.recommendedCrews > 1 ? 's' : ''}
                    </p>
                  )}
                </div>
              )}

              <div onClick={() => setCollapsed(prev => ({ ...prev, [r.gc]: !prev[r.gc] }))}
                className="flex items-center justify-between bg-gray-900 border border-gray-700 rounded-xl px-5 py-3 cursor-pointer hover:border-blue-500">
                <div className="flex items-center gap-3">
                  <span className="text-white font-bold">{r.gc}</span>
                  <span className="text-gray-500 text-xs">{r.numCrews} crew{r.numCrews > 1 ? 's' : ''} · {r.inProgress.length} active · {r.pipeline.length} pipeline · {r.held.length} on hold</span>
                </div>
                <span className="text-gray-500 text-sm">{isCollapsed ? '▼' : '▲'}</span>
              </div>

              {!isCollapsed && (
                <div className="bg-gray-900 border border-t-0 border-gray-700 rounded-b-xl p-4 space-y-6">

                  {/* In Progress */}
                  <div>
                    <h3 className="text-sm font-bold text-white mb-2">In Progress</h3>
                    {r.inProgress.length === 0 ? (
                      <p className="text-gray-500 text-xs">No active HOPs</p>
                    ) : (
                      <div className="overflow-x-auto">
                        <table className="w-full text-xs">
                          <thead>
                            <tr className="bg-gray-800 text-gray-400">
                              <th className="text-left p-2">HOP</th>
                              <th className="text-left p-2">Crew</th>
                              <th className="text-left p-2">MS15A</th>
                              <th className="text-left p-2">MS16F</th>
                              <th className="text-left p-2">Days Elapsed</th>
                              <th className="text-left p-2">Status</th>
                              <th className="text-left p-2">Hold</th>
                            </tr>
                          </thead>
                          <tbody>
                            {r.inProgress.map(row => (
                              <tr key={row.hop} className={`border-t border-gray-800 ${(row.daysElapsed ?? 0) > thresholds.durationAlertDays ? 'bg-red-950' : ''}`}>
                                <td className="p-2 font-semibold text-white whitespace-nowrap">{row.hop}</td>
                                <td className="p-2 text-gray-300">Crew {row.crew}</td>
                                <td className="p-2 text-gray-300 whitespace-nowrap">{fmtDate(row.ms15a)}</td>
                                <td className="p-2 text-gray-300 whitespace-nowrap">{fmtDate(row.ms16f)}</td>
                                <td className={`p-2 font-bold ${(row.daysElapsed ?? 0) > thresholds.durationAlertDays ? 'text-red-400' : 'text-green-400'}`}>{row.daysElapsed}d</td>
                                <td className="p-2">{(row.daysElapsed ?? 0) > thresholds.durationAlertDays ? <span className="text-red-400">⚠️ Over target</span> : <span className="text-green-400">On track</span>}</td>
                                <td className="p-2">
                                  <button onClick={() => toggleHold(row.hop)}
                                    className="text-xs bg-gray-700 hover:bg-red-700 text-gray-300 hover:text-white px-2 py-1 rounded font-semibold">
                                    🔴 HOLD
                                  </button>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>

                  {/* Pipeline */}
                  <div>
                    <h3 className="text-sm font-bold text-white mb-2">Pipeline</h3>
                    {r.pipeline.length === 0 ? (
                      <p className="text-gray-500 text-xs">No pipeline HOPs</p>
                    ) : (
                      <div className="overflow-x-auto">
                        <table className="w-full text-xs">
                          <thead>
                            <tr className="bg-gray-800 text-gray-400">
                              <th className="text-left p-2">HOP</th>
                              <th className="text-left p-2">Crew</th>
                              <th className="text-left p-2">MS15F</th>
                              <th className="text-left p-2">MS16F</th>
                              <th className="text-left p-2">NTP</th>
                              <th className="text-left p-2">Material</th>
                              <th className="text-left p-2">Vendor</th>
                              <th className="text-left p-2">Recommendation</th>
                              {hasOptimized && <th className="text-left p-2">Proposed Date</th>}
                              <th className="text-left p-2">Hold</th>
                            </tr>
                          </thead>
                          <tbody>
                            {r.pipeline.map(row => (
                              <tr key={row.hop} className="border-t border-gray-800">
                                <td className="p-2 font-semibold text-white whitespace-nowrap">{row.hop}</td>
                                <td className="p-2 text-gray-300">Crew {row.crew}</td>
                                <td className="p-2 text-gray-300 whitespace-nowrap">{fmtDate(row.ms15f)}</td>
                                <td className="p-2 text-gray-300 whitespace-nowrap">{fmtDate(row.ms16f)}</td>
                                <td className="p-2">{row.hasNtp ? <span className="text-green-400 font-bold">✓</span> : <span className="text-red-400 font-bold">✗</span>}</td>
                                <td className="p-2">{row.hasMat ? <span className="text-green-400 font-bold">✓</span> : <span className="text-red-400 font-bold">✗</span>}</td>
                                <td className="p-2 text-xs text-yellow-300 whitespace-nowrap">{row.vendorConflict || '—'}</td>
                                <td className="p-2 whitespace-nowrap">
                                  <span className={`text-xs font-bold ${RECOMMENDATION_COLOR[row.recommendation]}`}>{RECOMMENDATION_LABEL[row.recommendation]}</span>
                                  {row.reassignNote && <p className="text-gray-500 text-xs mt-0.5">{row.reassignNote}</p>}
                                </td>
                                {hasOptimized && (
                                  <td className="p-2 font-bold text-amber-400 whitespace-nowrap">{fmtDate(row.proposedStart)}</td>
                                )}
                                <td className="p-2">
                                  <button onClick={() => toggleHold(row.hop)}
                                    className="text-xs bg-gray-700 hover:bg-red-700 text-gray-300 hover:text-white px-2 py-1 rounded font-semibold">
                                    🔴 HOLD
                                  </button>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>

                  {/* Pull-In Queue */}
                  {r.pullInQueue.length > 0 && (
                    <div>
                      <h3 className="text-sm font-bold text-white mb-2">🚀 Pull-In Queue</h3>
                      <div className="overflow-x-auto">
                        <table className="w-full text-xs">
                          <thead>
                            <tr className="bg-gray-800 text-gray-400">
                              <th className="text-left p-2">HOP</th>
                              <th className="text-left p-2">Crew</th>
                              <th className="text-left p-2">Current MS15F</th>
                              <th className="text-left p-2">Suggested Pull-In Date</th>
                            </tr>
                          </thead>
                          <tbody>
                            {r.pullInQueue.map(row => (
                              <tr key={row.hop} className="border-t border-gray-800 bg-purple-950">
                                <td className="p-2 font-semibold text-white whitespace-nowrap">{row.hop}</td>
                                <td className="p-2 text-gray-300">Crew {row.crew}</td>
                                <td className="p-2 text-gray-300 whitespace-nowrap">{fmtDate(row.ms15f)}</td>
                                <td className="p-2 font-bold text-purple-300 whitespace-nowrap">{fmtDate(row.proposedStart)}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}

                  {/* Hold */}
                  {r.held.length > 0 && (
                    <div>
                      <h3 className="text-sm font-bold text-white mb-2">🔴 On Hold</h3>
                      <div className="flex flex-col gap-1">
                        {r.held.map(row => (
                          <div key={row.hop} className="flex items-center gap-3 bg-gray-800 rounded-lg p-2">
                            <span className="text-white text-xs font-semibold flex-1">{row.hop}</span>
                            <span className="text-gray-500 text-xs">{fmtDate(row.ms15f) || fmtDate(row.ms15a)}</span>
                            <button onClick={() => toggleHold(row.hop)}
                              className="text-xs bg-red-800 hover:bg-gray-700 text-white px-2 py-1 rounded font-semibold">
                              Release Hold
                            </button>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                </div>
              )}
            </div>
          )
        })}

        {!loaded && (
          <div className="bg-gray-900 rounded-xl border border-gray-700 p-12 text-center">
            <p className="text-gray-400 text-xl">📂 Upload your tracker on the Dashboard to begin schedule analysis</p>
          </div>
        )}

      </div>
    </div>
  )
}
