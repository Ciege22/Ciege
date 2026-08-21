'use client'

export const dynamic = 'force-dynamic'

import { useState, useCallback, useEffect } from 'react'
import * as XLSX from 'xlsx'
import { loadTrackerSnapshot } from '../lib/supabase'
import BackToDashboard from '../components/BackToDashboard'

interface HOP {
  hop: string
  gc: string
  cm: string
  ms15f: Date | null
  ms15a: Date | null
  ms16f: Date | null
  ms16a: Date | null
  hasNtp: boolean
  hasMat: boolean
  vendorConflicts: { vendor: string; start: Date; end: Date }[]
  siteNames: string[]
  inProgress: boolean
  complete: boolean
  daysOut: number | null
  nokiaPm: string
}

interface CrewAssignment {
  crewId: string
  gc: string
  hop: string
  start: Date
  end: Date
  isActive: boolean
  isComplete: boolean
}

interface ScheduleSuggestion {
  hop: string
  gc: string
  currentStart: string
  suggestedStart: string
  daysMoved: number
  direction: 'pull-in' | 'push-out' | 'reassign' | 'unassigned'
  reason: string
  blocker: string
  crewId: string
  readiness: string
  vendorClearsDate?: string
  crewAvailDate?: string
  allReasons?: string[]
  parallelHops?: string[]
}

interface GapInfo {
  gc: string
  crewId: string
  gapStart: Date
  gapEnd: Date
  gapDays: number
  nextHop: string
}

const GC_CREWS: Record<string, number> = {
  'MZI': 2,
  'NV Tel': 1,
  'Mastec': 1,
  'Vikor': 2,
  'Tech CX': 2,
}

const GC_MAX_CREWS: Record<string, number> = {
  'MZI': 2,
  'NV Tel': 2,
  'Mastec': 2,
  'Vikor': 2,
  'Tech CX': 4,
}

const HOP_DURATION = 18

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

function hasVendorConflict(date: Date, conflicts: { vendor: string; start: Date; end: Date }[]): string {
  for (const c of conflicts) {
    if (c.start <= date && date <= c.end) return `${c.vendor} on site thru ${fmtShort(c.end)}`
    const buf = daysBetween(c.end, date)
    if (buf >= 0 && buf <= 1) return `${c.vendor} clears ${fmtShort(c.end)} — only ${buf}d buffer`
  }
  return ''
}

export default function SchedulePage() {
  const [hops, setHops] = useState<HOP[]>([])
  const [loaded, setLoaded] = useState(false)
  const [fileName, setFileName] = useState('')
  const [suggestions, setSuggestions] = useState<ScheduleSuggestion[]>([])
  const [gaps, setGaps] = useState<GapInfo[]>([])
  const [crewAssignments, setCrewAssignments] = useState<CrewAssignment[]>([])
  const [activeTab, setActiveTab] = useState<'gantt' | 'suggestions' | 'gaps'>('suggestions')
  const [selectedGC, setSelectedGC] = useState<string>('ALL')
  const [sortAsc, setSortAsc] = useState(true)
  const [snapshotTime, setSnapshotTime] = useState<string>('')
  const today = new Date()

  useEffect(() => {
    const loadFromSnapshot = async () => {
      const snap = await loadTrackerSnapshot()
      if (!snap) return
      setSnapshotTime(new Date(snap.uploaded_at).toLocaleDateString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric' }) + ' at ' + new Date(snap.uploaded_at).toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' }))
      setFileName(snap.filename)
      processRows(snap.data, snap.filename)
    }
    loadFromSnapshot()
  }, [])

  const processRows = useCallback((rows: unknown[][], _filename: string) => {

    let headerRow = -1
    for (let i = 0; i < 10; i++) {
      if ((rows[i] as unknown[])?.some(c => String(c).trim() === 'HOP')) { headerRow = i; break }
    }
    if (headerRow === -1) { alert('Could not find header row'); return }

    const headers = rows[headerRow] as string[]
    const col = (name: string) => headers.findIndex(h => String(h).trim() === name)

    const hopCol      = col('HOP')
    const gcCol       = col('General Contractor')
    const nokiaPmCol  = col('Nokia PM')
    const newCmCol    = col('New CM')
    const don444Col   = col('DON 444')
    const ms15fCol    = col('MS15 Implementation Start F')
    const ms15aCol    = col('MS15 Implementation Start A')
    const ms16fCol    = col('MS16 Implementation Ends F')
    const ms16aCol    = col('MS16 Implementation Ends A')
    const ntpCol      = col('NTP A')
    const matCol      = headers.findIndex(h => String(h).trim() === 'Material Received A ')
    const itwSCol     = col('ITW Schedule Start')
    const itwECol     = col('ITW Schedule Complete')
    const ssSCol      = col('Samsung Schedule Start')
    const ssECol      = col('Samsung Schedule Complete')

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
      const row  = rows2[0]
      const ms15f = parseDate(row[ms15fCol])
      const ms15a = parseDate(row[ms15aCol])
      const ms16f = parseDate(row[ms16fCol])
      const ms16a = parseDate(row[ms16aCol])
      const ntpDate = parseDate(row[ntpCol])
      const matDate = parseDate(row[matCol])
      const started  = !!ms15a
      const complete = !!ms16a
      const hasNtp   = !!(ntpDate && ntpDate.getFullYear() >= 2025)
      const hasMat   = !!(matDate && matDate.getFullYear() >= 2020)

      // Collect vendor conflicts from ALL rows
      const vendorConflicts: { vendor: string; start: Date; end: Date }[] = []
      rows2.forEach(r => {
        const itwS = parseDate(r[itwSCol]); const itwE = parseDate(r[itwECol])
        const ssS  = parseDate(r[ssSCol]);  const ssE  = parseDate(r[ssECol])
        if (itwS && itwE) vendorConflicts.push({ vendor: 'ITW', start: itwS, end: itwE })
        if (ssS && ssE)   vendorConflicts.push({ vendor: 'Samsung', start: ssS, end: ssE })
      })

      const siteNames = rows2.map(r => String(r[col('Site Name')] || '').trim()).filter(Boolean)

      parsed.push({
        hop,
        gc:       String(row[gcCol] || '').trim(),
        cm:       String(row[newCmCol] || '').trim(),
        ms15f, ms15a, ms16f, ms16a,
        hasNtp, hasMat,
        vendorConflicts,
        siteNames,
        inProgress: started && !complete,
        complete,
        daysOut: ms15f ? daysBetween(today, ms15f) : null,
        nokiaPm: String(row[nokiaPmCol] || '').trim()
      })
    })

    // â”€â”€ Build crew assignments â”€â”€
    const assignments: CrewAssignment[] = []
    const gcList = ['MZI', 'NV Tel', 'Mastec', 'Vikor', 'Tech CX']

    gcList.forEach(gc => {
      const gcHopList = parsed
        .filter(h => h.gc === gc && !h.complete)
        .sort((a, b) => {
          if (a.inProgress && !b.inProgress) return -1
          if (!a.inProgress && b.inProgress) return 1
          const aDate = a.ms15a || a.ms15f
          const bDate = b.ms15a || b.ms15f
          if (!aDate && !bDate) return 0
          if (!aDate) return 1
          if (!bDate) return -1
          return aDate.getTime() - bDate.getTime()
        })

      const numCrews = GC_CREWS[gc] || 1
      const crewAvail: Date[] = Array(numCrews).fill(today).map(() => new Date(today))

      // Assign in-progress first
      const inProg = gcHopList.filter(h => h.inProgress)
      const pipeline = gcHopList.filter(h => !h.inProgress)

      inProg.forEach((h, i) => {
        const crewIdx = i % numCrews
        const start = h.ms15a || today
        const end   = h.ms16f || addDays(start, HOP_DURATION)
        assignments.push({ crewId: `${gc}-Crew${crewIdx + 1}`, gc, hop: h.hop, start, end, isActive: true, isComplete: false })
        crewAvail[crewIdx] = addDays(end, 1)
      })

      pipeline.forEach(h => {
        const earliest = crewAvail.reduce((min, d) => d < min ? d : min, crewAvail[0])
        const crewI = crewAvail.findIndex(d => d.getTime() === earliest.getTime())
        const start = h.ms15f || addDays(earliest, 0)
        const end   = addDays(start, HOP_DURATION)
        assignments.push({ crewId: `${gc}-Crew${crewI + 1}`, gc, hop: h.hop, start, end, isActive: false, isComplete: false })
        crewAvail[crewI] = addDays(end, 1)
      })
    })

    // â”€â”€ Detect gaps â”€â”€
    const gapList: GapInfo[] = []
    const crewMap = new Map<string, CrewAssignment[]>()
    assignments.forEach(a => {
      if (!crewMap.has(a.crewId)) crewMap.set(a.crewId, [])
      crewMap.get(a.crewId)!.push(a)
    })
    crewMap.forEach((crewJobs, crewId) => {
      const sorted = crewJobs.sort((a, b) => a.start.getTime() - b.start.getTime())
      for (let i = 0; i < sorted.length - 1; i++) {
        const gap = daysBetween(sorted[i].end, sorted[i + 1].start)
        if (gap > 3) {
          gapList.push({
            gc: sorted[i].gc, crewId,
            gapStart: addDays(sorted[i].end, 1),
            gapEnd: addDays(sorted[i + 1].start, -1),
            gapDays: gap,
            nextHop: sorted[i + 1].hop
          })
        }
      }
    })

    // â”€â”€ Generate suggestions â”€â”€
    const suggList: ScheduleSuggestion[] = []

    parsed.forEach(h => {
      if (h.complete) return

      // Show in-progress HOPs
      if (h.inProgress) {
        const elapsed = h.ms15a ? daysBetween(h.ms15a, today) : 0
        const ms16f = h.ms16f
        const daysRemaining = ms16f ? daysBetween(today, ms16f) : null
        const over18 = elapsed > 18
        const allReasons: string[] = []
        if (over18) allReasons.push(`⚠️ ${elapsed}d elapsed — over 18d target`)
        if (daysRemaining !== null && daysRemaining < 0) allReasons.push(`🔴 FC End passed ${Math.abs(daysRemaining)}d ago`)
        allReasons.push(`Started: ${fmtDate(h.ms15a)} | FC End: ${fmtDate(ms16f)}`)

        suggList.push({
          hop: h.hop, gc: h.gc,
          currentStart: fmtDate(h.ms15a),
          suggestedStart: fmtDate(ms16f),
          daysMoved: 0,
          direction: over18 ? 'push-out' : 'pull-in',
          reason: over18
            ? `🔨 IN PROGRESS — ${elapsed}d elapsed ⚠️ Over 18d target — get updated completion date`
            : `🔨 IN PROGRESS — ${elapsed}d elapsed — on track, FC end ${fmtDate(ms16f)}`,
          blocker: over18 ? `${elapsed}d elapsed` : 'None',
          crewId: '',
          readiness: `Started: ${fmtDate(h.ms15a)}`,
          vendorClearsDate: undefined,
          crewAvailDate: ms16f ? fmtDate(ms16f) : undefined,
          allReasons,
          parallelHops: []
        })
        return
      }

      // Flag unassigned HOPs
      if (!h.gc) {
        suggList.push({
          hop: h.hop, gc: 'UNASSIGNED', currentStart: fmtDate(h.ms15f),
          suggestedStart: '', daysMoved: 0, direction: 'unassigned',
          reason: '⚠️ No GC assigned — needs contractor assignment',
          blocker: 'No GC', crewId: '', readiness: h.hasNtp && h.hasMat ? '✅ NTP + Mat ready' : `NTP: ${h.hasNtp ? '✓' : '✗'} | Mat: ${h.hasMat ? '✓' : '✗'}`
        })
        return
      }

      if (!h.ms15f || h.inProgress) return

      const vcConflict = hasVendorConflict(h.ms15f, h.vendorConflicts)
      const readiness  = `NTP: ${h.hasNtp ? '✓' : '✗'} | Mat: ${h.hasMat ? '✓' : '✗'}`

      // Find earliest available crew date for this GC
      const gcAssignments = assignments.filter(a => a.gc === h.gc)
      const crewAvailDates = new Map<string, Date>()
      gcAssignments.forEach(a => {
        const cur = crewAvailDates.get(a.crewId)
        if (!cur || a.end > cur) crewAvailDates.set(a.crewId, addDays(a.end, 1))
      })

      let earliestCrew = ''
      let earliestDate: Date | null = null
      crewAvailDates.forEach((date, crewId) => {
        if (!earliestDate || date < earliestDate) {
          earliestDate = date; earliestCrew = crewId
        }
      })

      // Find vendor clear date
      let vendorClearDate: Date | null = null
      if (h.vendorConflicts.length > 0) {
        for (const c of h.vendorConflicts) {
          const clearCandidate = addDays(c.end, 1)
          if (!vendorClearDate || clearCandidate > vendorClearDate) vendorClearDate = clearCandidate
        }
      }

      // Effective earliest start = max of crew available and vendor clear
      const effectiveStart = earliestDate
        ? (vendorClearDate && vendorClearDate > (earliestDate as Date)
            ? vendorClearDate
            : earliestDate as Date)
        : vendorClearDate

      // Find parallel HOPs — same site names
      const parallelHops = parsed.filter(other =>
        other.hop !== h.hop &&
        other.gc === h.gc &&
        !other.complete &&
        !other.inProgress &&
        other.siteNames.some(s => h.siteNames.includes(s))
      ).map(o => o.hop)

      const allReasons: string[] = []

      // Pull-in opportunity
      if (h.hasNtp && h.hasMat && !vcConflict && effectiveStart && h.ms15f && effectiveStart < h.ms15f) {
        const daysMoved = daysBetween(effectiveStart, h.ms15f)
        if (daysMoved >= 1) {
          allReasons.push(`✅ Pull in ${daysMoved}d — crew available ${fmtDate(effectiveStart as Date)}, NTP + Mat confirmed`)
          if (parallelHops.length > 0) allReasons.push(`🔄 Can start together with: ${parallelHops.join(', ')}`)
          suggList.push({
            hop: h.hop, gc: h.gc,
            currentStart: fmtDate(h.ms15f),
            suggestedStart: fmtDate(effectiveStart as Date),
            daysMoved, direction: 'pull-in',
            reason: allReasons[0],
            blocker: 'None',
            crewId: earliestCrew,
            readiness,
            vendorClearsDate: vendorClearDate ? fmtDate(vendorClearDate) : undefined,
            crewAvailDate: earliestDate ? fmtDate(earliestDate as Date) : undefined,
            allReasons,
            parallelHops
          })
        }
      }

      // Vendor conflict
      if (vcConflict) {
        const suggestedAfterVendor = vendorClearDate || h.ms15f
        allReasons.push(`🔴 Vendor conflict — ${vcConflict}`)
        if (effectiveStart && h.ms15f && effectiveStart > h.ms15f) {
          allReasons.push(`📅 Crew available ${fmtDate(earliestDate as unknown as Date)} | Vendor clears ${vendorClearDate ? fmtDate(vendorClearDate) : '?'}`)
        }
        if (!suggList.find(s => s.hop === h.hop)) {
          suggList.push({
            hop: h.hop, gc: h.gc,
            currentStart: fmtDate(h.ms15f),
            suggestedStart: suggestedAfterVendor ? fmtDate(suggestedAfterVendor) : '',
            daysMoved: suggestedAfterVendor && h.ms15f ? Math.abs(daysBetween(h.ms15f, suggestedAfterVendor)) : 0,
            direction: 'push-out',
            reason: allReasons[0],
            blocker: vcConflict,
            crewId: earliestCrew || '',
            readiness,
            vendorClearsDate: vendorClearDate ? fmtDate(vendorClearDate) : undefined,
            crewAvailDate: earliestDate ? fmtDate(earliestDate as Date) : undefined,
            allReasons,
            parallelHops
          })
        }
      }

      // Missing NTP or material
      if ((!h.hasNtp || !h.hasMat) && h.daysOut !== null && h.daysOut <= 36) {
        if (!suggList.find(s => s.hop === h.hop)) {
          const blockers = []
          if (!h.hasNtp) blockers.push('NTP missing')
          if (!h.hasMat) blockers.push('Material not received')
          suggList.push({
            hop: h.hop, gc: h.gc,
            currentStart: fmtDate(h.ms15f),
            suggestedStart: '',
            daysMoved: 0, direction: 'push-out',
            reason: `⚠️ Starts in ${h.daysOut}d but blockers unresolved`,
            blocker: blockers.join(' | '),
            crewId: earliestCrew || '',
            readiness,
            allReasons: blockers
          })
        }
      }

      // Reassignment
      if (h.hasNtp && h.hasMat && !vcConflict && earliestDate) {
        const gapToStart = h.ms15f ? daysBetween(earliestDate as Date, h.ms15f) : 0
        if (gapToStart < 0 && Math.abs(gapToStart) > 36) {
          const otherGCs = ['MZI', 'NV Tel', 'Mastec', 'Tech CX'].filter(gc => gc !== h.gc)
          let bestGC = ''; let bestDate: Date | null = null
          otherGCs.forEach(gc => {
            const gcA = assignments.filter(a => a.gc === gc)
            const avail = new Map<string, Date>()
            gcA.forEach(a => {
              const cur = avail.get(a.crewId)
              if (!cur || a.end > cur) avail.set(a.crewId, addDays(a.end, 1))
            })
            avail.forEach((date) => {
              if (!bestDate || date < bestDate) { bestDate = date; bestGC = gc }
            })
          })
          if (bestGC && bestDate && (bestDate as Date) < (earliestDate as unknown as Date)) {
            if (!suggList.find(s => s.hop === h.hop && s.direction === 'reassign')) {
              suggList.push({
                hop: h.hop, gc: h.gc,
                currentStart: fmtDate(h.ms15f),
                suggestedStart: fmtDate(bestDate),
                daysMoved: h.ms15f ? daysBetween(bestDate, h.ms15f) : 0,
                direction: 'reassign',
                reason: `🔄 Consider reassigning to ${bestGC} — crew available ${fmtDate(bestDate)}`,
                blocker: `${h.gc} crew gap >36d`,
                crewId: bestGC,
                readiness,
                crewAvailDate: fmtDate(earliestDate as Date),
                allReasons: [`${h.gc} crew not available until ${fmtDate(earliestDate as Date)}`, `${bestGC} crew available ${fmtDate(bestDate)}`]
              })
            }
          }
        }
      }
    })

    // Sort suggestions by priority
    const priority = (s: ScheduleSuggestion) => {
      if (s.direction === 'unassigned') return 0
      if (s.direction === 'push-out' && s.blocker !== 'None') return 1
      if (s.direction === 'reassign') return 2
      if (s.direction === 'pull-in') return 3
      return 4
    }
    suggList.sort((a, b) => priority(a) - priority(b))

    setHops(parsed)
    setCrewAssignments(assignments)
    setGaps(gapList)
    setSuggestions(suggList)
    setLoaded(true)
  }, [today])

  const handleFile = useCallback((file: File) => {
    setFileName(file.name)
    const reader = new FileReader()
    reader.onload = (e) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer)
      const wb = XLSX.read(data, { type: 'array', cellDates: true })
      const ws = wb.Sheets['HOPs']
      if (!ws) { alert('HOPs tab not found'); return }
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 }) as unknown[][]
      processRows(rows, file.name)
    }
    reader.readAsArrayBuffer(file)
  }, [processRows])

  // â”€â”€ Gantt helpers â”€â”€
  const ganttStart = today
  const ganttWeeks = 16
  const weekStarts = Array.from({ length: ganttWeeks }, (_, i) => addDays(ganttStart, i * 7))

  const crewIds = Array.from(new Set(crewAssignments.map(a => a.crewId))).sort()

  return (
    <div className="min-h-screen bg-gray-950 text-white p-6">
      <div className="max-w-full mx-auto">

        <BackToDashboard />

        {/* Header */}
        <div className="mb-6">
          <h1 className="text-3xl font-bold">Schedule Optimizer</h1>
          <p className="text-gray-400 mt-1">Upload your tracker to analyze crew availability, gaps, and scheduling opportunities.</p>
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

        {loaded && (
          <>
            {/* KPI Summary */}
            <div className="grid grid-cols-2 md:grid-cols-5 gap-3 mb-6">
              {['MZI', 'NV Tel', 'Mastec', 'Vikor', 'Tech CX'].map(gc => {
                const gcHops   = hops.filter(h => h.gc === gc)
                const active   = gcHops.filter(h => h.inProgress).length
                const pipeline = gcHops.filter(h => !h.inProgress && !h.complete).length
                const gcGaps   = gaps.filter(g => g.gc === gc).length
                const crews    = GC_CREWS[gc] || 1
                const maxCrews = GC_MAX_CREWS[gc] || 1
                return (
                  <div key={gc} className="bg-gray-900 rounded-xl border border-gray-700 p-4">
                    <h3 className="font-bold text-white text-sm">{gc}</h3>
                    <p className="text-gray-400 text-xs mt-1">{crews} crew{crews > 1 ? 's' : ''} active / {maxCrews} max</p>
                    <div className="mt-2 space-y-1">
                      <p className="text-xs"><span className="text-blue-400 font-bold">{active}</span> <span className="text-gray-500">active</span></p>
                      <p className="text-xs"><span className="text-yellow-400 font-bold">{pipeline}</span> <span className="text-gray-500">pipeline</span></p>
                      {gcGaps > 0 && <p className="text-xs"><span className="text-red-400 font-bold">{gcGaps}</span> <span className="text-gray-500">gaps</span></p>}
                    </div>
                  </div>
                )
              })}
            </div>

            {/* GC Filter */}
            <div className="flex gap-2 mb-4 flex-wrap">
              {['ALL', 'MZI', 'NV Tel', 'Mastec', 'Vikor', 'Tech CX'].map(gc => (
                <button key={gc} onClick={() => setSelectedGC(gc)}
                  className={`px-4 py-2 rounded-lg text-sm font-semibold transition-all ${selectedGC === gc ? 'bg-blue-600 text-white scale-105' : 'bg-gray-800 text-gray-300 hover:bg-gray-700'}`}>
                  {gc}
                </button>
              ))}
            </div>

            {/* Tabs */}
            <div className="flex gap-2 mb-6">
              {(['suggestions', 'gaps', 'gantt'] as const).map(tab => (
                <button key={tab} onClick={() => setActiveTab(tab)}
                  className={`px-5 py-2 rounded-lg text-sm font-semibold transition-all capitalize ${activeTab === tab ? 'bg-blue-600 text-white' : 'bg-gray-800 text-gray-300 hover:bg-gray-700'}`}>
                  {tab === 'suggestions' ? `📋 Recommendations (${suggestions.length})` : tab === 'gaps' ? `⚡ Crew Gaps (${gaps.length})` : '📅 Gantt View'}
                </button>
              ))}
            </div>

            {/* SUGGESTIONS TAB */}
            {activeTab === 'suggestions' && (
              <div className="bg-gray-900 rounded-xl border border-gray-700 p-6">
                <div className="flex items-center justify-between mb-4">
                  <h2 className="text-lg font-bold">Schedule Recommendations</h2>
                  <button onClick={() => setSortAsc(prev => !prev)}
                    className="bg-gray-700 hover:bg-gray-600 text-white text-xs px-3 py-1 rounded font-semibold">
                    Sort by FC Start {sortAsc ? 'â†‘ Earliest' : 'â†“ Latest'}
                  </button>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-xs">
                    <thead>
                      <tr className="bg-gray-800 text-gray-400">
                        <th className="text-left p-2">Priority</th>
                        <th className="text-left p-2">HOP</th>
                        <th className="text-left p-2">GC</th>
                        <th className="text-left p-2">Current FC Start</th>
                        <th className="text-left p-2">Suggested Start</th>
                        <th className="text-left p-2">Days Moved</th>
                        <th className="text-left p-2">Readiness</th>
                        <th className="text-left p-2">Blocker</th>
                        <th className="text-left p-2">Vendor Clears</th>
                        <th className="text-left p-2">Crew Available</th>
                        <th className="text-left p-2">Parallel HOPs</th>
                        <th className="text-left p-2">Recommendation</th>
                      </tr>
                    </thead>
                    <tbody>
                      {[...suggestions]
                        .filter(s => selectedGC === 'ALL' || s.gc === selectedGC)
                        .sort((a, b) => {
                          const aDate = a.currentStart ? new Date(a.currentStart).getTime() : 0
                          const bDate = b.currentStart ? new Date(b.currentStart).getTime() : 0
                          return sortAsc ? aDate - bDate : bDate - aDate
                        })
                        .map((s, i) => {
                        const rowBg = s.direction === 'unassigned' ? 'bg-purple-950' :
                                      s.direction === 'push-out' ? 'bg-red-950' :
                                      s.direction === 'reassign' ? 'bg-yellow-950' : 'bg-green-950'
                        const badge = s.direction === 'unassigned' ? '🟣 No GC' :
                                      s.direction === 'push-out' ? '🔴 Push Out' :
                                      s.direction === 'reassign' ? '🔄 Reassign' : '✅ Pull In'
                        return (
                          <tr key={i} className={`border-t border-gray-800 ${rowBg}`}>
                            <td className="p-2 whitespace-nowrap">
                              <span className={`text-xs font-bold px-2 py-1 rounded ${
                                s.direction === 'unassigned' ? 'bg-purple-800 text-purple-200' :
                                s.direction === 'push-out' ? 'bg-red-800 text-red-200' :
                                s.direction === 'reassign' ? 'bg-yellow-800 text-yellow-200' :
                                'bg-green-800 text-green-200'}`}>
                                {badge}
                              </span>
                            </td>
                            <td className="p-2 font-semibold text-white whitespace-nowrap">{s.hop}</td>
                            <td className="p-2 text-gray-300 whitespace-nowrap">{s.gc}</td>
                            <td className="p-2 text-gray-300 whitespace-nowrap">{s.currentStart || '—'}</td>
                            <td className={`p-2 font-bold whitespace-nowrap ${s.suggestedStart ? (s.direction === 'pull-in' ? 'text-green-400' : 'text-yellow-400') : 'text-gray-500'}`}>
                              {s.suggestedStart || 'Review manually'}
                            </td>
                            <td className={`p-2 font-bold ${s.daysMoved > 0 ? (s.direction === 'pull-in' ? 'text-green-400' : 'text-red-400') : 'text-gray-500'}`}>
                              {s.daysMoved > 0 ? `${s.direction === 'pull-in' ? '-' : '+'}${s.daysMoved}d` : '—'}
                            </td>
                            <td className="p-2 text-gray-300 whitespace-nowrap">{s.readiness}</td>
                            <td className="p-2 text-xs text-red-400">{s.blocker !== 'None' ? s.blocker : <span className="text-green-400">None</span>}</td>
                            <td className="p-2 text-xs text-yellow-300 whitespace-nowrap">{s.vendorClearsDate || '—'}</td>
                            <td className="p-2 text-xs text-blue-300 whitespace-nowrap">{s.crewAvailDate || '—'}</td>
                            <td className="p-2 text-xs text-gray-300 whitespace-nowrap">
                              {s.parallelHops && s.parallelHops.length > 0
                                ? <span className="text-green-400" title={s.parallelHops.join(', ')}>🔄 {s.parallelHops.length} HOP{s.parallelHops.length > 1 ? 's' : ''}</span>
                                : <span className="text-gray-600">—</span>}
                            </td>
                            <td className="p-2 text-xs max-w-64 cursor-help" title={(s.allReasons || []).join(' | ')}>{s.reason}</td>
                          </tr>
                        )
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* GAPS TAB */}
            {activeTab === 'gaps' && (
              <div className="bg-gray-900 rounded-xl border border-gray-700 p-6">
                <h2 className="text-lg font-bold mb-4">Crew Gaps — Idle Time Between HOPs</h2>
                {gaps.length === 0
                  ? <p className="text-green-400">✅ No significant gaps detected — crews are well utilized</p>
                  : (
                    <div className="overflow-x-auto">
                      <table className="w-full text-xs">
                        <thead>
                          <tr className="bg-gray-800 text-gray-400">
                            <th className="text-left p-2">GC</th>
                            <th className="text-left p-2">Crew</th>
                            <th className="text-left p-2">Gap Start</th>
                            <th className="text-left p-2">Gap End</th>
                            <th className="text-left p-2">Idle Days</th>
                            <th className="text-left p-2">Next HOP</th>
                            <th className="text-left p-2">Action</th>
                          </tr>
                        </thead>
                        <tbody>
                          {gaps.filter(g => selectedGC === 'ALL' || g.gc === selectedGC).sort((a, b) => b.gapDays - a.gapDays).map((g, i) => (
                            <tr key={i} className={`border-t border-gray-800 ${g.gapDays >= 14 ? 'bg-red-950' : g.gapDays >= 7 ? 'bg-yellow-950' : 'bg-gray-900'}`}>
                              <td className="p-2 font-semibold text-white">{g.gc}</td>
                              <td className="p-2 text-gray-300">{g.crewId}</td>
                              <td className="p-2 text-gray-300">{fmtDate(g.gapStart)}</td>
                              <td className="p-2 text-gray-300">{fmtDate(g.gapEnd)}</td>
                              <td className={`p-2 font-bold ${g.gapDays >= 14 ? 'text-red-400' : g.gapDays >= 7 ? 'text-yellow-400' : 'text-gray-300'}`}>
                                {g.gapDays}d
                              </td>
                              <td className="p-2 text-gray-300 whitespace-nowrap">{g.nextHop}</td>
                              <td className="p-2 text-xs text-yellow-300">
                                {g.gapDays >= 14
                                  ? '🔴 Find a HOP to fill this gap or consider reducing crews'
                                  : g.gapDays >= 7
                                  ? '⚠️ Short gap — monitor for pull-in opportunities'
                                  : '✅ Minor gap — acceptable'}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )
                }

                {/* Crew expansion recommendations */}
                <div className="mt-6">
                  <h3 className="text-base font-bold mb-3">🚀 Crew Expansion Readiness</h3>
                  <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                    {['MZI', 'NV Tel', 'Mastec', 'Tech CX'].map(gc => {
                      const maxCrews  = GC_MAX_CREWS[gc] || 1
                      const curCrews  = GC_CREWS[gc] || 1
                      const gcHopList = hops.filter(h => h.gc === gc && !h.complete && !h.inProgress)
                      const readyHops = gcHopList.filter(h => h.hasNtp && h.hasMat)
                      const canExpand = maxCrews > curCrews
                      const hasEnoughPipeline = readyHops.length >= (curCrews + 1) * 2

                      return (
                        <div key={gc} className={`rounded-lg border p-4 ${hasEnoughPipeline && canExpand ? 'border-green-600 bg-green-950' : 'border-gray-700 bg-gray-900'}`}>
                          <h4 className="font-bold text-sm text-white">{gc}</h4>
                          <p className="text-xs text-gray-400 mt-1">{curCrews} â†’ {maxCrews} crews max</p>
                          <p className="text-xs mt-2">
                            <span className={readyHops.length >= 2 ? 'text-green-400' : 'text-yellow-400'}>
                              {readyHops.length} ready HOPs
                            </span>
                          </p>
                          <p className={`text-xs font-bold mt-1 ${hasEnoughPipeline && canExpand ? 'text-green-400' : 'text-gray-500'}`}>
                            {!canExpand ? '✅ At max crews' :
                             hasEnoughPipeline ? '✅ Ready to expand' :
                             `â³ Need ${(curCrews + 1) * 2 - readyHops.length} more ready HOPs`}
                          </p>
                        </div>
                      )
                    })}
                  </div>
                </div>
              </div>
            )}

            {/* GANTT TAB */}
            {activeTab === 'gantt' && (
              <div className="bg-gray-900 rounded-xl border border-gray-700 p-6">
                <h2 className="text-lg font-bold mb-4">Crew Schedule — 16 Week View</h2>
                <div className="overflow-x-auto">
                  <table className="w-full text-xs border-collapse">
                    <thead>
                      <tr className="bg-gray-800">
                        <th className="text-left p-2 w-40 sticky left-0 bg-gray-800 z-10">Crew</th>
                        {weekStarts.map((w, i) => (
                          <th key={i} className="p-1 text-center text-gray-400 min-w-16 border-l border-gray-700">
                            {fmtShort(w)}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {crewIds.filter(crewId => selectedGC === 'ALL' || crewId.startsWith(selectedGC)).map(crewId => {
                        const crewJobs = crewAssignments
                          .filter(a => a.crewId === crewId)
                          .sort((a, b) => a.start.getTime() - b.start.getTime())

                        return (
                          <tr key={crewId} className="border-t border-gray-800">
                            <td className="p-2 font-semibold text-white sticky left-0 bg-gray-900 z-10 whitespace-nowrap">
                              {crewId}
                            </td>
                            {weekStarts.map((weekStart, wi) => {
                              const weekEnd = addDays(weekStart, 6)
                              const job = crewJobs.find(j => j.start <= weekEnd && j.end >= weekStart)
                              const gap = gaps.find(g => g.crewId === crewId && g.gapStart <= weekEnd && g.gapEnd >= weekStart)

                              if (job) {
                                const isFirst = job.start >= weekStart && job.start <= weekEnd
                                return (
                                  <td key={wi} className={`p-1 border-l border-gray-700 ${job.isActive ? 'bg-blue-900' : 'bg-teal-900'}`}
                                    title={`${job.hop} | ${fmtDate(job.start)} - ${fmtDate(job.end)}`}>
                                    {isFirst && (
                                      <div className={`text-xs truncate font-semibold ${job.isActive ? 'text-blue-200' : 'text-teal-200'}`}>
                                        {job.hop.split('-').slice(-1)[0]}
                                      </div>
                                    )}
                                  </td>
                                )
                              }
                              if (gap) {
                                return (
                                  <td key={wi} className="p-1 border-l border-gray-700 bg-red-900" title={`Gap: ${gap.gapDays}d idle`}>
                                    <div className="text-xs text-red-300 text-center">GAP</div>
                                  </td>
                                )
                              }
                              return <td key={wi} className="p-1 border-l border-gray-700 bg-gray-950" />
                            })}
                          </tr>
                        )
                      })}
                    </tbody>
                  </table>

                  {/* Legend */}
                  <div className="flex gap-4 mt-4 text-xs">
                    <span className="flex items-center gap-1"><span className="w-4 h-4 bg-blue-900 rounded inline-block" /> Active HOP</span>
                    <span className="flex items-center gap-1"><span className="w-4 h-4 bg-teal-900 rounded inline-block" /> Pipeline HOP</span>
                    <span className="flex items-center gap-1"><span className="w-4 h-4 bg-red-900 rounded inline-block" /> Crew Gap</span>
                    <span className="flex items-center gap-1"><span className="w-4 h-4 bg-gray-950 rounded inline-block border border-gray-700" /> Available</span>
                  </div>
                </div>
              </div>
            )}
          </>
        )}

        {!loaded && (
          <div className="bg-gray-900 rounded-xl border border-gray-700 p-12 text-center">
            <p className="text-gray-400 text-xl">📂 Upload your tracker to begin schedule analysis</p>
          </div>
        )}

      </div>
    </div>
  )
}
