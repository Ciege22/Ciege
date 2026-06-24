'use client'

export const dynamic = 'force-dynamic'

import { useState, useCallback, useEffect } from 'react'
import * as XLSX from 'xlsx'
import { supabase } from '../lib/supabase'

interface HOP {
  hop: string
  gc: string
  cm: string
  ops: string
  ms15f: string
  ms15a: string
  ms16f: string
  ms16a: string
  mss: string
  powerUp: string
  mainCutover: string
  divCutover: string
  decom: string
  hasNtp: boolean
  hasMat: boolean
  wpApproved: boolean
  gcPickup: boolean
  gcPickupDate: string
  ntpOwner: string
  ntpWaitingOn: string
  matForecast: string
  matReceived: string
  matLocation: string
  steelFrom: string
  vendorWindow: string
  blockers: string[]
  daysOut: number | null
  daysElapsed: number | null
  inProgress: boolean
  complete: boolean
  statuses: string[]
  cmAction: string
}

interface CallNote {
  id: string
  hop_name: string
  note: string
  logged_at: string
}

interface PmUpdate {
  hop: string
  field: string
  oldValue: string
  newValue: string
  timestamp: string
  completed?: boolean
}

function parseDate(val: unknown): Date | null {
  if (!val) return null
  if (val instanceof Date) {
    if (val.getFullYear() >= 2025) return val
    return null
  }
  if (typeof val === 'number') {
    const d = new Date((val - 25569) * 86400 * 1000)
    if (d.getFullYear() >= 2025) return d
    return null
  }
  const d = new Date(String(val))
  return isNaN(d.getTime()) || d.getFullYear() < 2025 ? null : d
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

function getStatuses(h: HOP): string[] {
  const s: string[] = []
  if (h.daysElapsed !== null && h.daysElapsed > 18) s.push('⚠️ Over 18d')
  if (h.mss && !h.powerUp) s.push('📡 MSS done — awaiting power-up')
  if (h.mainCutover && !h.divCutover) s.push('🔗 Main cutover done — diversity pending')
  if (h.powerUp && !h.decom) s.push('♻️ Power-up done — decom pending')
  if (s.length === 0) s.push('✅ On track')
  return s
}

function getCmAction(h: HOP): string {
  if (h.complete) return '✅ Complete — verify close-out package uploaded to QB'
  if (h.inProgress) {
    const elapsed = h.daysElapsed ?? 0
    if (h.mss && !h.powerUp) return '📡 MSS/NMS done — chase Viaero for power-up, log wait time'
    if (h.mainCutover && !h.divCutover) return '🔗 Main cutover done — schedule diversity cutover'
    if (h.powerUp && !h.decom) return '♻️ Power-up complete — schedule decom pickup and return'
    if (elapsed > 18) return `⚠️ ${elapsed}d elapsed — provide updated completion date to CJ`
    return '🔨 Active — confirm crew on site, no access issues, update CJ on M/W/F'
  }
  const days = h.daysOut
  if (!h.hasNtp && days !== null && days <= 7) return '🔴 CRITICAL — NTP missing, starts this week, alert CJ immediately'
  if (!h.hasMat && days !== null && days <= 7) return '🔴 CRITICAL — material not received, starts this week, alert CJ'
  if (h.hasMat && !h.gcPickup && days !== null && days <= 7) return '🔴 URGENT — material in warehouse, coordinate GC pickup today'
  if (!h.hasNtp && days !== null && days <= 14) return '🟠 Chase NTP — starts in 2 weeks, alert CJ'
  if (!h.hasMat && days !== null && days <= 14) return `🟠 Material not received — forecast ${h.matForecast || 'TBD'}, alert CJ`
  if (h.hasMat && !h.gcPickup && days !== null && days <= 14) return '🟠 Material in warehouse — coordinate GC pickup this week'
  if (days !== null && days <= 30) return `🟡 Monitor — confirm readiness as start date approaches`
  return '👀 Pipeline — monitor'
}

export default function CMViewPage() {
  const [hops, setHops] = useState<HOP[]>([])
  const [loaded, setLoaded] = useState(false)
  const [fileName, setFileName] = useState('')
  const [selectedCM, setSelectedCM] = useState('')
  const [cmList, setCmList] = useState<string[]>([])
  const [sessionNotes, setSessionNotes] = useState<Record<string, string>>({})
  const [noteHistory, setNoteHistory] = useState<Record<string, CallNote[]>>({})
  const [pmUpdates, setPmUpdates] = useState<PmUpdate[]>([])
  const [showPmUpdates, setShowPmUpdates] = useState(false)
  const [editedDates, setEditedDates] = useState<Record<string, Record<string, string>>>({})
  const today = new Date()

  useEffect(() => {
    const loadNoteHistory = async () => {
      const { data, error } = await supabase
        .from('hop_call_notes')
        .select('*')
        .order('logged_at', { ascending: false })
      if (error) { console.error('Error loading notes:', error); return }
      if (data) {
        const map: Record<string, CallNote[]> = {}
        data.forEach((row: CallNote) => {
          if (!map[row.hop_name]) map[row.hop_name] = []
          map[row.hop_name].push(row)
        })
        setNoteHistory(map)
      }
    }
    loadNoteHistory()
  }, [])

  const saveCallNote = async (hop: string, note: string) => {
    if (!note.trim()) return
    const { data, error } = await supabase
      .from('hop_call_notes')
      .insert({ hop_name: hop, note: note.trim() })
      .select()
    if (error) { console.error('Error saving note:', error); return }
    if (data) {
      setNoteHistory(prev => ({ ...prev, [hop]: [data[0], ...(prev[hop] || [])] }))
      setSessionNotes(prev => ({ ...prev, [hop]: '' }))
    }
  }

  const logDateEdit = (hop: string, field: string, oldVal: string, newVal: string) => {
    if (!newVal || newVal === oldVal) return
    setEditedDates(prev => ({ ...prev, [hop]: { ...(prev[hop] || {}), [field]: newVal } }))
    setPmUpdates(prev => {
      const filtered = prev.filter(u => !(u.hop === hop && u.field === field))
      return [...filtered, { hop, field, oldValue: oldVal, newValue: newVal, timestamp: new Date().toLocaleTimeString() }]
    })
  }

  const logNA = (hop: string, field: string) => {
    logDateEdit(hop, field, '', 'N/A')
  }

  const EditableDate = ({ hop, field, value }: { hop: string, field: string, value: string }) => {
    const edited = editedDates[hop]?.[field]
    const display = edited || value
    const toInputFormat = (dateStr: string) => {
      if (!dateStr || dateStr === 'N/A') return ''
      const parts = dateStr.split('/')
      if (parts.length !== 3) return ''
      return `${parts[2]}-${parts[0].padStart(2, '0')}-${parts[1].padStart(2, '0')}`
    }
    const toDisplayFormat = (dateStr: string) => {
      if (!dateStr) return ''
      const parts = dateStr.split('-')
      if (parts.length !== 3) return ''
      return `${parseInt(parts[1])}/${parseInt(parts[2])}/${parts[0]}`
    }
    if (display === 'N/A') {
      return (
        <div className="flex items-center gap-1">
          <span className="text-gray-500 text-xs">N/A</span>
          <button onClick={() => logDateEdit(hop, field, 'N/A', '')}
            className="text-gray-600 hover:text-gray-400 text-xs">✕</button>
        </div>
      )
    }
    if (display) {
      return (
        <div className="flex flex-col gap-1">
          <span className={`text-xs ${edited ? 'text-yellow-300' : 'text-green-400'}`}>
            {edited ? `📝 ${display}` : `✓ ${display}`}
          </span>
        </div>
      )
    }
    return (
      <div className="flex flex-col gap-1">
        <input type="date" value={toInputFormat(display)}
          onChange={(e) => logDateEdit(hop, field, value, toDisplayFormat(e.target.value))}
          className="text-xs rounded px-1 py-1 border border-gray-600 bg-gray-800 text-gray-300 focus:outline-none focus:border-blue-500 cursor-pointer"
        />
        <button onClick={() => logNA(hop, field)}
          className="text-gray-500 hover:text-gray-300 text-xs text-left">N/A</button>
      </div>
    )
  }

  const NoteCell = ({ hop }: { hop: string }) => (
    <td className="p-2">
      <div className="flex gap-1">
        <input type="text" placeholder="Type note..."
          value={sessionNotes[hop] || ''}
          onChange={(e) => setSessionNotes(n => ({ ...n, [hop]: e.target.value }))}
          onKeyDown={(e) => { if (e.key === 'Enter') saveCallNote(hop, sessionNotes[hop] || '') }}
          className="w-32 bg-gray-800 text-white text-xs rounded px-2 py-1 border border-gray-600 focus:outline-none focus:border-blue-500"
        />
        <button onClick={() => saveCallNote(hop, sessionNotes[hop] || '')}
          className="bg-blue-700 hover:bg-blue-600 text-white text-xs px-2 py-1 rounded">💾</button>
      </div>
    </td>
  )

  const HistoryCell = ({ hop }: { hop: string }) => (
    <td className="p-2 max-w-48">
      {(noteHistory[hop] || []).length === 0
        ? <span className="text-gray-600 text-xs">No history</span>
        : <div className="flex flex-col gap-1 max-h-20 overflow-y-auto">
            {(noteHistory[hop] || []).slice(0, 5).map((n) => (
              <div key={n.id} className="text-xs">
                <span className="text-gray-500">{new Date(n.logged_at).toLocaleDateString('en-US', { month: 'numeric', day: 'numeric' })} </span>
                <span className="text-gray-300">{n.note}</span>
              </div>
            ))}
          </div>
      }
    </td>
  )

  const handleFile = useCallback((file: File) => {
    setFileName(file.name)
    const reader = new FileReader()
    reader.onload = (e) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer)
      const wb = XLSX.read(data, { type: 'array', cellDates: true })
      const ws = wb.Sheets['HOPs']
      if (!ws) { alert('HOPs tab not found'); return }
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 }) as unknown[][]

      let headerRow = -1
      for (let i = 0; i < 10; i++) {
        if ((rows[i] as unknown[])?.some(c => String(c).trim() === 'HOP')) { headerRow = i; break }
      }
      if (headerRow === -1) { alert('Could not find header row'); return }

      const headers = rows[headerRow] as string[]
      const col = (name: string) => headers.findIndex(h => String(h).trim() === name)

      const hopCol      = col('HOP')
      const gcCol       = col('General Contractor')
      const regionPmCol = col('Region PM')
      const newCmCol    = col('New CM')
      const opsCol      = col('Viaero Ops Field Ops')
      const don444Col   = col('DON 444')
      const ms15fCol    = col('MS15 Implementation Start F')
      const ms15aCol    = col('MS15 Implementation Start A')
      const ms16fCol    = col('MS16 Implementation Ends F')
      const ms16aCol    = col('MS16 Implementation Ends A')
      const mssCol      = col('MSS Completed NMS Ready ')
      const powerCol    = col('Power-Up Completion')
      const mainCutCol  = col('Main Path Cutover Completed')
      const divCutCol   = col('Diversity Cutover Completed')
      const decomCol    = col('Decom Complete')
      const ntpCol      = col('NTP A')
      const matCol      = headers.findIndex(h => String(h).trim().replace(/\s+$/, '') === 'Material Received A'.trim())
      const matFcCol    = col('Material Forecast +4ish')
      const wpCol       = col('Work Package Approved in QB')
      const pickupCol   = col('GC Material Pick-up (A)')
      const ntpOwnCol   = col('NTP Action Owner')
      const ntpWaitCol  = col('NTP is waiting on')
      const matLocCol   = col('Material Current Location')
      const steelCol    = headers.findIndex(h => String(h).trim().toLowerCase().includes('steel from'))
      const itwSCol     = col('ITW Schedule Start')
      const itwECol     = col('ITW Schedule Complete')
      const ssSCol      = col('Samsung Schedule Start')
      const ssECol      = col('Samsung Schedule Complete')
      const siteNameCol = col('Site Name')

      const hopRows = new Map<string, unknown[][]>()
      for (let i = headerRow + 1; i < rows.length; i++) {
        const row = rows[i] as unknown[]
        const don = String(row[don444Col] || '').trim().toUpperCase()
        if (don !== 'DON 444') continue
        const regionPm = String(row[regionPmCol] || '').trim().toUpperCase()
        if (regionPm !== 'CJ') continue
        const hop = String(row[hopCol] || '').trim()
        if (!hop || hop === 'undefined') continue
        if (!hopRows.has(hop)) hopRows.set(hop, [])
        hopRows.get(hop)!.push(row)
      }

      const parsed: HOP[] = []
      hopRows.forEach((rows2, hop) => {
        const row  = rows2[0]
        const row2 = rows2[1] || null

        const ms15f    = parseDateAny(row[ms15fCol])
        const ms15a    = parseDate(row[ms15aCol])
        const ms16a    = parseDate(row[ms16aCol])
        const ntpDate  = parseDate(row[ntpCol])
        const matDate  = parseDateAny(row[matCol])
        const wpDate   = parseDateAny(row[wpCol])
        const pickupD  = parseDateAny(row[pickupCol])
        const mssDate  = parseDateAny(row[mssCol])
        const powerDate= parseDateAny(row[powerCol])
        const mainDate = parseDateAny(row[mainCutCol])
        const divDate  = parseDateAny(row[divCutCol])
        const decomDate= parseDateAny(row[decomCol])

        const hasNtp     = !!(ntpDate && ntpDate.getFullYear() >= 2025)
        const hasMat     = !!(matDate && matDate.getFullYear() >= 2020)
        const wpApproved = !!wpDate
        const gcPickup   = !!pickupD
        const started    = !!ms15a
        const complete   = !!ms16a
        const inProgress = started && !complete
        const daysOut    = ms15f ? daysBetween(today, ms15f) : null
        const daysElapsed = inProgress && ms15a ? daysBetween(ms15a, today) : null

        // Vendor window — read both rows
        const allVendorParts: string[] = []
        rows2.forEach(r => {
          const rItwS  = parseDateAny(r[itwSCol])
          const rItwE  = parseDateAny(r[itwECol])
          const rSsS   = parseDateAny(r[ssSCol])
          const rSsE   = parseDateAny(r[ssECol])
          const rMs15f = parseDateAny(r[ms15fCol]) || ms15f
          if (!rMs15f) return
          if (!rSsS && !rSsE && !rItwS && !rItwE) return
          const siteName  = String(r[siteNameCol] || '').trim()
          const siteLabel = siteName ? ` (${siteName})` : ''
          const checkV = (name: string, start: Date | null, end: Date | null) => {
            if (!start || !end) return
            const ms15fTime = rMs15f.getTime()
            const startTime = start.getTime()
            const endTime   = end.getTime()
            if (startTime <= ms15fTime && ms15fTime <= endTime) {
              allVendorParts.push(`🔴 ${name} on site thru ${fmtDM(end)}${siteLabel}`)
            } else if (endTime < ms15fTime) {
              const buf = Math.round((ms15fTime - endTime) / (1000 * 60 * 60 * 24))
              if (buf <= 5)       allVendorParts.push(`🔴 ${name} clears ${fmtDM(end)} — only ${buf}d${siteLabel}`)
              else if (buf <= 10) allVendorParts.push(`⚠️ ${name} clears ${fmtDM(end)} — ${buf}d buffer${siteLabel}`)
              else                allVendorParts.push(`✅ ${name} clears ${fmtDM(end)}${siteLabel}`)
            } else {
              const buf = Math.round((startTime - ms15fTime) / (1000 * 60 * 60 * 24))
              if (buf <= 10) allVendorParts.push(`⚠️ ${name} starts ${fmtDM(start)} — ${buf}d after start${siteLabel}`)
              else           allVendorParts.push(`✅ ${name} starts ${fmtDM(start)}${siteLabel}`)
            }
          }
          checkV('ITW', rItwS, rItwE)
          checkV('Samsung', rSsS, rSsE)
        })

        const itwParts  = Array.from(new Set(allVendorParts.filter(p => p.includes('ITW'))))
        const ssParts   = Array.from(new Set(allVendorParts.filter(p => p.includes('Samsung'))))
        const sortParts = (parts: string[]) => [...parts.filter(p => p.includes('🔴')), ...parts.filter(p => p.includes('⚠️')), ...parts.filter(p => p.includes('✅'))]
        const vendorWindow = [...sortParts(itwParts), ...sortParts(ssParts)].filter(Boolean).join(' | ') || '✅ No conflicts'

        const blockers: string[] = []
        if (!hasNtp) blockers.push('🔴 NTP pending')
        if (!hasMat) blockers.push('🔴 Material not received')
        if (hasMat && !gcPickup) blockers.push('🟠 Mat in warehouse — GC pickup needed')
        if (!wpApproved) blockers.push('🟡 WP not approved')
        if (vendorWindow.includes('🔴')) blockers.push('🔴 Vendor conflict')

        const hopObj: HOP = {
          hop,
          gc:           String(row[gcCol] || '').trim(),
          cm:           String(row[newCmCol] || '').trim() || String(row2?.[newCmCol] || '').trim(),
          ops:          String(row[opsCol] || '').trim(),
          ms15f:        fmtDate(ms15f),
          ms15a:        fmtDate(ms15a),
          ms16f:        fmtDate(parseDateAny(row[ms16fCol])),
          ms16a:        fmtDate(ms16a),
          mss:          fmtDate(mssDate),
          powerUp:      fmtDate(powerDate),
          mainCutover:  fmtDate(mainDate),
          divCutover:   fmtDate(divDate),
          decom:        fmtDate(decomDate),
          hasNtp, hasMat, wpApproved, gcPickup,
          gcPickupDate: fmtDate(pickupD),
          ntpOwner:     String(row[ntpOwnCol] || '').trim() || String(row2?.[ntpOwnCol] || '').trim(),
          ntpWaitingOn: String(row[ntpWaitCol] || '').trim() || String(row2?.[ntpWaitCol] || '').trim(),
          matForecast:  fmtDate(parseDateAny(row[matFcCol])),
          matReceived:  hasMat ? fmtDate(matDate) : '',
          matLocation:  String(row[matLocCol] || '').trim() || String(row2?.[matLocCol] || '').trim(),
          steelFrom:    String(row[steelCol] || '').trim() || String(row2?.[steelCol] || '').trim(),
          vendorWindow, blockers,
          daysOut, daysElapsed, inProgress, complete,
          statuses: [],
          cmAction: ''
        }
        hopObj.statuses = getStatuses(hopObj)
        hopObj.cmAction = getCmAction(hopObj)
        parsed.push(hopObj)
      })

      const uniqueCMs = Array.from(new Set(parsed.map(h => h.cm).filter(Boolean))).sort()
      setCmList(uniqueCMs)
      setHops(parsed)
      setLoaded(true)
      setSelectedCM('')
    }
    reader.readAsArrayBuffer(file)
  }, [today])

  const cmHops    = hops.filter(h => h.cm === selectedCM)
  const active    = cmHops.filter(h => h.inProgress).sort((a, b) => (b.daysElapsed ?? 0) - (a.daysElapsed ?? 0))
  const thisWeek  = cmHops.filter(h => !h.inProgress && !h.complete && h.daysOut !== null && h.daysOut >= 0 && h.daysOut <= 7).sort((a, b) => (a.daysOut ?? 0) - (b.daysOut ?? 0))
  const next2Wks  = cmHops.filter(h => !h.inProgress && !h.complete && h.daysOut !== null && h.daysOut > 7 && h.daysOut <= 14).sort((a, b) => (a.daysOut ?? 0) - (b.daysOut ?? 0))
  const thisMonth = cmHops.filter(h => !h.inProgress && !h.complete && h.daysOut !== null && h.daysOut > 14 && h.daysOut <= 30).sort((a, b) => (a.daysOut ?? 0) - (b.daysOut ?? 0))
  const pipeline  = cmHops.filter(h => !h.inProgress && !h.complete && (h.daysOut === null || h.daysOut > 30)).sort((a, b) => (a.daysOut ?? 999) - (b.daysOut ?? 999))

  const generateEmail = () => {
    const date = today.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })
    const subj = `Viaero MW Program — CM Call Follow-Up | ${selectedCM} | ${date}`
    const div  = '─'.repeat(60)

    let body = `${selectedCM},\n\n`
    body += `Per our call today — here is your site summary and action items for ${date}.\n`
    body += `${div}\n\n`

    if (active.length > 0) {
      body += `ACTIVE SITES (${active.length})\n${div}\n`
      active.forEach(h => {
        body += `• ${h.hop}  |  GC: ${h.gc}\n`
        body += `  Started: ${h.ms15a}  |  FC Complete: ${h.ms16f}\n`
        body += `  Status: ${h.statuses.join('  |  ')}\n`
        if (h.mss)         body += `  MSS/NMS Ready: ${h.mss}\n`
        if (h.powerUp)     body += `  Power-Up: ${h.powerUp}\n`
        if (h.mainCutover) body += `  Main Cutover: ${h.mainCutover}\n`
        if (h.divCutover)  body += `  Diversity Cutover: ${h.divCutover}\n`
        if (sessionNotes[h.hop]) body += `  Note: ${sessionNotes[h.hop]}\n`
        body += '\n'
      })
    }

    const upcoming = [...thisWeek, ...next2Wks]
    if (upcoming.length > 0) {
      body += `STARTING WITHIN 2 WEEKS (${upcoming.length})\n${div}\n`
      upcoming.forEach(h => {
        body += `• ${h.hop}  |  GC: ${h.gc}  |  FC Start: ${h.ms15f}  |  ${h.daysOut}d out\n`
        body += `  NTP: ${h.hasNtp ? '✓' : '✗ Pending'}  |  Material: ${h.hasMat ? '✓' : '✗ Pending'}  |  GC Pickup: ${h.gcPickupDate ? '✓ ' + h.gcPickupDate : '✗'}\n`
        body += `  Steel From: ${h.steelFrom || '—'}  |  Mat Location: ${h.matLocation || '—'}\n`
        if (h.vendorWindow && !h.vendorWindow.includes('✅ No conflicts')) body += `  Vendor: ${h.vendorWindow}\n`
        if (h.blockers.length > 0) body += `  Blockers: ${h.blockers.join(' | ')}\n`
        if (sessionNotes[h.hop]) body += `  Note: ${sessionNotes[h.hop]}\n`
        body += '\n'
      })
    }

    // Action items from call notes
    const actionItems = Object.entries(sessionNotes).filter(([hop, note]) => note.trim() && cmHops.some(h => h.hop === hop))
    if (actionItems.length > 0) {
      body += `ACTION ITEMS FROM TODAY'S CALL\n${div}\n`
      actionItems.forEach(([hop, note], i) => {
        body += `${i + 1}. ${hop} — ${note}\n`
      })
      body += '\n'
    }

    body += `${div}\n`
    body += `Please confirm receipt and advise on any open items.\n\n`
    body += `Respectfully,\nCJ\nNokia Program Manager — Viaero MW Construction Program\nCC: Thomas M. — Lead CM`

    window.open(`mailto:?subject=${encodeURIComponent(subj)}&body=${encodeURIComponent(body)}`)
  }

  const PipelineSection = ({ title, rows }: { title: string, rows: HOP[] }) => (
    <div className="mb-8">
      <h3 className="text-base font-semibold text-white mb-3">{title} ({rows.length})</h3>
      {rows.length === 0
        ? <p className="text-gray-500 text-sm">No sites in this window</p>
        : (
          <div className="overflow-x-auto">
            <table className="w-full text-xs">
              <thead>
                <tr className="bg-gray-800 text-gray-400">
                  <th className="text-left p-2">HOP</th>
                  <th className="text-left p-2">GC</th>
                  <th className="text-left p-2">Days Out</th>
                  <th className="text-left p-2">NTP</th>
                  <th className="text-left p-2">Mat</th>
                  <th className="text-left p-2">Steel From</th>
                  <th className="text-left p-2">Mat Location</th>
                  <th className="text-left p-2">GC Pickup</th>
                  <th className="text-left p-2">FC Start</th>
                  <th className="text-left p-2">AC Start</th>
                  <th className="text-left p-2">Vendor Window</th>
                  <th className="text-left p-2">CM Action</th>
                  <th className="text-left p-2">Call Notes</th>
                  <th className="text-left p-2">Notes History</th>
                </tr>
              </thead>
              <tbody>
                {rows.map(h => {
                  const hasConflict = h.vendorWindow.includes('🔴')
                  const isUrgent    = h.blockers.length > 0 && (h.daysOut ?? 99) <= 7
                  const rowBg       = hasConflict ? 'bg-red-950' : isUrgent ? 'bg-yellow-950' : h.blockers.length === 0 ? 'bg-green-950' : 'bg-gray-900'
                  return (
                    <tr key={h.hop} className={`border-t border-gray-800 ${rowBg}`}>
                      <td className="p-2 font-semibold text-white whitespace-nowrap">{h.hop}</td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{h.gc}</td>
                      <td className={`p-2 font-bold whitespace-nowrap ${(h.daysOut ?? 99) <= 7 ? 'text-red-400' : (h.daysOut ?? 99) <= 14 ? 'text-yellow-400' : 'text-gray-300'}`}>
                        {h.daysOut !== null ? `${h.daysOut}d` : '—'}
                      </td>
                      <td className="p-2">
                        <span
                          title={!h.hasNtp ? (h.ntpWaitingOn || h.ntpOwner || 'NTP pending') : 'NTP confirmed'}
                          className="cursor-help">
                          {h.hasNtp
                            ? <span className="text-green-400 font-bold">✓</span>
                            : <span className="text-red-400 font-bold">✗</span>}
                        </span>
                      </td>
                      <td className="p-2">
                        <span
                          title={!h.hasMat ? (h.matForecast ? `Mat forecast: ${h.matForecast}` : 'No material forecast') : `Mat received: ${h.matReceived}`}
                          className="cursor-help">
                          {h.hasMat
                            ? <span className="text-green-400 font-bold">✓</span>
                            : <span className="text-red-400 font-bold">✗</span>}
                        </span>
                      </td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{h.steelFrom || '—'}</td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{h.matLocation || '—'}</td>
                      <td className="p-2">
                        {h.gcPickupDate
                          ? <span className="text-green-400 text-xs">✓ {h.gcPickupDate}</span>
                          : <EditableDate hop={h.hop} field="GC Material Pick-up (A)" value="" />
                        }
                      </td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{h.ms15f}</td>
                      <td className="p-2">
                        <EditableDate hop={h.hop} field="MS15 Implementation Start A" value={h.ms15a} />
                      </td>
                      <td className="p-2 text-xs max-w-40">
                        <span
                          className={h.vendorWindow.includes('🔴') ? 'text-red-400' : h.vendorWindow.includes('⚠️') ? 'text-yellow-400' : 'text-green-400'}
                          title={h.vendorWindow}>
                          {h.vendorWindow.length > 30 ? h.vendorWindow.slice(0, 30) + '...' : h.vendorWindow}
                        </span>
                      </td>
                      <td className="p-2 text-xs max-w-48">
                        <span
                          className={h.cmAction.includes('🔴') || h.cmAction.includes('CRITICAL') ? 'text-red-400' : h.cmAction.includes('🟠') ? 'text-orange-400' : h.cmAction.includes('🟡') ? 'text-yellow-400' : h.cmAction.includes('✅') ? 'text-green-400' : 'text-gray-300'}
                          title={h.cmAction}>
                          {h.cmAction.length > 40 ? h.cmAction.slice(0, 40) + '...' : h.cmAction}
                        </span>
                      </td>
                      <NoteCell hop={h.hop} />
                      <HistoryCell hop={h.hop} />
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>
        )
      }
    </div>
  )

  return (
    <div className="min-h-screen bg-gray-950 text-white p-6">
      <div className="max-w-full mx-auto">

        {/* Header */}
        <div className="flex items-center justify-between mb-6">
          <div>
            <h1 className="text-3xl font-bold">CM View</h1>
            <p className="text-gray-400 mt-1">Select a Site CM to view their pipeline, log field updates, and generate a follow-up email.</p>
          </div>
          {pmUpdates.length > 0 && (
            <button onClick={() => setShowPmUpdates(!showPmUpdates)}
              className="bg-yellow-600 hover:bg-yellow-700 text-white px-4 py-2 rounded-lg text-sm font-semibold">
              📋 PM Updates ({pmUpdates.length})
            </button>
          )}
        </div>

        {/* PM Daily Updates Panel */}
        {showPmUpdates && pmUpdates.length > 0 && (
          <div className="mb-6 bg-yellow-950 border border-yellow-600 rounded-xl p-5">
            <h2 className="text-yellow-300 font-bold text-lg mb-3">📋 PM Daily Updates — Update These in Your Tracker</h2>
            <table className="w-full text-sm">
              <thead>
                <tr className="text-yellow-400 text-xs">
                  <th className="text-left p-2">Done</th>
                  <th className="text-left p-2">HOP</th>
                  <th className="text-left p-2">Field</th>
                  <th className="text-left p-2">Old Value</th>
                  <th className="text-left p-2">New Value</th>
                  <th className="text-left p-2">Logged At</th>
                </tr>
              </thead>
              <tbody>
                {pmUpdates.map((u, i) => (
                  <tr key={i} className={`border-t border-yellow-800 ${u.completed ? 'opacity-40' : ''}`}>
                    <td className="p-2">
                      <input type="checkbox" checked={u.completed || false}
                        onChange={() => setPmUpdates(prev => prev.map((item, idx) => idx === i ? { ...item, completed: !item.completed } : item))}
                        className="w-4 h-4 cursor-pointer accent-green-500" />
                    </td>
                    <td className={`p-2 font-semibold ${u.completed ? 'line-through text-gray-500' : 'text-white'}`}>{u.hop}</td>
                    <td className={`p-2 ${u.completed ? 'line-through text-gray-500' : 'text-yellow-300'}`}>{u.field}</td>
                    <td className="p-2 text-gray-400">{u.oldValue || '—'}</td>
                    <td className={`p-2 font-bold ${u.completed ? 'text-gray-500' : 'text-green-400'}`}>{u.newValue}</td>
                    <td className="p-2 text-gray-500">{u.timestamp}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            <div className="mt-4 flex gap-3">
              <button onClick={() => { const r = pmUpdates.filter(u => !u.completed); setPmUpdates(r); if (r.length === 0) setShowPmUpdates(false) }}
                className="bg-green-700 hover:bg-green-600 text-white px-4 py-2 rounded text-sm font-semibold">
                ✅ Clear Completed
              </button>
              <button onClick={() => { setPmUpdates([]); setShowPmUpdates(false) }}
                className="bg-gray-700 hover:bg-gray-600 text-white px-4 py-2 rounded text-sm">
                🗑 Clear All
              </button>
            </div>
          </div>
        )}

        {/* File Upload */}
        <div
          className="mb-6 border-2 border-dashed border-gray-600 rounded-xl p-5 text-center cursor-pointer hover:border-blue-500 transition-colors"
          onDragOver={(e) => e.preventDefault()}
          onDrop={(e) => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) handleFile(f) }}
          onClick={() => document.getElementById('cm-tracker-upload')?.click()}
        >
          <input id="cm-tracker-upload" type="file" accept=".xlsx" className="hidden"
            onChange={(e) => { const f = e.target.files?.[0]; if (f) handleFile(f) }} />
          {loaded
            ? <p className="text-green-400 font-semibold">✅ {fileName} — {hops.length} HOPs loaded</p>
            : <p className="text-gray-400">📂 Drop your tracker here or click to upload</p>}
        </div>

        {/* CM Selector */}
        <div className="flex gap-3 mb-6 flex-wrap">
          {cmList.map((cm) => (
            <button key={cm} onClick={() => setSelectedCM(cm)}
              className={`px-6 py-3 rounded-lg font-semibold text-sm transition-all ${selectedCM === cm ? 'bg-blue-600 text-white shadow-lg scale-105' : 'bg-gray-800 text-gray-300 hover:bg-gray-700'}`}>
              {cm}
            </button>
          ))}
        </div>

        {/* CM Panel */}
        {selectedCM && (
          <div className="bg-gray-900 rounded-xl border border-gray-700 p-6">
            <div className="flex items-center justify-between mb-6">
              <div>
                <h2 className="text-2xl font-bold">{selectedCM}</h2>
                {loaded && (
                  <p className="text-gray-400 mt-1">
                    <span className="text-gray-500">
                      {cmHops.length} HOPs · {active.length} active · {thisWeek.length} this week · {next2Wks.length} next 2 wks · {thisMonth.length} this month · {pipeline.length} pipeline
                    </span>
                  </p>
                )}
              </div>
              <button onClick={generateEmail}
                className="bg-blue-600 hover:bg-blue-700 text-white px-5 py-2 rounded-lg text-sm font-semibold">
                ✉️ Generate Follow-Up Email
              </button>
            </div>

            {loaded && (
              <div className="space-y-8">

                {/* Active Sites */}
                <div>
                  <h3 className="text-lg font-semibold text-white mb-3">🔨 Active Sites ({active.length})</h3>
                  {active.length === 0
                    ? <p className="text-gray-500 text-sm">No active sites</p>
                    : (
                      <div className="overflow-x-auto">
                        <table className="w-full text-xs">
                          <thead>
                            <tr className="bg-gray-800 text-gray-400">
                              <th className="text-left p-2">HOP</th>
                              <th className="text-left p-2">GC</th>
                              <th className="text-left p-2">Started</th>
                              <th className="text-left p-2">FC End</th>
                              <th className="text-left p-2">Days Elapsed</th>
                              <th className="text-left p-2">Status</th>
                              <th className="text-left p-2">MSS</th>
                              <th className="text-left p-2">Power-Up</th>
                              <th className="text-left p-2">Main Cutover</th>
                              <th className="text-left p-2">Diversity Cutover</th>
                              <th className="text-left p-2">MS16 Fc</th>
                              <th className="text-left p-2">Edit MS16 Fc</th>
                              <th className="text-left p-2">MS16 Act</th>
                              <th className="text-left p-2">Decom</th>
                              <th className="text-left p-2">Call Notes</th>
                              <th className="text-left p-2">Notes History</th>
                            </tr>
                          </thead>
                          <tbody>
                            {active.map(h => (
                              <tr key={h.hop} className={`border-t border-gray-800 ${h.statuses.some(s => s.includes('⚠️')) ? 'bg-red-950' : 'bg-gray-900'}`}>
                                <td className="p-2 font-semibold text-white whitespace-nowrap">{h.hop}</td>
                                <td className="p-2 text-gray-300 whitespace-nowrap">{h.gc}</td>
                                <td className="p-2 text-gray-300 whitespace-nowrap">{h.ms15a}</td>
                                <td className="p-2 text-gray-300 whitespace-nowrap">{h.ms16f}</td>
                                <td className={`p-2 font-bold ${(h.daysElapsed ?? 0) > 18 ? 'text-red-400' : 'text-green-400'}`}>
                                  {h.daysElapsed}d
                                </td>
                                <td className="p-2">
                                  <div className="flex flex-col gap-1">
                                    {h.statuses.map((s, i) => (
                                      <span key={i} className={`text-xs ${s.includes('⚠️') ? 'text-red-400' : s.includes('📡') || s.includes('🔗') || s.includes('♻️') ? 'text-yellow-400' : 'text-green-400'}`}>
                                        {s}
                                      </span>
                                    ))}
                                  </div>
                                </td>
                                <td className="p-2">
                                  <EditableDate hop={h.hop} field="MSS Completed NMS Ready " value={h.mss} />
                                </td>
                                <td className="p-2">
                                  <EditableDate hop={h.hop} field="Power-Up Completion" value={h.powerUp} />
                                </td>
                                <td className="p-2">
                                  <EditableDate hop={h.hop} field="Main Path Cutover Completed" value={h.mainCutover} />
                                </td>
                                <td className="p-2">
                                  <EditableDate hop={h.hop} field="Diversity Cutover Completed" value={h.divCutover} />
                                </td>
                                <td className="p-2 text-gray-300 whitespace-nowrap">{h.ms16f}</td>
                                <td className="p-2">
                                  <EditableDate hop={h.hop} field="MS16 Implementation Ends F" value={h.ms16f} />
                                </td>
                                <td className="p-2">
                                  <EditableDate hop={h.hop} field="MS16 Implementation Ends A" value={h.ms16a} />
                                </td>
                                <td className="p-2">
                                  <EditableDate hop={h.hop} field="Decom Complete" value={h.decom} />
                                </td>
                                <NoteCell hop={h.hop} />
                                <HistoryCell hop={h.hop} />
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    )
                  }
                </div>

                {/* Pipeline Sections */}
                <PipelineSection title="⚡ This Week (0–7 days)" rows={thisWeek} />
                <PipelineSection title="🟠 Next 2 Weeks (8–14 days)" rows={next2Wks} />
                <PipelineSection title="🟡 This Month (15–30 days)" rows={thisMonth} />
                <PipelineSection title="🔵 Full Pipeline (30d+)" rows={pipeline} />

              </div>
            )}

            {!loaded && (
              <div className="bg-gray-800 rounded-lg p-8 text-center border border-dashed border-gray-600">
                <p className="text-gray-400">📂 Upload your tracker above to load {selectedCM} pipeline</p>
              </div>
            )}
          </div>
        )}

        {!selectedCM && (
          <div className="bg-gray-900 rounded-xl border border-gray-700 p-12 text-center">
            <p className="text-gray-400 text-xl">👆 Select a CM above to begin</p>
          </div>
        )}

      </div>
    </div>
  )
}
