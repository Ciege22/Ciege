'use client'

export const dynamic = 'force-dynamic'

import { useState, useCallback, useEffect } from 'react'
import * as XLSX from 'xlsx'
import { supabase, loadTrackerSnapshot } from '../lib/supabase'


const GC_CM_MAP: Record<string, string> = {
  'MZI': 'Steve',
  'NV Tel': 'Steve',
  'Mastec': 'Benny',
  'Vikor': 'Benny',
  'Tech CX': 'Hap',
}

interface HOP {
  hop: string
  gc: string
  ops: string
  ms15f: string
  ms15a: string
  ms16f: string
  ms16a: string
  hasNtp: boolean
  hasMat: boolean
  wpApproved: boolean
  gcPickup: boolean
  ntpOwner: string
  ntpWaitingOn: string
  matForecast: string
  matReceived: string
  gcPickupDate: string
  hasSpo: boolean
  spoIssued: string
  steelFrom: string
  itwStart: string
  itwEnd: string
  ssStart: string
  ssEnd: string
  daysOut: number | null
  daysElapsed: number | null
  inProgress: boolean
  complete: boolean
  vendorWindow: string
  blockers: string[]
  pullInReady: boolean
  pullInStatus: string
  internalConflict: string
  siteAConflict: string
  siteBConflict: string
}

interface PmUpdate {
  hop: string
  field: string
  oldValue: string
  newValue: string
  timestamp: string
  completed?: boolean
}

interface CallNote {
  id: string
  hop_name: string
  note: string
  logged_at: string
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

function getVendorWindow(ms15f: Date | null, itwS: Date | null, itwE: Date | null, ssS: Date | null, ssE: Date | null): string {
  if (!ms15f) return ''
  const parts: string[] = []

  const checkVendor = (name: string, start: Date | null, end: Date | null) => {
    if (!start || !end) return
    const ms15fTime = ms15f!.getTime()
    const startTime = start.getTime()
    const endTime = end.getTime()

    if (startTime <= ms15fTime && ms15fTime <= endTime) {
      parts.push(`🔴 ${name} on site thru ${fmtDM(end)}`)
      return
    }

    if (endTime < ms15fTime) {
      const bufferDays = Math.round((ms15fTime - endTime) / (1000 * 60 * 60 * 24))
      if (bufferDays <= 5) {
        parts.push(`🔴 ${name} clears ${fmtDM(end)} — only ${bufferDays}d before our start`)
      } else if (bufferDays <= 10) {
        parts.push(`⚠️ ${name} clears ${fmtDM(end)} — ${bufferDays}d buffer, monitor`)
      } else {
        parts.push(`✅ ${name} clears ${fmtDM(end)} — ${bufferDays}d buffer`)
      }
      return
    }

    if (startTime > ms15fTime) {
      const bufferDays = Math.round((startTime - ms15fTime) / (1000 * 60 * 60 * 24))
      if (bufferDays <= 10) {
        parts.push(`⚠️ ${name} starts ${fmtDM(start)} — ${bufferDays}d after our start`)
      } else {
        parts.push(`✅ ${name} starts ${fmtDM(start)} — ${bufferDays}d after our start`)
      }
    }
  }

  checkVendor('ITW', itwS, itwE)
  checkVendor('Samsung', ssS, ssE)

  return parts.join(' | ') || '✅ No conflicts'
}

function getBlockers(h: HOP): string[] {
  const b: string[] = []
  if (!h.hasNtp) b.push('🔴 NTP pending')
  if (!h.hasMat) b.push('🔴 Material not received')
  if (h.hasMat && !h.gcPickup) b.push('🟠 Mat in warehouse — GC pickup needed')
  if (!h.wpApproved) b.push('🟡 WP not approved')
  if (h.vendorWindow.includes('🔴')) b.push(`🔴 Vendor conflict`)
  return b
}

function getPullInStatus(h: HOP): string {
  if (h.inProgress || h.complete) return '—'
  if (!h.hasNtp && !h.hasMat) return '🔴 Not ready — NTP + Mat missing'
  if (!h.hasNtp) return '🔴 Not ready — NTP missing'
  if (!h.hasMat) return '🔴 Not ready — Mat missing'
  if (h.blockers.some(b => b.includes('Vendor conflict'))) return '🔴 Cannot pull in — vendor conflict'
  if (h.vendorWindow.includes('🔴')) return '🔴 Cannot pull in — vendor conflict'
  if (h.vendorWindow.includes('⚠️')) return '⚠️ Risky — monitor vendor window'
  if (h.siteAConflict && h.siteBConflict) return '🔴 Cannot pull in — both sites occupied'
  if (h.siteAConflict) return '⚠️ Site B available — Site A has internal conflict'
  if (h.siteBConflict) return '⚠️ Site A available — Site B has internal conflict'
  if (!h.gcPickup) return '🟡 Ready — GC pickup needed first'
  return '✅ Ready to pull in'
}

export default function GCCallPage() {
  const [selectedGC, setSelectedGC] = useState('')
  const [hops, setHops] = useState<HOP[]>([])
  const [loaded, setLoaded] = useState(false)
  const [fileName, setFileName] = useState('')
  const [gcList, setGcList] = useState<string[]>([])
  const [noteHistory, setNoteHistory] = useState<Record<string, CallNote[]>>({})
  const [sessionNotes, setSessionNotes] = useState<Record<string, string>>({})
  const [editedDates, setEditedDates] = useState<Record<string, Record<string, string>>>({})
  const [pmUpdates, setPmUpdates] = useState<PmUpdate[]>([])
  const [showPmUpdates, setShowPmUpdates] = useState(false)
  const [pmSortAsc, setPmSortAsc] = useState(true)
  const [snapshotTime, setSnapshotTime] = useState<string>('')
  const today = new Date()

  useEffect(() => {
    const loadNoteHistory = async () => {
      const { data, error } = await supabase
        .from('hop_call_notes')
        .select('id, hop_name, note, logged_at')
        .order('logged_at', { ascending: false })
      if (error) { console.error('Error loading note history:', error); return }
      if (data) {
        const historyMap: Record<string, CallNote[]> = {}
        data.forEach((row: CallNote) => {
          if (!historyMap[row.hop_name]) historyMap[row.hop_name] = []
          historyMap[row.hop_name].push(row)
        })
        setNoteHistory(historyMap)
      }
    }
    loadNoteHistory()
  }, [])

  const saveCallNote = async (hop: string) => {
    const note = sessionNotes[hop]
    if (!note?.trim()) return
    const logged_at = new Date().toISOString()
    const { data, error } = await supabase
      .from('hop_call_notes')
      .insert({ hop_name: hop, note: note.trim(), logged_at })
      .select()
      .single()
    if (error) { console.error('Error saving note:', error); return }
    if (data) {
      setNoteHistory(h => ({ ...h, [hop]: [data as CallNote, ...(h[hop] || [])] }))
      setSessionNotes(s => ({ ...s, [hop]: '' }))
    }
  }

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
      const row = rows[i] as unknown[]
      if (row && row.some(cell => String(cell).trim() === 'HOP')) {
        headerRow = i; break
      }
    }
    if (headerRow === -1) { alert('Could not find header row'); return }

    const headers = rows[headerRow] as string[]
    const col = (name: string) => headers.findIndex(h => String(h).trim() === name)

    const hopCol    = col('HOP')
    const gcCol     = col('General Contractor')
    const nokiaPmCol= col('Nokia PM')
    const opsCol    = col('Viaero Ops Field Ops')
    const ms15fCol  = col('MS15 Implementation Start F')
    const ms15aCol  = col('MS15 Implementation Start A')
    const ms16fCol  = col('MS16 Implementation Ends F')
    const ms16aCol  = col('MS16 Implementation Ends A')
    const ntpCol    = col('NTP A')
    const matCol    = headers.findIndex(h => String(h).trim() === 'Material Received A')
    const matFcCol  = col('Material Forecast +4ish')
    const wpCol     = col('Work Package Approved in QB')
    const pickupCol = col('GC Material Pick-up (A)')
    const spoCol    = headers.findIndex(h => String(h).trim().toLowerCase() === 'cx spo issued')
    const steelCol  = headers.findIndex(h => String(h).trim().toLowerCase().includes('steel from'))
    const ntpOwnCol = col('NTP Action Owner')
    const ntpWaitCol= col('NTP is waiting on')
    const don444Col = col('DON 444')
    const siteNameCol = col('Site Name')
    const siteNumCol  = col('Site Number')
    const itwSCol   = col('ITW Schedule Start')
    const itwECol   = col('ITW Schedule Complete')
    const ssSCol    = col('Samsung Schedule Start')
    const ssECol    = col('Samsung Schedule Complete')

    // First pass — collect all rows per HOP
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

    // Build site occupancy map from all in-progress HOPs
    // Maps "SiteName|SiteNumber" -> { gc, hop, ms16f }
    const siteOccupancy = new Map<string, { gc: string, hop: string, ms16f: string }>()

    hopRows.forEach((rows2, hop) => {
      rows2.forEach(r => {
        const ms15a = parseDate(r[ms15aCol])
        const ms16a = parseDate(r[ms16aCol])
        const inProg = !!ms15a && !ms16a
        if (!inProg) return
        const siteName   = String(r[siteNameCol] || '').trim()
        const siteNumber = String(r[siteNumCol]  || '').trim()
        const gc         = String(r[gcCol]        || '').trim()
        const ms16f      = parseDateAny(r[ms16fCol])
        if (siteName && siteNumber) {
          const key = `${siteName}|${siteNumber}`.toLowerCase()
          siteOccupancy.set(key, { gc, hop, ms16f: fmtDate(ms16f) })
        }
      })
    })

    const parsed: HOP[] = []

    hopRows.forEach((rows2, hop) => {
      const row  = rows2[0]
      const row2 = rows2[1] || null

      const gc      = String(row[gcCol] || '').trim()
      const ms15f   = parseDateAny(row[ms15fCol])
      const ms15a   = parseDate(row[ms15aCol])
      const ms16f   = parseDateAny(row[ms16fCol])
      const ms16a   = parseDate(row[ms16aCol])
      const ntpDate = parseDate(row[ntpCol])
      const matDate = parseDateAny(row[matCol])
      const matFc   = parseDateAny(row[matFcCol])
      const wpDate  = parseDateAny(row[wpCol])
      const pickupD = parseDateAny(row[pickupCol])
      const spoDate  = parseDateAny(row[spoCol]) || (row2 ? parseDateAny(row2[spoCol]) : null)
      const steelFrom = String(row[steelCol] || row2?.[steelCol] || '').trim()
      const itwS    = parseDateAny(row[itwSCol]) || (row2 ? parseDateAny(row2[itwSCol]) : null)
      const itwE    = parseDateAny(row[itwECol]) || (row2 ? parseDateAny(row2[itwECol]) : null)
      const ssS     = parseDateAny(row[ssSCol])  || (row2 ? parseDateAny(row2[ssSCol])  : null)
      const ssE     = parseDateAny(row[ssECol])  || (row2 ? parseDateAny(row2[ssECol])  : null)

      const hasNtp     = !!(ntpDate && ntpDate.getFullYear() >= 2025)
      const hasMat     = !!(matDate && matDate.getFullYear() >= 2020)
      const wpApproved = !!wpDate
      const gcPickup   = !!pickupD
      const started    = !!ms15a
      const complete   = !!ms16a
      const inProgress = started && !complete
      const daysOut    = ms15f ? daysBetween(today, ms15f) : null
      const daysElapsed = inProgress && ms15a ? daysBetween(ms15a, today) : null

      // Build vendor window by checking ALL site rows for worst case conflict
      const allVendorParts: string[] = []

      rows2.forEach(r => {
        const rItwS  = parseDateAny(r[itwSCol])
        const rItwE  = parseDateAny(r[itwECol])
        const rSsS   = parseDateAny(r[ssSCol])
        const rSsE   = parseDateAny(r[ssECol])
        const rMs15f = parseDateAny(r[ms15fCol]) || ms15f

        if (!rMs15f) return
        if (!rSsS && !rSsE && !rItwS && !rItwE) return

        const siteName = String(r[siteNameCol] || '').trim()
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
            if (buf <= 5)       allVendorParts.push(`🔴 ${name} clears ${fmtDM(end)} — only ${buf}d before start${siteLabel}`)
            else if (buf <= 10) allVendorParts.push(`⚠️ ${name} clears ${fmtDM(end)} — ${buf}d buffer${siteLabel}`)
            else                allVendorParts.push(`✅ ${name} clears ${fmtDM(end)} — ${buf}d buffer${siteLabel}`)
          } else {
            const buf = Math.round((startTime - ms15fTime) / (1000 * 60 * 60 * 24))
            if (buf <= 10) allVendorParts.push(`⚠️ ${name} starts ${fmtDM(start)} — ${buf}d after start${siteLabel}`)
            else           allVendorParts.push(`✅ ${name} starts ${fmtDM(start)} — ${buf}d after start${siteLabel}`)
          }
        }
        checkV('ITW', rItwS, rItwE)
        checkV('Samsung', rSsS, rSsE)
      })

      const itwParts = Array.from(new Set(allVendorParts.filter(p => p.includes('ITW'))))
      const ssParts  = Array.from(new Set(allVendorParts.filter(p => p.includes('Samsung'))))

      const sortParts = (parts: string[]) => {
        const red    = parts.filter(p => p.includes('🔴'))
        const yellow = parts.filter(p => p.includes('⚠️'))
        const green  = parts.filter(p => p.includes('✅'))
        return [...red, ...yellow, ...green]
      }

      const itwSorted = sortParts(itwParts)
      const ssSorted  = sortParts(ssParts)

      const allParts = [...itwSorted, ...ssSorted].filter(Boolean)
      const vendorWindow = allParts.length > 0 ? allParts.join(' | ') : '✅ No conflicts'

      // Check internal conflicts at site level
      let siteAConflict = ''
      let siteBConflict = ''

      rows2.forEach((r, idx) => {
        const siteName   = String(r[siteNameCol] || '').trim()
        const siteNumber = String(r[siteNumCol]  || '').trim()
        const thisGc     = String(r[gcCol] || '').trim()
        if (!siteName || !siteNumber) return
        const key = `${siteName}|${siteNumber}`.toLowerCase()
        const occupant = siteOccupancy.get(key)
        if (occupant && occupant.gc !== thisGc && occupant.hop !== hop) {
          const conflictMsg = `⚠️ ${occupant.gc} on site — ${occupant.hop} completing ${occupant.ms16f}`
          if (idx === 0) siteAConflict = conflictMsg
          else siteBConflict = conflictMsg
        }
      })

      let internalConflict = ''
      if (siteAConflict && siteBConflict) {
        internalConflict = `🔴 Both sites occupied — Site A: ${siteAConflict} | Site B: ${siteBConflict}`
      } else if (siteAConflict) {
        internalConflict = `⚠️ Site A occupied — ${siteAConflict} | Site B available`
      } else if (siteBConflict) {
        internalConflict = `⚠️ Site B occupied — ${siteBConflict} | Site A available`
      }

      const hopObj: HOP = {
        hop, gc,
        ops:          String(row[opsCol] || '').trim(),
        ms15f:        fmtDate(ms15f),
        ms15a:        fmtDate(ms15a),
        ms16f:        fmtDate(ms16f),
        ms16a:        fmtDate(ms16a),
        hasNtp, hasMat, wpApproved, gcPickup,
        ntpOwner:     String(row[ntpOwnCol] || '').trim() || String(row2?.[ntpOwnCol] || '').trim(),
        ntpWaitingOn: String(row[ntpWaitCol] || '').trim() || String(row2?.[ntpWaitCol] || '').trim(),
        matForecast:  fmtDate(matFc),
        matReceived:  matDate ? fmtDate(matDate) : '',
        gcPickupDate: fmtDate(pickupD),
        hasSpo:       !!(spoDate && spoDate.getFullYear() >= 2020),
        spoIssued:    spoDate ? fmtDate(spoDate) : '',
        steelFrom:    steelFrom === 'nan' ? '' : steelFrom,
        itwStart:     fmtDate(itwS),
        itwEnd:       fmtDate(itwE),
        ssStart:      fmtDate(ssS),
        ssEnd:        fmtDate(ssE),
        daysOut, daysElapsed, inProgress, complete,
        vendorWindow,
        internalConflict,
        siteAConflict,
        siteBConflict,
        blockers: [],
        pullInReady: hasNtp && hasMat && !vendorWindow.includes('🔴') && !vendorWindow.includes('⚠️') && !inProgress && !complete,
        pullInStatus: ''
      }
      hopObj.blockers    = getBlockers(hopObj)
      hopObj.pullInStatus = getPullInStatus(hopObj)
      parsed.push(hopObj)
    })

    setHops(parsed)

    // Build GC list dynamically from parsed HOPs
    const uniqueGCs = Array.from(new Set(parsed.map(h => h.gc).filter(Boolean))).sort()
    setGcList(uniqueGCs)
    setSelectedGC('')
    setLoaded(true)
  }, [today])

  const handleFile = useCallback((file: File) => {
    setFileName(file.name)
    const reader = new FileReader()
    reader.onload = (e) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer)
      const wb = XLSX.read(data, { type: 'array', cellDates: true })
      const ws = wb.Sheets['HOPs']
      if (!ws) { alert('HOPs tab not found in tracker'); return }
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 }) as unknown[][]
      processRows(rows, file.name)
    }
    reader.readAsArrayBuffer(file)
  }, [processRows])

  const logDateEdit = (hop: string, field: string, oldVal: string, newVal: string) => {
    if (!newVal || newVal === oldVal) return
    setEditedDates(prev => ({ ...prev, [hop]: { ...(prev[hop] || {}), [field]: newVal } }))
    setPmUpdates(prev => {
      const filtered = prev.filter(u => !(u.hop === hop && u.field === field))
      return [...filtered, { hop, field, oldValue: oldVal, newValue: newVal, timestamp: new Date().toLocaleTimeString() }]
    })
  }

  const gcHops      = hops.filter(h => h.gc === selectedGC)
  const active      = gcHops.filter(h => h.inProgress).sort((a, b) => (b.daysElapsed ?? 0) - (a.daysElapsed ?? 0))
  const thisWeek    = gcHops.filter(h => !h.inProgress && !h.complete && h.daysOut !== null && h.daysOut >= 0 && h.daysOut <= 7).sort((a, b) => (a.daysOut ?? 0) - (b.daysOut ?? 0))
  const next2Weeks  = gcHops.filter(h => !h.inProgress && !h.complete && h.daysOut !== null && h.daysOut > 7 && h.daysOut <= 14).sort((a, b) => (a.daysOut ?? 0) - (b.daysOut ?? 0))
  const thisMonth   = gcHops.filter(h => !h.inProgress && !h.complete && h.daysOut !== null && h.daysOut > 14 && h.daysOut <= 30).sort((a, b) => (a.daysOut ?? 0) - (b.daysOut ?? 0))
  const pullIns     = gcHops.filter(h => !h.inProgress && !h.complete && h.daysOut !== null && h.daysOut > 30).sort((a, b) => (a.daysOut ?? 0) - (b.daysOut ?? 0))
  const pullInReady = pullIns.filter(h => h.pullInReady)

  const generateEmail = () => {
    const cm   = GC_CM_MAP[selectedGC] || 'CM'
    const date = today.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })
    const subj = `Viaero MW Program — Weekly Site Update | ${selectedGC} | ${date}`

    let body = `Dear ${selectedGC} Team,\n\n`
    body += `Please find below your weekly site update for the Viaero Wireless Microwave Construction Program — ${date}.\n`
    body += `Site CM: ${cm}  ·  Nokia PM: CJ  ·  Lead CM: Thomas M.\n`
    body += `${'─'.repeat(60)}\n\n`

    if (active.length > 0) {
      body += `ACTIVE SITES (${active.length})\n`
      body += `${'─'.repeat(60)}\n`
      active.forEach(h => {
        const status = (h.daysElapsed ?? 0) > 18
          ? `⚠️ OVER TARGET — ${h.daysElapsed}d elapsed — please confirm completion date with your crew and advise`
          : `✅ On track — ${h.daysElapsed}d elapsed`
        body += `• ${h.hop}\n`
        body += `  Started: ${h.ms15a}  |  FC Complete: ${h.ms16f}\n`
        body += `  Status: ${status}\n`
        if (sessionNotes[h.hop]) body += `  Note: ${sessionNotes[h.hop]}\n`
        body += '\n'
      })
    }

    const upcoming = [...thisWeek, ...next2Weeks]
    if (upcoming.length > 0) {
      body += `STARTING WITHIN 2 WEEKS (${upcoming.length})\n`
      body += `${'─'.repeat(60)}\n`
      upcoming.forEach(h => {
        body += `• ${h.hop}\n`
        body += `  FC Start: ${h.ms15f}  |  Days Out: ${h.daysOut}d\n`
        body += `  NTP: ${h.hasNtp ? '✓ Confirmed' : '✗ Pending'}  |  Material: ${h.hasMat ? '✓ Received' : '✗ Pending'}  |  GC Pickup: ${h.gcPickupDate ? '✓ ' + h.gcPickupDate : '✗ Not yet'}\n`
        if (h.blockers.length > 0) body += `  Blockers: ${h.blockers.join('  |  ')}\n`
        if (h.ntpWaitingOn) body += `  NTP Waiting On: ${h.ntpWaitingOn}\n`
        if (h.internalConflict) body += `  Internal Conflict: ${h.internalConflict}\n`
        if (sessionNotes[h.hop]) body += `  Note: ${sessionNotes[h.hop]}\n`
        body += '\n'
      })
    }

    if (thisMonth.length > 0) {
      body += `THIS MONTH — 15 TO 30 DAYS (${thisMonth.length})\n`
      body += `${'─'.repeat(60)}\n`
      thisMonth.forEach(h => {
        body += `• ${h.hop}\n`
        body += `  FC Start: ${h.ms15f}  |  Days Out: ${h.daysOut}d\n`
        body += `  NTP: ${h.hasNtp ? '✓' : '✗ Pending'}  |  Material: ${h.hasMat ? '✓' : '✗ Pending'}  |  GC Pickup: ${h.gcPickupDate ? '✓' : '✗'}\n`
        if (h.blockers.length > 0) body += `  Blockers: ${h.blockers.join('  |  ')}\n`
        if (h.internalConflict) body += `  Internal Conflict: ${h.internalConflict}\n`
        if (sessionNotes[h.hop]) body += `  Note: ${sessionNotes[h.hop]}\n`
        body += '\n'
      })
    }

    if (pullInReady.length > 0) {
      body += `PULL-IN OPPORTUNITIES (${pullInReady.length})\n`
      body += `${'─'.repeat(60)}\n`
      body += `The following sites are ready to accelerate if schedule allows:\n\n`
      pullInReady.forEach(h => {
        body += `• ${h.hop}  |  FC Start: ${h.ms15f}  |  ${h.daysOut}d out  |  NTP ✓  |  Mat ✓\n`
        if (sessionNotes[h.hop]) body += `  Note: ${sessionNotes[h.hop]}\n`
        body += '\n'
      })
    }

    // Action items — anything with blockers starting soon
    const actionItems = [...thisWeek, ...next2Weeks].filter(h => h.blockers.length > 0)
    if (actionItems.length > 0) {
      body += `ACTION ITEMS REQUIRED\n`
      body += `${'─'.repeat(60)}\n`
      actionItems.forEach((h, i) => {
        body += `${i + 1}. ${h.hop} (starts ${h.ms15f})\n`
        h.blockers.forEach(b => body += `   ${b}\n`)
        body += '\n'
      })
    }

    body += `${'─'.repeat(60)}\n`
    body += `Please coordinate with your Site CM ${cm} for all field questions.\n`
    body += `For schedule, finance, or contract matters contact CJ directly.\n\n`
    body += `Respectfully,\n`
    body += `CJ\n`
    body += `Nokia Program Manager — Viaero MW Construction Program\n`
    body += `CC: ${cm} — Site CM  |  Thomas M. — Lead CM`

    window.open(`mailto:?subject=${encodeURIComponent(subj)}&body=${encodeURIComponent(body)}`)
  }

  const EditableDate = ({ hop, field, value }: { hop: string, field: string, value: string }) => {
    const edited = editedDates[hop]?.[field]

    // Convert mm/dd/yyyy to yyyy-mm-dd for the date input
    const toInputFormat = (dateStr: string) => {
      if (!dateStr) return ''
      const parts = dateStr.split('/')
      if (parts.length !== 3) return ''
      return `${parts[2]}-${parts[0].padStart(2,'0')}-${parts[1].padStart(2,'0')}`
    }

    // Convert yyyy-mm-dd back to mm/dd/yyyy
    const toDisplayFormat = (dateStr: string) => {
      if (!dateStr) return ''
      const parts = dateStr.split('-')
      if (parts.length !== 3) return ''
      return `${parseInt(parts[1])}/${parseInt(parts[2])}/${parts[0]}`
    }

    const currentValue = toInputFormat(edited || value)

    return (
      <div className="flex flex-col gap-1">
        <input
          type="date"
          value={currentValue}
          onChange={(e) => {
            const newDisplay = toDisplayFormat(e.target.value)
            logDateEdit(hop, field, value, newDisplay)
          }}
          className={`text-xs rounded px-2 py-1 border focus:outline-none focus:border-blue-500 cursor-pointer ${edited ? 'bg-yellow-900 border-yellow-500 text-yellow-200' : 'bg-gray-800 border-gray-600 text-gray-300'}`}
        />
        {edited && (
          <span className="text-yellow-400 text-xs">📝 {edited}</span>
        )}
      </div>
    )
  }

  const PipelineTable = ({ title, rows }: { title: string, rows: HOP[] }) => (
    <div className="mb-6">
      <h3 className="text-base font-semibold text-white mb-3">{title} ({rows.length})</h3>
      {rows.length === 0
        ? <p className="text-gray-500 text-sm">No sites in this window</p>
        : (
          <div className="overflow-x-auto">
            <table className="w-full text-xs">
              <thead>
                <tr className="bg-gray-800 text-gray-400">
                  <th className="text-left p-2">HOP</th>
                  <th className="text-left p-2">FC Start</th>
                  <th className="text-left p-2">FC End</th>
                  <th className="text-left p-2">Days Out</th>
                  <th className="text-left p-2">SPO Issued</th>
                  <th className="text-left p-2">NTP</th>
                  <th className="text-left p-2">NTP Owner</th>
                  <th className="text-left p-2">NTP Waiting On</th>
                  <th className="text-left p-2">Mat</th>
                  <th className="text-left p-2">Mat Forecast</th>
                  <th className="text-left p-2">Mat Received</th>
                  <th className="text-left p-2">GC Pickup</th>
                  <th className="text-left p-2">Steel From</th>
                  <th className="text-left p-2">Pull-In Status</th>
                  <th className="text-left p-2">Vendor Window</th>
                  <th className="text-left p-2">Internal Conflict</th>
                  <th className="text-left p-2">Blockers</th>
                  <th className="text-left p-2">Edit MS15 Fc</th>
                  <th className="text-left p-2">Log MS15 Act</th>
                  <th className="text-left p-2">Call Notes (Today)</th>
                  <th className="text-left p-2">Notes History</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((h) => {
                  const hasConflict = h.vendorWindow.includes('🔴')
                  const isUrgent = h.blockers.length > 0 && (h.daysOut ?? 99) <= 7
                  const rowBg = hasConflict ? 'bg-red-950' : isUrgent ? 'bg-yellow-950' : h.blockers.length === 0 ? 'bg-green-950' : 'bg-gray-900'
                  return (
                    <tr key={h.hop} className={`border-t border-gray-800 ${rowBg}`}>
                      <td className="p-2 font-semibold text-white whitespace-nowrap">{h.hop}</td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{h.ms15f}</td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{h.ms16f}</td>
                      <td className={`p-2 font-bold whitespace-nowrap ${(h.daysOut ?? 99) <= 7 ? 'text-red-400' : (h.daysOut ?? 99) <= 14 ? 'text-yellow-400' : 'text-gray-300'}`}>{h.daysOut}d</td>
                      <td className="p-2">
                        {h.hasSpo
                          ? <span className="text-green-400 font-bold text-sm" title={h.spoIssued}>✓</span>
                          : <span className="text-red-400 font-bold text-sm">✗</span>
                        }
                      </td>
                      <td className="p-2">{h.hasNtp ? <span className="text-green-400 font-bold text-sm">✓</span> : <span className="text-red-400 font-bold text-sm">✗</span>}</td>
                      <td className="p-2 text-gray-300 text-xs">
                        <div className="cursor-help" title={h.ntpOwner}>
                          {h.ntpOwner ? (h.ntpOwner.length > 12 ? h.ntpOwner.slice(0, 12) + '...' : h.ntpOwner) : '—'}
                        </div>
                      </td>
                      <td className="p-2 text-gray-300 text-xs">
                        <div className="cursor-help" title={h.ntpWaitingOn}>
                          {h.ntpWaitingOn ? (h.ntpWaitingOn.length > 12 ? h.ntpWaitingOn.slice(0, 12) + '...' : h.ntpWaitingOn) : '—'}
                        </div>
                      </td>
                      <td className="p-2">{h.hasMat ? <span className="text-green-400 font-bold text-sm">✓</span> : <span className="text-red-400 font-bold text-sm">✗</span>}</td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{h.matForecast || '—'}</td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{h.matReceived || '—'}</td>
                      <td className="p-2">
                        {h.gcPickupDate
                          ? <span className="text-green-400 text-xs">✓ {h.gcPickupDate}</span>
                          : <EditableDate hop={h.hop} field="GC Material Pick-up (A)" value="" />
                        }
                      </td>
                      <td className="p-2 text-gray-300 text-xs whitespace-nowrap">
                        {h.steelFrom || '—'}
                      </td>
                      <td className="p-2 text-xs whitespace-nowrap">
                        <span className={h.pullInStatus.includes('✅') ? 'text-green-400' : h.pullInStatus.includes('⚠️') ? 'text-yellow-400' : h.pullInStatus.includes('🔴') ? 'text-red-400' : h.pullInStatus.includes('🟡') ? 'text-yellow-300' : 'text-gray-500'}>
                          {h.pullInStatus || '—'}
                        </span>
                      </td>
                      <td className="p-2 text-xs max-w-36 truncate" title={h.vendorWindow}>
                        <span className={h.vendorWindow.includes('🔴') ? 'text-red-400' : h.vendorWindow.includes('⚠️') ? 'text-yellow-400' : 'text-green-400'}>
                          {h.vendorWindow || '—'}
                        </span>
                      </td>
                      <td className="p-2 text-xs">
                        {h.internalConflict
                          ? <span className={h.internalConflict.includes('🔴') ? 'text-red-400' : 'text-yellow-400'} title={h.internalConflict}>
                              {h.internalConflict.length > 30 ? h.internalConflict.slice(0, 30) + '...' : h.internalConflict}
                            </span>
                          : <span className="text-green-400">✅ Clear</span>
                        }
                      </td>
                      <td className="p-2 text-xs">
                        {h.blockers.length === 0
                          ? <span className="text-green-400">✅ Clear</span>
                          : <div className="flex flex-col gap-1">
                              {h.blockers.map((b, i) => (
                                <span key={i} className={`whitespace-nowrap ${b.includes('🔴') ? 'text-red-400' : b.includes('🟠') ? 'text-orange-400' : 'text-yellow-400'}`}>
                                  {b}
                                </span>
                              ))}
                            </div>
                        }
                      </td>
                      <td className="p-2">
                        <EditableDate hop={h.hop} field="MS15 Fc Start" value={h.ms15f} />
                      </td>
                      <td className="p-2">
                        <EditableDate hop={h.hop} field="MS15 Implementation Start A" value={h.ms15a} />
                      </td>
                      <td className="p-2">
                        <div className="flex gap-1">
                          <input type="text" placeholder="Note..." value={sessionNotes[h.hop] || ''}
                            onChange={(e) => setSessionNotes(s => ({ ...s, [h.hop]: e.target.value }))}
                            onKeyDown={(e) => { if (e.key === 'Enter') saveCallNote(h.hop) }}
                            className="w-36 bg-gray-800 text-white text-xs rounded px-2 py-1 border border-gray-600 focus:outline-none focus:border-blue-500" />
                          <button onClick={() => saveCallNote(h.hop)} className="text-xs bg-blue-700 hover:bg-blue-600 text-white px-2 py-1 rounded">💾</button>
                        </div>
                      </td>
                      <td className="p-2 max-w-48">
                        <div className="max-h-20 overflow-y-auto flex flex-col gap-1">
                          {(noteHistory[h.hop] || []).slice(0, 5).map((n, i) => (
                            <div key={i} className="text-xs text-gray-300 border-b border-gray-700 pb-1">
                              <span className="text-gray-500 text-xs">{new Date(n.logged_at).toLocaleDateString()}</span>
                              <span className="ml-1">{n.note}</span>
                            </div>
                          ))}
                          {!noteHistory[h.hop]?.length && <span className="text-gray-600 text-xs">No history</span>}
                        </div>
                      </td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>
        )}
    </div>
  )

  return (
    <div className="min-h-screen bg-gray-950 text-white p-6">
      <div className="max-w-full mx-auto">

        {/* Header */}
        <div className="flex items-center justify-between mb-6">
          <div>
            <h1 className="text-3xl font-bold">GC Call View</h1>
            <p className="text-gray-400 mt-1">Select a contractor to view their full pipeline, blockers, and generate a follow-up email.</p>
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
            <div className="flex items-center justify-between mb-3">
              <h2 className="text-yellow-300 font-bold text-lg">📋 PM Daily Updates — Update These in Your Tracker</h2>
              <button
                onClick={() => setPmSortAsc(prev => !prev)}
                className="bg-yellow-800 hover:bg-yellow-700 text-yellow-200 text-xs px-3 py-1 rounded font-semibold">
                Sort by Field {pmSortAsc ? '↑ A→Z' : '↓ Z→A'}
              </button>
            </div>
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
                {[...pmUpdates].sort((a, b) => pmSortAsc ? a.field.localeCompare(b.field) : b.field.localeCompare(a.field)).map((u, i) => (
                  <tr key={i} className={`border-t border-yellow-800 ${u.completed ? 'opacity-40' : ''}`}>
                    <td className="p-2">
                      <input
                        type="checkbox"
                        checked={u.completed || false}
                        onChange={() => {
                          setPmUpdates(prev => prev.map((item, idx) =>
                            idx === i ? { ...item, completed: !item.completed } : item
                          ))
                        }}
                        className="w-4 h-4 cursor-pointer accent-green-500"
                      />
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
              <button onClick={() => {
                const remaining = pmUpdates.filter(u => !u.completed)
                setPmUpdates(remaining)
                if (remaining.length === 0) setShowPmUpdates(false)
              }}
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

        {/* GC Selector */}
        <div className="flex gap-3 mb-6 flex-wrap">
          {gcList.map((gc) => (
            <button key={gc} onClick={() => setSelectedGC(gc)}
              className={`px-6 py-3 rounded-lg font-semibold text-sm transition-all ${selectedGC === gc ? 'bg-blue-600 text-white shadow-lg scale-105' : 'bg-gray-800 text-gray-300 hover:bg-gray-700'}`}>
              {gc}
            </button>
          ))}
        </div>

        {/* GC Panel */}
        {selectedGC && (
          <div className="bg-gray-900 rounded-xl border border-gray-700 p-6">
            <div className="flex items-center justify-between mb-6">
              <div>
                <h2 className="text-2xl font-bold">{selectedGC}</h2>
                <p className="text-gray-400 mt-1">
                  Site CM: <span className="text-blue-400 font-semibold">{GC_CM_MAP[selectedGC] || 'See Contacts'}</span>
                  {loaded && (
                    <span className="text-gray-500">
                      {' · '}{gcHops.length} HOPs{' · '}
                      {active.length} active{' · '}
                      {thisWeek.length} this week{' · '}
                      {next2Weeks.length} next 2 wks{' · '}
                      {thisMonth.length} this month{' · '}
                      <span className="text-green-400">{pullInReady.length} pull-in ready</span>
                    </span>
                  )}
                </p>
              </div>
              <button onClick={generateEmail}
                className="bg-blue-600 hover:bg-blue-700 text-white px-5 py-2 rounded-lg text-sm font-semibold">
                ✉️ Generate Email
              </button>
            </div>

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
                              <th className="text-left p-2">Started</th>
                              <th className="text-left p-2">FC End</th>
                              <th className="text-left p-2">Days Elapsed</th>
                              <th className="text-left p-2">Status</th>
                              <th className="text-left p-2">SPO Issued</th>
                              <th className="text-left p-2">MS16 Fc</th>
                              <th className="text-left p-2">Edit MS16 Fc</th>
                              <th className="text-left p-2">MS16 Act</th>
                              <th className="text-left p-2">Call Notes (Today)</th>
                              <th className="text-left p-2">Notes History</th>
                            </tr>
                          </thead>
                          <tbody>
                            {active.map((h) => (
                              <tr key={h.hop} className={`border-t border-gray-800 ${(h.daysElapsed ?? 0) > 18 ? 'bg-red-950' : 'bg-gray-900'}`}>
                                <td className="p-2 font-semibold text-white whitespace-nowrap">{h.hop}</td>
                                <td className="p-2 text-gray-300 text-xs whitespace-nowrap">{h.ms15a || '—'}</td>
                                <td className="p-2 text-gray-300 text-xs whitespace-nowrap">{h.ms16f || '—'}</td>
                                <td className={`p-2 font-bold ${(h.daysElapsed ?? 0) > 18 ? 'text-red-400' : 'text-green-400'}`}>{h.daysElapsed}d</td>
                                <td className="p-2">{(h.daysElapsed ?? 0) > 18 ? <span className="text-red-400">⚠️ Over 18d</span> : <span className="text-green-400">On track</span>}</td>
                                <td className="p-2">
                                  {h.hasSpo
                                    ? <span className="text-green-400 font-bold text-sm" title={h.spoIssued}>✓</span>
                                    : <span className="text-red-400 font-bold text-sm">✗</span>
                                  }
                                </td>
                                <td className="p-2 text-gray-300 text-xs whitespace-nowrap">{h.ms16f || '—'}</td>
                                <td className="p-2">
                                  <EditableDate hop={h.hop} field="MS16 Implementation Ends F" value={h.ms16f} />
                                </td>
                                <td className="p-2">
                                  <EditableDate hop={h.hop} field="MS16 Implementation Ends A" value={h.ms16a} />
                                </td>
                                <td className="p-2">
                                  <div className="flex gap-1">
                                    <input type="text" placeholder="Note..." value={sessionNotes[h.hop] || ''}
                                      onChange={(e) => setSessionNotes(s => ({ ...s, [h.hop]: e.target.value }))}
                                      onKeyDown={(e) => { if (e.key === 'Enter') saveCallNote(h.hop) }}
                                      className="w-36 bg-gray-800 text-white text-xs rounded px-2 py-1 border border-gray-600 focus:outline-none focus:border-blue-500" />
                                    <button onClick={() => saveCallNote(h.hop)} className="text-xs bg-blue-700 hover:bg-blue-600 text-white px-2 py-1 rounded">💾</button>
                                  </div>
                                </td>
                                <td className="p-2 max-w-48">
                                  <div className="max-h-20 overflow-y-auto flex flex-col gap-1">
                                    {(noteHistory[h.hop] || []).slice(0, 5).map((n, i) => (
                                      <div key={i} className="text-xs text-gray-300 border-b border-gray-700 pb-1">
                                        <span className="text-gray-500 text-xs">{new Date(n.logged_at).toLocaleDateString()}</span>
                                        <span className="ml-1">{n.note}</span>
                                      </div>
                                    ))}
                                    {!noteHistory[h.hop]?.length && <span className="text-gray-600 text-xs">No history</span>}
                                  </div>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    )}
                </div>

                {/* Pipeline Sections */}
                <PipelineTable title="⚡ This Week (0–7 days)" rows={thisWeek} />
                <PipelineTable title="🟠 Next 2 Weeks (8–14 days)" rows={next2Weeks} />
                <PipelineTable title="🟡 This Month (15–30 days)" rows={thisMonth} />

                {/* Pull-In Opportunities */}
                <div>
                  <div className="flex items-center gap-3 mb-3">
                    <h3 className="text-lg font-semibold text-white">🚀 Full Pipeline + Pull-In Opportunities (30d+)</h3>
                    {pullInReady.length > 0 && (
                      <span className="bg-green-800 text-green-200 text-xs px-3 py-1 rounded-full font-semibold">
                        {pullInReady.length} ready to pull in
                      </span>
                    )}
                  </div>
                  <PipelineTable title="" rows={pullIns} />
                </div>

              </div>
          </div>
        )}

        {!selectedGC && (
          <div className="bg-gray-900 rounded-xl border border-gray-700 p-12 text-center">
            <p className="text-gray-400 text-xl">👆 Select a GC above to begin</p>
          </div>
        )}

      </div>
    </div>
  )
}
