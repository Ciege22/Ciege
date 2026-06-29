'use client'

export const dynamic = 'force-dynamic'

import { useState, useCallback, useEffect } from 'react'
import * as XLSX from 'xlsx'
import { supabase, loadTrackerSnapshot } from '../lib/supabase'

interface HOP {
  hop: string
  gc: string
  cm: string
  nokiaPm: string
  ms15f: string
  ms15fDate: Date | null
  ms16f: string
  hasNtp: boolean
  ntpOwner: string
  ntpWaitingOn: string
  hasMat: boolean
  matForecast: string
  vendorWindow: string
  month: string
  daysOut: number | null
  ownerCategory: string
  waitingOnBucket: string
}

interface CallNote {
  id: string
  hop_name: string
  note: string
  logged_at: string
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

function parseNtpDate(val: unknown): Date | null {
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

function getOwnerCategory(owner: string): string {
  const o = owner.trim()
  if (!o) return 'No Owner'
  if (o === 'On Hold') return 'Viaero'
  if (o === 'Viaero') return 'Viaero'
  if (o === 'Viaero & Nokia') return 'Viaero & Nokia'
  if (o === 'ITW') return 'ITW'
  if (o === 'Samsung/ITW') return 'Samsung / ITW'
  if (o === 'Nokia/ITW') return 'Nokia / ITW'
  if (o === 'Nokia') return 'Nokia'
  if (o === 'Commsearch') return 'Commsearch'
  return 'Other'
}

function getWaitingOnBucket(waitingOn: string, ntpOwner: string): string {
  const w = waitingOn.toLowerCase()
  const o = ntpOwner.toLowerCase()

  if (!w || w === 'blank') return '❓ No Info — Review'

  // On Hold always goes to Viaero
  if (w.includes('on hold')) return '🔴 On Hold per Viaero'

  // Check for BOTH submittal and approval in same entry → Viaero & Nokia
  const hasSubmittal = w.includes('submittal')
  const hasApproval  = w.includes('approval')
  if (hasSubmittal && hasApproval) return '🤝 Pending Submittal & Approval — Viaero & Nokia'

  // Approval only → Viaero
  if (hasApproval && !hasSubmittal) {
    if (w.includes('service quote') || w.includes('wp') || w.includes('lld') || w.includes('work package')) {
      return '✅ Pending Approval — Viaero Action'
    }
  }

  // Submittal only → Nokia
  if (hasSubmittal && !hasApproval) {
    if (w.includes('service quote') || w.includes('wp') || w.includes('lld') || w.includes('work package')) {
      return '📄 Pending Submittal — Nokia Action'
    }
  }

  // Cage match → Nokia
  if (w.includes('cage match')) return '📋 Pending Cage Match — Viaero & Nokia'

  // PCN → Nokia
  if (w.includes('pcn')) return '📡 Pending PCN — Nokia'

  // Samsung CDs → ITW
  if (w.includes('samsung') || w.includes('combo cds') || w.includes('revised cds')) return '📡 Pending Samsung CDs — ITW'

  // Landlord → ITW
  if (w.includes(' ll ') || w.includes('ll ') || w.includes('landlord') || w.includes('lease') || w.startsWith('ll ') || w.includes('pending ll')) return '🏢 Pending Landlord (LL) — ITW'

  // SA / MA → ITW
  if (w.includes('structural') || w.includes('mount analysis') || w.includes('failing sa') || w.includes(' sa ') || w.includes('sa ') || w.startsWith('sa ') || w.includes(',sa') || w.includes('ma ') || w.includes(' ma,')) return '🏗️ Pending SA / MA — ITW'

  // ITW specific
  if (w.includes('itw') || w.includes('lisco tg')) return '🔧 Pending ITW Approval'

  // Commsearch
  if (w.includes('commsearch')) return '🔍 Pending Commsearch'

  // Scoping
  if (w.includes('scoping')) return '🔭 Pending Scoping — Nokia'

  // New tower
  if (w.includes('new tower')) return '🗼 Pending New Tower — ITW'

  // Catch all
  return '❓ Other — Review'
}

const OWNER_CATEGORIES = ['All', 'Nokia', 'Viaero & Nokia', 'Viaero', 'On Hold', 'ITW', 'Samsung / ITW', 'Nokia / ITW', 'Commsearch', 'No Owner', 'Other']

const OWNER_COLORS: Record<string, string> = {
  'Nokia': 'bg-blue-900 border-blue-600',
  'Viaero & Nokia': 'bg-purple-900 border-purple-600',
  'Viaero': 'bg-amber-900 border-amber-600',
  'On Hold': 'bg-red-900 border-red-600',
  'ITW': 'bg-orange-900 border-orange-600',
  'Samsung / ITW': 'bg-orange-900 border-orange-600',
  'Nokia / ITW': 'bg-indigo-900 border-indigo-600',
  'Commsearch': 'bg-teal-900 border-teal-600',
  'No Owner': 'bg-gray-800 border-gray-600',
  'Other': 'bg-gray-800 border-gray-600',
}

const OWNER_TEXT: Record<string, string> = {
  'Nokia': 'text-blue-300',
  'Viaero & Nokia': 'text-purple-300',
  'Viaero': 'text-amber-300',
  'On Hold': 'text-red-300',
  'ITW': 'text-orange-300',
  'Samsung / ITW': 'text-orange-300',
  'Nokia / ITW': 'text-indigo-300',
  'Commsearch': 'text-teal-300',
  'No Owner': 'text-gray-400',
  'Other': 'text-gray-400',
}

export default function NTPTrackerPage() {
  const [hops, setHops] = useState<HOP[]>([])
  const [loaded, setLoaded] = useState(false)
  const [fileName, setFileName] = useState('')
  const [selectedOwner, setSelectedOwner] = useState('All')
  const [selectedMonth, setSelectedMonth] = useState('All')
  const [noteHistory, setNoteHistory] = useState<Record<string, CallNote[]>>({})
  const [sessionNotes, setSessionNotes] = useState<Record<string, string>>({})
  const [expandedBuckets, setExpandedBuckets] = useState<Set<string>>(new Set())
  const [snapshotTime, setSnapshotTime] = useState<string>('')
  const today = new Date()

  const loadNotes = async () => {
    const { data } = await supabase.from('hop_call_notes').select('*').order('logged_at', { ascending: false })
    if (data) {
      const map: Record<string, CallNote[]> = {}
      data.forEach((row: CallNote) => {
        if (!map[row.hop_name]) map[row.hop_name] = []
        map[row.hop_name].push(row)
      })
      setNoteHistory(map)
    }
  }

  const saveNote = async (hop: string, note: string) => {
    if (!note.trim()) return
    const { data } = await supabase.from('hop_call_notes').insert({ hop_name: hop, note: note.trim() }).select()
    if (data) {
      setNoteHistory(prev => ({ ...prev, [hop]: [data[0], ...(prev[hop] || [])] }))
      setSessionNotes(prev => ({ ...prev, [hop]: '' }))
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

  const processRows = useCallback(async (rows: unknown[][], _filename: string) => {

    let headerRow = -1
    for (let i = 0; i < 10; i++) {
      if ((rows[i] as unknown[])?.some(c => String(c).trim() === 'HOP')) { headerRow = i; break }
    }
    if (headerRow === -1) { alert('Could not find header row'); return }

    const headers = rows[headerRow] as string[]
    const col = (name: string) => headers.findIndex(h => String(h).trim() === name)

    const hopCol      = col('HOP')
    const gcCol       = col('General Contractor')
    const newCmCol    = col('New CM')
    const nokiaPmCol  = col('Nokia PM')
    const don444Col   = col('DON 444')
    const ms15fCol    = col('MS15 Implementation Start F')
    const ms16fCol    = col('MS16 Implementation Ends F')
    const ms16aCol    = col('MS16 Implementation Ends A')
    const ntpCol      = col('NTP A')
    const ntpOwnCol   = col('NTP Action Owner')
    const ntpWaitCol  = col('NTP is waiting on')
    const matCol      = headers.findIndex(h => String(h).trim() === 'Material Received A ')
    const matFcCol    = col('Material Forecast +4ish')
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
      const row   = rows2[0]
      const ntpDate = parseNtpDate(row[ntpCol])
      const ms16a   = parseNtpDate(row[ms16aCol])
      if (ms16a) return // skip complete HOPs
      const hasNtp  = !!ntpDate
      if (hasNtp) return // skip NTP complete

      const ms15f    = parseDateAny(row[ms15fCol])
      const matDate  = parseDateAny(row[matCol])
      const ntpOwner = String(row[ntpOwnCol] || '').trim()
      const ntpWait  = String(row[ntpWaitCol] || '').trim()

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

      const monthLabel = ms15f
        ? ms15f.toLocaleDateString('en-US', { month: 'long', year: 'numeric' })
        : 'No Date'

      parsed.push({
        hop,
        gc:            String(row[gcCol] || '').trim(),
        cm:            String(row[newCmCol] || '').trim(),
        nokiaPm:       String(row[nokiaPmCol] || '').trim(),
        ms15f:         fmtDate(ms15f),
        ms15fDate:     ms15f,
        ms16f:         fmtDate(parseDateAny(row[ms16fCol])),
        hasNtp,
        ntpOwner,
        ntpWaitingOn:  ntpWait,
        hasMat:        !!(matDate && matDate.getFullYear() >= 2020),
        matForecast:   fmtDate(parseDateAny(row[matFcCol])),
        vendorWindow,
        month:         monthLabel,
        daysOut:       ms15f ? daysBetween(today, ms15f) : null,
        ownerCategory: getOwnerCategory(ntpOwner),
        waitingOnBucket: getWaitingOnBucket(ntpWait, ntpOwner)
      })
    })

    setHops(parsed)
    setLoaded(true)
    await loadNotes()
  }, [today])

  const handleFile = useCallback((file: File) => {
    setFileName(file.name)
    const reader = new FileReader()
    reader.onload = async (e) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer)
      const wb = XLSX.read(data, { type: 'array', cellDates: true })
      const ws = wb.Sheets['HOPs']
      if (!ws) { alert('HOPs tab not found'); return }
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 }) as unknown[][]
      await processRows(rows, file.name)
    }
    reader.readAsArrayBuffer(file)
  }, [processRows])

  // Filtered hops
  const filtered = hops.filter(h => {
    const ownerMatch = selectedOwner === 'All' || h.ownerCategory === selectedOwner
    const monthMatch = selectedMonth === 'All' || h.month === selectedMonth
    return ownerMatch && monthMatch
  })

  // Unique months
  const months = ['All', ...Array.from(new Set(hops.map(h => h.month).filter(m => m !== 'No Date'))).sort((a, b) => new Date(a) > new Date(b) ? 1 : -1)]

  // Pivot data
  const pivotData = OWNER_CATEGORIES.filter(o => o !== 'All').map(cat => {
    const catHops = hops.filter(h => h.ownerCategory === cat)
    return { cat, total: catHops.length }
  }).filter(d => d.total > 0)

  // Monthly breakdown
  const monthlyData = months.filter(m => m !== 'All').map(month => {
    const monthHops = hops.filter(h => h.month === month)
    return {
      month,
      total: monthHops.length,
      nokia: monthHops.filter(h => h.ownerCategory === 'Nokia').length,
      viaNok: monthHops.filter(h => h.ownerCategory === 'Viaero & Nokia').length,
      viaero: monthHops.filter(h => h.ownerCategory === 'Viaero').length,
      onHold: monthHops.filter(h => h.ownerCategory === 'On Hold').length,
      itw: monthHops.filter(h => ['ITW', 'Samsung / ITW', 'Nokia / ITW'].includes(h.ownerCategory)).length,
      other: monthHops.filter(h => ['Commsearch', 'No Owner', 'Other'].includes(h.ownerCategory)).length,
    }
  })

  // Group filtered by owner then by waiting on bucket
  const grouped = new Map<string, Map<string, HOP[]>>()
  filtered.forEach(h => {
    if (!grouped.has(h.ownerCategory)) grouped.set(h.ownerCategory, new Map())
    const ownerMap = grouped.get(h.ownerCategory)!
    if (!ownerMap.has(h.waitingOnBucket)) ownerMap.set(h.waitingOnBucket, [])
    ownerMap.get(h.waitingOnBucket)!.push(h)
  })

  const toggleBucket = (key: string) => {
    setExpandedBuckets(prev => {
      const next = new Set(prev)
      if (next.has(key)) next.delete(key)
      else next.add(key)
      return next
    })
  }

  const generateEmail = (ownerCat: string) => {
    const catHops = filtered.filter(h => h.ownerCategory === ownerCat)
      .sort((a, b) => (a.daysOut ?? 999) - (b.daysOut ?? 999))
    if (catHops.length === 0) return

    const date = today.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })
    const div  = '─'.repeat(60)

    const isViaero = ownerCat.includes('Viaero') || ownerCat === 'On Hold'
    const isNokia  = ownerCat.includes('Nokia')
    const isITW    = ownerCat.includes('ITW') || ownerCat.includes('Samsung')

    const buildEmailBody = (party: 'Viaero' | 'Nokia' | 'ITW') => {
      const to = ''
      const greeting = 'Team,'

      const relevantHops = party === 'Nokia'
        ? catHops.filter(h => h.ownerCategory.includes('Nokia'))
        : party === 'Viaero'
        ? catHops.filter(h => h.ownerCategory.includes('Viaero') || h.ownerCategory === 'On Hold')
        : catHops

      const subj = `Viaero MW Program — NTP Action Required | ${party} | ${date}`
      let body = `${greeting}\n\n`
      body += `The following HOPs require ${party} action to issue NTP. These sites are scheduled for construction and NTP is on the critical path.\n`
      body += `${div}\n\n`

      // Group by month
      const byMonth = new Map<string, HOP[]>()
      relevantHops.forEach(h => {
        if (!byMonth.has(h.month)) byMonth.set(h.month, [])
        byMonth.get(h.month)!.push(h)
      })

      byMonth.forEach((monthHops, month) => {
        body += `${month.toUpperCase()} (${monthHops.length} HOPs)\n${div}\n`
        monthHops.forEach(h => {
          body += `• ${h.hop}  |  GC: ${h.gc || '—'}  |  FC Start: ${h.ms15f || '—'}\n`
          body += `  Waiting On: ${h.ntpWaitingOn || '—'}\n`
          if (sessionNotes[h.hop]) body += `  Note: ${sessionNotes[h.hop]}\n`
          body += '\n'
        })
      })

      body += `${div}\n`
      body += `Please advise on status for each item. Nokia is ready to mobilize once NTP is issued.\n\n`
      body += `Respectfully,\nCJ\nNokia Program Manager — Viaero MW Construction Program`

      window.open(`mailto:${to}?subject=${encodeURIComponent(subj)}&body=${encodeURIComponent(body)}`)
    }

    if (ownerCat === 'Viaero & Nokia') {
      buildEmailBody('Viaero')
      setTimeout(() => buildEmailBody('Nokia'), 500)
    } else if (isViaero) buildEmailBody('Viaero')
    else if (isNokia)   buildEmailBody('Nokia')
    else if (isITW)     buildEmailBody('ITW')
  }

  const downloadExcel = () => {
    const wb2 = XLSX.utils.book_new()

    // Summary sheet
    const summaryData = [
      ['NTP Tracker Summary', '', '', '', '', '', '', ''],
      [`Generated: ${today.toLocaleDateString()}`, '', '', '', '', '', '', ''],
      ['', '', '', '', '', '', '', ''],
      ['Month', 'Total Pending', 'Nokia', 'Viaero & Nokia', 'Viaero', 'On Hold', 'ITW / Samsung', 'Other'],
      ...monthlyData.map(m => [m.month, m.total, m.nokia, m.viaNok, m.viaero, m.onHold, m.itw, m.other])
    ]
    XLSX.utils.book_append_sheet(wb2, XLSX.utils.aoa_to_sheet(summaryData), 'Summary')

    // Per owner sheets
    const ownerGroups: Record<string, HOP[]> = {}
    hops.forEach(h => {
      if (!ownerGroups[h.ownerCategory]) ownerGroups[h.ownerCategory] = []
      ownerGroups[h.ownerCategory].push(h)
    })

    Object.entries(ownerGroups).forEach(([owner, ownerHops]) => {
      const sheetName = owner.slice(0, 31)
      const sheetData = [
        ['HOP', 'GC', 'CM', 'Nokia PM', 'FC Start', 'FC End', 'NTP Owner', 'NTP Waiting On', 'Material', 'Mat Forecast', 'Vendor Window', 'Month'],
        ...ownerHops.sort((a, b) => (a.daysOut ?? 999) - (b.daysOut ?? 999)).map(h => [
          h.hop, h.gc, h.cm, h.nokiaPm, h.ms15f, h.ms16f,
          h.ntpOwner, h.ntpWaitingOn,
          h.hasMat ? 'Received' : 'Pending',
          h.matForecast, h.vendorWindow, h.month
        ])
      ]
      XLSX.utils.book_append_sheet(wb2, XLSX.utils.aoa_to_sheet(sheetData), sheetName)
    })

    XLSX.writeFile(wb2, `NTP_Tracker_${today.toLocaleDateString('en-US').replace(/\//g, '-')}.xlsx`)
  }

  return (
    <div className="min-h-screen bg-gray-950 text-white p-6">
      <div className="max-w-full mx-auto">

        {/* Header */}
        <div className="flex items-center justify-between mb-6">
          <div>
            <h1 className="text-3xl font-bold">NTP Tracker</h1>
            <p className="text-gray-400 mt-1">Full program view — DON 444 filter — all pending NTPs by owner and month</p>
          </div>
          {loaded && (
            <button onClick={downloadExcel}
              className="bg-green-700 hover:bg-green-600 text-white px-4 py-2 rounded-lg text-sm font-semibold">
              📥 Download Excel
            </button>
          )}
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
            {/* Pivot Cards */}
            <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-5 gap-3 mb-6">
              <div className="bg-gray-800 rounded-xl border border-gray-600 p-4 text-center cursor-pointer hover:border-blue-500"
                onClick={() => setSelectedOwner('All')}>
                <p className="text-gray-400 text-xs font-semibold">TOTAL PENDING</p>
                <p className="text-white text-3xl font-bold mt-1">{hops.length}</p>
                <p className="text-gray-500 text-xs mt-1">of {hops.length + 144} program HOPs</p>
              </div>
              {pivotData.map(({ cat, total }) => (
                <div key={cat}
                  onClick={() => setSelectedOwner(selectedOwner === cat ? 'All' : cat)}
                  className={`rounded-xl border p-4 text-center cursor-pointer transition-all hover:scale-105 ${selectedOwner === cat ? 'ring-2 ring-white' : ''} ${OWNER_COLORS[cat] || 'bg-gray-800 border-gray-600'}`}>
                  <p className={`text-xs font-semibold ${OWNER_TEXT[cat] || 'text-gray-300'}`}>{cat.toUpperCase()}</p>
                  <p className="text-white text-3xl font-bold mt-1">{total}</p>
                  <p className="text-gray-400 text-xs mt-1">{Math.round(total / hops.length * 100)}% of pending</p>
                </div>
              ))}
            </div>

            {/* Monthly Breakdown */}
            <div className="bg-gray-900 rounded-xl border border-gray-700 p-5 mb-6 overflow-x-auto">
              <h2 className="text-base font-bold mb-3">📅 Pending NTPs by FC Start Month</h2>
              <table className="w-full text-xs">
                <thead>
                  <tr className="bg-gray-800 text-gray-400">
                    <th className="text-left p-2">Month</th>
                    <th className="text-center p-2">Total</th>
                    <th className="text-center p-2 text-blue-300">Nokia</th>
                    <th className="text-center p-2 text-purple-300">Viaero & Nokia</th>
                    <th className="text-center p-2 text-amber-300">Viaero</th>
                    <th className="text-center p-2 text-red-300">On Hold</th>
                    <th className="text-center p-2 text-orange-300">ITW / Samsung</th>
                    <th className="text-center p-2 text-gray-400">Other</th>
                  </tr>
                </thead>
                <tbody>
                  {monthlyData.map((m, i) => (
                    <tr key={i}
                      onClick={() => setSelectedMonth(selectedMonth === m.month ? 'All' : m.month)}
                      className={`border-t border-gray-800 cursor-pointer hover:bg-gray-800 transition-colors ${selectedMonth === m.month ? 'bg-gray-800 ring-1 ring-blue-500' : ''}`}>
                      <td className="p-2 font-semibold text-white">{m.month}</td>
                      <td className="p-2 text-center font-bold text-white">{m.total}</td>
                      <td className="p-2 text-center text-blue-300">{m.nokia || '—'}</td>
                      <td className="p-2 text-center text-purple-300">{m.viaNok || '—'}</td>
                      <td className="p-2 text-center text-amber-300">{m.viaero || '—'}</td>
                      <td className="p-2 text-center text-red-300">{m.onHold || '—'}</td>
                      <td className="p-2 text-center text-orange-300">{m.itw || '—'}</td>
                      <td className="p-2 text-center text-gray-400">{m.other || '—'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {selectedMonth !== 'All' && (
                <button onClick={() => setSelectedMonth('All')} className="mt-2 text-xs text-gray-500 hover:text-gray-300">
                  ✕ Clear month filter
                </button>
              )}
            </div>

            {/* Owner Filter Pills */}
            <div className="flex gap-2 mb-4 flex-wrap items-center">
              <span className="text-gray-500 text-xs">Filter:</span>
              {OWNER_CATEGORIES.map(cat => (
                <button key={cat} onClick={() => setSelectedOwner(cat)}
                  className={`px-3 py-1 rounded-full text-xs font-semibold transition-all ${selectedOwner === cat ? 'bg-blue-600 text-white' : 'bg-gray-800 text-gray-400 hover:bg-gray-700'}`}>
                  {cat} {cat !== 'All' ? `(${hops.filter(h => h.ownerCategory === cat).length})` : `(${hops.length})`}
                </button>
              ))}
            </div>

            {/* Email Buttons */}
            <div className="flex gap-3 mb-6 flex-wrap">
              {Array.from(grouped.keys()).map(owner => (
                <button key={owner} onClick={() => generateEmail(owner)}
                  className={`px-4 py-2 rounded-lg text-xs font-semibold transition-all hover:scale-105 ${OWNER_COLORS[owner] || 'bg-gray-800 border border-gray-600'} ${OWNER_TEXT[owner] || 'text-gray-300'}`}>
                  ✉️ {owner === 'Viaero & Nokia' ? 'Draft 2 Emails' : 'Draft Email'} — {owner}
                </button>
              ))}
            </div>

            {/* Detail Table — grouped by owner → bucket */}
            <div className="space-y-4">
              {Array.from(grouped.entries()).map(([owner, bucketMap]) => (
                <div key={owner} className={`rounded-xl border p-4 ${OWNER_COLORS[owner] || 'bg-gray-900 border-gray-700'}`}>
                  <div className="flex items-center justify-between mb-3">
                    <h3 className={`text-base font-bold ${OWNER_TEXT[owner] || 'text-white'}`}>
                      {owner} — {Array.from(bucketMap.values()).flat().length} HOPs
                    </h3>
                    <button onClick={() => generateEmail(owner)}
                      className="text-xs bg-gray-800 hover:bg-gray-700 text-white px-3 py-1 rounded font-semibold">
                      ✉️ {owner === 'Viaero & Nokia' ? 'Draft 2 Emails' : 'Draft Email'}
                    </button>
                  </div>

                  {Array.from(bucketMap.entries()).sort((a, b) => b[1].length - a[1].length).map(([bucket, bucketHops]) => {
                    const key = `${owner}|${bucket}`
                    const isExpanded = expandedBuckets.has(key)
                    const urgent = bucketHops.filter(h => h.daysOut !== null && h.daysOut <= 30).length

                    return (
                      <div key={bucket} className="mb-2 bg-gray-950 rounded-lg overflow-hidden">
                        <div
                          className="flex items-center justify-between p-3 cursor-pointer hover:bg-gray-900"
                          onClick={() => toggleBucket(key)}>
                          <div className="flex items-center gap-3">
                            <span className="text-white text-sm font-semibold">{bucket}</span>
                            <span className="bg-gray-800 text-gray-300 text-xs px-2 py-0.5 rounded-full">{bucketHops.length} HOPs</span>
                            {urgent > 0 && <span className="bg-red-800 text-red-200 text-xs px-2 py-0.5 rounded-full">🔴 {urgent} starting ≤30d</span>}
                          </div>
                          <span className="text-gray-500 text-xs">{isExpanded ? '▲ Collapse' : '▼ Expand'}</span>
                        </div>

                        {isExpanded && (
                          <div className="overflow-x-auto border-t border-gray-800">
                            <table className="w-full text-xs">
                              <thead>
                                <tr className="bg-gray-900 text-gray-400">
                                  <th className="text-left p-2">HOP</th>
                                  <th className="text-left p-2">GC</th>
                                  <th className="text-left p-2">CM</th>
                                  <th className="text-left p-2">Nokia PM</th>
                                  <th className="text-left p-2">Days Out</th>
                                  <th className="text-left p-2">FC Start</th>
                                  <th className="text-left p-2">FC End</th>
                                  <th className="text-left p-2">Month</th>
                                  <th className="text-left p-2">NTP Owner</th>
                                  <th className="text-left p-2">NTP Waiting On</th>
                                  <th className="text-left p-2">Mat</th>
                                  <th className="text-left p-2">Mat Forecast</th>
                                  <th className="text-left p-2">Vendor</th>
                                  <th className="text-left p-2">Notes</th>
                                  <th className="text-left p-2">History</th>
                                </tr>
                              </thead>
                              <tbody>
                                {bucketHops.sort((a, b) => (a.daysOut ?? 999) - (b.daysOut ?? 999)).map(h => {
                                  const rowBg = h.daysOut !== null && h.daysOut <= 14 ? 'bg-red-950' :
                                               h.daysOut !== null && h.daysOut <= 30 ? 'bg-yellow-950' : 'bg-gray-900'
                                  return (
                                    <tr key={h.hop} className={`border-t border-gray-800 ${rowBg}`}>
                                      <td className="p-2 font-semibold text-white whitespace-nowrap">{h.hop}</td>
                                      <td className="p-2 text-gray-300 whitespace-nowrap">{h.gc || '—'}</td>
                                      <td className="p-2 text-gray-300 whitespace-nowrap">{h.cm || '—'}</td>
                                      <td className="p-2 text-gray-300 whitespace-nowrap">{h.nokiaPm || '—'}</td>
                                      <td className={`p-2 font-bold whitespace-nowrap ${h.daysOut !== null && h.daysOut <= 14 ? 'text-red-400' : h.daysOut !== null && h.daysOut <= 30 ? 'text-yellow-400' : 'text-gray-300'}`}>
                                        {h.daysOut !== null ? `${h.daysOut}d` : '—'}
                                      </td>
                                      <td className="p-2 text-gray-300 whitespace-nowrap">{h.ms15f || '—'}</td>
                                      <td className="p-2 text-gray-300 whitespace-nowrap">{h.ms16f || '—'}</td>
                                      <td className="p-2 text-gray-300 whitespace-nowrap">{h.month}</td>
                                      <td className="p-2 text-gray-300 whitespace-nowrap">{h.ntpOwner || '—'}</td>
                                      <td className="p-2 text-gray-300 text-xs max-w-48" title={h.ntpWaitingOn}>
                                        {h.ntpWaitingOn.length > 40 ? h.ntpWaitingOn.slice(0, 40) + '...' : h.ntpWaitingOn || '—'}
                                      </td>
                                      <td className="p-2">
                                        {h.hasMat
                                          ? <span className="text-green-400 font-bold">✓</span>
                                          : <span className="text-red-400 font-bold" title={`Forecast: ${h.matForecast || 'No forecast'}`}>✗</span>}
                                      </td>
                                      <td className="p-2 text-gray-300 whitespace-nowrap">{h.matForecast || '—'}</td>
                                      <td className="p-2 text-xs">
                                        <span className={h.vendorWindow.includes('🔴') ? 'text-red-400' : 'text-green-400'}
                                          title={h.vendorWindow}>
                                          {h.vendorWindow.includes('🔴') ? '🔴 Conflict' : '✅ Clear'}
                                        </span>
                                      </td>
                                      <td className="p-2">
                                        <div className="flex gap-1">
                                          <input type="text" placeholder="Note..."
                                            value={sessionNotes[h.hop] || ''}
                                            onChange={(e) => setSessionNotes(n => ({ ...n, [h.hop]: e.target.value }))}
                                            onKeyDown={(e) => { if (e.key === 'Enter') saveNote(h.hop, sessionNotes[h.hop] || '') }}
                                            className="w-28 bg-gray-800 text-white text-xs rounded px-2 py-1 border border-gray-600 focus:outline-none focus:border-blue-500"
                                          />
                                          <button onClick={() => saveNote(h.hop, sessionNotes[h.hop] || '')}
                                            className="bg-blue-700 hover:bg-blue-600 text-white text-xs px-2 py-1 rounded">💾</button>
                                        </div>
                                      </td>
                                      <td className="p-2 max-w-40">
                                        {(noteHistory[h.hop] || []).length === 0
                                          ? <span className="text-gray-600 text-xs">—</span>
                                          : <div className="flex flex-col gap-1 max-h-16 overflow-y-auto">
                                              {(noteHistory[h.hop] || []).slice(0, 3).map(n => (
                                                <div key={n.id} className="text-xs">
                                                  <span className="text-gray-500">{new Date(n.logged_at).toLocaleDateString('en-US', { month: 'numeric', day: 'numeric' })} </span>
                                                  <span className="text-gray-300">{n.note}</span>
                                                </div>
                                              ))}
                                            </div>
                                        }
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
                  })}
                </div>
              ))}
            </div>

            {/* Print button */}
            <div className="mt-6 flex justify-end">
              <button onClick={() => window.print()}
                className="bg-gray-700 hover:bg-gray-600 text-white px-4 py-2 rounded-lg text-sm font-semibold">
                🖨️ Print / Save as PDF
              </button>
            </div>
          </>
        )}

        {!loaded && (
          <div className="bg-gray-900 rounded-xl border border-gray-700 p-12 text-center">
            <p className="text-gray-400 text-xl">📂 Upload your tracker to view NTP status</p>
          </div>
        )}

      </div>
    </div>
  )
}
