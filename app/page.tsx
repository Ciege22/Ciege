'use client'

import { useEffect, useState, useCallback } from 'react'
import * as XLSX from 'xlsx'
import { saveTrackerSnapshot, loadTrackerSnapshot } from './lib/supabase'

const navItems = [
  { label: "HOP Readiness", href: "/weekly-focus", active: false },
  { label: "Deck Builder", href: "/deck-builder", active: true },
  { label: "GC Call View", href: "/gc-call", active: false },
  { label: "CM Call View", href: "/cm-view", active: false },
  { label: "Schedule Optimizer", href: "/schedule", active: false },
  { label: "NTP Tracker", href: "/ntp-tracker", active: false },
];

const stats = [
  { label: "Active Decks", value: "12", delta: "+8%" },
  { label: "Open Actions", value: "34", delta: "-4%" },
  { label: "NTPs Pending", value: "7", delta: "+14%" },
  { label: "SCOP Invoices", value: "18", delta: "+22%" },
];

function WeatherWidget() {
  const [weather, setWeather] = useState<{
    time: string
    temp: number
    condition: string
    wind: number
    windDir: number
    precip: number
    hourly: { time: string; temp: number; condition: string; precip: number; wind: number }[]
  } | null>(null)

  const getWindDir = (deg: number) => {
    const dirs = ['N','NE','E','SE','S','SW','W','NW']
    return dirs[Math.round(deg / 45) % 8]
  }

  useEffect(() => {
    const fetchWeather = async () => {
      try {
        const res = await fetch(
          'https://api.open-meteo.com/v1/forecast?latitude=40.2508&longitude=-103.7996&current=temperature_2m,weathercode,windspeed_10m,winddirection_10m,precipitation&hourly=temperature_2m,weathercode,precipitation_probability,windspeed_10m&temperature_unit=fahrenheit&timezone=America%2FDenver&forecast_days=1'
        )
        const data = await res.json()
        const now = new Date()
        const currentHour = now.getHours()

        const getCondition = (code: number) => {
          if (code === 0) return '☀️ Clear'
          if (code <= 3) return '⛅ Partly Cloudy'
          if (code <= 49) return '🌫️ Foggy'
          if (code <= 69) return '🌧️ Rain'
          if (code <= 79) return '🌨️ Snow'
          if (code <= 99) return '⛈️ Thunderstorm'
          return '🌤️ Mixed'
        }

        const hourly = data.hourly.time
          .slice(currentHour, currentHour + 6)
          .map((t: string, i: number) => ({
            time: new Date(t).toLocaleTimeString('en-US', { hour: 'numeric', hour12: true }),
            temp: Math.round(data.hourly.temperature_2m[currentHour + i]),
            condition: getCondition(data.hourly.weathercode[currentHour + i]),
            precip: data.hourly.precipitation_probability[currentHour + i] ?? 0,
            wind: Math.round(data.hourly.windspeed_10m[currentHour + i] ?? 0)
          }))

        setWeather({
          time: now.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit', hour12: true }),
          temp: Math.round(data.current.temperature_2m),
          condition: getCondition(data.current.weathercode),
          wind: Math.round(data.current.windspeed_10m ?? 0),
          windDir: data.current.winddirection_10m ?? 0,
          precip: data.current.precipitation ?? 0,
          hourly
        })
      } catch (e) {
        console.error('Weather fetch failed', e)
      }
    }
    fetchWeather()
    const interval = setInterval(fetchWeather, 300000)
    return () => clearInterval(interval)
  }, [])

  if (!weather) return <p className="text-gray-400 text-sm">Loading weather...</p>

  return (
    <div className="flex flex-col gap-2">
      <div className="flex items-center gap-4 flex-wrap">
        <span className="text-gray-300 text-sm font-semibold">{weather.time}</span>
        <span className="text-white text-lg font-bold">{weather.temp}°F</span>
        <span className="text-gray-300 text-sm">{weather.condition}</span>
        <span className="text-gray-400 text-sm">💨 {weather.wind} mph {getWindDir(weather.windDir)}</span>
        <span className="text-gray-400 text-sm">🌧️ {weather.precip} mm</span>
        <span className="text-gray-500 text-xs">Fort Morgan, CO</span>
      </div>
      <div className="flex gap-4 flex-wrap">
        {weather.hourly.map((h, i) => (
          <div key={i} className="text-center bg-gray-800 rounded-lg px-3 py-2">
            <p className="text-gray-500 text-xs">{h.time}</p>
            <p className="text-white text-sm font-semibold">{h.temp}°F</p>
            <p className="text-gray-400 text-xs">{h.condition.split(' ')[0]}</p>
            <p className="text-blue-400 text-xs">💧 {h.precip}%</p>
            <p className="text-gray-400 text-xs">💨 {h.wind} mph</p>
          </div>
        ))}
      </div>
    </div>
  )
}

export default function Home() {
  const [snapshotInfo, setSnapshotInfo] = useState<{ filename: string; uploaded_at: string; hop_count: number } | null>(null)
  const [uploading, setUploading] = useState(false)
  const [kpis, setKpis] = useState<{
    totalHops: number
    ntpComplete: number
    materialsReceived: number
    startsToDate: number
    completesToDate: number
    activeNow: number
    over18d: number
    startingThisWeek: number
    ntpUrgent: number
    spoNeeded: number
    openActions: number
    cutSpoNow: number
    materialWatch: number
    vendorConflicts: number
    currentMonthFcComplete: number
    currentMonthActComplete: number
  } | null>(null)
  const [modalTitle, setModalTitle] = useState('')
  const [modalHops, setModalHops] = useState<unknown[]>([])
  const [showModal, setShowModal] = useState(false)
  const [hopDetails, setHopDetails] = useState<{
    hop: string, gc: string, nokiaPm: string,
    ms15f: string, ms15a: string, ms16f: string, ms16a: string,
    hasNtp: boolean, ntpWaitingOn: string,
    hasMat: boolean, hasSpo: boolean, hasCpo: boolean,
    daysOut: number | null, daysElapsed: number | null,
    inProgress: boolean, complete: boolean, over18d: boolean,
    startingThisWeek: boolean, ntpUrgent: boolean,
    cutSpoNow: boolean, spoNeeded: boolean,
    materialWatch: boolean, vendorConflict: boolean,
    currentMonthFc: boolean, currentMonthAct: boolean,
  }[]>([])

  useEffect(() => {
    const checkSnapshot = async () => {
      const snap = await loadTrackerSnapshot()
      if (snap) setSnapshotInfo({ filename: snap.filename, uploaded_at: snap.uploaded_at, hop_count: snap.hop_count })
    }
    checkSnapshot()
  }, [])

  const computeKPIs = useCallback((rows: unknown[][]) => {
    const today = new Date()
    const currentMonth = today.getMonth()
    const currentYear = today.getFullYear()

    let headerRow = -1
    for (let i = 0; i < 10; i++) {
      if ((rows[i] as unknown[])?.some(c => String(c).trim() === 'HOP')) { headerRow = i; break }
    }
    if (headerRow === -1) return

    const headers = rows[headerRow] as string[]
    const col = (name: string) => headers.findIndex(h => String(h).trim() === name)

    const hopCol      = col('HOP')
    const don444Col   = col('DON 444')
    const nokiaPmCol  = col('Nokia PM')
    const ntpWaitCol  = col('NTP is waiting on')
    const ms15fCol    = col('MS15 Implementation Start F')
    const ms15aCol    = col('MS15 Implementation Start A')
    const ms16fCol    = col('MS16 Implementation Ends F')
    const ms16aCol    = col('MS16 Implementation Ends A')
    const ntpCol      = col('NTP A')
    const matCol      = headers.findIndex(h => String(h).trim().replace(/\s+/g, ' ').startsWith('Material Received A'))
    const spoCol      = headers.findIndex(h => String(h).trim().toLowerCase() === 'cx spo issued')
    const cpoCol      = headers.findIndex(h => String(h).trim().toLowerCase() === 'service cpo received')
    const itwSCol     = col('ITW Schedule Start')
    const itwECol     = col('ITW Schedule Complete')
    const ssSCol      = col('Samsung Schedule Start')
    const ssECol      = col('Samsung Schedule Complete')

    const parseD = (val: unknown): Date | null => {
      if (!val) return null
      if (val instanceof Date) return isNaN(val.getTime()) ? null : val
      if (typeof val === 'number') {
        if (val > 40000 && val < 60000) {
          const d = new Date((val - 25569) * 86400 * 1000)
          return isNaN(d.getTime()) ? null : d
        }
        return null
      }
      const s = String(val).trim()
      if (!s || s === 'null' || s === 'undefined' || s === 'NaN') return null
      const d = new Date(s)
      if (!isNaN(d.getTime())) return d
      const parts = s.split('/')
      if (parts.length === 3) {
        const d2 = new Date(parseInt(parts[2]), parseInt(parts[0])-1, parseInt(parts[1]))
        return isNaN(d2.getTime()) ? null : d2
      }
      return null
    }

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

    let totalHops = 0, ntpComplete = 0, materialsReceived = 0
    let startsToDate = 0, completesToDate = 0, activeNow = 0, over18d = 0
    let startingThisWeek = 0, ntpUrgent = 0, spoNeeded = 0
    let cutSpoNow = 0, materialWatch = 0, vendorConflicts = 0
    let currentMonthFcComplete = 0, currentMonthActComplete = 0

    hopRows.forEach((rows2) => {
      const row = rows2[0]
      totalHops++

      const ntpDate = parseD(row[ntpCol])
      const matDate = parseD(row[matCol])
      const ms15f   = parseD(row[ms15fCol])
      const ms15a   = parseD(row[ms15aCol])
      const ms16f   = parseD(row[ms16fCol])
      const ms16a   = parseD(row[ms16aCol])
      const spoDate = parseD(row[spoCol])
      const cpoVal  = String(row[cpoCol] || '').trim()

      const hasNtp     = !!(ntpDate && ntpDate.getFullYear() >= 2025)
      const hasMat     = !!(matDate && !isNaN(matDate.getTime()) && matDate.getFullYear() >= 2020) || !!ms16a
      const hasSpo     = !!spoDate
      const hasCpo     = cpoVal.length > 0 && cpoVal.toLowerCase() !== 'nan'
      const started    = !!ms15a
      const complete   = !!ms16a
      const inProgress = started && !complete
      const daysOut    = ms15f ? Math.round((ms15f.getTime() - today.getTime()) / (1000 * 60 * 60 * 24)) : null
      const daysElapsed = inProgress && ms15a ? Math.round((today.getTime() - ms15a.getTime()) / (1000 * 60 * 60 * 24)) : null

      if (hasNtp) ntpComplete++
      if (hasMat) materialsReceived++
      if (started) startsToDate++
      if (complete) completesToDate++
      if (inProgress) activeNow++
      if (inProgress && (daysElapsed ?? 0) > 18) over18d++
      if (!started && !complete && daysOut !== null && daysOut >= 0 && daysOut <= 7) startingThisWeek++
      if (!hasNtp && !complete && daysOut !== null && daysOut <= 14) ntpUrgent++
      if (!hasSpo && !complete) {
        if (hasCpo) cutSpoNow++
        else spoNeeded++
      }
      if (!hasMat && !complete && daysOut !== null && daysOut <= 14) materialWatch++

      if (ms16f && ms16f.getMonth() === currentMonth && ms16f.getFullYear() === currentYear) currentMonthFcComplete++
      if (ms16a && ms16a.getMonth() === currentMonth && ms16a.getFullYear() === currentYear) currentMonthActComplete++

      let hasConflict = false
      rows2.forEach(r => {
        const itwS = parseD(r[itwSCol]); const itwE = parseD(r[itwECol])
        const ssS  = parseD(r[ssSCol]);  const ssE  = parseD(r[ssECol])
        if (ms15f) {
          if (itwS && itwE && itwS <= ms15f && ms15f <= itwE) hasConflict = true
          if (ssS && ssE && ssS <= ms15f && ms15f <= ssE) hasConflict = true
        }
      })
      if (hasConflict && !complete) vendorConflicts++
    })

    const details: {
      hop: string, gc: string, nokiaPm: string,
      ms15f: string, ms15a: string, ms16f: string, ms16a: string,
      hasNtp: boolean, ntpWaitingOn: string,
      hasMat: boolean, hasSpo: boolean, hasCpo: boolean,
      daysOut: number | null, daysElapsed: number | null,
      inProgress: boolean, complete: boolean, over18d: boolean,
      startingThisWeek: boolean, ntpUrgent: boolean,
      cutSpoNow: boolean, spoNeeded: boolean,
      materialWatch: boolean, vendorConflict: boolean,
      currentMonthFc: boolean, currentMonthAct: boolean,
    }[] = []
    hopRows.forEach((rows2) => {
      const row = rows2[0]
      const ntpDate = parseD(row[ntpCol])
      const matDate = parseD(row[matCol])
      const ms15f   = parseD(row[ms15fCol])
      const ms15a   = parseD(row[ms15aCol])
      const ms16f   = parseD(row[ms16fCol])
      const ms16a   = parseD(row[ms16aCol])
      const spoDate = parseD(row[spoCol])
      const cpoVal  = String(row[cpoCol] || '').trim()
      const hasNtp  = !!(ntpDate && ntpDate.getFullYear() >= 2025)
      const hasMat  = !!(matDate && !isNaN(matDate.getTime()) && matDate.getFullYear() >= 2020) || !!ms16a
      const hasSpo  = !!spoDate
      const hasCpo  = cpoVal.length > 0 && cpoVal.toLowerCase() !== 'nan'
      const started    = !!ms15a
      const complete   = !!ms16a
      const inProgress = started && !complete
      const daysOut    = ms15f ? Math.round((ms15f.getTime() - today.getTime()) / (1000*60*60*24)) : null
      const daysElapsed = inProgress && ms15a ? Math.round((today.getTime() - ms15a.getTime()) / (1000*60*60*24)) : null
      const fmtD = (d: Date | null) => d ? `${d.getMonth()+1}/${d.getDate()}/${d.getFullYear()}` : ''
      let hasConflict2 = false
      rows2.forEach(r => {
        const itwS = parseD(r[itwSCol]); const itwE = parseD(r[itwECol])
        const ssS  = parseD(r[ssSCol]);  const ssE  = parseD(r[ssECol])
        if (ms15f) {
          if (itwS && itwE && itwS <= ms15f && ms15f <= itwE) hasConflict2 = true
          if (ssS && ssE && ssS <= ms15f && ms15f <= ssE) hasConflict2 = true
        }
      })
      details.push({
        hop: String(row[hopCol] || '').trim(),
        gc: String(row[col('General Contractor')] || '').trim(),
        nokiaPm: String(row[nokiaPmCol] || '').trim(),
        ms15f: fmtD(ms15f), ms15a: fmtD(ms15a),
        ms16f: fmtD(ms16f), ms16a: fmtD(ms16a),
        hasNtp, ntpWaitingOn: String(row[ntpWaitCol] || '').trim(),
        hasMat, hasSpo, hasCpo,
        daysOut, daysElapsed, inProgress, complete,
        over18d: inProgress && (daysElapsed??0) > 18,
        startingThisWeek: !started && !complete && daysOut !== null && daysOut >= 0 && daysOut <= 7,
        ntpUrgent: !hasNtp && !complete && daysOut !== null && daysOut <= 14,
        cutSpoNow: !hasSpo && !complete && hasCpo,
        spoNeeded: !hasSpo && !complete && !hasCpo,
        materialWatch: !hasMat && !complete && daysOut !== null && daysOut <= 14,
        vendorConflict: hasConflict2 && !complete,
        currentMonthFc: !!(ms16f && ms16f.getMonth() === currentMonth && ms16f.getFullYear() === currentYear),
        currentMonthAct: !!(ms16a && ms16a.getMonth() === currentMonth && ms16a.getFullYear() === currentYear),
      })
    })
    setHopDetails(details)

    setKpis({
      totalHops, ntpComplete, materialsReceived, startsToDate, completesToDate,
      activeNow, over18d, startingThisWeek, ntpUrgent, spoNeeded,
      openActions: 0, cutSpoNow, materialWatch, vendorConflicts,
      currentMonthFcComplete, currentMonthActComplete
    })
  }, [])

  useEffect(() => {
    const loadKPIs = async () => {
      const snap = await loadTrackerSnapshot()
      if (snap) computeKPIs(snap.data)
    }
    loadKPIs()
  }, [computeKPIs])

  const handleTrackerUpload = useCallback(async (file: File) => {
    setUploading(true)
    const reader = new FileReader()
    reader.onload = async (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer)
        const wb = XLSX.read(data, { type: 'array', cellDates: true })
        const ws = wb.Sheets['HOPs']
        if (!ws) { alert('HOPs tab not found'); setUploading(false); return }
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1 }) as unknown[][]
        const hopCount = rows.filter((r: unknown[]) => String(r[4] || '').trim() && String(r[4]).trim() !== 'HOP').length
        await saveTrackerSnapshot(file.name, hopCount, rows)
        setSnapshotInfo({ filename: file.name, uploaded_at: new Date().toISOString(), hop_count: hopCount })
        computeKPIs(rows)
      } catch (err) {
        console.error('Upload error:', err)
      }
      setUploading(false)
    }
    reader.readAsArrayBuffer(file)
  }, [computeKPIs])

  const openModal = (title: string, filter: (h: typeof hopDetails[0]) => boolean) => {
    setModalTitle(title)
    setModalHops(hopDetails.filter(filter))
    setShowModal(true)
  }

  return (
    <div className="min-h-screen bg-zinc-950 text-zinc-100">
      <div className="mx-auto max-w-full px-12 py-12">
        <div className="grid gap-10 lg:grid-cols-[360px_minmax(0,1fr)]">
          <aside className="rounded-[32px] border border-white/10 bg-white/5 p-6 shadow-[0_24px_120px_-80px_rgba(0,0,0,0.55)] backdrop-blur-xl">
            <div className="mb-8">
              <p className="text-xs uppercase tracking-[0.4em] text-emerald-300/80">Ciege</p>
              <h1 className="mt-3 text-3xl font-semibold text-white">PM Dashboard</h1>
              <p className="mt-2 text-sm leading-6 text-zinc-400">
                Plan Execute Conquer
              </p>
            </div>

            <div className="mb-4">
              <p className="text-xs text-gray-500 uppercase font-semibold mb-2">Tracker</p>
              <div
                className="border border-dashed border-gray-600 rounded-lg p-3 cursor-pointer hover:border-blue-500 transition-colors"
                onDragOver={(e) => e.preventDefault()}
                onDrop={(e) => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) handleTrackerUpload(f) }}
                onClick={() => document.getElementById('dashboard-tracker-upload')?.click()}
              >
                <input id="dashboard-tracker-upload" type="file" accept=".xlsx" className="hidden"
                  onChange={(e) => { const f = e.target.files?.[0]; if (f) handleTrackerUpload(f) }} />
                {uploading
                  ? <p className="text-blue-400 text-xs text-center">⏳ Uploading...</p>
                  : snapshotInfo
                  ? <div>
                      <p className="text-green-400 text-xs font-semibold">✅ {snapshotInfo.filename}</p>
                      <p className="text-gray-500 text-xs mt-1">{snapshotInfo.hop_count} HOPs · {new Date(snapshotInfo.uploaded_at).toLocaleDateString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric' })} at {new Date(snapshotInfo.uploaded_at).toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })}</p>
                    </div>
                  : <p className="text-gray-400 text-xs text-center">📂 Upload tracker here</p>
                }
              </div>
            </div>

            <nav className="space-y-2">
              {navItems.map((item) =>
                item.href ? (
                  <a
                    key={item.label}
                    href={item.href}
                    className={`block w-full rounded-2xl px-4 py-3 text-left text-sm font-medium transition hover:bg-white/10 ${
                      item.active ? "bg-emerald-400/10 text-emerald-300" : "text-zinc-300"
                    }`}
                  >
                    {item.label}
                  </a>
                ) : (
                  <button
                    key={item.label}
                    className={`w-full rounded-2xl px-4 py-3 text-left text-sm font-medium transition hover:bg-white/10 ${
                      item.active ? "bg-emerald-400/10 text-emerald-300" : "text-zinc-300"
                    }`}
                    type="button"
                  >
                    {item.label}
                  </button>
                )
              )}
            </nav>
          </aside>

          <main className="space-y-6">
            <section className="rounded-[32px] border border-white/10 bg-white/5 p-8 shadow-[0_24px_120px_-80px_rgba(0,0,0,0.55)] backdrop-blur-xl">
              <div className="flex flex-col gap-6 lg:flex-row lg:items-end lg:justify-between">
                <div className="max-w-2xl">
                  <p className="text-sm uppercase tracking-[0.4em] text-emerald-300/80">
                    Welcome back, team
                  </p>
                  <h2 className="mt-4 text-4xl font-semibold text-white">
                    Let's get to work!
                  </h2>
                  <div className="mt-4">
                    <WeatherWidget />
                  </div>
                </div>

                <img
                  src="/hylian-crest.png"
                  alt="Hylian Crest"
                  className="w-64 h-auto"
                  style={{
                    filter: 'invert(72%) sepia(98%) saturate(346%) hue-rotate(106deg) brightness(95%) contrast(88%)'
                  }}
                />
              </div>
            </section>

            {kpis && (
              <div className="space-y-3 mb-6">
                {/* Row 1 — Program Health */}
                <div className="grid grid-cols-2 md:grid-cols-7 gap-3">
                  {[
                    { label: 'Total HOPs', value: kpis.totalHops, color: 'text-white', sub: 'DON 444 program', filter: () => true },
                    { label: 'NTP Complete', value: kpis.ntpComplete, color: 'text-green-400', sub: `${Math.round(kpis.ntpComplete/kpis.totalHops*100)}% of program`, filter: (h: typeof hopDetails[0]) => h.hasNtp },
                    { label: 'Materials Received', value: kpis.materialsReceived, color: 'text-green-400', sub: `${Math.round(kpis.materialsReceived/kpis.totalHops*100)}% of program`, filter: (h: typeof hopDetails[0]) => h.hasMat },
                    { label: 'Starts to Date', value: kpis.startsToDate, color: 'text-blue-400', sub: 'MS15 A confirmed', filter: (h: typeof hopDetails[0]) => !!h.ms15a },
                    { label: 'Completes to Date', value: kpis.completesToDate, color: 'text-teal-400', sub: 'MS16 A confirmed', filter: (h: typeof hopDetails[0]) => h.complete },
                    { label: `${new Date().toLocaleString('default',{month:'short'})} FC Completes`, value: kpis.currentMonthFcComplete, color: 'text-yellow-400', sub: 'MS16 F this month', filter: (h: typeof hopDetails[0]) => h.currentMonthFc },
                    { label: `${new Date().toLocaleString('default',{month:'short'})} Act Completes`, value: kpis.currentMonthActComplete, color: 'text-emerald-400', sub: 'MS16 A this month', filter: (h: typeof hopDetails[0]) => h.currentMonthAct },
                  ].map(({ label, value, color, sub, filter }) => (
                    <div key={label} onClick={() => openModal(label, filter)}
                      className="bg-gray-900 rounded-xl border border-gray-700 p-3 text-center cursor-pointer hover:border-blue-500 hover:bg-gray-800 transition-all">
                      <p className="text-gray-500 text-xs">{label}</p>
                      <p className={`text-2xl font-bold ${color}`}>{value}</p>
                      <p className="text-gray-600 text-xs mt-1">{sub}</p>
                    </div>
                  ))}
                </div>

                {/* Row 2 — Needs Action Today */}
                <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
                  {[
                    { label: 'Active Sites', value: kpis.activeNow, color: 'text-blue-400', sub: 'crews on site now', filter: (h: typeof hopDetails[0]) => h.inProgress },
                    { label: 'Over 18 Days', value: kpis.over18d, color: kpis.over18d > 0 ? 'text-red-400' : 'text-green-400', sub: 'past target duration', filter: (h: typeof hopDetails[0]) => h.over18d },
                    { label: 'Starting This Week', value: kpis.startingThisWeek, color: 'text-orange-400', sub: '0–7 days out', filter: (h: typeof hopDetails[0]) => h.startingThisWeek },
                    { label: 'NTP Urgent', value: kpis.ntpUrgent, color: kpis.ntpUrgent > 0 ? 'text-red-400' : 'text-green-400', sub: 'missing NTP ≤14d', filter: (h: typeof hopDetails[0]) => h.ntpUrgent },
                    { label: 'SPO Needed', value: kpis.spoNeeded, color: kpis.spoNeeded > 0 ? 'text-red-400' : 'text-green-400', sub: 'no CPO yet', filter: (h: typeof hopDetails[0]) => h.spoNeeded },
                  ].map(({ label, value, color, sub, filter }) => (
                    <div key={label} onClick={() => openModal(label, filter)}
                      className="bg-gray-900 rounded-xl border border-gray-700 p-3 text-center cursor-pointer hover:border-blue-500 hover:bg-gray-800 transition-all">
                      <p className="text-gray-500 text-xs">{label}</p>
                      <p className={`text-2xl font-bold ${color}`}>{value}</p>
                      <p className="text-gray-600 text-xs mt-1">{sub}</p>
                    </div>
                  ))}
                </div>

                {/* Row 3 — Actions */}
                <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                  {[
                    { label: 'Cut SPO Now', value: kpis.cutSpoNow, color: kpis.cutSpoNow > 0 ? 'text-yellow-400' : 'text-green-400', sub: 'CPO ready — no SPO', filter: (h: typeof hopDetails[0]) => h.cutSpoNow },
                    { label: 'Material Watch', value: kpis.materialWatch, color: kpis.materialWatch > 0 ? 'text-orange-400' : 'text-green-400', sub: 'no mat ≤14d to start', filter: (h: typeof hopDetails[0]) => h.materialWatch },
                    { label: 'Vendor Conflicts', value: kpis.vendorConflicts, color: kpis.vendorConflicts > 0 ? 'text-red-400' : 'text-green-400', sub: 'Samsung/ITW on site at start', filter: (h: typeof hopDetails[0]) => h.vendorConflict },
                    { label: 'Open Actions', value: kpis.openActions, color: 'text-blue-400', sub: 'from HOP Readiness', filter: () => false },
                  ].map(({ label, value, color, sub, filter }) => (
                    <div key={label} onClick={() => openModal(label, filter)}
                      className="bg-gray-900 rounded-xl border border-gray-700 p-3 text-center cursor-pointer hover:border-blue-500 hover:bg-gray-800 transition-all">
                      <p className="text-gray-500 text-xs">{label}</p>
                      <p className={`text-2xl font-bold ${color}`}>{value}</p>
                      <p className="text-gray-600 text-xs mt-1">{sub}</p>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </main>
        </div>
      </div>

      {/* KPI Drill-Down Modal */}
      {showModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/70" onClick={() => setShowModal(false)}>
          <div className="bg-gray-900 border border-gray-700 rounded-xl shadow-2xl w-full max-w-5xl max-h-[80vh] flex flex-col mx-4" onClick={e => e.stopPropagation()}>
            <div className="flex items-center justify-between px-6 py-4 border-b border-gray-700">
              <h2 className="text-white font-semibold text-lg">{modalTitle} <span className="text-gray-400 text-sm ml-2">({(modalHops as typeof hopDetails).length} sites)</span></h2>
              <button onClick={() => setShowModal(false)} className="text-gray-400 hover:text-white text-xl leading-none">&times;</button>
            </div>
            <div className="overflow-auto flex-1">
              {(modalHops as typeof hopDetails).length === 0 ? (
                <p className="text-gray-500 text-center py-10">No sites match this filter.</p>
              ) : (
                <table className="w-full text-xs text-left">
                  <thead className="sticky top-0 bg-gray-800 text-gray-400">
                    <tr>
                      <th className="px-3 py-2">HOP</th>
                      <th className="px-3 py-2">GC</th>
                      <th className="px-3 py-2">Nokia PM</th>
                      <th className="px-3 py-2">FC Start</th>
                      <th className="px-3 py-2">AC Start</th>
                      <th className="px-3 py-2">FC End</th>
                      <th className="px-3 py-2">AC End</th>
                      <th className="px-3 py-2">NTP</th>
                      <th className="px-3 py-2">NTP Waiting On</th>
                      <th className="px-3 py-2">Mat</th>
                      <th className="px-3 py-2">SPO</th>
                    </tr>
                  </thead>
                  <tbody>
                    {(modalHops as typeof hopDetails).map((h, i) => (
                      <tr key={i} className={i % 2 === 0 ? 'bg-gray-900' : 'bg-gray-850'}>
                        <td className="px-3 py-1.5 font-mono text-white">{h.hop}</td>
                        <td className="px-3 py-1.5 text-gray-300">{h.gc}</td>
                        <td className="px-3 py-1.5 text-gray-300">{h.nokiaPm}</td>
                        <td className="px-3 py-1.5 text-gray-400">{h.ms15f}</td>
                        <td className="px-3 py-1.5 text-blue-400">{h.ms15a}</td>
                        <td className="px-3 py-1.5 text-gray-400">{h.ms16f}</td>
                        <td className="px-3 py-1.5 text-teal-400">{h.ms16a}</td>
                        <td className="px-3 py-1.5">{h.hasNtp ? <span className="text-green-400">✓</span> : <span className="text-red-400">✗</span>}</td>
                        <td className="px-3 py-1.5 text-gray-400 max-w-[180px] truncate">{h.ntpWaitingOn}</td>
                        <td className="px-3 py-1.5">{h.hasMat ? <span className="text-green-400">✓</span> : <span className="text-red-400">✗</span>}</td>
                        <td className="px-3 py-1.5">{h.hasSpo ? <span className="text-green-400">✓</span> : h.hasCpo ? <span className="text-yellow-400">CPO</span> : <span className="text-red-400">✗</span>}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
