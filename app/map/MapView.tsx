'use client'

import { useEffect, useMemo, useState } from 'react'
import { MapContainer, TileLayer, CircleMarker, Polyline, Popup, useMap } from 'react-leaflet'
import type { LatLngBoundsExpression } from 'leaflet'
import 'leaflet/dist/leaflet.css'
import { useIsAuthenticated } from '@azure/msal-react'
import { useSharePointFile } from '../components/SharePointDataProvider'
import SharePointSourceLine from '../components/SharePointSourceLine'
import MicrosoftLogin from '../components/MicrosoftLogin'
import { DEFAULT_THRESHOLDS, loadThresholdSettings, type ThresholdSettings } from '../lib/settings'
import { MAP_STATUS_ORDER, parseMapHops, type MapHop, type MapStatus } from '../lib/mapHops'

// HOP Network Map v2. Data still comes only from the shared SharePoint
// provider (live, in memory) — nothing here stores Nokia data anywhere. The
// status rules live in app/lib/mapHops.ts and mirror the Dashboard tiles.
//
// Default view = every HOP that is active or not yet started (anything that
// is NOT complete). Completed HOPs are one click away via the Complete chip.

const FALLBACK_CENTER: [number, number] = [39.8, -101.3]

const STATUS_STYLE: Record<MapStatus, { label: string; color: string; dash?: string; weight: number }> = {
  ready:    { label: 'Ready to start',        color: '#3DDC84', weight: 4 },
  progress: { label: 'In progress',           color: '#00C2D6', weight: 4 },
  ntp:      { label: 'Needs NTP',             color: '#FFB020', dash: '10 7', weight: 4 },
  mat:      { label: 'Needs material',        color: '#B58CFF', dash: '2 7',  weight: 4 },
  blocked:  { label: 'NTP + material missing', color: '#FF5C5C', dash: '10 7', weight: 5 },
  complete: { label: 'Complete',              color: '#8FA3C8', weight: 3 },
}

const GC_PALETTE = ['#00C2D6', '#FFB020', '#B58CFF', '#3DDC84', '#FF7AB6', '#6FA8FF', '#F2E55C', '#FF8A4C', '#9BE15D', '#C9C9D6']

const PHASE_LABEL: Record<string, string> = {
  downloading: 'Downloading tracker from SharePoint…',
  parsing: 'Parsing tracker…',
  building: 'Building map…',
}

function whenText(daysOut: number | null): string {
  if (daysOut === null) return 'no forecast start'
  if (daysOut < 0) return `${-daysOut} day${daysOut === -1 ? '' : 's'} overdue`
  if (daysOut === 0) return 'today'
  return `in ${daysOut} day${daysOut === 1 ? '' : 's'}`
}

function alertFor(h: MapHop, t: ThresholdSettings): { text: string; color: string } {
  switch (h.status) {
    case 'complete':
      return { text: `Construction complete${h.ms16a ? ` (${h.ms16a})` : ''}.`, color: STATUS_STYLE.complete.color }
    case 'progress':
      return {
        text: h.over18d
          ? `In progress ${h.daysElapsed} days — over the ${t.durationAlertDays}-day duration alert.`
          : `In progress, day ${h.daysElapsed ?? 0}.`,
        color: h.over18d ? '#FFB020' : STATUS_STYLE.progress.color,
      }
    case 'ready':
      return { text: `Ready. NTP and material are both in. Forecast start ${whenText(h.daysOut)}.`, color: STATUS_STYLE.ready.color }
    case 'ntp':
      return {
        text: h.ntpUrgent
          ? `NTP urgent: forecast start ${whenText(h.daysOut)}. Chase the Program Team now.`
          : `NTP still missing. Forecast start ${whenText(h.daysOut)}.`,
        color: STATUS_STYLE.ntp.color,
      }
    case 'mat':
      return {
        text: h.materialWatch
          ? `Material watch: forecast start ${whenText(h.daysOut)} and material is not received. Escalate to the GC.`
          : `Material not received yet. Forecast start ${whenText(h.daysOut)}.`,
        color: STATUS_STYLE.mat.color,
      }
    default:
      return { text: `Neither NTP nor material is in. Forecast start ${whenText(h.daysOut)}.`, color: STATUS_STYLE.blocked.color }
  }
}

// Fits once when the data first arrives (not on every filter click, which
// would yank the map around while the user is panning).
function MapController({ fitTo, focus }: { fitTo: LatLngBoundsExpression | null; focus: { bounds: LatLngBoundsExpression; n: number } | null }) {
  const map = useMap()
  useEffect(() => {
    if (fitTo) map.fitBounds(fitTo, { padding: [30, 30] })
  }, [fitTo, map])
  useEffect(() => {
    if (focus) map.flyToBounds(focus.bounds, { padding: [90, 90], maxZoom: 11, duration: 0.8 })
  }, [focus, map])
  return null
}

function boundsOf(points: [number, number][]): LatLngBoundsExpression | null {
  if (points.length === 0) return null
  const lats = points.map(p => p[0])
  const lons = points.map(p => p[1])
  return [[Math.min(...lats), Math.min(...lons)], [Math.max(...lats), Math.max(...lons)]]
}

export default function MapView() {
  const isAuthenticated = useIsAuthenticated()
  const { rows, loading, phase, error: loadError } = useSharePointFile('tracker')

  const [thresholds, setThresholds] = useState<ThresholdSettings>(DEFAULT_THRESHOLDS)
  useEffect(() => {
    let cancelled = false
    loadThresholdSettings().then(t => { if (!cancelled) setThresholds(t) }).catch(() => {})
    return () => { cancelled = true }
  }, [])

  const [hidden, setHidden] = useState<Record<MapStatus, boolean>>({
    ready: false, progress: false, ntp: false, mat: false, blocked: false, complete: true,
  })
  const [soonOnly, setSoonOnly] = useState(false)
  const [gcFilter, setGcFilter] = useState('All')
  const [query, setQuery] = useState('')
  const [colorBy, setColorBy] = useState<'status' | 'gc'>('status')
  const [selected, setSelected] = useState<string | null>(null)
  const [focus, setFocus] = useState<{ bounds: LatLngBoundsExpression; n: number } | null>(null)

  const parsed = useMemo(() => (rows ? parseMapHops(rows, thresholds) : null), [rows, thresholds])
  const allHops = useMemo(() => parsed?.hops ?? [], [parsed])

  const gcList = useMemo(() => {
    const set = new Set<string>()
    allHops.forEach(h => { if (h.gc) set.add(h.gc) })
    return Array.from(set).sort()
  }, [allHops])
  const gcColor = useMemo(() => {
    const m = new Map<string, string>()
    gcList.forEach((g, i) => m.set(g, GC_PALETTE[i % GC_PALETTE.length]))
    return m
  }, [gcList])

  const statusCounts = useMemo(() => {
    const c: Record<MapStatus, number> = { ready: 0, progress: 0, ntp: 0, mat: 0, blocked: 0, complete: 0 }
    allHops.forEach(h => { c[h.status]++ })
    return c
  }, [allHops])

  const q = query.trim().toLowerCase()
  const matchesQuery = (h: MapHop) =>
    !q || h.hop.toLowerCase().includes(q) || h.a.name.toLowerCase().includes(q) || h.b.name.toLowerCase().includes(q)

  const visible = useMemo(() => allHops.filter(h =>
    !hidden[h.status]
    && (!soonOnly || h.startingSoon)
    && (gcFilter === 'All' || h.gc === gcFilter)
    && matchesQuery(h)
  // eslint-disable-next-line react-hooks/exhaustive-deps
  ), [allHops, hidden, soonOnly, gcFilter, q])

  const visibleSites = useMemo(() => {
    const m = new Map<string, { lat: number; lon: number; count: number }>()
    visible.forEach(h => {
      for (const s of [h.a, h.b]) {
        const e = m.get(s.name)
        if (e) e.count++
        else m.set(s.name, { lat: s.lat, lon: s.lon, count: 1 })
      }
    })
    return m
  }, [visible])

  // Initial framing = the default view (everything not complete).
  const fitTo = useMemo(() => {
    const pts: [number, number][] = []
    allHops.filter(h => h.status !== 'complete').forEach(h => { pts.push([h.a.lat, h.a.lon], [h.b.lat, h.b.lon]) })
    if (pts.length === 0) allHops.forEach(h => { pts.push([h.a.lat, h.a.lon], [h.b.lat, h.b.lon]) })
    return boundsOf(pts)
  }, [allHops])

  const selectedHop = useMemo(() => allHops.find(h => h.hop === selected) ?? null, [allHops, selected])

  const count = (f: (h: MapHop) => boolean) => visible.filter(f).length
  const counters = [
    { label: 'HOPs shown', value: visible.length, sub: `of ${allHops.length} in the program` },
    { label: 'Ready to start', value: count(h => h.status === 'ready'), sub: 'NTP and material both in' },
    { label: 'Blocked', value: count(h => h.status === 'ntp' || h.status === 'mat' || h.status === 'blocked'), sub: 'missing NTP and/or material' },
    { label: 'In progress', value: count(h => h.status === 'progress'), sub: 'started, not yet complete' },
    { label: 'Starting in 14 days', value: count(h => h.startingSoon), sub: 'forecast start inside the window' },
  ]

  const colorFor = (h: MapHop) => colorBy === 'gc' ? (gcColor.get(h.gc) ?? '#C9C9D6') : STATUS_STYLE[h.status].color
  const dashFor = (h: MapHop) => colorBy === 'gc' ? undefined : STATUS_STYLE[h.status].dash

  function submitSearch() {
    if (!q) return
    const hit = allHops.find(h => h.hop.toLowerCase().includes(q) || h.a.name.toLowerCase().includes(q) || h.b.name.toLowerCase().includes(q))
    if (!hit) return
    // A search hit that a status chip is hiding gets that chip turned back on,
    // otherwise Enter would "find" something that never appears.
    if (hidden[hit.status]) setHidden(prev => ({ ...prev, [hit.status]: false }))
    setSelected(hit.hop)
    const b = boundsOf([[hit.a.lat, hit.a.lon], [hit.b.lat, hit.b.lon]])
    if (b) setFocus(prev => ({ bounds: b, n: (prev?.n ?? 0) + 1 }))
  }

  if (!isAuthenticated) {
    return (
      <div className="rounded-2xl border border-white/10 bg-white/5 p-6">
        <p className="text-sm text-zinc-400 mb-3">Sign in with your Nokia account to load the tracker from SharePoint.</p>
        <MicrosoftLogin />
      </div>
    )
  }
  // Only a full-page block when there's nothing to show yet — a cached entry
  // (stale-while-revalidate) renders the map immediately instead.
  if (loading && !rows) {
    return <div className="flex items-center justify-center h-[70vh] text-zinc-400 text-sm">{PHASE_LABEL[phase] ?? 'Loading tracker from SharePoint…'}</div>
  }
  if (loadError && !rows) {
    return (
      <div className="rounded-2xl border border-red-500/20 bg-red-500/10 px-4 py-3 text-sm text-red-300">
        {loadError}
      </div>
    )
  }
  if (parsed?.error) {
    return (
      <div className="rounded-2xl border border-red-500/20 bg-red-500/10 px-4 py-3 text-sm text-red-300">
        {parsed.error}
      </div>
    )
  }

  // Active chips get their own border color via inline style; inactive ones a plain grey border.
  const chip = (on: boolean) =>
    `inline-flex items-center gap-2 h-8 px-3 rounded-full text-xs font-semibold border transition-colors ${on ? 'bg-white/10 text-zinc-100' : 'bg-transparent text-zinc-500 hover:text-zinc-300 border-white/15'}`
  const hiddenCompleted = hidden.complete ? statusCounts.complete : 0

  return (
    <div>
      <div className="mb-1 text-xs text-zinc-400">
        <SharePointSourceLine fileKey="tracker" />
      </div>

      {/* Live counters */}
      <div className="grid grid-cols-2 md:grid-cols-5 gap-3 mb-3">
        {counters.map(c => (
          <div key={c.label} className="rounded-xl border border-white/10 bg-white/5 px-4 py-3">
            <div className="text-xs text-zinc-400 font-medium">{c.label}</div>
            <div className="text-2xl font-bold text-white leading-tight">{c.value}</div>
            <div className="text-[11px] text-zinc-500">{c.sub}</div>
          </div>
        ))}
      </div>

      <div className="flex flex-col lg:flex-row gap-4 items-start">
        <section className="w-full min-w-0 flex-1 rounded-2xl border border-white/10 bg-white/5 overflow-hidden">

          {/* Filter bar */}
          <div className="flex flex-col gap-2 p-3 border-b border-white/10">
            <div className="flex flex-wrap items-center gap-2">
              <span className="text-xs text-zinc-400 font-semibold mr-1">Status</span>
              {MAP_STATUS_ORDER.map(k => {
                const on = !hidden[k]
                const s = STATUS_STYLE[k]
                return (
                  <button
                    key={k}
                    type="button"
                    aria-pressed={on}
                    onClick={() => setHidden(prev => ({ ...prev, [k]: !prev[k] }))}
                    className={chip(on)}
                    style={on ? { borderColor: s.color } : undefined}
                  >
                    <span
                      className="inline-block w-2.5 h-2.5 rounded-full box-border"
                      style={{ background: on ? s.color : 'transparent', border: `2px solid ${s.color}` }}
                    />
                    {s.label}
                    <span className="text-zinc-400 font-medium">{statusCounts[k]}</span>
                  </button>
                )
              })}
            </div>
            <div className="flex flex-wrap items-center gap-2">
              <span className="text-xs text-zinc-400 font-semibold mr-1">Filter</span>
              <button
                type="button"
                aria-pressed={soonOnly}
                onClick={() => setSoonOnly(v => !v)}
                className={chip(soonOnly)}
                style={soonOnly ? { borderColor: '#00A0B0' } : undefined}
              >
                Starts within 14 days
              </button>
              <span className="w-px h-5 bg-white/15 mx-1" />
              {['All', ...gcList].map(g => (
                <button
                  key={g}
                  type="button"
                  aria-pressed={gcFilter === g}
                  onClick={() => setGcFilter(g)}
                  className={chip(gcFilter === g)}
                  style={gcFilter === g ? { borderColor: '#00A0B0' } : undefined}
                >
                  {g === 'All' ? 'All GCs' : g}
                </button>
              ))}
            </div>
            <div className="flex flex-wrap items-center gap-3">
              <label className="sr-only" htmlFor="map-search">Search HOP or site</label>
              <input
                id="map-search"
                type="search"
                value={query}
                onChange={e => setQuery(e.target.value)}
                onKeyDown={e => { if (e.key === 'Enter') submitSearch() }}
                placeholder="Search HOP or site, Enter to zoom…"
                className="h-9 w-72 max-w-full rounded-lg border border-white/15 bg-zinc-950 px-3 text-sm text-zinc-100 placeholder:text-zinc-500"
              />
              <div className="flex items-center gap-1 rounded-lg border border-white/15 p-0.5 text-xs font-semibold">
                <span className="px-2 text-zinc-400">Color by</span>
                {(['status', 'gc'] as const).map(m => (
                  <button
                    key={m}
                    type="button"
                    aria-pressed={colorBy === m}
                    onClick={() => setColorBy(m)}
                    className={`h-7 px-3 rounded-md ${colorBy === m ? 'bg-[#124191] text-white' : 'text-zinc-400 hover:text-zinc-200'}`}
                  >
                    {m === 'status' ? 'Status' : 'GC'}
                  </button>
                ))}
              </div>
              {hiddenCompleted > 0 && (
                <span className="text-xs text-zinc-500">{hiddenCompleted} completed HOPs hidden. Turn on Complete to show them.</span>
              )}
              {parsed && parsed.unmapped > 0 && (
                <span className="text-xs text-amber-400">{parsed.unmapped} HOP{parsed.unmapped === 1 ? '' : 's'} not drawn: missing coordinates on a site.</span>
              )}
            </div>
          </div>

          {/* Map */}
          <div className="relative" style={{ height: '68vh', minHeight: 420 }}>
            <MapContainer center={FALLBACK_CENTER} zoom={7} style={{ height: '100%', width: '100%' }}>
              <TileLayer
                attribution='&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
                url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png"
              />
              <MapController fitTo={fitTo} focus={focus} />
              {selectedHop && visible.some(h => h.hop === selectedHop.hop) && (
                <Polyline
                  key={`halo-${selectedHop.hop}`}
                  positions={[[selectedHop.a.lat, selectedHop.a.lon], [selectedHop.b.lat, selectedHop.b.lon]]}
                  pathOptions={{ color: '#FFFFFF', weight: 11, opacity: 0.55 }}
                  interactive={false}
                />
              )}
              {visible.map(h => (
                <Polyline
                  key={h.hop}
                  positions={[[h.a.lat, h.a.lon], [h.b.lat, h.b.lon]]}
                  pathOptions={{
                    color: colorFor(h),
                    weight: STATUS_STYLE[h.status].weight + (h.hop === selected ? 2 : 0),
                    opacity: 0.95,
                    dashArray: dashFor(h),
                  }}
                  eventHandlers={{ click: () => setSelected(h.hop) }}
                />
              ))}
              {Array.from(visibleSites.entries()).map(([name, s]) => (
                <CircleMarker
                  key={name}
                  center={[s.lat, s.lon]}
                  radius={5}
                  pathOptions={{
                    color: selectedHop && (selectedHop.a.name === name || selectedHop.b.name === name) ? '#FFFFFF' : '#00A0B0',
                    fillColor: '#0B1220',
                    fillOpacity: 1,
                    weight: 2.5,
                  }}
                >
                  <Popup>
                    <div className="text-sm">
                      <div className="font-bold">{name}</div>
                      <div className="text-gray-600">{s.count} HOP{s.count === 1 ? '' : 's'} shown</div>
                    </div>
                  </Popup>
                </CircleMarker>
              ))}
            </MapContainer>

            {/* Legend */}
            <div className="absolute left-3 bottom-6 z-[1000] rounded-xl border border-white/15 bg-zinc-950/90 px-3 py-2 flex flex-col gap-1.5 max-h-[60%] overflow-auto">
              <span className="text-[11px] font-semibold text-zinc-400">{colorBy === 'status' ? 'HOP status' : 'General contractor'}</span>
              {colorBy === 'status'
                ? MAP_STATUS_ORDER.map(k => (
                    <div key={k} className="flex items-center gap-2 text-xs text-zinc-200">
                      <svg width="26" height="8" aria-hidden="true">
                        <line x1="0" y1="4" x2="26" y2="4" stroke={STATUS_STYLE[k].color} strokeWidth={STATUS_STYLE[k].weight} strokeDasharray={STATUS_STYLE[k].dash} />
                      </svg>
                      {STATUS_STYLE[k].label}
                    </div>
                  ))
                : gcList.map(g => (
                    <div key={g} className="flex items-center gap-2 text-xs text-zinc-200">
                      <svg width="26" height="8" aria-hidden="true">
                        <line x1="0" y1="4" x2="26" y2="4" stroke={gcColor.get(g)} strokeWidth={4} />
                      </svg>
                      {g}
                    </div>
                  ))}
            </div>
          </div>
        </section>

        {/* Side panel */}
        <aside className="w-full lg:w-96 shrink-0 rounded-2xl border border-white/10 bg-white/5 p-5 flex flex-col gap-4">
          {selectedHop ? (
            <>
              <div>
                <div className="text-xs text-zinc-400 font-semibold mb-1">Selected HOP</div>
                <div className="flex flex-wrap items-center gap-2">
                  <h2 className="text-lg font-bold text-white break-all">{selectedHop.hop}</h2>
                  <span
                    className="text-xs font-semibold px-2.5 py-1 rounded-full text-zinc-950"
                    style={{ background: STATUS_STYLE[selectedHop.status].color }}
                  >
                    {STATUS_STYLE[selectedHop.status].label}
                  </span>
                </div>
                <div className="text-xs text-zinc-400 mt-1">{selectedHop.a.name}  ↔  {selectedHop.b.name}</div>
              </div>

              {(() => {
                const a = alertFor(selectedHop, thresholds)
                return (
                  <div className="text-sm leading-snug rounded-lg bg-white/5 px-3 py-2.5 text-zinc-100" style={{ border: `1px solid ${a.color}` }}>
                    {a.text}
                  </div>
                )
              })()}

              <dl className="grid grid-cols-2 gap-3 text-sm">
                {[
                  ['Path ID', selectedHop.pathId || '—'],
                  ['Forecast start', selectedHop.ms15f ? `${selectedHop.ms15f}${selectedHop.daysOut !== null && !selectedHop.ms15a ? ` (${selectedHop.daysOut}d)` : ''}` : '—'],
                  ['Forecast end', selectedHop.ms16f || '—'],
                  ['Actual start', selectedHop.ms15a || '—'],
                  ['General contractor', selectedHop.gc || '—'],
                  ['Site CM', selectedHop.cm || '—'],
                ].map(([k, v]) => (
                  <div key={k}>
                    <dt className="text-[11px] text-zinc-400">{k}</dt>
                    <dd className="font-semibold text-zinc-100 break-words">{v}</dd>
                  </div>
                ))}
              </dl>

              <div className="rounded-xl border border-white/10 divide-y divide-white/10 overflow-hidden text-sm">
                {[
                  { label: 'NTP', ok: selectedHop.hasNtp, value: selectedHop.hasNtp ? 'Received' : (selectedHop.ntpWaitingOn ? `Waiting on ${selectedHop.ntpWaitingOn}` : 'Missing') },
                  { label: 'Material', ok: selectedHop.hasMat, value: selectedHop.hasMat ? 'Received' : 'Not received' },
                  ...selectedHop.serviceCpos.map(p => ({ label: `Service CPO · ${p.siteName || '—'}`, ok: !!p.number, value: p.number || 'Not yet received' })),
                  ...selectedHop.hardwarePos.map(p => ({ label: `Hardware PO · ${p.siteName || '—'}`, ok: !!p.number, value: p.number || 'Not yet issued' })),
                ].map((r, i) => (
                  <div key={`${r.label}-${i}`} className="flex items-center gap-3 px-3 py-2.5">
                    <span
                      className={`inline-flex w-5 h-5 shrink-0 items-center justify-center rounded-full text-xs font-bold ${r.ok ? 'bg-green-500/20 text-green-400' : 'bg-red-500/20 text-red-400'}`}
                      aria-label={r.ok ? 'done' : 'missing'}
                    >
                      {r.ok ? '✓' : '✗'}
                    </span>
                    <span className="flex-1 text-zinc-400">{r.label}</span>
                    <span className="font-semibold text-zinc-100 text-right break-all">{r.value}</span>
                  </div>
                ))}
              </div>
            </>
          ) : (
            <div className="text-sm text-zinc-400 leading-relaxed">
              Click a HOP line on the map to see its readiness, POs and what is blocking it.
            </div>
          )}
        </aside>
      </div>
    </div>
  )
}
