'use client'

import { useMemo } from 'react'
import { MapContainer, TileLayer, CircleMarker, Polyline, Popup, useMap } from 'react-leaflet'
import type { LatLngBoundsExpression } from 'leaflet'
import 'leaflet/dist/leaflet.css'
import { useIsAuthenticated } from '@azure/msal-react'
import { useSharePointFile } from '../components/SharePointDataProvider'
import SharePointSourceLine from '../components/SharePointSourceLine'
import MicrosoftLogin from '../components/MicrosoftLogin'

// Stage 2 Part B of the SharePoint migration — first real page moved off
// Supabase (loadTrackerSnapshot) onto the shared SharePointDataProvider.
// The Dashboard's own upload-to-Supabase path (app/page.tsx) is completely
// untouched — this page just no longer reads from what that path saves.

const NAVY = '#124191'
const TEAL = '#00A0B0'

// Fallback center (roughly the middle of the CO/NE/KS service area) — only
// used before data loads or if the tracker has zero valid coordinates.
const FALLBACK_CENTER: [number, number] = [39.8, -101.3]

interface SitePoint {
  name: string
  lat: number
  lon: number
  hopCount: number
}

interface HopConnection {
  hop: string
  gc: string
  siteA: string
  siteB: string
  a: [number, number]
  b: [number, number]
}

function isValidCoord(lat: unknown, lon: unknown): lat is number {
  const latN = Number(lat)
  const lonN = Number(lon)
  return Number.isFinite(latN) && Number.isFinite(lonN) && latN !== 0 && lonN !== 0
    && Math.abs(latN) <= 90 && Math.abs(lonN) <= 180
}

// Fits the map to the data's bounding box once it's available — MapContainer
// itself mounts before the tracker finishes loading, so this has to happen
// as a child effect via useMap() rather than a static prop.
function FitBounds({ bounds }: { bounds: LatLngBoundsExpression | null }) {
  const map = useMap()
  useMemo(() => {
    if (bounds) map.fitBounds(bounds, { padding: [30, 30] })
  }, [bounds, map])
  return null
}

// Same header-scan + column-resolve + byHop grouping every other tracker
// consumer in this app uses — unchanged from before this migration, just
// now operating on SharePointDataProvider's rows instead of a Supabase
// snapshot's .data (the two are the same shape by construction, see
// app/lib/sharepointParsers.ts).
function parseTrackerRows(rows: unknown[][]): { sites: SitePoint[]; connections: HopConnection[]; error: string | null } {
  let headerRowIdx = -1
  for (let i = 0; i < 10; i++) {
    const row = rows[i] as unknown[]
    if (row && row.some(c => String(c ?? '').trim() === 'HOP')) { headerRowIdx = i; break }
  }
  if (headerRowIdx === -1) return { sites: [], connections: [], error: 'Could not find a HOP header row in the tracker.' }

  const headers = (rows[headerRowIdx] as unknown[]).map(h => String(h ?? '').trim())
  const col = (name: string) => headers.findIndex(h => h === name)
  const hopCol = col('HOP')
  const siteNameCol = col('Site Name')
  const latCol = col('Latt.')
  const lonCol = col('Long.')
  const gcCol = col('General Contractor')
  const don444Col = col('DON 444')

  if (latCol === -1 || lonCol === -1) {
    return { sites: [], connections: [], error: 'This tracker export has no Latt./Long. columns.' }
  }

  // hop -> its site rows (a HOP spans two physical site rows, same
  // convention as every other page that reads this tracker)
  const byHop = new Map<string, unknown[][]>()
  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i] as unknown[]
    if (!row) continue
    if (String(row[don444Col] ?? '').trim().toUpperCase() !== 'DON 444') continue
    const hop = String(row[hopCol] ?? '').trim()
    if (!hop || hop === 'undefined') continue
    if (!byHop.has(hop)) byHop.set(hop, [])
    byHop.get(hop)!.push(row)
  }

  const siteCoords = new Map<string, { lat: number; lon: number; hopCount: number }>()
  const conns: HopConnection[] = []

  byHop.forEach((hopRows, hop) => {
    const withCoords = hopRows
      .map(r => ({
        site: String(r[siteNameCol] ?? '').trim(),
        lat: Number(r[latCol]),
        lon: Number(r[lonCol]),
      }))
      .filter(r => r.site && isValidCoord(r.lat, r.lon))
    // Dedupe to one entry per distinct site within this HOP (a HOP's two
    // rows are normally two different sites, but guard against dupes).
    const uniqueSites = Array.from(new Map(withCoords.map(s => [s.site, s])).values())
    if (uniqueSites.length < 2) return // skip HOPs missing a coordinate on either end

    const [siteA, siteB] = uniqueSites
    for (const s of uniqueSites) {
      if (!siteCoords.has(s.site)) siteCoords.set(s.site, { lat: s.lat, lon: s.lon, hopCount: 0 })
      siteCoords.get(s.site)!.hopCount++
    }
    const gc = String(hopRows.find(r => String(r[gcCol] ?? '').trim())?.[gcCol] ?? '').trim()
    conns.push({
      hop, gc, siteA: siteA.site, siteB: siteB.site,
      a: [siteA.lat, siteA.lon], b: [siteB.lat, siteB.lon],
    })
  })

  const siteList: SitePoint[] = Array.from(siteCoords.entries()).map(([name, v]) => ({
    name, lat: v.lat, lon: v.lon, hopCount: v.hopCount,
  }))

  return { sites: siteList, connections: conns, error: null }
}

const PHASE_LABEL: Record<string, string> = {
  downloading: 'Downloading tracker from SharePoint…',
  parsing: 'Parsing tracker…',
  building: 'Building map…',
}

export default function MapView() {
  const isAuthenticated = useIsAuthenticated()
  const { rows, loading, phase, error: loadError } = useSharePointFile('tracker')

  const parsed = useMemo(() => (rows ? parseTrackerRows(rows) : null), [rows])
  const sites = parsed?.sites ?? []
  const connections = parsed?.connections ?? []
  const parseError = parsed?.error ?? null

  const bounds = useMemo<LatLngBoundsExpression | null>(() => {
    if (!parsed || parsed.sites.length === 0) return null
    const lats = parsed.sites.map(s => s.lat)
    const lons = parsed.sites.map(s => s.lon)
    return [[Math.min(...lats), Math.min(...lons)], [Math.max(...lats), Math.max(...lons)]]
  }, [parsed])

  if (!isAuthenticated) {
    return (
      <div className="rounded-2xl border border-white/10 bg-white/5 p-6">
        <p className="text-sm text-zinc-400 mb-3">Sign in with your Nokia account to load the tracker from SharePoint.</p>
        <MicrosoftLogin />
      </div>
    )
  }
  // Only a full-page block when there's nothing to show yet — a cached
  // entry (stale-while-revalidate) renders the map immediately instead,
  // with SharePointSourceLine's "refreshing…" indicator covering the
  // background fetch instead of blocking the page on it.
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
  if (parseError) {
    return (
      <div className="rounded-2xl border border-red-500/20 bg-red-500/10 px-4 py-3 text-sm text-red-300">
        {parseError}
      </div>
    )
  }

  return (
    <div>
      <div className="flex flex-wrap items-center gap-3 mb-1 text-xs text-zinc-400">
        <span>{connections.length} HOPs</span>
        <span>·</span>
        <span>{sites.length} sites</span>
      </div>
      <div className="mb-3">
        <SharePointSourceLine fileKey="tracker" />
      </div>
      <div className="rounded-2xl overflow-hidden border border-white/10" style={{ height: '75vh' }}>
        <MapContainer center={FALLBACK_CENTER} zoom={7} style={{ height: '100%', width: '100%' }}>
          <TileLayer
            attribution='&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
            url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png"
          />
          <FitBounds bounds={bounds} />
          {connections.map(c => (
            <Polyline key={c.hop} positions={[c.a, c.b]} pathOptions={{ color: NAVY, weight: 2, opacity: 0.75 }}>
              <Popup>
                <div className="text-sm">
                  <div className="font-bold">{c.hop}</div>
                  {c.gc && <div className="text-gray-600">GC: {c.gc}</div>}
                </div>
              </Popup>
            </Polyline>
          ))}
          {sites.map(s => (
            <CircleMarker
              key={s.name}
              center={[s.lat, s.lon]}
              radius={5}
              pathOptions={{ color: TEAL, fillColor: TEAL, fillOpacity: 0.9, weight: 1.5 }}
            >
              <Popup>
                <div className="text-sm">
                  <div className="font-bold">{s.name}</div>
                  <div className="text-gray-600">{s.hopCount} connected HOP{s.hopCount === 1 ? '' : 's'}</div>
                </div>
              </Popup>
            </CircleMarker>
          ))}
        </MapContainer>
      </div>
    </div>
  )
}
