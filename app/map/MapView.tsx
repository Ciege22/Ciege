'use client'

import { useEffect, useMemo, useState } from 'react'
import { MapContainer, TileLayer, CircleMarker, Polyline, Popup, useMap } from 'react-leaflet'
import type { LatLngBoundsExpression } from 'leaflet'
import 'leaflet/dist/leaflet.css'
import { loadTrackerSnapshot } from '../lib/supabase'

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
  useEffect(() => {
    if (bounds) map.fitBounds(bounds, { padding: [30, 30] })
  }, [bounds, map])
  return null
}

export default function MapView() {
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)
  const [sites, setSites] = useState<SitePoint[]>([])
  const [connections, setConnections] = useState<HopConnection[]>([])
  const [trackerInfo, setTrackerInfo] = useState<{ filename: string; uploaded_at: string } | null>(null)

  useEffect(() => {
    let cancelled = false
    loadTrackerSnapshot().then(snap => {
      if (cancelled) return
      if (!snap) { setError('No tracker found — upload one on the Dashboard first.'); setLoading(false); return }
      setTrackerInfo({ filename: snap.filename, uploaded_at: snap.uploaded_at })

      const rows = snap.data
      let headerRowIdx = -1
      for (let i = 0; i < 10; i++) {
        const row = rows[i] as unknown[]
        if (row && row.some(c => String(c ?? '').trim() === 'HOP')) { headerRowIdx = i; break }
      }
      if (headerRowIdx === -1) { setError('Could not find a HOP header row in the tracker.'); setLoading(false); return }

      const headers = (rows[headerRowIdx] as unknown[]).map(h => String(h ?? '').trim())
      const col = (name: string) => headers.findIndex(h => h === name)
      const hopCol = col('HOP')
      const siteNameCol = col('Site Name')
      const latCol = col('Latt.')
      const lonCol = col('Long.')
      const gcCol = col('General Contractor')
      const don444Col = col('DON 444')

      if (latCol === -1 || lonCol === -1) {
        setError('This tracker export has no Latt./Long. columns.')
        setLoading(false)
        return
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

      if (!cancelled) {
        setSites(siteList)
        setConnections(conns)
        setLoading(false)
      }
    }).catch(err => {
      if (!cancelled) { setError(err instanceof Error ? err.message : 'Failed to load tracker.'); setLoading(false) }
    })
    return () => { cancelled = true }
  }, [])

  const bounds = useMemo<LatLngBoundsExpression | null>(() => {
    if (sites.length === 0) return null
    const lats = sites.map(s => s.lat)
    const lons = sites.map(s => s.lon)
    return [[Math.min(...lats), Math.min(...lons)], [Math.max(...lats), Math.max(...lons)]]
  }, [sites])

  if (loading) {
    return <div className="flex items-center justify-center h-[70vh] text-zinc-400 text-sm">Loading tracker data…</div>
  }
  if (error) {
    return (
      <div className="rounded-2xl border border-red-500/20 bg-red-500/10 px-4 py-3 text-sm text-red-300">
        {error}
      </div>
    )
  }

  return (
    <div>
      <div className="flex flex-wrap items-center gap-3 mb-3 text-xs text-zinc-400">
        <span>{connections.length} HOPs</span>
        <span>·</span>
        <span>{sites.length} sites</span>
        {trackerInfo && (
          <>
            <span>·</span>
            <span>{trackerInfo.filename} ({new Date(trackerInfo.uploaded_at).toLocaleDateString('en-US')})</span>
          </>
        )}
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
