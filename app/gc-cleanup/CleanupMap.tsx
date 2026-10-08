'use client'

import { Fragment, useEffect, useMemo } from 'react'
import { MapContainer, TileLayer, CircleMarker, Circle, Popup, useMap } from 'react-leaflet'
import type { LatLngBoundsExpression } from 'leaflet'
import 'leaflet/dist/leaflet.css'
import { fmtMoney } from '../lib/grTracker'
import type { CleanupHop } from '../lib/gcCleanup'
import type { GeoCluster } from '../lib/hopCoords'

const FALLBACK_CENTER: [number, number] = [39.8, -101.3]
const MILES_TO_METERS = 1609.34

// Cycled by cluster index — enough distinct hues to tell adjacent clusters
// apart at a glance without reading the legend every time.
const PALETTE = ['#00A0B0', '#E67E22', '#9B59B6', '#2ECC71', '#E74C3C', '#F1C40F', '#3498DB', '#1ABC9C', '#D35400', '#8E44AD']

function FitBounds({ bounds }: { bounds: LatLngBoundsExpression | null }) {
  const map = useMap()
  useEffect(() => {
    if (bounds) map.fitBounds(bounds, { padding: [30, 30] })
  }, [bounds, map])
  return null
}

export default function CleanupMap({ clusters, radiusMiles }: { clusters: GeoCluster<CleanupHop>[]; radiusMiles: number }) {
  const bounds = useMemo<LatLngBoundsExpression | null>(() => {
    const pts = clusters.flatMap(c => c.members.map(h => h.coord).filter((c): c is NonNullable<typeof c> => !!c))
    if (pts.length === 0) return null
    const lats = pts.map(p => p.lat)
    const lons = pts.map(p => p.lon)
    return [[Math.min(...lats), Math.min(...lons)], [Math.max(...lats), Math.max(...lons)]]
  }, [clusters])

  return (
    <MapContainer center={FALLBACK_CENTER} zoom={6} style={{ height: '60vh', width: '100%', borderRadius: '0.75rem' }}>
      <TileLayer attribution='&copy; OpenStreetMap contributors' url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png" />
      <FitBounds bounds={bounds} />
      {clusters.map((cluster, ci) => {
        const color = PALETTE[ci % PALETTE.length]
        return (
          <Fragment key={cluster.id}>
            {cluster.members.length > 1 && (
              <Circle center={[cluster.center.lat, cluster.center.lon]} radius={radiusMiles * MILES_TO_METERS}
                pathOptions={{ color, fillColor: color, fillOpacity: 0.05, weight: 1, dashArray: '4 4' }} />
            )}
            {cluster.members.map(hop => hop.coord && (
              <CircleMarker key={hop.hop} center={[hop.coord.lat, hop.coord.lon]} radius={7}
                pathOptions={{ color: '#0a0a0a', weight: 1, fillColor: color, fillOpacity: 0.9 }}>
                <Popup>
                  <div className="text-xs">
                    <p className="font-semibold">{hop.hopDisplay}</p>
                    <p>Cluster {cluster.id} ({cluster.members.length} HOP{cluster.members.length === 1 ? '' : 's'})</p>
                    <p>Original GC: {hop.gc}</p>
                    <p>Unpaid: {fmtMoney(hop.unpaidValue)}</p>
                  </div>
                </Popup>
              </CircleMarker>
            ))}
          </Fragment>
        )
      })}
    </MapContainer>
  )
}
