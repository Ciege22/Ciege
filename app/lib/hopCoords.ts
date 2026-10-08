// HOP-level coordinates and distance clustering, for plotting GC Clean-Up
// HOPs on a map and grouping ones close enough together to hand one GC a
// single bundle of work instead of scattering it across the territory.
import { normalizeTrackerHop } from './grTracker'

export interface HopCoord {
  lat: number
  lon: number
}

function isValidCoord(lat: unknown, lon: unknown): lat is number {
  const latN = Number(lat)
  const lonN = Number(lon)
  return Number.isFinite(latN) && Number.isFinite(lonN) && latN !== 0 && lonN !== 0
    && Math.abs(latN) <= 90 && Math.abs(lonN) <= 180
}

// Same normalization grTracker.ts uses to key the SPO report's HOP name
// against the tracker's '<>'-separated one — matchKey() there is just
// trim+lowercase, inlined here rather than exporting a private helper.
function matchKey(s: string): string {
  return s.trim().toLowerCase()
}

// One point per HOP — the midpoint of its one or two site coordinates. A
// GC crew works both ends of a HOP anyway, so the midpoint is close enough
// for "is this worth bundling with that other HOP" territory decisions;
// it is not meant to be a precise site location.
export function buildHopCoordMap(rows: unknown[][]): Map<string, HopCoord> {
  const out = new Map<string, HopCoord>()
  let headerRow = -1
  for (let i = 0; i < 10; i++) {
    if ((rows[i] as unknown[])?.some(c => String(c).trim() === 'HOP')) { headerRow = i; break }
  }
  if (headerRow === -1) return out

  const headers = (rows[headerRow] as unknown[]).map(h => String(h ?? '').trim())
  const col = (name: string) => headers.findIndex(h => h === name)
  const hopCol = col('HOP')
  const siteNameCol = col('Site Name')
  const latCol = col('Latt.')
  const lonCol = col('Long.')
  const don444Col = col('DON 444')
  if (hopCol === -1 || latCol === -1 || lonCol === -1) return out

  const byHop = new Map<string, unknown[][]>()
  for (let i = headerRow + 1; i < rows.length; i++) {
    const row = rows[i] as unknown[]
    if (!row) continue
    if (String(row[don444Col] ?? '').trim().toUpperCase() !== 'DON 444') continue
    const hop = String(row[hopCol] ?? '').trim()
    if (!hop || hop === 'undefined') continue
    if (!byHop.has(hop)) byHop.set(hop, [])
    byHop.get(hop)!.push(row)
  }

  byHop.forEach((hopRows, rawHop) => {
    const siteCoords = new Map<string, HopCoord>()
    hopRows.forEach(r => {
      const site = String(r[siteNameCol] ?? '').trim()
      const lat = Number(r[latCol])
      const lon = Number(r[lonCol])
      if (site && isValidCoord(lat, lon)) siteCoords.set(site, { lat, lon })
    })
    const pts = Array.from(siteCoords.values())
    if (pts.length === 0) return
    const lat = pts.reduce((s, p) => s + p.lat, 0) / pts.length
    const lon = pts.reduce((s, p) => s + p.lon, 0) / pts.length
    out.set(matchKey(normalizeTrackerHop(rawHop)), { lat, lon })
  })

  return out
}

const EARTH_RADIUS_MILES = 3958.8

export function haversineMiles(a: HopCoord, b: HopCoord): number {
  const toRad = (d: number) => (d * Math.PI) / 180
  const dLat = toRad(b.lat - a.lat)
  const dLon = toRad(b.lon - a.lon)
  const la1 = toRad(a.lat)
  const la2 = toRad(b.lat)
  const h = Math.sin(dLat / 2) ** 2 + Math.cos(la1) * Math.cos(la2) * Math.sin(dLon / 2) ** 2
  return 2 * EARTH_RADIUS_MILES * Math.asin(Math.sqrt(h))
}

export interface GeoCluster<T> {
  id: number
  center: HopCoord
  members: T[]
}

// Greedy radius clustering, not a true optimal grouping: seed a cluster with
// one unclustered point, then keep pulling in any remaining point within
// radiusMiles of the cluster's running centroid (recomputed as it grows)
// until nothing more qualifies. Good enough to turn "HOPs near each other"
// into discrete, explainable bundles to hand a GC — not a claim that it's
// the minimum number of clusters or the tightest possible grouping.
export function clusterByDistance<T>(points: { item: T; coord: HopCoord }[], radiusMiles: number): GeoCluster<T>[] {
  const remaining = points.slice()
  const clusters: GeoCluster<T>[] = []

  while (remaining.length > 0) {
    const members = [remaining.shift()!]
    let changed = true
    while (changed) {
      changed = false
      const centerLat = members.reduce((s, m) => s + m.coord.lat, 0) / members.length
      const centerLon = members.reduce((s, m) => s + m.coord.lon, 0) / members.length
      for (let i = remaining.length - 1; i >= 0; i--) {
        if (haversineMiles({ lat: centerLat, lon: centerLon }, remaining[i].coord) <= radiusMiles) {
          members.push(remaining[i])
          remaining.splice(i, 1)
          changed = true
        }
      }
    }
    const center = {
      lat: members.reduce((s, m) => s + m.coord.lat, 0) / members.length,
      lon: members.reduce((s, m) => s + m.coord.lon, 0) / members.length,
    }
    clusters.push({ id: clusters.length + 1, center, members: members.map(m => m.item) })
  }

  clusters.sort((a, b) => b.members.length - a.members.length)
  clusters.forEach((c, i) => { c.id = i + 1 })
  return clusters
}
