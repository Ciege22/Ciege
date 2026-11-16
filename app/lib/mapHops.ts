// HOP Network Map — turns the tracker rows into one record per HOP with a
// readiness status, so the map can color, filter and count them. Pure and
// framework-free (no React, no Supabase) so it can be tested on its own.
//
// Every definition below deliberately mirrors the Dashboard's computeKPIs
// (app/page.tsx) so a number on the map never disagrees with the same tile:
//   NTP received   = 'NTP A' is a date in 2025 or later
//   Material in    = 'Material Received A' is a date in 2020 or later, or
//                    the HOP is already complete
//   Started        = 'MS15 Implementation Start A' has a date
//   Complete       = 'MS16 Implementation Ends A' has a date
//   Days out       = whole days from today to 'MS15 Implementation Start F'
// and the HOP's row for those fields is picked the same way (the row that has
// both a GC and a CM, else the first row), because a HOP spans two site rows
// and one of them can be a blank duplicate.
import type { ThresholdSettings } from './settings'

export type MapStatus = 'ready' | 'progress' | 'ntp' | 'mat' | 'blocked' | 'complete'

export const MAP_STATUS_ORDER: MapStatus[] = ['ready', 'progress', 'ntp', 'mat', 'blocked', 'complete']

export interface MapSite {
  name: string
  lat: number
  lon: number
}

export interface SitePo {
  siteName: string
  number: string // blank when nothing is on file for this site
}

export interface MapHop {
  hop: string
  pathId: string
  gc: string
  cm: string
  a: MapSite
  b: MapSite
  status: MapStatus
  hasNtp: boolean
  hasMat: boolean
  ms15f: string
  ms15a: string
  ms16f: string
  ms16a: string
  daysOut: number | null
  daysElapsed: number | null
  ntpWaitingOn: string
  serviceCpos: SitePo[]
  hardwarePos: SitePo[]
  // Same triggers as the Dashboard's NTP Urgent / Material Watch tiles.
  ntpUrgent: boolean
  materialWatch: boolean
  over18d: boolean
  // Forecast start inside the next 14 days and not started yet.
  startingSoon: boolean
}

export interface ParsedMapHops {
  hops: MapHop[]
  // HOPs in the tracker that could not be drawn (a site has no usable
  // latitude/longitude), so the page can say so instead of silently dropping them.
  unmapped: number
  error: string | null
}

function isValidCoord(lat: unknown, lon: unknown): boolean {
  const latN = Number(lat)
  const lonN = Number(lon)
  return Number.isFinite(latN) && Number.isFinite(lonN) && latN !== 0 && lonN !== 0
    && Math.abs(latN) <= 90 && Math.abs(lonN) <= 180
}

// Same blank/epoch-zero guard the Dashboard's parseD uses: a blank Excel date
// can round-trip as 1899/1970 instead of null, so anything before 1990 is "no date".
function parseD(val: unknown): Date | null {
  if (!val) return null
  if (val instanceof Date) return isNaN(val.getTime()) || val.getFullYear() < 1990 ? null : val
  if (typeof val === 'number') {
    if (val > 40000 && val < 60000) {
      const d = new Date((val - 25569) * 86400 * 1000)
      return isNaN(d.getTime()) || d.getFullYear() < 1990 ? null : d
    }
    return null
  }
  const s = String(val).trim()
  if (!s || s === 'null' || s === 'undefined' || s === 'NaN') return null
  const d = new Date(s)
  if (!isNaN(d.getTime())) return d.getFullYear() < 1990 ? null : d
  const parts = s.split('/')
  if (parts.length === 3) {
    const d2 = new Date(parseInt(parts[2]), parseInt(parts[0]) - 1, parseInt(parts[1]))
    return isNaN(d2.getTime()) || d2.getFullYear() < 1990 ? null : d2
  }
  return null
}

function fmt(d: Date | null): string {
  return d ? `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}` : ''
}

function cleanText(v: unknown): string {
  const s = String(v ?? '').trim()
  return s.toLowerCase() === 'nan' || s === 'undefined' || s === 'null' ? '' : s
}

export function parseMapHops(
  rows: unknown[][],
  thresholds: Pick<ThresholdSettings, 'ntpUrgentDays' | 'materialWatchDays' | 'durationAlertDays'>,
  now: Date = new Date(),
): ParsedMapHops {
  let headerRowIdx = -1
  for (let i = 0; i < 10; i++) {
    const row = rows[i] as unknown[]
    if (row && row.some(c => String(c ?? '').trim() === 'HOP')) { headerRowIdx = i; break }
  }
  if (headerRowIdx === -1) return { hops: [], unmapped: 0, error: 'Could not find a HOP header row in the tracker.' }

  const headers = (rows[headerRowIdx] as unknown[]).map(h => String(h ?? '').trim())
  const col = (name: string) => headers.findIndex(h => h === name)
  const hopCol = col('HOP')
  const pathIdCol = col('Path ID')
  const siteNameCol = col('Site Name')
  const latCol = col('Latt.')
  const lonCol = col('Long.')
  const gcCol = col('General Contractor')
  const cmCol = col('New CM')
  const don444Col = col('DON 444')
  const ms15fCol = col('MS15 Implementation Start F')
  const ms15aCol = col('MS15 Implementation Start A')
  const ms16fCol = col('MS16 Implementation Ends F')
  const ms16aCol = col('MS16 Implementation Ends A')
  const ntpCol = col('NTP A')
  const ntpWaitCol = col('NTP is waiting on')
  const matCol = headers.findIndex(h => h.replace(/\s+/g, ' ').startsWith('Material Received A'))
  const serviceCpoCol = col('Service CPO Received')
  const hardwareCpoCol = col('Hardware CPO Received')

  if (latCol === -1 || lonCol === -1) {
    return { hops: [], unmapped: 0, error: 'This tracker export has no Latt./Long. columns.' }
  }

  const today = new Date(now)
  today.setHours(0, 0, 0, 0)

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

  const hops: MapHop[] = []
  let unmapped = 0

  byHop.forEach((hopRows, hop) => {
    const withCoords = hopRows
      .map(r => ({
        name: String(r[siteNameCol] ?? '').trim(),
        lat: Number(r[latCol]),
        lon: Number(r[lonCol]),
      }))
      .filter(s => s.name && isValidCoord(s.lat, s.lon))
    // One entry per distinct site (a HOP's two rows are normally two different sites).
    const uniqueSites = Array.from(new Map(withCoords.map(s => [s.name, s])).values())
    if (uniqueSites.length < 2) { unmapped++; return }
    const [a, b] = uniqueSites

    const row = (hopRows.find(r => String(r[gcCol] ?? '').trim() && String(r[cmCol] ?? '').trim()) ?? hopRows[0]) as unknown[]

    const ntpDate = parseD(row[ntpCol])
    const matDate = parseD(row[matCol])
    const ms15f = parseD(row[ms15fCol])
    const ms15a = parseD(row[ms15aCol])
    const ms16f = parseD(row[ms16fCol])
    const ms16a = parseD(row[ms16aCol])

    const hasNtp = !!(ntpDate && ntpDate.getFullYear() >= 2025)
    const hasMat = !!(matDate && matDate.getFullYear() >= 2020) || !!ms16a
    const started = !!ms15a
    const complete = !!ms16a
    const inProgress = started && !complete
    const daysOut = ms15f ? Math.round((ms15f.getTime() - today.getTime()) / 86400000) : null
    const daysElapsed = inProgress && ms15a ? Math.round((today.getTime() - ms15a.getTime()) / 86400000) : null

    let status: MapStatus
    if (complete) status = 'complete'
    else if (inProgress) status = 'progress'
    else if (hasNtp && hasMat) status = 'ready'
    else if (hasNtp) status = 'mat' // NTP is in, material is the gap
    else if (hasMat) status = 'ntp' // material is in, NTP is the gap
    else status = 'blocked'

    const perSite = (colIdx: number): SitePo[] =>
      colIdx === -1 ? [] : hopRows.map(r => ({
        siteName: String(r[siteNameCol] ?? '').trim(),
        number: cleanText(r[colIdx]),
      }))

    hops.push({
      hop,
      pathId: cleanText(row[pathIdCol]),
      gc: cleanText(row[gcCol]),
      cm: cleanText(row[cmCol]),
      a: { name: a.name, lat: a.lat, lon: a.lon },
      b: { name: b.name, lat: b.lat, lon: b.lon },
      status,
      hasNtp,
      hasMat,
      ms15f: fmt(ms15f),
      ms15a: fmt(ms15a),
      ms16f: fmt(ms16f),
      ms16a: fmt(ms16a),
      daysOut,
      daysElapsed,
      ntpWaitingOn: cleanText(row[ntpWaitCol]),
      serviceCpos: perSite(serviceCpoCol),
      hardwarePos: perSite(hardwareCpoCol),
      ntpUrgent: !hasNtp && !complete && daysOut !== null && daysOut <= thresholds.ntpUrgentDays,
      materialWatch: !hasMat && !complete && daysOut !== null && daysOut <= thresholds.materialWatchDays,
      over18d: inProgress && (daysElapsed ?? 0) > thresholds.durationAlertDays,
      startingSoon: !started && !complete && daysOut !== null && daysOut <= 14,
    })
  })

  return { hops, unmapped, error: null }
}
