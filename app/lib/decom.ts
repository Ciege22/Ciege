import { GC_CONFIG, matches } from './gcConfig'
import { lookupContactEmail } from './settings'

export interface DecomRow {
  // Unique per site row — Path ID (falling back to Site Name + Site Number
  // when Path ID is blank). HOP is NOT unique per row (a HOP can cover
  // multiple sites), so it must never be used as a key or dedup identifier.
  rowKey: string
  appPathId: string
  hop: string
  pathId: string
  siteName: string
  siteNumber: string
  cxStart: Date | null
  cxComplete: Date | null
  dropOffDate: Date | null
  comment: string
  cm: string
  gc: string
  podPathwave: boolean
  podQuickBase: boolean
  aging: number | null
  status: 'complete' | 'pod_gap' | 'outstanding' | 'pending'
}

export interface DecomGcSummary {
  gc: string
  total: number
  complete: number
  // Total pod_gap status count (Pathwave pending OR QuickBase pending) —
  // pendingPathwave + pendingQuickBase always sum to this.
  podGap: number
  outstanding: number
  pending: number
  // Drop Off done, Pathwave = No/blank (regardless of QuickBase).
  pendingPathwave: number
  // Pathwave done, QuickBase = No/blank. Mutually exclusive with
  // pendingPathwave — together they partition podGap with no overlap.
  pendingQuickBase: number
  // Completed HOPs from the master tracker with no matching decom entry —
  // not part of `total` (which only counts rows that exist in the decom
  // file). total + missing = the full decom-eligible universe for this GC.
  missing: number
  avgAging: number | null
}

export interface TrackerHop {
  hop: string
  pathId: string
  siteName: string
  siteNumber: string
  gc: string
  cm: string
  nokiaPm: string
  ms16a: Date
}

export interface MissingDecomSite {
  hop: string
  pathId: string
  siteName: string
  gc: string
  cm: string
  nokiaPm: string
  ms16a: Date
  daysElapsed: number
}

export const STATUS_DISPLAY_LABEL: Record<DecomRow['status'], string> = {
  complete: 'Complete',
  pod_gap: 'Pending POD in Pathwave',
  outstanding: 'Pending Decom Drop Off',
  pending: 'Pending Decom Drop Off',
}

function findColCaseInsensitive(headers: unknown[], name: string): number {
  const target = name.trim().toLowerCase()
  return headers.findIndex(h => String(h).trim().toLowerCase() === target)
}

function parseDate(val: unknown): Date | null {
  if (!val) return null
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val
  if (typeof val === 'number') {
    const d = new Date((val - 25569) * 86400 * 1000)
    return isNaN(d.getTime()) ? null : d
  }
  const d = new Date(String(val))
  return isNaN(d.getTime()) ? null : d
}

function yearOk(d: Date | null): d is Date {
  return !!d && d.getFullYear() >= 2020
}

function daysBetween(a: Date, b: Date): number {
  return Math.round((b.getTime() - a.getTime()) / (1000 * 60 * 60 * 24))
}

// DECOM Drop Off cells mix a date with free-text notes in the same string
// (e.g. "8/15/2025 - waiting on GC"). Pull the date out; whatever's left,
// trimmed of separator punctuation, becomes the comment. Native Excel date
// cells (no text) come through as a pure date with an empty comment.
const DATE_PATTERNS = [
  /\d{1,2}\/\d{1,2}\/\d{2,4}/,
  /\d{4}-\d{1,2}-\d{1,2}/,
]

export function parseDropOff(raw: unknown): { date: Date | null; comment: string } {
  if (raw === null || raw === undefined) return { date: null, comment: '' }
  if (raw instanceof Date) {
    return isNaN(raw.getTime()) ? { date: null, comment: '' } : { date: raw, comment: '' }
  }
  const str = String(raw).trim()
  if (!str) return { date: null, comment: '' }

  for (const pattern of DATE_PATTERNS) {
    const match = str.match(pattern)
    if (match) {
      const d = new Date(match[0])
      if (!isNaN(d.getTime())) {
        const comment = str.replace(match[0], '').replace(/^[\s\-–—:,.]+|[\s\-–—:,.]+$/g, '').trim()
        return { date: d, comment }
      }
    }
  }
  // No date found in the cell — treat the whole string as a comment; the
  // row is still "outstanding" since no drop-off date was ever extracted.
  return { date: null, comment: str }
}

function parseYesNo(val: unknown): boolean {
  const s = String(val || '').trim().toLowerCase()
  return s === 'yes' || s === 'y' || s === 'true'
}

// GC matching reuses the existing GC_CONFIG alias table (same one SPO/CR
// vendor matching uses) — tries the GC column first (either a direct short
// name like "Vikor" or a known vendor alias like "Sioux Falls Tower
// Specialists Inc."), then falls back to the CG column the same way.
export function resolveGc(gcRaw: string, cgRaw: string): string {
  const tryMatch = (raw: string): string => {
    const trimmed = (raw || '').trim()
    if (!trimmed) return ''
    const lower = trimmed.toLowerCase()
    const direct = GC_CONFIG.find(cfg => cfg.gc.toLowerCase() === lower)
    if (direct) return direct.gc
    const viaAlias = GC_CONFIG.find(cfg => matches(raw, cfg.spo_match) || matches(raw, cfg.cr_match))
    return viaAlias ? viaAlias.gc : ''
  }
  return tryMatch(gcRaw) || tryMatch(cgRaw) || gcRaw.trim() || cgRaw.trim()
}

export function fmtDecomDate(d: Date | null): string {
  if (!d) return ''
  return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`
}

// Some sites get physically dropped off at the warehouse before their CX
// Complete date gets logged in the tracker — the two fields are updated
// independently, not in a guaranteed order. parseDecomRows excludes any row
// without CX Complete (it isn't "decom-eligible" yet), which also silently
// excludes a real Drop Off date on that row from every count. This counts
// those specifically — rows with a valid Drop Off date but no CX Complete —
// so the Decom Funnel's "Dropped Off" stage can include them without
// changing what counts as decom-eligible everywhere else in the app.
export function countDroppedOffWithoutCxComplete(allRows: unknown[][]): number {
  if (allRows.length === 0) return 0
  const headers = allRows[0] as unknown[]
  const col = (name: string) => findColCaseInsensitive(headers, name)
  const hopCol = col('HOP')
  const cxCompleteCol = col('CX Complete - MS16')
  const dropOffCol = col('DECOM Drop Off')

  let count = 0
  for (let i = 1; i < allRows.length; i++) {
    const row = allRows[i]
    if (!row || !row.some(v => v !== null && v !== undefined && String(v).trim() !== '')) continue
    const hop = String(row[hopCol] ?? '').trim()
    if (!hop) continue

    const cxCompleteRaw = parseDate(row[cxCompleteCol])
    const cxComplete = yearOk(cxCompleteRaw) ? cxCompleteRaw : null
    if (cxComplete) continue // already counted normally via parseDecomRows

    const dropOffParsed = parseDropOff(row[dropOffCol])
    const dropOffDate = yearOk(dropOffParsed.date) ? dropOffParsed.date : null
    if (dropOffDate) count++
  }
  return count
}

// `allRows` is the raw chunked-and-reassembled sheet: row 0 is the header
// row, everything after is data. Header matching is case-insensitive per
// spec, since the source sheet's casing isn't guaranteed to match exactly.
export function parseDecomRows(allRows: unknown[][]): DecomRow[] {
  if (allRows.length === 0) return []
  const headers = allRows[0] as unknown[]
  const col = (name: string) => findColCaseInsensitive(headers, name)

  const appPathIdCol = col('App Path ID')
  const hopCol = col('HOP')
  const pathIdCol = col('Path ID')
  const siteNameCol = col('Site Name')
  const siteNumberCol = col('Site Number')
  const cxStartCol = col('CX Start - MS15')
  const cxCompleteCol = col('CX Complete - MS16')
  const dropOffCol = col('DECOM Drop Off')
  const cmCol = col('CM')
  const gcCol = col('GC')
  const cgCol = col('CG')
  const podPathwaveCol = col('POD In Pathwave')
  const podQuickBaseCol = col('POD In QuickBase')

  console.log('[decom-parse] header row:', headers)
  console.log('[decom-parse] resolved columns:', {
    appPathIdCol, hopCol, pathIdCol, siteNameCol, siteNumberCol,
    cxStartCol, cxCompleteCol, dropOffCol, cmCol, gcCol, cgCol,
    podPathwaveCol, podQuickBaseCol,
  })

  const today = new Date()
  const rows: DecomRow[] = []
  let blankRowSkipped = 0
  let blankHopSkipped = 0
  let noCxCompleteSkipped = 0

  for (let i = 1; i < allRows.length; i++) {
    const row = allRows[i]
    if (!row || !row.some(v => v !== null && v !== undefined && String(v).trim() !== '')) { blankRowSkipped++; continue }

    const hop = String(row[hopCol] ?? '').trim()
    if (!hop) { blankHopSkipped++; continue }

    const cxStartRaw = parseDate(row[cxStartCol])
    const cxStart = yearOk(cxStartRaw) ? cxStartRaw : null
    const cxCompleteRaw = parseDate(row[cxCompleteCol])
    const cxComplete = yearOk(cxCompleteRaw) ? cxCompleteRaw : null

    // Sites without a construction complete date aren't ready for decom yet
    // — excluded entirely, not just left with unknown aging.
    if (!cxComplete) { noCxCompleteSkipped++; continue }

    const dropOffParsed = parseDropOff(row[dropOffCol])
    // Same year >= 2020 sanity filter as every other date field — a stray
    // typo'd date (e.g. 1/1/1900) shouldn't be treated as a real drop-off.
    const dropOffDate = yearOk(dropOffParsed.date) ? dropOffParsed.date : null
    const comment = dropOffParsed.comment

    const podPathwave = parseYesNo(row[podPathwaveCol])
    const podQuickBase = parseYesNo(row[podQuickBaseCol])

    const gc = resolveGc(String(row[gcCol] ?? ''), String(row[cgCol] ?? ''))

    // Aging = days from CX Complete to today. No fallback to CX Start —
    // rows without CX Complete are excluded above before this runs.
    const aging = daysBetween(cxComplete, today)

    const status: DecomRow['status'] = dropOffDate
      ? ((podPathwave && podQuickBase) ? 'complete' : 'pod_gap')
      : (aging >= 7 ? 'outstanding' : 'pending')

    const pathId = String(row[pathIdCol] ?? '').trim()
    const siteName = String(row[siteNameCol] ?? '').trim()
    const siteNumber = String(row[siteNumberCol] ?? '').trim()
    // Path ID is the real per-row identifier — a HOP can span multiple
    // sites, so it can't disambiguate rows on its own. Fall back to
    // Site Name + Site Number only when Path ID itself is blank.
    const rowKey = pathId || (siteName || siteNumber ? `${siteName}|${siteNumber}` : `row-${i}`)

    rows.push({
      rowKey,
      appPathId: String(row[appPathIdCol] ?? '').trim(),
      hop,
      pathId,
      siteName,
      siteNumber,
      cxStart, cxComplete,
      dropOffDate, comment,
      cm: String(row[cmCol] ?? '').trim(),
      gc,
      podPathwave, podQuickBase,
      aging,
      status,
    })
  }

  console.log(`[decom-parse] ${allRows.length - 1} data rows in → ${rows.length} parsed out (${blankRowSkipped} fully blank, ${blankHopSkipped} blank HOP, ${noCxCompleteSkipped} no CX Complete)`)

  return rows
}

export function decomRowsForGc(rows: DecomRow[], gc: string): DecomRow[] {
  return rows.filter(r => r.gc?.trim().toLowerCase() === gc?.trim().toLowerCase())
}

// Case/whitespace-insensitive GC dedup — decomRowsForGc matches this same
// way, so the GC name list driving summarizeDecomByGc must dedupe the same
// way too. A plain `Set` on raw r.gc values doesn't: any GC not in
// GC_CONFIG's known aliases falls through resolveGc() to its raw, un-
// normalized cell text, so two rows spelled "Vantage" and "VANTAGE" produce
// two "different" entries — decomRowsForGc then matches the same rows under
// both, double-counting that GC's sites when the table is summed.
export function uniqueDecomGcNames(rows: DecomRow[]): string[] {
  const seen = new Map<string, string>() // normalized key -> first-seen display spelling
  rows.forEach(r => {
    const gc = r.gc?.trim()
    if (!gc) return
    const key = gc.toLowerCase()
    if (!seen.has(key)) seen.set(key, gc)
  })
  return Array.from(seen.values())
}

export function summarizeDecomByGc(rows: DecomRow[], gcNames: string[], missingSites: MissingDecomSite[] = []): DecomGcSummary[] {
  return gcNames.map(gc => {
    const gcRows = decomRowsForGc(rows, gc)
    const complete = gcRows.filter(r => r.status === 'complete').length
    const podGap = gcRows.filter(r => r.status === 'pod_gap').length
    const outstanding = gcRows.filter(r => r.status === 'outstanding').length
    const pending = gcRows.filter(r => r.status === 'pending').length
    const pendingPathwave = gcRows.filter(r => r.dropOffDate && !r.podPathwave).length
    const pendingQuickBase = gcRows.filter(r => r.dropOffDate && r.podPathwave && !r.podQuickBase).length
    const missing = missingSites.filter(m => m.gc?.trim().toLowerCase() === gc?.trim().toLowerCase()).length
    const agingVals = gcRows
      .filter(r => r.status === 'outstanding' || r.status === 'pending')
      .map(r => r.aging)
      .filter((a): a is number => a !== null)
    const avgAging = agingVals.length > 0 ? Math.round(agingVals.reduce((s, a) => s + a, 0) / agingVals.length) : null
    return { gc, total: gcRows.length, complete, podGap, outstanding, pending, pendingPathwave, pendingQuickBase, missing, avgAging }
  })
}

// Cross-references completed HOPs from the master tracker against the decom
// file to find sites nobody has entered into decom tracking yet. Matches
// Path ID first, HOP name as fallback — each match consumes one decom row so
// a HOP with two tracker sites needs two distinct decom entries to be fully
// covered (matching one decom row against both sites would hide a real gap).
export function findMissingDecom(decomRows: DecomRow[], trackerHops: TrackerHop[]): MissingDecomSite[] {
  const availableDecom = [...decomRows]
  const missing: MissingDecomSite[] = []
  const today = new Date()

  trackerHops.forEach(t => {
    let matchIdx = t.pathId ? availableDecom.findIndex(d => d.pathId && d.pathId === t.pathId) : -1
    if (matchIdx === -1) {
      matchIdx = availableDecom.findIndex(d => d.hop === t.hop)
    }
    if (matchIdx !== -1) {
      availableDecom.splice(matchIdx, 1)
    } else {
      missing.push({
        hop: t.hop,
        pathId: t.pathId,
        siteName: t.siteName,
        gc: t.gc,
        cm: t.cm,
        nokiaPm: t.nokiaPm,
        ms16a: t.ms16a,
        daysElapsed: daysBetween(t.ms16a, today),
      })
    }
  })

  return missing
}

// Parses the master tracker snapshot for completed HOPs (MS16A actualized,
// year >= 2025) across ALL Nokia PMs — deliberately not scoped to CJ like
// every other tracker parse in this app, since decom coverage needs to be
// checked program-wide. Each qualifying row becomes its own entry (a HOP can
// have two site rows), matching the per-site model used everywhere else in
// this file.
export function parseTrackerHopsForDecom(rows: unknown[][]): TrackerHop[] {
  let headerRow = -1
  for (let i = 0; i < 10; i++) {
    if ((rows[i] as unknown[])?.some(c => String(c).trim() === 'HOP')) { headerRow = i; break }
  }
  if (headerRow === -1) return []

  const headers = rows[headerRow] as string[]
  const col = (name: string) => headers.findIndex(h => String(h).trim() === name)

  const hopCol = col('HOP')
  const don444Col = col('DON 444')
  const gcCol = col('General Contractor')
  const cmCol = col('New CM')
  const nokiaPmCol = col('Nokia PM')
  const ms16aCol = col('MS16 Implementation Ends A')
  const siteNameCol = col('Site Name')
  const siteNumberCol = col('Site Number')
  const pathIdCol = headers.findIndex(h => String(h).trim().replace(/^'+|'+$/g, '') === 'Path ID')

  const result: TrackerHop[] = []

  for (let i = headerRow + 1; i < rows.length; i++) {
    const row = rows[i] as unknown[]
    if (!row) continue

    const don = String(row[don444Col] || '').trim().toUpperCase()
    if (don !== 'DON 444') continue

    const hop = String(row[hopCol] || '').trim()
    if (!hop || hop === 'undefined') continue

    const ms16aRaw = parseDate(row[ms16aCol])
    const ms16a = (ms16aRaw && ms16aRaw.getFullYear() >= 2025) ? ms16aRaw : null
    if (!ms16a) continue

    result.push({
      hop,
      pathId: String(row[pathIdCol] || '').trim().replace(/^'+|'+$/g, ''),
      siteName: String(row[siteNameCol] || '').trim(),
      siteNumber: String(row[siteNumberCol] || '').trim(),
      gc: String(row[gcCol] || '').trim(),
      cm: String(row[cmCol] || '').trim(),
      nokiaPm: String(row[nokiaPmCol] || '').trim(),
      ms16a,
    })
  }

  return result
}

// ─────────────────────────────────────────────
// DECOM STATUS EMAILS — shared summary + top-priority body builder used by
// both the per-GC email (buildDecomEmailMailto) and the program-wide email
// (buildDecomCmEmailMailto, all GCs combined). Aging tiers: 🔴 Critical = 7+
// days, 🟡 Urgent = 4-6 days, 🟢 On Track = 0-3 days — applied to "days
// outstanding" (CX Complete → today) in the Pending Drop Off section, and to
// "days since drop off" in the Pathwave/QuickBase sections.
// ─────────────────────────────────────────────

type AgingTier = 'critical' | 'urgent' | 'onTrack'

function classifyAgingTier(days: number | null): AgingTier {
  const n = days ?? 0
  if (n >= 7) return 'critical'
  if (n >= 4) return 'urgent'
  return 'onTrack'
}

const TIER_ICON: Record<AgingTier, string> = { critical: '🔴', urgent: '🟡', onTrack: '🟢' }

function buildDecomStatusBody(gcRows: DecomRow[], headerLabel: string, dateStr: string): string {
  const today = new Date()
  const pct = (n: number, total: number) => (total > 0 ? Math.round((n / total) * 100) : 0)
  const byAgingDesc = (a: DecomRow, b: DecomRow) => (b.aging ?? -1) - (a.aging ?? -1)

  const total = gcRows.length
  const completeCount = gcRows.filter(r => r.status === 'complete').length
  // pod_gap always implies a Drop Off date exists (that's the branch that
  // produces it) — pathwave/quickbase here are the same non-overlapping split
  // summarizeDecomByGc uses, so a row only ever counts toward one bucket.
  const pendingDropOff = gcRows.filter(r => r.status === 'outstanding' || r.status === 'pending')
  const pendingPathwave = gcRows.filter(r => r.status === 'pod_gap' && !r.podPathwave)
  const pendingQuickBase = gcRows.filter(r => r.status === 'pod_gap' && r.podPathwave && !r.podQuickBase)

  // outstanding/pending already partition Pending Drop Off at the aging >= 7
  // line (see parseDecomRows), so these three tiers exactly re-slice it.
  const criticalCount = pendingDropOff.filter(r => (r.aging ?? 0) >= 7).length
  const urgentCount = pendingDropOff.filter(r => (r.aging ?? 0) >= 4 && (r.aging ?? 0) < 7).length
  const onTrackCount = pendingDropOff.filter(r => (r.aging ?? 0) < 4).length
  const maxAging = pendingDropOff.reduce((max, r) => Math.max(max, r.aging ?? 0), 0)

  const top5DropOff = [...pendingDropOff].sort(byAgingDesc).slice(0, 5)
  const top5Pathwave = [...pendingPathwave].sort(byAgingDesc).slice(0, 5)
  const top5QuickBase = [...pendingQuickBase].sort(byAgingDesc).slice(0, 5)

  const thinDiv = `${'─'.repeat(41)}\n`
  const thickDiv = `${'═'.repeat(41)}\n`

  let body = thickDiv
  body += `📊 DECOM PROGRAM SUMMARY — ${headerLabel}\n`
  body += thickDiv
  body += `Total Sites Tracked:        ${total}\n`
  body += `✅ Complete:                ${completeCount} (${pct(completeCount, total)}%)\n\n`
  body += `⏳ Pending Drop Off:        ${pendingDropOff.length} sites\n`
  body += `   🔴 Critical (7+ days):  ${criticalCount} sites\n`
  body += `   🟡 Urgent (4-6 days):   ${urgentCount} sites\n`
  body += `   🟢 On Track (0-3 days): ${onTrackCount} sites\n\n`
  body += `⚠️  Pending POD Pathwave:  ${pendingPathwave.length} sites\n`
  body += `📋 Pending POD QuickBase:  ${pendingQuickBase.length} sites\n\n`
  body += `⏱️  Oldest Outstanding:    ${maxAging} days\n`
  body += `📅 Report Date:            ${dateStr}\n\n`

  body += thinDiv
  body += `★★★ TOP PRIORITY — PENDING DROP OFF ★★★\n`
  body += thinDiv
  body += `\n`
  top5DropOff.forEach(r => {
    const tier = classifyAgingTier(r.aging)
    const days = r.aging ?? 0
    const label = tier === 'onTrack' ? `ON TRACK — ${days} DAYS` : `${tier === 'critical' ? 'CRITICAL' : 'URGENT'} — ${days} DAYS OUTSTANDING`
    body += `${TIER_ICON[tier]} ${label}\n`
    body += `★ ${r.hop} ★  |  Path ID: ${r.pathId || '—'}  |  CM: ${r.cm || '—'}\n`
    body += `Site: ${r.siteName || '—'}  |  CX Complete: ${fmtDecomDate(r.cxComplete) || '—'}\n`
    if (tier === 'critical') body += `⚠️  DROP OFF OVERDUE — IMMEDIATE ACTION REQUIRED\n`
    else if (tier === 'urgent') body += `⏳ Drop Off Due — Action needed soon\n`
    if (r.comment) body += `💬 Note: ${r.comment}\n`
    body += `\n`
  })

  body += thinDiv
  body += `★★★ TOP PRIORITY — PENDING POD PATHWAVE ★★★\n`
  body += thinDiv
  body += `\n`
  top5Pathwave.forEach(r => {
    const daysSince = r.dropOffDate ? daysBetween(r.dropOffDate, today) : 0
    const tier = classifyAgingTier(daysSince)
    const label = tier === 'onTrack' ? `ON TRACK — ${daysSince} DAYS SINCE DROP OFF` : `${tier === 'critical' ? 'CRITICAL' : 'URGENT'} — ${daysSince} DAYS SINCE DROP OFF`
    body += `${TIER_ICON[tier]} ${label}\n`
    body += `★ ${r.hop} ★  |  Path ID: ${r.pathId || '—'}  |  CM: ${r.cm || '—'}\n`
    body += `Site: ${r.siteName || '—'}  |  Drop Off: ${fmtDecomDate(r.dropOffDate) || '—'}\n`
    if (tier === 'critical') body += `⚠️  POD PATHWAVE OVERDUE — IMMEDIATE ACTION REQUIRED\n`
    else if (tier === 'urgent') body += `⏳ POD Pathwave pending — Action needed soon\n`
    body += `\n`
  })

  body += thinDiv
  body += `★★★ TOP PRIORITY — PENDING POD QUICKBASE ★★★\n`
  body += thinDiv
  body += `\n`
  top5QuickBase.forEach(r => {
    const daysSince = r.dropOffDate ? daysBetween(r.dropOffDate, today) : 0
    const tier = classifyAgingTier(daysSince)
    const label = tier === 'onTrack' ? `ON TRACK — ${daysSince} DAYS SINCE DROP OFF` : `${tier === 'critical' ? 'CRITICAL' : 'URGENT'} — ${daysSince} DAYS SINCE DROP OFF`
    body += `${TIER_ICON[tier]} ${label}\n`
    body += `★ ${r.hop} ★  |  Path ID: ${r.pathId || '—'}  |  CM: ${r.cm || '—'}\n`
    body += `Site: ${r.siteName || '—'}  |  Drop Off: ${fmtDecomDate(r.dropOffDate) || '—'}  |  POD Pathwave: ✅\n`
    if (tier === 'critical') body += `⚠️  POD QUICKBASE OVERDUE — IMMEDIATE ACTION REQUIRED\n`
    else if (tier === 'urgent') body += `⏳ POD QuickBase pending — Action needed soon\n`
    body += `\n`
  })

  return body
}

export function buildDecomEmailMailto(
  gc: string,
  gcRows: DecomRow[],
  emailSettings: { ccList: string[]; gcContactEmails: Record<string, string> }
): string {
  const today = new Date()
  const pad = (n: number) => String(n).padStart(2, '0')
  const dateStr = `${pad(today.getMonth() + 1)}/${pad(today.getDate())}/${today.getFullYear()}`
  const subject = `Viaero Decom Analysis — ${gc} — ${dateStr}`

  let body = `Dear ${gc} Team,\n\n`
  body += `Please find below your current decom status requiring immediate attention.\n\n`
  body += buildDecomStatusBody(gcRows, gc.toUpperCase(), dateStr)
  body += `${'═'.repeat(41)}\n`
  body += `Please see the attached Excel for full site detail.\n\n`
  body += `Thank you,\nCJ`

  const to = lookupContactEmail(emailSettings.gcContactEmails, gc)
  const cc = emailSettings.ccList.join(',')

  return `mailto:${to}?cc=${encodeURIComponent(cc)}&subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`
}

// Program-wide decom status digest — same summary + top-priority format as
// buildDecomEmailMailto, but built across every GC's rows combined instead
// of scoping to one GC.
export function buildDecomCmEmailMailto(
  decomRows: DecomRow[],
  emailSettings: { ccList: string[]; cmContactEmails: Record<string, string> }
): string {
  const today = new Date()
  const pad = (n: number) => String(n).padStart(2, '0')
  const dateStr = `${pad(today.getMonth() + 1)}/${pad(today.getDate())}/${today.getFullYear()}`
  const subject = `Viaero Decom Analysis — ${dateStr}`

  let body = `Dear Team,\n\n`
  body += `Please find below the current decom status requiring immediate attention.\n\n`
  body += buildDecomStatusBody(decomRows, 'ALL GCs', dateStr)
  body += `${'═'.repeat(41)}\n`
  body += `Please see the attached Excel for the full decom detail by site and CM.\n\n`
  body += `Thank you,\nCJ`

  const cc = Array.from(new Set([...emailSettings.ccList, ...Object.values(emailSettings.cmContactEmails)].filter(Boolean))).join(',')

  return `mailto:?cc=${encodeURIComponent(cc)}&subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`
}
