// Dashboard → "Material Watch" tile email: one confirmation request to the
// material team for every site forecast to start within the watch window
// (default 14 days, ThresholdSettings.materialWatchDays) that doesn't yet
// have a Material Received date. A plain program-wide digest, not per-GC —
// the material team needs the whole upcoming list in one place to confirm
// against, not a GC-by-GC split.
import { applyEmailRouting, type EmailRouting } from './settings'

export interface MaterialWatchHop {
  hop: string
  gc: string
  pathId: string
  ms15f: string // forecast start date, already formatted M/D/YYYY
  daysOut: number | null
  // One entry per physical site row for this HOP (a HOP normally has two —
  // see buildHopHardwarePoMap below for why this is a per-HOP array rather
  // than two fixed fields).
  hardwarePos: HardwarePoEntry[]
}

export interface HardwarePoEntry {
  siteName: string
  poNumber: string // blank when the tracker has nothing on file yet
}

// 'Hardware CPO Received' is tracked per physical site, same one-row-per-
// site structure as every other per-site tracker column in this app (see
// app/lib/gcCleanup.ts's buildHopCpoMap for 'Service CPO Received' — a
// different column/business meaning, identical shape: a HOP with two
// sites carries two entries here, one PO number per site). Headers are
// trimmed before matching so a trailing space on the real column name
// (CJ flagged this as worth checking) can't break the lookup.
export function buildHopHardwarePoMap(rows: unknown[][]): Map<string, HardwarePoEntry[]> {
  const out = new Map<string, HardwarePoEntry[]>()
  let headerRow = -1
  for (let i = 0; i < 10; i++) {
    if ((rows[i] as unknown[])?.some(c => String(c).trim() === 'HOP')) { headerRow = i; break }
  }
  if (headerRow === -1) return out
  const headers = (rows[headerRow] as unknown[]).map(h => String(h ?? '').trim())
  const col = (name: string) => headers.indexOf(name)
  const hopCol = col('HOP')
  const don444Col = col('DON 444')
  const siteCol = col('Site Name')
  const poCol = col('Hardware CPO Received')
  if (hopCol === -1) return out
  for (let i = headerRow + 1; i < rows.length; i++) {
    const row = rows[i] as unknown[]
    if (!row) continue
    if (String(row[don444Col] ?? '').trim().toUpperCase() !== 'DON 444') continue
    const hop = String(row[hopCol] ?? '').trim()
    if (!hop || hop === 'undefined') continue
    const siteName = String(row[siteCol] ?? '').trim()
    const poNumber = poCol !== -1 ? String(row[poCol] ?? '').trim() : ''
    out.set(hop, [...(out.get(hop) ?? []), { siteName, poNumber }])
  }
  return out
}

function todayStr(): string {
  const d = new Date()
  return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`
}

export function buildMaterialWatchEmailMailto(
  hops: MaterialWatchHop[],
  ccList: string[],
  routing: Record<string, EmailRouting> | undefined,
): string {
  const dateStr = todayStr()
  const subject = `Upcoming Starts | Material Watch | Needing Confirmation | ${dateStr}`

  // Soonest-to-start first — that's the one the material team needs to
  // confirm on first.
  const sorted = [...hops].sort((a, b) => (a.daysOut ?? 999) - (b.daysOut ?? 999))

  let body = `Hi Team,\n\n`
  body += `These sites are expected to start within the next 14 days. Are we on track with the forecasted material dates — if not, please advise so I can coordinate schedules accordingly.\n\n`
  body += `${hops.length} site${hops.length === 1 ? '' : 's'}:\n\n`
  // Plain stacked lines, not the old one-row "• HOP | GC: X | ..." format —
  // CJ asked for this to read top-to-bottom per HOP, not scanned across a
  // long line, now that it also carries Path ID and two per-site Hardware
  // PO lines. Plain "\n" line breaks (no HTML table, no tabs) so this still
  // opens cleanly as a plain-text mailto in Outlook.
  sorted.forEach(h => {
    body += `• ${h.hop}\n`
    body += `GC: ${h.gc || '—'}\n`
    body += `Path ID: ${h.pathId || '—'}\n`
    body += `Forecast Start: ${h.ms15f || '—'}${h.daysOut != null ? ` (${h.daysOut}d out)` : ''}\n`
    h.hardwarePos.forEach(po => {
      body += `Hardware PO (${po.siteName || '—'}): ${po.poNumber || 'Not yet issued'}\n`
    })
    body += `\n`
  })
  body += `Thank you,\nCJ`

  const { to, cc } = applyEmailRouting(routing, 'materialWatchEmail', [], ccList)
  return `mailto:${to}?cc=${encodeURIComponent(cc)}&subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`
}
