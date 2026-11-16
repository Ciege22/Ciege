// Dashboard → "Pending CPO" tile email: every HOP still waiting on a Service
// CPO (Customer PO, from Viaero to Nokia — see app/lib/spoStatus.ts's doc
// comment for the full procurement chain) before Nokia can even request
// the SPO. Recipients aren't set up yet; CJ adds them in Settings → Email
// Routing (pendingCpoEmail) after this ships. Same self-contained pattern
// as the Dashboard's other tile emails (materialWatchEmail.ts, etc.).
import { applyEmailRouting, type EmailRouting } from './settings'

export interface PendingCpoHop {
  hop: string
  pathId: string
  ms15f: string // forecast start date, already formatted M/D/YYYY
  daysOut: number | null
  // One entry per physical site row for this HOP (normally two), so the email
  // can show which side has a Service CPO and which is still missing.
  serviceCpos: ServiceCpoEntry[]
}

export interface ServiceCpoEntry {
  siteName: string
  cpoNumber: string // blank when the tracker has nothing on file for this site
}

// 'Service CPO Received' is tracked per physical site, the same one-row-per-
// site structure as 'Hardware CPO Received' (see materialWatchEmail.ts). Unlike
// gcCleanup.ts's buildHopCpoMap, which skips blank cells because it only
// wants the numbers that exist, this keeps blank sites too — the point of this
// email is to show which side is still missing. Headers are trimmed before
// matching so a trailing space on the real column name can't break the lookup.
export function buildHopServiceCpoMap(rows: unknown[][]): Map<string, ServiceCpoEntry[]> {
  const out = new Map<string, ServiceCpoEntry[]>()
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
  const cpoCol = col('Service CPO Received')
  if (hopCol === -1) return out
  for (let i = headerRow + 1; i < rows.length; i++) {
    const row = rows[i] as unknown[]
    if (!row) continue
    if (String(row[don444Col] ?? '').trim().toUpperCase() !== 'DON 444') continue
    const hop = String(row[hopCol] ?? '').trim()
    if (!hop || hop === 'undefined') continue
    const siteName = String(row[siteCol] ?? '').trim()
    const raw = cpoCol !== -1 ? String(row[cpoCol] ?? '').trim() : ''
    const cpoNumber = raw.toLowerCase() === 'nan' ? '' : raw
    out.set(hop, [...(out.get(hop) ?? []), { siteName, cpoNumber }])
  }
  return out
}

function todayStr(): string {
  const d = new Date()
  return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`
}

export function buildPendingCpoEmailMailto(
  hops: PendingCpoHop[],
  ccList: string[],
  routing: Record<string, EmailRouting> | undefined,
): string {
  const dateStr = todayStr()
  const subject = `Needed Service CPOs | SPO Requests on Hold | ${dateStr}`

  // Soonest-to-start first, same convention as every other Dashboard tile
  // email — a HOP closer to construction is more urgent to unblock.
  const sorted = [...hops].sort((a, b) => (a.daysOut ?? 999) - (b.daysOut ?? 999))

  // CJ asked for a one-sentence intro with no explanation of why the CPO is
  // needed, and no GC line per HOP (recipients here are the CPO owners, not
  // the contractors). Plain "\n" line breaks (no HTML table, no tabs) so this
  // opens cleanly as a plain-text mailto in Outlook.
  let body = `Hi Team,\n\n`
  body += `Please advise on the status of the Service CPOs needed for the HOPs below.\n\n`
  body += `${hops.length} HOP${hops.length === 1 ? '' : 's'} needing a Service CPO:\n\n`
  sorted.forEach(h => {
    body += `• ${h.hop}\n`
    body += `Path ID: ${h.pathId || '—'}\n`
    body += `Forecast Start: ${h.ms15f || '—'}${h.daysOut != null ? ` (${h.daysOut}d out)` : ''}\n`
    h.serviceCpos.forEach(c => {
      body += `Service CPO (${c.siteName || '—'}): ${c.cpoNumber || 'Not yet received'}\n`
    })
    body += `\n`
  })
  body += `Thank you,\nCJ`

  const { to, cc } = applyEmailRouting(routing, 'pendingCpoEmail', [], ccList)
  return `mailto:${to}?cc=${encodeURIComponent(cc)}&subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`
}
