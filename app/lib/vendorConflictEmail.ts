// Dashboard → "Vendor Conflicts" tile email: flags every HOP whose forecast
// start date lands inside a Samsung/ITW/Allwave on-site window, with the
// real conflict detail (which vendor, through what date) — not just a bare
// flagged list. Recipients aren't set up yet; CJ adds them in Settings →
// Email Routing (vendorConflictEmail) after this ships.
import { applyEmailRouting, type EmailRouting } from './settings'

export interface VendorConflictHop {
  hop: string
  gc: string
  pathId: string
  ms15f: string // forecast start date, already formatted M/D/YYYY
  daysOut: number | null
  vendorConflictDetail: string // e.g. "ITW on site through 11/20/2026"
}

function todayStr(): string {
  const d = new Date()
  return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`
}

export function buildVendorConflictEmailMailto(
  hops: VendorConflictHop[],
  ccList: string[],
  routing: Record<string, EmailRouting> | undefined,
): string {
  const dateStr = todayStr()
  const subject = `Vendor Conflicts — Schedule Review Needed | ${dateStr}`

  // Soonest-to-start first — the most time-sensitive conflict to resolve.
  const sorted = [...hops].sort((a, b) => (a.daysOut ?? 999) - (b.daysOut ?? 999))

  let body = `Hi Team,\n\n`
  body += `These sites have a vendor on-site window that overlaps the forecast construction start date. Please review and advise if the schedule needs to shift.\n\n`
  body += `${hops.length} HOP${hops.length === 1 ? '' : 's'}:\n\n`
  sorted.forEach(h => {
    body += `• ${h.hop}\n`
    body += `GC: ${h.gc || '—'}\n`
    body += `Path ID: ${h.pathId || '—'}\n`
    body += `Forecast Start: ${h.ms15f || '—'}${h.daysOut != null ? ` (${h.daysOut}d out)` : ''}\n`
    body += `Conflict: ${h.vendorConflictDetail || '—'}\n`
    body += `\n`
  })
  body += `Thank you,\nCJ`

  const { to, cc } = applyEmailRouting(routing, 'vendorConflictEmail', [], ccList)
  return `mailto:${to}?cc=${encodeURIComponent(cc)}&subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`
}
