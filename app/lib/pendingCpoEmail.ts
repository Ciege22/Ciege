// Dashboard → "Pending CPO" tile email: every HOP still waiting on a CPO
// (Customer PO, from Viaero to Nokia — see app/lib/spoStatus.ts's doc
// comment for the full procurement chain) before Nokia can even request
// the SPO. Recipients aren't set up yet; CJ adds them in Settings → Email
// Routing (pendingCpoEmail) after this ships. Same self-contained pattern
// as the Dashboard's other tile emails (materialWatchEmail.ts, etc.).
import { applyEmailRouting, type EmailRouting } from './settings'

export interface PendingCpoHop {
  hop: string
  gc: string
  pathId: string
  ms15f: string // forecast start date, already formatted M/D/YYYY
  daysOut: number | null
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
  const subject = `Pending CPO — SPO Request Blocked | ${dateStr}`

  // Soonest-to-start first, same convention as every other Dashboard tile
  // email — a HOP closer to construction is more urgent to unblock.
  const sorted = [...hops].sort((a, b) => (a.daysOut ?? 999) - (b.daysOut ?? 999))

  let body = `Hi Team,\n\n`
  body += `These sites don't have a CPO on file yet, so Nokia can't request the SPO to the contractor. Please advise on status so we can keep procurement moving.\n\n`
  body += `${hops.length} HOP${hops.length === 1 ? '' : 's'}:\n\n`
  sorted.forEach(h => {
    body += `• ${h.hop}\n`
    body += `GC: ${h.gc || '—'}\n`
    body += `Path ID: ${h.pathId || '—'}\n`
    body += `Forecast Start: ${h.ms15f || '—'}${h.daysOut != null ? ` (${h.daysOut}d out)` : ''}\n`
    body += `\n`
  })
  body += `Thank you,\nCJ`

  const { to, cc } = applyEmailRouting(routing, 'pendingCpoEmail', [], ccList)
  return `mailto:${to}?cc=${encodeURIComponent(cc)}&subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`
}
