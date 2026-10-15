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
  ms15f: string // forecast start date, already formatted M/D/YYYY
  daysOut: number | null
  hasNtp: boolean
  ntpWaitingOn: string
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
  sorted.forEach(h => {
    body += `• ${h.hop}  |  GC: ${h.gc || '—'}  |  Forecast Start: ${h.ms15f || '—'}${h.daysOut != null ? `  (${h.daysOut}d out)` : ''}\n`
    body += `    NTP: ${h.hasNtp ? '✅ Received' : `⏳ Pending${h.ntpWaitingOn ? ` — waiting on ${h.ntpWaitingOn}` : ''}`}\n\n`
  })
  body += `Thank you,\nCJ`

  const { to, cc } = applyEmailRouting(routing, 'materialWatchEmail', [], ccList)
  return `mailto:${to}?cc=${encodeURIComponent(cc)}&subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`
}
