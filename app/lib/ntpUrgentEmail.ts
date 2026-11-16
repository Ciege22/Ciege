// Dashboard → "NTP Urgent" tile email: a single combined chase email for
// every HOP missing NTP within the urgent window (default 14 days,
// ThresholdSettings.ntpUrgentDays), regardless of who owns the blocker.
// NTP Tracker (app/ntp-tracker/page.tsx) already has a mature chase-email
// system split by owner (Viaero/Nokia/ITW routing) — CJ confirmed this
// Dashboard tile should stay a single combined list instead of reusing
// that split, since it's a narrower, more urgent slice than NTP Tracker's
// full pending-NTP view.
import { applyEmailRouting, type EmailRouting } from './settings'

export interface NtpUrgentHop {
  hop: string
  gc: string
  pathId: string
  ms15f: string // forecast start date, already formatted M/D/YYYY
  daysOut: number | null
  ntpWaitingOn: string
}

function todayStr(): string {
  const d = new Date()
  return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`
}

export function buildNtpUrgentEmailMailto(
  hops: NtpUrgentHop[],
  ccList: string[],
  routing: Record<string, EmailRouting> | undefined,
): string {
  const dateStr = todayStr()
  const subject = `NTP Urgent — Missing NTP ≤14 Days to Start | ${dateStr}`

  // Soonest-to-start first — that's the one that needs NTP chased first.
  const sorted = [...hops].sort((a, b) => (a.daysOut ?? 999) - (b.daysOut ?? 999))

  let body = `Hi Team,\n\n`
  body += `These sites are forecast to start within 14 days and are still missing NTP. Please advise on status so we can keep these on schedule.\n\n`
  body += `${hops.length} HOP${hops.length === 1 ? '' : 's'}:\n\n`
  sorted.forEach(h => {
    body += `• ${h.hop}\n`
    body += `GC: ${h.gc || '—'}\n`
    body += `Path ID: ${h.pathId || '—'}\n`
    body += `Forecast Start: ${h.ms15f || '—'}${h.daysOut != null ? ` (${h.daysOut}d out)` : ''}\n`
    body += `NTP Waiting On: ${h.ntpWaitingOn || '—'}\n`
    body += `\n`
  })
  body += `Thank you,\nCJ`

  const { to, cc } = applyEmailRouting(routing, 'ntpUrgentEmail', [], ccList)
  return `mailto:${to}?cc=${encodeURIComponent(cc)}&subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`
}
