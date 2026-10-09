// GC Clean-Up emails: a quote-confirmation email per new GC (one mailto per
// GC, opened one click at a time — matches the SPO batch bar's convention,
// since browsers block several mailto: navigations fired at once), and one
// combined SPO-cancellation email to the internal finance team naming the
// ORIGINAL (departed) GC's SPOs to cancel, never the new GC's.
import { type EmailRouting, applyEmailRouting, lookupContactEmail } from './settings'
import { GR_EMAIL_TO, GR_EMAIL_CC_BASE, fmtMoney } from './grTracker'
import type { CleanupHop, CleanupAssignment } from './gcCleanup'

export interface QuotedHop {
  hop: CleanupHop
  assignment: CleanupAssignment
}

function todayStr(): string {
  const d = new Date()
  return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`
}

// Every HOP with both a quote and a new GC assigned — the set either email
// draws from.
export function quotedHops(hops: CleanupHop[], assignments: Record<string, CleanupAssignment>): QuotedHop[] {
  return hops
    .filter(h => {
      const a = assignments[h.hop]
      return !!a && a.quote != null && a.newGc.trim() !== ''
    })
    .map(h => ({ hop: h, assignment: assignments[h.hop] }))
}

export function groupByNewGc(items: QuotedHop[]): { newGc: string; items: QuotedHop[] }[] {
  const byGc = new Map<string, QuotedHop[]>()
  items.forEach(i => {
    const key = i.assignment.newGc.trim()
    byGc.set(key, [...(byGc.get(key) ?? []), i])
  })
  return Array.from(byGc.entries())
    .map(([newGc, groupItems]) => ({ newGc, items: groupItems }))
    .sort((a, b) => a.newGc.localeCompare(b.newGc))
}

export function openGcQuoteEmail(
  newGc: string,
  items: QuotedHop[],
  gcContactEmails: Record<string, string>,
  ccList: string[],
  routing: Record<string, EmailRouting> | undefined,
): void {
  const total = items.reduce((s, i) => s + (i.assignment.quote ?? 0), 0)
  const subject = `Site Clean-Up — Quote Confirmation — ${newGc} — ${todayStr()}`
  let body = `Hi ${newGc},\n\nPlease confirm you're ready to proceed on the following site clean-up at the quoted price:\n\n`
  items.forEach(({ hop, assignment }) => {
    body += `• ${hop.hopDisplay}  |  Path ID: ${hop.pathId || '—'}  |  Quoted Price: ${fmtMoney(assignment.quote ?? 0)}\n`
  })
  body += `\nTotal: ${fmtMoney(total)}\n\nThank you,\nCJ`

  const contact = lookupContactEmail(gcContactEmails, newGc)
  const { to, cc } = applyEmailRouting(routing, 'gcCleanupQuoteEmail', contact ? [contact] : [], ccList)
  window.location.href = `mailto:${encodeURIComponent(to)}?cc=${encodeURIComponent(cc)}&subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`
}

// Only HOPs that actually have an unpaid SPO left to cancel — a HOP that was
// already paid in full before reassignment has nothing for Finance to act on.
export function cancellableHops(items: QuotedHop[]): QuotedHop[] {
  return items.filter(i => i.hop.unpaidTiers.length > 0)
}

export function openCancellationEmail(
  items: QuotedHop[],
  financeEmails: string[],
  routing: Record<string, EmailRouting> | undefined,
): void {
  const subject = `SPO Cancellation Request — Reassigned Clean-Up Work — ${todayStr()}`
  let body = `Please cancel the SPOs below — this scope is being reassigned to a new contractor for clean-up, `
    + `so the original vendor's SPO should not also be paid out.\n\n`
  items.forEach(({ hop }) => {
    hop.unpaidTiers.forEach(t => {
      body += `• ${hop.hopDisplay}  |  Path ID: ${hop.pathId || '—'}  |  SPO #: ${t.spoNumber || '—'}  |  Vendor: ${hop.gc}  |  Tier: ${t.tierLabel}  |  Value: ${fmtMoney(t.value)}\n`
    })
  })
  body += `\nThank you,\nCJ`

  const baseCc = financeEmails.length > 0 ? financeEmails : GR_EMAIL_CC_BASE
  const { to, cc } = applyEmailRouting(routing, 'gcCleanupCancelEmail', GR_EMAIL_TO, baseCc)
  window.location.href = `mailto:${encodeURIComponent(to)}?cc=${encodeURIComponent(cc)}&subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`
}
