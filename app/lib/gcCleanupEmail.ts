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
  // Bullet is just the HOP so it reads at a glance; every detail stacks
  // underneath it instead of running into one long line.
  items.forEach(({ hop, assignment }) => {
    body += `• ${hop.hopDisplay}  |  Path ID: ${hop.pathId || '—'}\n`
    body += `    Quoted Price: ${fmtMoney(assignment.quote ?? 0)}\n\n`
  })
  body += `Total: ${fmtMoney(total)}\n\nThank you,\nCJ`

  const contact = lookupContactEmail(gcContactEmails, newGc)
  const { to, cc } = applyEmailRouting(routing, 'gcCleanupQuoteEmail', contact ? [contact] : [], ccList)
  window.location.href = `mailto:${encodeURIComponent(to)}?cc=${encodeURIComponent(cc)}&subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`
}

// Only HOPs that actually have an unpaid SPO left to cancel — a HOP that was
// already paid in full before reassignment has nothing for Finance to act on.
export function cancellableHops(items: QuotedHop[]): QuotedHop[] {
  return items.filter(i => i.hop.unpaidTiers.length > 0)
}

// The clean percentage a tier code represents, separate from tierLabel
// (which is the fuller "60% — MS16A (CX Complete)" description) — Finance
// needs the plain % to know which split of the SPO is being cancelled.
export function tierPercent(tier: string): string {
  switch (tier) {
    case 'init20': return '20%'
    case '60': return '60%'
    case '70': return '70%'
    case '20': return '20%'
    case '30': return '30%'
    case 'CR': return 'CR (no %)'
    default: return tier || '—'
  }
}

export function openCancellationEmail(
  items: QuotedHop[],
  financeEmails: string[],
  routing: Record<string, EmailRouting> | undefined,
): void {
  const subject = `SPO Cancellation Request — Reassigned Clean-Up Work — ${todayStr()}`
  let body = `Please cancel the SPOs below — this scope is being reassigned to a new contractor for clean-up, `
    + `so the original vendor's SPO should not also be paid out.\n\n`
  // Bullet is the HOP + Path ID + the HOP's current (departed) GC per the
  // tracker; each unpaid tier stacks underneath with its OWN vendor, since a
  // HOP can carry SPO lines from more than one vendor (a small CR issued to
  // an unrelated sub, say) even though the tracker attributes the HOP itself
  // to one GC — Finance needs to know exactly whose SPO each line cancels.
  items.forEach(({ hop }) => {
    body += `• ${hop.hopDisplay}  |  Path ID: ${hop.pathId || '—'}  |  HOP's GC: ${hop.gc}\n`
    hop.unpaidTiers.forEach(t => {
      body += `    SPO #: ${t.spoNumber || '—'}  |  Vendor: ${t.vendor}  |  %: ${tierPercent(t.tier)}  |  Tier: ${t.tierLabel}  |  Value: ${fmtMoney(t.value)}\n`
    })
    body += `\n`
  })
  body += `Thank you,\nCJ`

  const baseCc = financeEmails.length > 0 ? financeEmails : GR_EMAIL_CC_BASE
  const { to, cc } = applyEmailRouting(routing, 'gcCleanupCancelEmail', GR_EMAIL_TO, baseCc)
  window.location.href = `mailto:${encodeURIComponent(to)}?cc=${encodeURIComponent(cc)}&subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`
}
