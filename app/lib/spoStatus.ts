// Shared "what's the SPO status" classification, used by the Dashboard, CM
// View, GC Call View's Pipeline tab, and Weekly Focus — one place so the
// hierarchy only has to be defined once instead of four slightly-diverging
// copies of the same ternary.
//
// Reads three tracker columns:
//   'CX SPO issued'        -> hasSpo: the SPO itself has been created
//   'Service CPO Received' -> hasCpo: CPO's in, ready to cut the SPO
//   'CX SPO Request'       -> hasSpoRequest: the request was logged, but the
//                             SPO hasn't actually been created/issued yet
export type SpoStatus = 'issued' | 'cpo_ready' | 'requested' | 'needed'

export interface SpoStatusResult {
  status: SpoStatus
  label: string       // e.g. "Requested — Pending SPO Creation"
  shortLabel: string  // for compact badges, e.g. "📨 Requested"
}

export function computeSpoStatus(hasSpo: boolean, hasCpo: boolean, hasSpoRequest: boolean): SpoStatusResult {
  if (hasSpo) return { status: 'issued', label: 'Issued', shortLabel: '✓ Issued' }
  if (hasCpo) return { status: 'cpo_ready', label: 'Cut SPO Now', shortLabel: '⚡ Cut Now' }
  if (hasSpoRequest) return { status: 'requested', label: 'Requested — Pending SPO Creation', shortLabel: '📨 Requested' }
  return { status: 'needed', label: 'Pending SPO Request', shortLabel: '🔴 Pending Request' }
}
