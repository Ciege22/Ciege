// Shared "what's the SPO status" classification, used by the Dashboard, CM
// View, GC Call View's Pipeline tab, and Weekly Focus — one place so the
// hierarchy only has to be defined once instead of four slightly-diverging
// copies of the same ternary.
//
// Reads three tracker columns:
//   'CX SPO issued'        -> hasSpo: the SPO itself has been created
//   'Service CPO Received' -> hasCpo: CPO's in
//   'CX SPO Request'       -> hasSpoRequest: the request was logged, but the
//                             SPO hasn't actually been created/issued yet
//
// Matches the real procurement flow (per CJ): the CPO is the Customer PO to
// Nokia, and it's what lets Nokia turn around and issue the SPO (Supplier
// PO) to the contractor — CJ can't meaningfully request an SPO without a
// CPO in hand, so "has a CPO" isn't its own separate signal once a request
// exists; it's the gate that made the request possible in the first place.
// Four states, in priority order:
//   issued        - the SPO itself exists. Done.
//   requested     - CX SPO Request logged, SPO not created yet — waiting on
//                   Nokia. (Previously split into two tiles depending on
//                   whether a CPO was also present — collapsed into one,
//                   since by the time a request exists the CPO already did
//                   its job of enabling it.)
//   needs_request - CPO's in, nobody's requested the SPO yet. This is the
//                   action item: CJ needs to request it.
//   pending_cpo   - nothing yet. Waiting on the CPO to arrive.
export type SpoStatus = 'issued' | 'requested' | 'needs_request' | 'pending_cpo'

export interface SpoStatusResult {
  status: SpoStatus
  label: string       // e.g. "Requested — Pending SPO Creation"
  shortLabel: string  // for compact badges, e.g. "📨 Requested"
}

export function computeSpoStatus(hasSpo: boolean, hasCpo: boolean, hasSpoRequest: boolean): SpoStatusResult {
  if (hasSpo) return { status: 'issued', label: 'Issued', shortLabel: '✓ Issued' }
  if (hasSpoRequest) return { status: 'requested', label: 'Requested — Pending SPO Creation', shortLabel: '📨 Requested' }
  if (hasCpo) return { status: 'needs_request', label: 'Needs SPO Requested', shortLabel: '⚡ Needs Request' }
  return { status: 'pending_cpo', label: 'Pending CPO', shortLabel: '🔴 Pending CPO' }
}
