import { supabase } from './supabase'

// ─────────────────────────────────────────────
// TYPES
// ─────────────────────────────────────────────

export interface GcProgramEntry {
  gc: string
  cm: string
  crewCount: number
}

export interface ProgramSettings {
  nokiaPMs: string[]
  gcs: GcProgramEntry[]
}

export interface ThresholdSettings {
  ntpUrgentDays: number
  materialWatchDays: number
  durationAlertDays: number
  pullInBufferDays: number
  // Schedule Optimizer
  pushWindow: number
  pushAmount: number
  hopDuration: number
  rampUpThreshold: number
  rampUpWindow: number
}

// One row per distinct email this app generates — lets Settings expose a
// per-type To/CC override instead of everything sharing one CC List. `to`
// entries here are ADDED on top of whatever a type already resolves as its
// primary recipient (a per-GC/per-CM contact lookup, or a hardcoded team
// address like the GR processors) — they never replace it, since for a
// GC/CM-specific email a fixed To wouldn't make sense across different
// GCs/CMs. `cc`, when non-empty, REPLACES the shared ccList (or, for GR,
// the financeEmails default) for that type only — see applyEmailRouting.
export interface EmailTypeConfig {
  id: string
  label: string
  // Shown in the Settings UI as a hint — true if this type already fills in
  // a primary recipient on its own (per-GC/per-CM contact, or a fixed team
  // address) before any per-type "To" here gets added on top.
  hasBaseTo: boolean
}

export const EMAIL_TYPES: EmailTypeConfig[] = [
  { id: 'cmEmail', label: 'CM View — CM Email', hasBaseTo: false },
  { id: 'cmPipelineEmail', label: 'CM View — CM Pipeline Email (All CMs)', hasBaseTo: false },
  { id: 'cmDailyEmail', label: 'CM View — CM Daily Email', hasBaseTo: true },
  { id: 'gcPipelineEmail', label: 'GC Call View — Pipeline Weekly Email', hasBaseTo: true },
  { id: 'decomGcEmail', label: 'Decom — GC Email', hasBaseTo: true },
  { id: 'decomCmDigestEmail', label: 'Decom — Program-Wide CM Digest', hasBaseTo: false },
  { id: 'scopGcEmail', label: 'SCOP — GC Email', hasBaseTo: true },
  { id: 'grEmail', label: 'GR / Invoicing — Release Request', hasBaseTo: true },
  { id: 'ntpEmailViaero', label: 'NTP Tracker — Viaero Action Required', hasBaseTo: false },
  { id: 'ntpEmailNokia', label: 'NTP Tracker — Nokia Action Required', hasBaseTo: false },
  { id: 'ntpEmailItw', label: 'NTP Tracker — ITW/Samsung Action Required', hasBaseTo: false },
]

export interface EmailRouting {
  to: string[]
  cc: string[]
}

export interface EmailSettings {
  ccList: string[]
  financeEmails: string[]
  gcContactEmails: Record<string, string>
  // Keyed by CM name (not GC) — used to CC the right people on the decom
  // per-CM digest email.
  cmContactEmails: Record<string, string>
  // Keyed by EmailTypeConfig.id — see applyEmailRouting for how these merge
  // with each email's existing to/cc logic.
  routing: Record<string, EmailRouting>
}

// Merges a per-type routing override (Settings → Email Routing) onto an
// email's already-computed base to/cc. baseTo/baseCc are whatever that
// email builder already resolves today (a per-GC contact lookup, a
// hardcoded team address, the shared ccList, etc.) — this never changes
// that logic, it only layers the override on top: custom "to" entries are
// added to baseTo (never replacing it), and a non-empty custom "cc" fully
// replaces baseCc (falls back to baseCc when unset, per that trade-off).
export function applyEmailRouting(
  routing: Record<string, EmailRouting> | undefined,
  typeId: string,
  baseTo: string[],
  baseCc: string[],
): { to: string; cc: string } {
  const entry = routing?.[typeId]
  const to = [...baseTo, ...(entry?.to ?? [])].filter(Boolean)
  const cc = entry?.cc && entry.cc.length > 0 ? entry.cc : baseCc
  return { to: to.filter(Boolean).join(','), cc: cc.filter(Boolean).join(',') }
}

export interface DisplaySettings {
  defaultPmFilter: string
  defaultSortOrder: string
}

// SCOP (Site Close-Out Package) tunables. Additive — see
// docs/scop/scop_settings_schema.js. The GC alias map and GC contact list
// are deliberately NOT here; SCOP layers gcAliasMap on top of GC_CONFIG and
// reuses EmailSettings.gcContactEmails / ccList, since a GC is the same GC
// program-wide.
export interface ScopSettings {
  // Days since Construction Complete before the aging number turns amber,
  // then red. `< green` = green, `green..<amber` = amber, `>= amber` = red.
  agingColorThresholds: { green: number; amber: number }
  // Min aging for a HOP to appear in a GC email's "Top Priority" callout.
  emailPriorityThresholdDays: number
  // Max HOPs shown in that callout (oldest first).
  emailPriorityCap: number
  // Case-insensitive substrings against `One and Done` that reclassify a HOP
  // from In Progress into the separate OAD bucket.
  oadKeywords: string[]
  // 1-based header-row number in the uploaded tracker's HOPs tab. null = auto
  // (scan for a row containing both "HOP" and "General Contractor").
  headerRowOverride: number | null
  // Overrides "today" for every aging calc. 'YYYY-MM-DD' or null (= real today).
  asOfDateOverride: string | null
  // Extra raw-name -> canonical-name entries, layered on top of GC_CONFIG's
  // roster. Keyed lowercase. Seeded with the known casing collisions.
  gcAliasMap: Record<string, string>
  // Which of the 8 Pathwave checklist item labels count as GC action items.
  pathwaveGcOwnedItems: string[]
  // Confirmed false — all QuickBase items are Nokia-owned. A settings flip,
  // not a redeploy, if that business rule ever changes.
  quickbaseHasGcOwnedItems: boolean
}

// ─────────────────────────────────────────────
// DEFAULTS — match the values that were hardcoded before this settings
// page existed, so nothing changes in behavior until someone edits them.
// ─────────────────────────────────────────────

export const DEFAULT_PROGRAM: ProgramSettings = {
  nokiaPMs: [],
  gcs: [],
}

// Fallback crew counts used by the Schedule Optimizer when a GC has no
// crewCount configured (or is 0) in the Program Settings GC roster.
export const DEFAULT_CREW_COUNTS: Record<string, number> = {
  Mastec: 2,
  MZI: 2,
  'NV Tel': 2,
  'Tech CX': 4,
  Vikor: 3,
  TCE: 1,
  InSite: 1,
}

// GC/CM contact-email lookup, case/whitespace-insensitive — names get typed
// by hand both in the source tracker (General Contractor / CM columns) and
// again in Settings, so exact casing between the two isn't guaranteed. A
// plain `map[name]` lookup silently misses on any casing/spacing mismatch
// (e.g. tracker has "Vikor", Settings has "vikor "), which reads as "the
// contact email never gets used" even though it's saved correctly.
export function lookupContactEmail(map: Record<string, string>, name: string): string {
  const target = name?.trim().toLowerCase()
  if (!target) return ''
  const key = Object.keys(map).find(k => k.trim().toLowerCase() === target)
  return key ? map[key] : ''
}

export function crewCountForGc(program: ProgramSettings, gc: string): number {
  const entry = program.gcs.find(g => g.gc?.trim().toLowerCase() === gc?.trim().toLowerCase())
  if (entry && entry.crewCount > 0) return entry.crewCount
  return DEFAULT_CREW_COUNTS[gc] || 1
}

export const DEFAULT_THRESHOLDS: ThresholdSettings = {
  ntpUrgentDays: 14,
  materialWatchDays: 14,
  durationAlertDays: 18,
  pullInBufferDays: 10,
  pushWindow: 7,
  pushAmount: 5,
  hopDuration: 14,
  rampUpThreshold: 3,
  rampUpWindow: 30,
}

// Same 8 addresses that were previously hardcoded in app/lib/grTracker.ts
// (GR_EMAIL_CC_BASE) and duplicated in app/gc-call/page.tsx's NTP email CC list.
export const DEFAULT_EMAIL: EmailSettings = {
  ccList: [
    'thomas.meinke.ext@nokia.com',
    'steve.jahr.ext@nokia.com',
    'christopher.seebach@nokia.com',
    'george.anson@nokia.com',
    'curtiss.lindsey.ext@nokia.com',
    'emily.rudolph@nokia.com',
    'scott.tomlinson.ext@nokia.com',
    'paul.1.barlow.ext@nokia.com',
  ],
  financeEmails: [
    'thomas.meinke.ext@nokia.com',
    'steve.jahr.ext@nokia.com',
    'christopher.seebach@nokia.com',
    'george.anson@nokia.com',
    'curtiss.lindsey.ext@nokia.com',
    'emily.rudolph@nokia.com',
    'scott.tomlinson.ext@nokia.com',
    'paul.1.barlow.ext@nokia.com',
  ],
  gcContactEmails: {},
  cmContactEmails: {},
  routing: {},
}

export const DEFAULT_DISPLAY: DisplaySettings = {
  defaultPmFilter: 'ALL',
  defaultSortOrder: 'trigger',
}

// SCOP checklist item labels — kept here so DEFAULT_SCOP and the settings UI
// can reference the full 8-item Pathwave set without importing app/lib/scop.ts
// (which pulls in the whole calc engine).
export const SCOP_PATHWAVE_ITEM_LABELS = [
  'Install Photos', 'Decom Photos NQR', 'Install Photos NQR',
  'Red-Line CD', 'Decom Asset Form', 'POD', 'Asset Form', 'Packing Slip',
]

export const DEFAULT_SCOP: ScopSettings = {
  agingColorThresholds: { green: 30, amber: 60 },
  emailPriorityThresholdDays: 60,
  emailPriorityCap: 5,
  oadKeywords: ['oad'],
  headerRowOverride: null,
  asOfDateOverride: null,
  gcAliasMap: { wavelink: 'WaveLink', 'viking/capital tower': 'Viking' },
  pathwaveGcOwnedItems: [
    'Install Photos', 'Decom Photos NQR', 'Install Photos NQR',
    'Red-Line CD', 'Decom Asset Form', 'POD',
  ],
  quickbaseHasGcOwnedItems: false,
}

// ─────────────────────────────────────────────
// LOAD / SAVE — mirrors the existing pm_updates_cache upsert/select
// pattern used elsewhere in the app (id / updates / updated_at columns).
// ─────────────────────────────────────────────

async function loadSection<T>(section: string, fallback: T): Promise<T> {
  const { data } = await supabase
    .from('pm_updates_cache')
    .select('updates')
    .eq('id', `settings-${section}`)
    .single()
  if (data?.updates) {
    try {
      return { ...fallback, ...JSON.parse(data.updates) }
    } catch {
      return fallback
    }
  }
  return fallback
}

async function saveSection(section: string, value: unknown) {
  await supabase.from('pm_updates_cache').upsert({
    id: `settings-${section}`,
    updates: JSON.stringify(value),
    updated_at: new Date().toISOString(),
  })
}

export const loadProgramSettings = () => loadSection('program', DEFAULT_PROGRAM)
export const saveProgramSettings = (v: ProgramSettings) => saveSection('program', v)

export const loadThresholdSettings = () => loadSection('thresholds', DEFAULT_THRESHOLDS)
export const saveThresholdSettings = (v: ThresholdSettings) => saveSection('thresholds', v)

export const loadEmailSettings = () => loadSection('email', DEFAULT_EMAIL)
export const saveEmailSettings = (v: EmailSettings) => saveSection('email', v)

export const loadDisplaySettings = () => loadSection('display', DEFAULT_DISPLAY)
export const saveDisplaySettings = (v: DisplaySettings) => saveSection('display', v)

export const loadScopSettings = () => loadSection('scop', DEFAULT_SCOP)
export const saveScopSettings = (v: ScopSettings) => saveSection('scop', v)
