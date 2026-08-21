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
}

export interface EmailSettings {
  ccList: string[]
  financeEmails: string[]
  gcContactEmails: Record<string, string>
}

export interface DisplaySettings {
  defaultPmFilter: string
  defaultSortOrder: string
}

// ─────────────────────────────────────────────
// DEFAULTS — match the values that were hardcoded before this settings
// page existed, so nothing changes in behavior until someone edits them.
// ─────────────────────────────────────────────

export const DEFAULT_PROGRAM: ProgramSettings = {
  nokiaPMs: [],
  gcs: [],
}

export const DEFAULT_THRESHOLDS: ThresholdSettings = {
  ntpUrgentDays: 14,
  materialWatchDays: 14,
  durationAlertDays: 18,
  pullInBufferDays: 10,
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
}

export const DEFAULT_DISPLAY: DisplaySettings = {
  defaultPmFilter: 'ALL',
  defaultSortOrder: 'trigger',
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
