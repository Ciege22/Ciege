'use client'

export const dynamic = 'force-dynamic'

import React, { useState, useCallback, useEffect } from 'react'
import * as XLSX from 'xlsx'
import { supabase, loadTrackerSnapshot } from '../lib/supabase'
import BackToDashboard from '../components/BackToDashboard'
import { ThresholdSettings, DEFAULT_THRESHOLDS, loadThresholdSettings, EmailSettings, DEFAULT_EMAIL, loadEmailSettings } from '../lib/settings'
import { PendingUpdate, SOURCE_LABELS, SOURCE_BADGE_CLASSES, loadPendingUpdates, persistPendingUpdates, upsertPendingUpdate } from '../lib/pendingUpdates'

interface HOP {
  hop: string
  pathId: string
  gc: string
  cm: string
  nokiaPm: string
  regionPm: string
  ops: string
  ms15f: string
  ms15a: string
  ms16f: string
  ms16a: string
  mss: string
  powerUp: string
  mainCutover: string
  divCutover: string
  decom: string
  hasNtp: boolean
  hasMat: boolean
  wpApproved: boolean
  gcPickup: boolean
  gcPickupDate: string
  ntpOwner: string
  ntpWaitingOn: string
  matForecast: string
  matReceived: string
  matLocation: string
  steelFrom: string
  gcPickupF: string
  gcPickupA: string
  cxNotes: string
  ms16fEdited: string
  hasSpo: boolean
  hasCpo: boolean
  vendorWindow: string
  blockers: string[]
  daysOut: number | null
  daysElapsed: number | null
  inProgress: boolean
  complete: boolean
  statuses: string[]
  cmAction: string
}

interface CallNote {
  id: string
  hop_name: string
  note: string
  logged_at: string
}

function parseDate(val: unknown): Date | null {
  if (!val) return null
  if (val instanceof Date) {
    if (val.getFullYear() >= 2025) return val
    return null
  }
  if (typeof val === 'number') {
    const d = new Date((val - 25569) * 86400 * 1000)
    if (d.getFullYear() >= 2025) return d
    return null
  }
  const d = new Date(String(val))
  return isNaN(d.getTime()) || d.getFullYear() < 2025 ? null : d
}

function parseDateAny(val: unknown): Date | null {
  if (!val) return null
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val
  if (typeof val === 'number') {
    // Excel serial date — must be > 40000 to be a valid modern date
    if (val > 40000 && val < 60000) {
      const d = new Date((val - 25569) * 86400 * 1000)
      return isNaN(d.getTime()) ? null : d
    }
    return null
  }
  const s = String(val).trim()
  if (!s || s === 'null' || s === 'undefined' || s === 'NaN') return null
  // Handle ISO strings from Supabase JSON
  const d = new Date(s)
  if (!isNaN(d.getTime())) {
    // Reject dates before 2000 — likely a parsing error
    if (d.getFullYear() < 2000) return null
    return d
  }
  return null
}

function fmtDate(d: Date | null): string {
  if (!d) return ''
  return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`
}

function fmtDM(d: Date | null): string {
  if (!d) return ''
  return `${d.getMonth() + 1}/${d.getDate()}`
}

function daysBetween(a: Date, b: Date): number {
  return Math.round((b.getTime() - a.getTime()) / (1000 * 60 * 60 * 24))
}

function getStatuses(h: HOP, durationAlertDays: number = DEFAULT_THRESHOLDS.durationAlertDays): string[] {
  // Single next-step status based on milestone sequence
  if (h.decom) return ['✅ Complete — verify close-out package']
  if (h.divCutover) return ['♻️ Diversity done — schedule decom pickup and return']
  if (h.mainCutover) return ['🔗 Main cutover done — schedule diversity cutover']
  if (h.powerUp) return ['🔗 Powered up — schedule main cutover']
  if (h.mss) return ['⚡ MSS done — chase Viaero power-up']
  if ((h.daysElapsed ?? 0) > durationAlertDays) return [`⚠️ Over target — ${h.daysElapsed}d elapsed — get updated completion date`]
  return [`🔨 Active — ${h.daysElapsed ?? 0}d elapsed — drive to completion`]
}

function getCmAction(h: HOP, t: ThresholdSettings = DEFAULT_THRESHOLDS): string {
  if (h.complete) return '✅ Complete — verify close-out package uploaded to QB'
  if (h.inProgress) {
    const elapsed = h.daysElapsed ?? 0
    if (h.mss && !h.powerUp) return '📡 MSS/NMS done — chase Viaero for power-up, log wait time'
    if (h.mainCutover && !h.divCutover) return '🔗 Main cutover done — schedule diversity cutover'
    if (h.powerUp && !h.decom) return '♻️ Power-up complete — schedule decom pickup and return'
    if (elapsed > t.durationAlertDays) return `⚠️ ${elapsed}d elapsed — provide updated completion date to CJ`
    return '🔨 Active — confirm crew on site, no access issues, update CJ on M/W/F'
  }
  const days = h.daysOut
  if (!h.hasNtp && days !== null && days <= 7) return '🔴 CRITICAL — NTP missing, starts this week, alert CJ immediately'
  if (!h.hasMat && days !== null && days <= 7) return '🔴 CRITICAL — material not received, starts this week, alert CJ'
  if (h.hasMat && !h.gcPickup && days !== null && days <= 7) return '🔴 URGENT — material in warehouse, coordinate GC pickup today'
  if (!h.hasNtp && days !== null && days <= t.ntpUrgentDays) return '🟠 Chase NTP — starts in 2 weeks, alert CJ'
  if (!h.hasMat && days !== null && days <= t.materialWatchDays) return `🟠 Material not received — forecast ${h.matForecast || 'TBD'}, alert CJ`
  if (h.hasMat && !h.gcPickup && days !== null && days <= t.materialWatchDays) return '🟠 Material in warehouse — coordinate GC pickup this week'
  if (days !== null && days <= 30) return `🟡 Monitor — confirm readiness as start date approaches`
  return '👀 Pipeline — monitor'
}

interface NoteCellProps {
  hop: string
  noteValue: string
  onNoteChange: (hop: string, val: string) => void
  onSave: (hop: string, note: string) => void
}

function NoteCell({ hop, noteValue, onNoteChange, onSave }: NoteCellProps) {
  return (
    <td className="p-2">
      <div className="flex gap-1">
        <input type="text" placeholder="Type note..."
          value={noteValue}
          onChange={(e) => onNoteChange(hop, e.target.value)}
          onKeyDown={(e) => { if (e.key === 'Enter') onSave(hop, noteValue) }}
          className="w-32 bg-gray-800 text-white text-xs rounded px-2 py-1 border border-gray-600 focus:outline-none focus:border-blue-500"
        />
        <button onClick={() => onSave(hop, noteValue)}
          className="bg-blue-700 hover:bg-blue-600 text-white text-xs px-2 py-1 rounded">💾</button>
      </div>
    </td>
  )
}

interface HistoryCellProps {
  hop: string
  noteHistory: Record<string, { id: string; hop_name: string; note: string; logged_at: string }[]>
}

function HistoryCell({ hop, noteHistory }: HistoryCellProps) {
  return (
    <td className="p-2 max-w-48">
      {(noteHistory[hop] || []).length === 0
        ? <span className="text-gray-600 text-xs">No history</span>
        : <div className="flex flex-col gap-1 max-h-20 overflow-y-auto">
            {(noteHistory[hop] || []).slice(0, 5).map((n) => (
              <div key={n.id} className="text-xs">
                <span className="text-gray-500">{new Date(n.logged_at).toLocaleDateString('en-US', { month: 'numeric', day: 'numeric' })} </span>
                <span className="text-gray-300">{n.note}</span>
              </div>
            ))}
          </div>
      }
    </td>
  )
}

interface EditableDateProps {
  hop: string
  field: string
  value: string
  alwaysEditable?: boolean
  editedDates: Record<string, Record<string, string>>
  logDateEdit: (hop: string, field: string, oldVal: string, newVal: string) => void
}

function EditableDate({ hop, field, value, alwaysEditable = false, editedDates, logDateEdit }: EditableDateProps) {
  const edited = editedDates[hop]?.[field]

  const toInputFormat = (dateStr: string) => {
    if (!dateStr || dateStr === 'N/A') return ''
    const parts = dateStr.split('/')
    if (parts.length !== 3) return ''
    return `${parts[2]}-${parts[0].padStart(2, '0')}-${parts[1].padStart(2, '0')}`
  }

  const toDisplayFormat = (dateStr: string) => {
    if (!dateStr) return ''
    const parts = dateStr.split('-')
    if (parts.length !== 3) return ''
    return `${parseInt(parts[1])}/${parseInt(parts[2])}/${parts[0]}`
  }

  if (value && value !== 'N/A' && !edited && !alwaysEditable) {
    return <span className="text-green-400 text-xs font-semibold">{value}</span>
  }

  const display = edited || ''

  if (display === 'N/A') {
    return (
      <div className="flex items-center gap-1">
        <span className="text-gray-500 text-xs">N/A</span>
        <button onClick={() => logDateEdit(hop, field, value, '')}
          className="text-gray-600 hover:text-gray-400 text-xs">✕</button>
      </div>
    )
  }

  return (
    <div className="flex flex-col gap-1">
      {display && (
        <span className="text-yellow-300 text-xs font-semibold">📝 {display}</span>
      )}
      <div className="flex items-center gap-1">
        <input
          type="date"
          value={toInputFormat(display)}
          onChange={(e) => {
            const val = e.target.value
            if (val) {
              logDateEdit(hop, field, value, toDisplayFormat(val))
            }
          }}
          className="text-xs rounded px-1 py-1 border border-gray-600 bg-gray-800 text-gray-300 focus:outline-none focus:border-blue-500 cursor-pointer"
        />
        <button
          onClick={() => logDateEdit(hop, field, value, 'N/A')}
          title="Mark as N/A"
          className="text-gray-500 hover:text-red-400 text-xs px-1 py-1 rounded border border-gray-700 hover:border-red-500 transition-colors">
          N/A
        </button>
      </div>
    </div>
  )
}

interface PipelineSectionProps {
  title: string
  rows: HOP[]
  sessionNotes: Record<string, string>
  setSessionNotes: React.Dispatch<React.SetStateAction<Record<string, string>>>
  saveCallNote: (hop: string, note: string) => void
  noteHistory: Record<string, { id: string; hop_name: string; note: string; logged_at: string }[]>
  editedDates: Record<string, Record<string, string>>
  logDateEdit: (hop: string, field: string, oldVal: string, newVal: string) => void
  setCxNotesModal: (val: { hop: string; notes: string } | null) => void
  showNokiaPm?: boolean
  // Read-only here — crew is assigned/edited in GC Call View and the Tracker
  // Grid (same crew-assign-{hop} Supabase rows), this just displays it for
  // reference.
  crewAssignments: Record<string, string>
}

function PipelineSection({ title, rows, sessionNotes, setSessionNotes, saveCallNote, noteHistory, editedDates, logDateEdit, setCxNotesModal, showNokiaPm, crewAssignments }: PipelineSectionProps) {
  return (
    <div className="mb-8">
      <h3 className="text-base font-semibold text-white mb-3">{title} ({rows.length})</h3>
      {rows.length === 0
        ? <p className="text-gray-500 text-sm">No sites in this window</p>
        : (
          <div className="overflow-x-auto">
            <table className="w-full text-xs">
              <thead>
                <tr className="bg-gray-800 text-gray-400">
                  <th className="text-left p-2">HOP</th>
                  <th className="text-left p-2">Path ID</th>
                  {showNokiaPm && <th className="text-left p-2">Nokia PM</th>}
                  <th className="text-left p-2">GC</th>
                  <th className="text-left p-2">Crew</th>
                  <th className="text-left p-2">Days Out</th>
                  <th className="text-left p-2">NTP</th>
                  <th className="text-left p-2">Mat</th>
                  <th className="text-left p-2">NTP Waiting On</th>
                  <th className="text-left p-2">SPO</th>
                  <th className="text-left p-2">Steel From</th>
                  <th className="text-left p-2">Mat Location</th>
                  <th className="text-left p-2">GC Pickup F</th>
                  <th className="text-left p-2">Edit GC Pickup F</th>
                  <th className="text-left p-2">GC Pickup A</th>
                  <th className="text-left p-2">FC Start</th>
                  <th className="text-left p-2">Edit FC Start</th>
                  <th className="text-left p-2">AC Start</th>
                  <th className="text-left p-2">Vendor Window</th>
                  <th className="text-left p-2">CM Action</th>
                  <th className="text-left p-2">Call Notes</th>
                  <th className="text-left p-2">Notes History</th>
                  <th className="text-left p-2">CX Notes</th>
                </tr>
              </thead>
              <tbody>
                {rows.map(h => {
                  const hasConflict = h.vendorWindow.includes('🔴')
                  const isUrgent    = h.blockers.length > 0 && (h.daysOut ?? 99) <= 7
                  const rowBg       = hasConflict ? 'bg-red-950' : isUrgent ? 'bg-yellow-950' : h.blockers.length === 0 ? 'bg-green-950' : 'bg-gray-900'
                  return (
                    <tr key={h.hop} className={`border-t border-gray-800 ${rowBg}`}>
                      <td className="p-2 font-semibold text-white whitespace-nowrap">{h.hop}</td>
                      <td className="p-2 text-gray-400 text-xs whitespace-nowrap">{h.pathId || '—'}</td>
                      {showNokiaPm && <td className="p-2 text-gray-300 whitespace-nowrap">{h.nokiaPm || '—'}</td>}
                      <td className="p-2 text-gray-300 whitespace-nowrap">{h.gc}</td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{crewAssignments[h.hop] || '—'}</td>
                      <td className={`p-2 font-bold whitespace-nowrap ${(h.daysOut ?? 99) <= 7 ? 'text-red-400' : (h.daysOut ?? 99) <= 14 ? 'text-yellow-400' : 'text-gray-300'}`}>
                        {h.daysOut !== null ? `${h.daysOut}d` : '—'}
                      </td>
                      <td className="p-2">
                        <span
                          title={!h.hasNtp ? (h.ntpWaitingOn || h.ntpOwner || 'NTP pending') : 'NTP confirmed'}
                          className="cursor-help">
                          {h.hasNtp
                            ? <span className="text-green-400 font-bold">✓</span>
                            : <span className="text-red-400 font-bold">✗</span>}
                        </span>
                      </td>
                      <td className="p-2">
                        <span
                          title={!h.hasMat ? (h.matForecast ? `Mat forecast: ${h.matForecast}` : 'No material forecast') : `Mat received: ${h.matReceived}`}
                          className="cursor-help">
                          {h.hasMat
                            ? <span className="text-green-400 font-bold">✓</span>
                            : <span className="text-red-400 font-bold">✗</span>}
                        </span>
                      </td>
                      <td className="p-2 text-xs max-w-48">
                        <span className="text-yellow-300 text-xs" title={h.ntpWaitingOn}>
                          {h.ntpWaitingOn
                            ? (h.ntpWaitingOn.length > 30 ? h.ntpWaitingOn.slice(0, 30) + '...' : h.ntpWaitingOn)
                            : '—'}
                        </span>
                      </td>
                      <td className="p-2 text-xs whitespace-nowrap">
                        <span className={
                          h.hasSpo ? 'text-green-400 font-bold' :
                          h.hasCpo ? 'text-yellow-400 font-bold' :
                          'text-red-400 font-bold'
                        }>
                          {h.hasSpo ? '✓ Issued' : h.hasCpo ? '⚡ Cut Now' : '🔴 Chase CPO'}
                        </span>
                      </td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{h.steelFrom || '—'}</td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{h.matLocation || '—'}</td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">
                        {h.gcPickupF || '—'}
                      </td>
                      <td className="p-2">
                        <EditableDate hop={h.hop} field="GC Material Pick-up (F)" value={h.gcPickupF} alwaysEditable={true} editedDates={editedDates} logDateEdit={logDateEdit} />
                      </td>
                      <td className="p-2">
                        {h.gcPickupA
                          ? <span className="text-green-400 text-xs">✓ {h.gcPickupA}</span>
                          : <EditableDate hop={h.hop} field="GC Material Pick-up (A)" value="" editedDates={editedDates} logDateEdit={logDateEdit} />
                        }
                      </td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{h.ms15f}</td>
                      <td className="p-2">
                        <EditableDate hop={h.hop} field="MS15 Implementation Start F" value={h.ms15f} alwaysEditable={true} editedDates={editedDates} logDateEdit={logDateEdit} />
                      </td>
                      <td className="p-2">
                        <EditableDate hop={h.hop} field="MS15 Implementation Start A" value={h.ms15a} editedDates={editedDates} logDateEdit={logDateEdit} />
                      </td>
                      <td className="p-2 text-xs max-w-40">
                        <span
                          className={h.vendorWindow.includes('🔴') ? 'text-red-400' : h.vendorWindow.includes('⚠️') ? 'text-yellow-400' : 'text-green-400'}
                          title={h.vendorWindow}>
                          {h.vendorWindow.length > 30 ? h.vendorWindow.slice(0, 30) + '...' : h.vendorWindow}
                        </span>
                      </td>
                      <td className="p-2 text-xs max-w-48">
                        <span
                          className={h.cmAction.includes('🔴') || h.cmAction.includes('CRITICAL') ? 'text-red-400' : h.cmAction.includes('🟠') ? 'text-orange-400' : h.cmAction.includes('🟡') ? 'text-yellow-400' : h.cmAction.includes('✅') ? 'text-green-400' : 'text-gray-300'}
                          title={h.cmAction}>
                          {h.cmAction.length > 40 ? h.cmAction.slice(0, 40) + '...' : h.cmAction}
                        </span>
                      </td>
                      <NoteCell
                        hop={h.hop}
                        noteValue={sessionNotes[h.hop] || ''}
                        onNoteChange={(hop, val) => setSessionNotes(n => ({ ...n, [hop]: val }))}
                        onSave={saveCallNote}
                      />
                      <HistoryCell hop={h.hop} noteHistory={noteHistory} />
                      <td className="p-2">
                        {h.cxNotes ? (
                          <button
                            onClick={() => setCxNotesModal({ hop: h.hop, notes: h.cxNotes })}
                            className="bg-gray-700 hover:bg-gray-600 text-white text-xs px-2 py-1 rounded flex items-center gap-1 whitespace-nowrap">
                            📝 {h.cxNotes.split('\n').filter(Boolean).length || 1}
                          </button>
                        ) : (
                          <span className="text-gray-600 text-xs">—</span>
                        )}
                      </td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>
        )
      }
    </div>
  )
}

export default function CMViewPage() {
  const [hops, setHops] = useState<HOP[]>([])
  const [loaded, setLoaded] = useState(false)
  const [fileName, setFileName] = useState('')
  const [selectedCM, setSelectedCM] = useState('')
  const [workloadMode, setWorkloadMode] = useState<'mine' | 'full'>('mine')
  const [thresholds, setThresholds] = useState<ThresholdSettings>(DEFAULT_THRESHOLDS)
  const [emailSettings, setEmailSettings] = useState<EmailSettings>(DEFAULT_EMAIL)

  useEffect(() => {
    loadThresholdSettings().then(setThresholds)
    loadEmailSettings().then(setEmailSettings)
  }, [])
  const [sessionNotes, setSessionNotes] = useState<Record<string, string>>({})
  const [noteHistory, setNoteHistory] = useState<Record<string, CallNote[]>>({})
  const [pmUpdates, setPmUpdates] = useState<PendingUpdate[]>([])
  const [showPmUpdates, setShowPmUpdates] = useState(false)
  const [pmSortField, setPmSortField] = useState<'hop' | 'field'>('hop')
  const [pmSearch, setPmSearch] = useState('')
  const [editedDates, setEditedDates] = useState<Record<string, Record<string, string>>>({})
  const [snapshotTime, setSnapshotTime] = useState<string>('')
  const [cxNotesModal, setCxNotesModal] = useState<{ hop: string; notes: string } | null>(null)
  // hop -> 'Crew 1' | 'Crew 2' | ... — read-only reference here; assigned/
  // edited in GC Call View and the Tracker Grid, same crew-assign-{hop}
  // Supabase rows.
  const [crewAssignments, setCrewAssignments] = useState<Record<string, string>>({})

  useEffect(() => {
    const loadCrewAssignments = async () => {
      const { data } = await supabase.from('pm_updates_cache').select('id, updates').like('id', 'crew-assign-%')
      const map: Record<string, string> = {}
      ;(data || []).forEach(row => {
        const hop = row.id.slice('crew-assign-'.length)
        try {
          const parsed = JSON.parse(row.updates)
          if (parsed?.crew) map[hop] = parsed.crew
        } catch {}
      })
      setCrewAssignments(map)
    }
    loadCrewAssignments()
  }, [])

  const today = new Date()
  // Zero out the time-of-day so daysOut is a clean whole-day count — same fix
  // as the dashboard's computeKPIs (commit 8a32098).
  today.setHours(0, 0, 0, 0)

  useEffect(() => {
    const loadNoteHistory = async () => {
      const { data, error } = await supabase
        .from('hop_call_notes')
        .select('*')
        .order('logged_at', { ascending: false })
      if (error) { console.error('Error loading notes:', error); return }
      if (data) {
        const map: Record<string, CallNote[]> = {}
        data.forEach((row: CallNote) => {
          if (!map[row.hop_name]) map[row.hop_name] = []
          map[row.hop_name].push(row)
        })
        setNoteHistory(map)
      }
    }
    loadNoteHistory()
  }, [])

  // Persisted to the shared pending-updates row (app/lib/pendingUpdates.ts) —
  // same list Tracker and GC Call View read/write, so an edit made here shows
  // up on their "Pending Updates" panel too, and vice versa.
  const persistPmUpdates = persistPendingUpdates

  useEffect(() => {
    loadPendingUpdates().then(setPmUpdates).catch(e => console.error('Error loading PM updates:', e))
  }, [])

  // Dedup on hop+field before appending (source: 'cm' tags where this entry
  // came from) — a second edit/comment for the same HOP+field today replaces
  // the pending entry instead of stacking a duplicate. The full note history
  // itself (hop_call_notes, below) is untouched by this — only what's queued
  // as "still needs to go in the tracker" collapses to the latest one.
  const upsertPmUpdate = (next: { hop: string; field: string; oldValue: string; newValue: string; timestamp: string }) => {
    setPmUpdates(prev => {
      const updated = upsertPendingUpdate(prev, { ...next, source: 'cm' })
      persistPmUpdates(updated)
      return updated
    })
  }

  const toggleUpdateCompleted = (hop: string, field: string, timestamp: string) => {
    setPmUpdates(prev => {
      const next = prev.map(u => (u.hop === hop && u.field === field && u.timestamp === timestamp ? { ...u, completed: !u.completed } : u))
      persistPmUpdates(next)
      return next
    })
  }

  const clearCompletedUpdates = () => {
    setPmUpdates(prev => {
      const next = prev.filter(u => !u.completed)
      persistPmUpdates(next)
      if (next.length === 0) setShowPmUpdates(false)
      return next
    })
  }

  const clearAllUpdates = () => {
    setPmUpdates([])
    persistPmUpdates([])
    setShowPmUpdates(false)
  }

  // CX note entries flow into the SAME unified pending-updates table as
  // milestone edits (field: 'CX Notes') so there's one single "what needs
  // to go back into the tracker" list — not a separate comments queue with
  // its own copy/clear buttons. The full history (hop_call_notes, feeding
  // the inline "Notes History" column) is untouched — this is additive.
  const saveCallNote = async (hop: string, note: string) => {
    if (!note.trim()) return
    const logged_at = new Date().toISOString()
    const { data, error } = await supabase
      .from('hop_call_notes')
      .insert({ hop_name: hop, note: note.trim(), logged_at })
      .select()
    if (error) { console.error('Error saving note:', error); return }
    if (data) {
      setNoteHistory(prev => ({ ...prev, [hop]: [data[0], ...(prev[hop] || [])] }))
      setSessionNotes(prev => ({ ...prev, [hop]: '' }))
    }
    upsertPmUpdate({ hop, field: 'CX Notes', oldValue: '—', newValue: note.trim(), timestamp: logged_at })
  }

  const logDateEdit = (hop: string, field: string, oldVal: string, newVal: string) => {
    if (!newVal || newVal === oldVal) return
    setEditedDates(prev => ({ ...prev, [hop]: { ...(prev[hop] || {}), [field]: newVal } }))
    upsertPmUpdate({ hop, field, oldValue: oldVal, newValue: newVal, timestamp: new Date().toISOString() })
  }

  const logNA = (hop: string, field: string) => {
    logDateEdit(hop, field, '', 'N/A')
  }


  const processRows = useCallback((rows: unknown[][], _filename: string) => {
    // Computed locally (shadowing the render-scoped `today` above) so this
    // callback's identity doesn't change on every render — `new Date()` is a
    // fresh object reference each time, and putting it in the dependency
    // array below made processRows (and the effect that depends on it)
    // re-fire on every single render, including ones from unrelated state
    // changes like typing a call note.
    const today = new Date()
    today.setHours(0, 0, 0, 0)

    let headerRow = -1
    for (let i = 0; i < 10; i++) {
      if ((rows[i] as unknown[])?.some(c => String(c).trim() === 'HOP')) { headerRow = i; break }
    }
    if (headerRow === -1) { alert('Could not find header row'); return }

    const headers = rows[headerRow] as string[]
    const col = (name: string) => headers.findIndex(h => String(h).trim() === name)

    const hopCol      = col('HOP')
    const gcCol       = col('General Contractor')
    const regionPmCol = col('Region PM')
    const nokiaPmCol  = col('Nokia PM')
    const siteCmCol   = col('New CM')
    const opsCol      = col('Viaero Ops Field Ops')
    const don444Col   = col('DON 444')
    const ms15fCol    = col('MS15 Implementation Start F')
    const ms15aCol    = col('MS15 Implementation Start A')
    const ms16fCol    = col('MS16 Implementation Ends F')
    const ms16aCol    = col('MS16 Implementation Ends A')
    const mssCol      = headers.findIndex(h => String(h).trim().replace(/\s+$/, '') === 'MSS Completed NMS Ready'.trim())
    const powerCol    = headers.findIndex(h => String(h).trim() === 'Power-Up Completion')
    const mainCutCol  = col('Main Path Cutover Completed')
    const divCutCol   = col('Diversity Cutover Completed')
    const decomCol    = col('Decom Complete')
    const ntpCol      = col('NTP A')
    const matCol      = headers.findIndex(h => String(h).trim().replace(/\s+$/, '') === 'Material Received A'.trim())
    const matFcCol    = col('Material Forecast +4ish')
    const wpCol       = col('Work Package Approved in QB')
    const pickupCol   = col('GC Material Pick-up (A)')
    const ntpOwnCol   = col('NTP Action Owner')
    const ntpWaitCol  = col('NTP is waiting on')
    const matLocCol   = col('Material Current Location')
      // Steel From (Nokia/ITW) is at a known position — find by exact index match
      const steelCol = (() => {
        // First try exact match with newline
        let idx = headers.findIndex(h => String(h) === 'Steel From\n(Nokia/ITW)')
        if (idx !== -1) return idx
        // Try without newline
        idx = headers.findIndex(h => String(h) === 'Steel From (Nokia/ITW)')
        if (idx !== -1) return idx
        // Try contains Steel From
        idx = headers.findIndex(h => String(h).includes('Steel From'))
        if (idx !== -1) return idx
        // Fallback — known position from tracker
        return 59
      })()
    const gcPickupFCol = headers.findIndex(h => String(h).trim() === 'GC Material Pick-up (F)')
    const gcPickupACol = headers.findIndex(h => String(h).trim() === 'GC Material Pick-up (A)')
    const cxNotesCol   = headers.findIndex(h => String(h).trim().replace(/^'+|'+$/g, '') === 'CX Notes:')
    const pathIdCol = headers.findIndex(h => String(h).trim().replace(/^'+|'+$/g, '') === 'Path ID')
    const spoCol       = headers.findIndex(h => String(h).trim().toLowerCase() === 'cx spo issued')
    const cpoCol       = headers.findIndex(h => String(h).trim().toLowerCase() === 'service cpo received')
    const itwSCol     = col('ITW Schedule Start')
    const itwECol     = col('ITW Schedule Complete')
    const ssSCol      = col('Samsung Schedule Start')
    const ssECol      = col('Samsung Schedule Complete')
    const siteNameCol = col('Site Name')

    const hopRows = new Map<string, unknown[][]>()
    for (let i = headerRow + 1; i < rows.length; i++) {
      const row = rows[i] as unknown[]
      const don = String(row[don444Col] || '').trim().toUpperCase()
      if (don !== 'DON 444') continue
      // Region PM is no longer gated here — collect rows for every PM so the
      // "Full CM Workload" toggle can show them. The CJ-only default is
      // applied at render time via cmHops instead.
      const hop = String(row[hopCol] || '').trim()
      if (!hop || hop === 'undefined') continue
      if (!hopRows.has(hop)) hopRows.set(hop, [])
      hopRows.get(hop)!.push(row)
    }

    const parsed: HOP[] = []
    hopRows.forEach((rows2, hop) => {
      // Prefer the row where GC and Site CM are both populated — the other
      // row for this HOP may be a blank/partial duplicate.
      const row  = rows2.find(r => String(r[gcCol] || '').trim() && String(r[siteCmCol] || '').trim()) || rows2[0]
      const row2 = rows2.find(r => r !== row) || null

      const ms15f    = parseDateAny(row[ms15fCol])
      const ms15a    = parseDate(row[ms15aCol])
      const ms16a    = parseDate(row[ms16aCol])
      const ntpDate  = parseDate(row[ntpCol])
      const matDate  = parseDateAny(row[matCol])
      const wpDate   = parseDateAny(row[wpCol])
      const pickupD  = parseDateAny(row[pickupCol])
      const mssDate  = parseDateAny(row[mssCol])
      const powerDate= parseDateAny(row[powerCol])
      const mainDate = parseDateAny(row[mainCutCol])
      const divDate  = parseDateAny(row[divCutCol])
      const decomDate= parseDateAny(row[decomCol])

      const hasNtp     = !!(ntpDate && ntpDate.getFullYear() >= 2025)
      const hasMat     = !!(matDate && matDate.getFullYear() >= 2020)
      const wpApproved = !!wpDate
      const gcPickup   = !!pickupD
      const started    = !!ms15a
      const complete   = !!ms16a
      const inProgress = started && !complete
      const daysOut    = ms15f ? daysBetween(today, ms15f) : null
      const daysElapsed = inProgress && ms15a ? daysBetween(ms15a, today) : null

      // Vendor window — read both rows
      const allVendorParts: string[] = []
      rows2.forEach(r => {
        const rItwS  = parseDateAny(r[itwSCol])
        const rItwE  = parseDateAny(r[itwECol])
        const rSsS   = parseDateAny(r[ssSCol])
        const rSsE   = parseDateAny(r[ssECol])
        const rMs15f = parseDateAny(r[ms15fCol]) || ms15f
        if (!rMs15f) return
        if (!rSsS && !rSsE && !rItwS && !rItwE) return
        const siteName  = String(r[siteNameCol] || '').trim()
        const siteLabel = siteName ? ` (${siteName})` : ''
        const checkV = (name: string, start: Date | null, end: Date | null) => {
          if (!start || !end) return
          const ms15fTime = rMs15f.getTime()
          const startTime = start.getTime()
          const endTime   = end.getTime()
          if (startTime <= ms15fTime && ms15fTime <= endTime) {
            allVendorParts.push(`🔴 ${name} on site thru ${fmtDM(end)}${siteLabel}`)
          } else if (endTime < ms15fTime) {
            const buf = Math.round((ms15fTime - endTime) / (1000 * 60 * 60 * 24))
            if (buf <= 5)                            allVendorParts.push(`🔴 ${name} clears ${fmtDM(end)} — only ${buf}d${siteLabel}`)
            else if (buf <= thresholds.pullInBufferDays) allVendorParts.push(`⚠️ ${name} clears ${fmtDM(end)} — ${buf}d buffer${siteLabel}`)
            else                                      allVendorParts.push(`✅ ${name} clears ${fmtDM(end)}${siteLabel}`)
          } else {
            const buf = Math.round((startTime - ms15fTime) / (1000 * 60 * 60 * 24))
            if (buf <= thresholds.pullInBufferDays) allVendorParts.push(`⚠️ ${name} starts ${fmtDM(start)} — ${buf}d after start${siteLabel}`)
            else                                  allVendorParts.push(`✅ ${name} starts ${fmtDM(start)}${siteLabel}`)
          }
        }
        checkV('ITW', rItwS, rItwE)
        checkV('Samsung', rSsS, rSsE)
      })

      const itwParts  = Array.from(new Set(allVendorParts.filter(p => p.includes('ITW'))))
      const ssParts   = Array.from(new Set(allVendorParts.filter(p => p.includes('Samsung'))))
      const sortParts = (parts: string[]) => [...parts.filter(p => p.includes('🔴')), ...parts.filter(p => p.includes('⚠️')), ...parts.filter(p => p.includes('✅'))]
      const vendorWindow = [...sortParts(itwParts), ...sortParts(ssParts)].filter(Boolean).join(' | ') || '✅ No conflicts'

      const blockers: string[] = []
      if (!hasNtp) blockers.push('🔴 NTP pending')
      if (!hasMat) blockers.push('🔴 Material not received')
      if (hasMat && !gcPickup) blockers.push('🟠 Mat in warehouse — GC pickup needed')
      if (!wpApproved) blockers.push('🟡 WP not approved')
      if (vendorWindow.includes('🔴')) blockers.push('🔴 Vendor conflict')

      const hopObj: HOP = {
        hop,
        pathId:       String(row[pathIdCol] || '').trim().replace(/^'+|'+$/g, ''),
        gc:           String(row[gcCol] || '').trim() || String(row2?.[gcCol] || '').trim(),
        cm:           String(row[siteCmCol] || '').trim() || String(row2?.[siteCmCol] || '').trim(),
        nokiaPm:      String(row[nokiaPmCol] || '').trim() || String(row2?.[nokiaPmCol] || '').trim(),
        regionPm:     String(row[regionPmCol] || '').trim() || String(row2?.[regionPmCol] || '').trim(),
        ops:          String(row[opsCol] || '').trim(),
        ms15f:        fmtDate(ms15f),
        ms15a:        fmtDate(ms15a),
        ms16f:        fmtDate(parseDateAny(row[ms16fCol])),
        ms16a:        fmtDate(ms16a),
        mss:          fmtDate(mssDate),
        powerUp:      fmtDate(powerDate),
        mainCutover:  fmtDate(mainDate),
        divCutover:   fmtDate(divDate),
        decom:        fmtDate(decomDate),
        hasNtp, hasMat, wpApproved, gcPickup,
        gcPickupDate: fmtDate(pickupD),
        ntpOwner:     String(row[ntpOwnCol] || '').trim() || String(row2?.[ntpOwnCol] || '').trim(),
        ntpWaitingOn: String(row[ntpWaitCol] || '').trim() || String(row2?.[ntpWaitCol] || '').trim(),
        matForecast:  fmtDate(parseDateAny(row[matFcCol])),
        matReceived:  hasMat ? fmtDate(matDate) : '',
        matLocation:  String(row[matLocCol] || '').trim() || String(row2?.[matLocCol] || '').trim(),
        steelFrom: (() => {
            for (let i = 0; i < headers.length; i++) {
              const h = String(headers[i])
              if (h.includes('Steel From') || h.includes('steel from')) {
                const v1 = String(row[i] || '').trim().replace(/^'+|'+$/g, '').trim()
                const v2 = String(row2?.[i] || '').trim().replace(/^'+|'+$/g, '').trim()
                for (const v of [v1, v2]) {
                  if (!v || v === 'nan' || v === 'undefined' || v === ' ') continue
                  if (v.match(/^\d{4}-/) || v.includes('T00:00') || v.includes('T04:00') || !isNaN(Number(v))) continue
                  if (v.match(/^\d+\/\d+\//)) continue
                  return v
                }
              }
            }
            return ''
          })(),
        gcPickupF:    fmtDate(parseDateAny(row[gcPickupFCol])),
        gcPickupA:    fmtDate(parseDateAny(row[gcPickupACol])),
        cxNotes:      String(row[cxNotesCol] || '').trim(),
        ms16fEdited:  '',
        hasSpo: spoCol !== -1 ? !!parseDateAny(row[spoCol]) : false,
        hasCpo: (() => {
          if (cpoCol === -1) return false
          const v = String(row[cpoCol] || '').trim()
          return v.length > 0 && v.toLowerCase() !== 'nan' && v !== ''
        })(),
        vendorWindow, blockers,
        daysOut, daysElapsed, inProgress, complete,
        statuses: [],
        cmAction: ''
      }
      hopObj.statuses = getStatuses(hopObj, thresholds.durationAlertDays)
      hopObj.cmAction = getCmAction(hopObj, thresholds)
      parsed.push(hopObj)
    })

    setHops(parsed)
    setLoaded(true)
  }, [thresholds])

  useEffect(() => {
    const loadFromSnapshot = async () => {
      const snap = await loadTrackerSnapshot()
      if (!snap) return
      console.log('[cm-view] fetched', snap.data.length, 'rows from Supabase')
      console.log('[cm-view] NE-SQUAW_MOUND-NE-CHADRON in fetched data:', snap.data.some(row => row.some(cell => String(cell).trim() === 'NE-SQUAW_MOUND-NE-CHADRON')))
      setSnapshotTime(new Date(snap.uploaded_at).toLocaleDateString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric' }) + ' at ' + new Date(snap.uploaded_at).toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' }))
      setFileName(snap.filename)
      processRows(snap.data, snap.filename)
    }
    loadFromSnapshot()
  }, [processRows])

  const handleFile = useCallback((file: File) => {
    setFileName(file.name)
    const reader = new FileReader()
    reader.onload = (e) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer)
      const wb = XLSX.read(data, { type: 'array', cellDates: true })
      const ws = wb.Sheets['HOPs']
      if (!ws) { alert('HOPs tab not found'); return }
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 }) as unknown[][]
      setSelectedCM('')
      processRows(rows, file.name)
    }
    reader.readAsArrayBuffer(file)
  }, [processRows])

  // "My HOPs Only" preserves the original behavior of scoping to CJ's HOPs;
  // bulk all-CM actions (download/email) always use cjHops regardless of the
  // toggle, since that toggle is scoped to the selected-CM view only.
  const cjHops    = hops.filter(h => h.regionPm?.trim().toUpperCase() === 'CJ')
  // Only list a CM if they have at least one HOP that isn't complete —
  // active + thisWeek + next2Wks + thisMonth + pipeline below is exactly
  // the set of !h.complete HOPs (inProgress and complete are mutually
  // exclusive), so a CM whose every HOP is already complete would show
  // every section empty anyway. Filtering the tab out here avoids
  // presenting names with genuinely nothing to call about.
  const cmList    = Array.from(new Set(
    (workloadMode === 'full' ? hops : cjHops).filter(h => !h.complete).map(h => h.cm?.trim().toLowerCase()).filter(Boolean)
  )).sort()
  // Per-CM active-site count for the tab badge — "active" is the same
  // in-progress set the panel's ACTIVE SITES section uses (started, not yet
  // complete), scoped to whatever the workload toggle currently shows, so
  // the badge tracks the PM filter just like the GC view's outstanding count.
  const cmActiveCounts = (() => {
    const counts = new Map<string, number>()
    ;(workloadMode === 'full' ? hops : cjHops).forEach(h => {
      if (!h.inProgress) return
      const key = h.cm?.trim().toLowerCase()
      if (!key) return
      counts.set(key, (counts.get(key) || 0) + 1)
    })
    return counts
  })()
  const cmHops    = (workloadMode === 'full' ? hops : cjHops)
    .filter(h => h.cm?.trim().toLowerCase() === selectedCM?.trim().toLowerCase())
  const active    = cmHops.filter(h => h.inProgress).sort((a, b) => {
    const aTime = a.ms16f ? new Date(a.ms16f).getTime() : Infinity
    const bTime = b.ms16f ? new Date(b.ms16f).getTime() : Infinity
    return aTime - bTime
  })
  const thisWeek  = cmHops.filter(h => !h.inProgress && !h.complete && h.daysOut !== null && h.daysOut >= 0 && h.daysOut <= 7).sort((a, b) => (a.daysOut ?? 0) - (b.daysOut ?? 0))
  const next2Wks  = cmHops.filter(h => !h.inProgress && !h.complete && h.daysOut !== null && h.daysOut > 7 && h.daysOut <= 14).sort((a, b) => (a.daysOut ?? 0) - (b.daysOut ?? 0))
  const thisMonth = cmHops.filter(h => !h.inProgress && !h.complete && h.daysOut !== null && h.daysOut > 14 && h.daysOut <= 30).sort((a, b) => (a.daysOut ?? 0) - (b.daysOut ?? 0))
  const pipeline  = cmHops.filter(h => !h.inProgress && !h.complete && (h.daysOut === null || h.daysOut > 30)).sort((a, b) => (a.daysOut ?? 999) - (b.daysOut ?? 999))

  const generateEmail = () => {
    const date = today.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })
    const subj = `Viaero MW Program — CM Call Follow-Up | ${selectedCM} | ${date}`
    const div  = '─'.repeat(60)

    let body = `${selectedCM},\n\n`
    body += `Per our call today — here is your site summary and action items for ${date}.\n`
    body += `${div}\n\n`

    if (active.length > 0) {
      body += `ACTIVE SITES (${active.length})\n${div}\n`
      active.forEach(h => {
        body += `• ${h.hop}  |  GC: ${h.gc}\n`
        body += `  Started: ${h.ms15a}  |  FC Complete: ${h.ms16f}\n`
        body += `  Status: ${h.statuses.join('  |  ')}\n`
        if (h.mss)         body += `  MSS/NMS Ready: ${h.mss}\n`
        if (h.powerUp)     body += `  Power-Up: ${h.powerUp}\n`
        if (h.mainCutover) body += `  Main Cutover: ${h.mainCutover}\n`
        if (h.divCutover)  body += `  Diversity Cutover: ${h.divCutover}\n`
        if (sessionNotes[h.hop]) body += `  Note: ${sessionNotes[h.hop]}\n`
        body += '\n'
      })
    }

    const upcoming = [...thisWeek, ...next2Wks]
    if (upcoming.length > 0) {
      body += `STARTING WITHIN 2 WEEKS (${upcoming.length})\n${div}\n`
      upcoming.forEach(h => {
        body += `• ${h.hop}  |  GC: ${h.gc}  |  FC Start: ${h.ms15f}  |  ${h.daysOut}d out\n`
        body += `  NTP: ${h.hasNtp ? '✓' : '✗ Pending'}  |  Material: ${h.hasMat ? '✓' : '✗ Pending'}  |  GC Pickup: ${h.gcPickupDate ? '✓ ' + h.gcPickupDate : '✗'}\n`
        body += `  Steel From: ${h.steelFrom || '—'}  |  Mat Location: ${h.matLocation || '—'}\n`
        if (h.vendorWindow && !h.vendorWindow.includes('✅ No conflicts')) body += `  Vendor: ${h.vendorWindow}\n`
        if (h.blockers.length > 0) body += `  Blockers: ${h.blockers.join(' | ')}\n`
        if (sessionNotes[h.hop]) body += `  Note: ${sessionNotes[h.hop]}\n`
        body += '\n'
      })
    }

    // Action items from call notes
    const actionItems = Object.entries(sessionNotes).filter(([hop, note]) => note.trim() && cmHops.some(h => h.hop === hop))
    if (actionItems.length > 0) {
      body += `ACTION ITEMS FROM TODAY'S CALL\n${div}\n`
      actionItems.forEach(([hop, note], i) => {
        body += `${i + 1}. ${hop} — ${note}\n`
      })
      body += '\n'
    }

    body += `${div}\n`
    body += `Please confirm receipt and advise on any open items.\n\n`
    body += `Respectfully,\nCJ\nNokia Program Manager — Viaero MW Construction Program\nCC: Thomas M. — Lead CM`

    window.open(`mailto:?subject=${encodeURIComponent(subj)}&body=${encodeURIComponent(body)}`)
  }

  const downloadAllCMs = () => {
    try {
      const wb = XLSX.utils.book_new()
      // Always uses the full, PM-unfiltered workload regardless of the toggle.
      const cmNames = Array.from(new Set(hops.map(h => h.cm).filter(Boolean))).sort()
      const date = today.toLocaleDateString('en-US').replace(/\//g, '-')

      cmNames.forEach(cm => {
        const cmHops = hops.filter(h => h.cm === cm && !h.complete)
        if (cmHops.length === 0) return

        const active    = cmHops.filter(h => h.inProgress).sort((a, b) => {
          const aTime = a.ms16f ? new Date(a.ms16f).getTime() : Infinity
          const bTime = b.ms16f ? new Date(b.ms16f).getTime() : Infinity
          return aTime - bTime
        })
        const thisWeek  = cmHops.filter(h => !h.inProgress && h.daysOut !== null && h.daysOut >= 0 && h.daysOut <= 7).sort((a, b) => (a.daysOut ?? 0) - (b.daysOut ?? 0))
        const next2Wks  = cmHops.filter(h => !h.inProgress && h.daysOut !== null && h.daysOut > 7 && h.daysOut <= 14).sort((a, b) => (a.daysOut ?? 0) - (b.daysOut ?? 0))
        const thisMonth = cmHops.filter(h => !h.inProgress && h.daysOut !== null && h.daysOut > 14 && h.daysOut <= 30).sort((a, b) => (a.daysOut ?? 0) - (b.daysOut ?? 0))
        const pipeline  = cmHops.filter(h => !h.inProgress && (h.daysOut === null || h.daysOut > 30)).sort((a, b) => (a.daysOut ?? 999) - (b.daysOut ?? 999))

        const headers = ['HOP', 'Path ID', 'Nokia PM', 'GC', 'Status', 'Days Elapsed', 'Days Out', 'FC Start', 'AC Start', 'FC End', 'AC End', 'NTP', 'Material', 'Steel From', 'Mat Location', 'GC Pickup F', 'GC Pickup A', 'SPO', 'Vendor Window', 'CM Action', 'Latest Note']

        const rows: unknown[][] = []

        const getLatestNote = (hop: string) => {
          const notes = noteHistory[hop] || []
          if (notes.length === 0) return ''
          return `${new Date(notes[0].logged_at).toLocaleDateString('en-US', {month:'numeric',day:'numeric'})}: ${notes[0].note}`
        }

        const addSection = (label: string, sectionHops: typeof cmHops, isActive: boolean) => {
          if (sectionHops.length === 0) return
          rows.push([`--- ${label} (${sectionHops.length}) ---`])
          rows.push(headers)
          sectionHops.forEach(h => {
            const spoStatus = h.hasSpo ? '✓ Issued' : h.hasCpo ? '⚡ Cut Now' : '🔴 Chase CPO'
            const elapsed = isActive && h.daysElapsed !== null ? `${h.daysElapsed}d` : ''
            const daysOut = !isActive && h.daysOut !== null ? `${h.daysOut}d` : ''
            const status = isActive
              ? ((h.daysElapsed ?? 0) > thresholds.durationAlertDays ? '⚠️ OVER TARGET' : '🔨 Active')
              : h.daysOut !== null && h.daysOut <= 7 ? '🔴 This Week'
              : h.daysOut !== null && h.daysOut <= 14 ? '🟠 2 Weeks'
              : h.daysOut !== null && h.daysOut <= 30 ? '🟡 This Month'
              : '🔵 Pipeline'
            rows.push([
              h.hop, h.pathId || '—', h.nokiaPm || '—', h.gc, status, elapsed, daysOut,
              h.ms15f || '—', h.ms15a || '—', h.ms16f || '—', h.ms16a || '—',
              h.hasNtp ? '✓' : '✗',
              h.hasMat ? '✓' : '✗',
              h.steelFrom || '—',
              h.matLocation || '—',
              h.gcPickupF || '—',
              h.gcPickupA || '—',
              spoStatus,
              h.vendorWindow.includes('🔴') ? h.vendorWindow : '✅ Clear',
              h.cmAction,
              getLatestNote(h.hop)
            ])
          })
          rows.push([])
        }

        // Header info
        rows.push([`Site CM: ${cm} — Full Pipeline Report — ${today.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })}`])
        rows.push([`Total HOPs: ${cmHops.length} | Active: ${active.length} | This Week: ${thisWeek.length} | Next 2 Wks: ${next2Wks.length} | This Month: ${thisMonth.length} | Pipeline: ${pipeline.length}`])
        rows.push([])

        addSection('ACTIVE SITES', active, true)
        addSection('THIS WEEK (0-7 days)', thisWeek, false)
        addSection('NEXT 2 WEEKS (8-14 days)', next2Wks, false)
        addSection('THIS MONTH (15-30 days)', thisMonth, false)
        addSection('30D+ PIPELINE', pipeline, false)

        const ws = XLSX.utils.aoa_to_sheet(rows)
        ws['!cols'] = [
          { wch: 36 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 16 }, { wch: 12 }, { wch: 10 },
          { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 },
          { wch: 6 }, { wch: 6 }, { wch: 12 }, { wch: 20 }, { wch: 14 }, { wch: 14 },
          { wch: 14 }, { wch: 30 }, { wch: 45 }, { wch: 50 }
        ]
        XLSX.utils.book_append_sheet(wb, ws, cm.slice(0, 31))
      })

      XLSX.writeFile(wb, `Viaero_CM_Report_${date}.xlsx`)
    } catch (err) {
      console.error('Download error:', err)
      alert('Download failed — please try again')
    }
  }

  const generateAllCMsEmail = () => {
    const fullWorkload = workloadMode === 'full'
    const sourceHops = fullWorkload ? hops : cjHops
    const cmNames = Array.from(new Set(sourceHops.map(h => h.cm).filter(Boolean))).sort()
    const pad = (n: number) => String(n).padStart(2, '0')
    const dateSlash = `${pad(today.getMonth() + 1)}/${pad(today.getDate())}/${today.getFullYear()}`
    const div     = '─'.repeat(60)
    const starDiv = '═'.repeat(60)
    const subj = `Viaero/Nokia MW Program — CM Pipeline Report | ${dateSlash}`

    let body = `Hap, Steve, Benny,\n\n`
    body += `Please find below and attached your weekly updates to include Cx Pipeline, active sites, and action items.\n`
    body += `${div}\n\n`

    cmNames.forEach(cm => {
      const cmHops = sourceHops.filter(h => h.cm === cm && !h.complete)
      if (cmHops.length === 0) return

      const active   = cmHops.filter(h => h.inProgress).sort((a, b) => {
        const aTime = a.ms16f ? new Date(a.ms16f).getTime() : Infinity
        const bTime = b.ms16f ? new Date(b.ms16f).getTime() : Infinity
        return aTime - bTime
      })
      const upcoming = cmHops.filter(h => !h.inProgress && h.daysOut !== null && h.daysOut <= 14).sort((a, b) => (a.daysOut ?? 0) - (b.daysOut ?? 0))

      body += `${starDiv}\n`
      body += `★★★  ${cm.toUpperCase()}  ★★★\n`
      body += `${starDiv}\n\n`

      if (active.length > 0) {
        body += `★★ Active Sites (${active.length}) ★★\n\n`
        active.forEach(h => {
          const status = (h.daysElapsed ?? 0) > thresholds.durationAlertDays
            ? `⚠️ OVER TARGET — ${h.daysElapsed}d elapsed — confirm completion date with crew`
            : `✅ On track — ${h.daysElapsed}d elapsed`
          const spoStatus = h.hasSpo ? '✓ Issued' : h.hasCpo ? '⚡ Cut Now' : '🔴 Chase CPO'
          const latestNote = (noteHistory[h.hop] || []).length > 0
            ? `  💬 Latest Note: ${new Date(noteHistory[h.hop][0].logged_at).toLocaleDateString('en-US', {month:'numeric',day:'numeric'})} — ${noteHistory[h.hop][0].note}`
            : ''
          body += `★ ${h.hop} ★`
          if (h.pathId) body += `  |  Path ID: ${h.pathId}`
          if (fullWorkload) body += `  |  Nokia PM: ${h.nokiaPm || '—'}`
          body += '\n'
          body += `  SPO: ${spoStatus}\n`
          body += `  AC Start: ${h.ms15a || '—'}  |  FC End: ${h.ms16f || '—'}\n`
          body += `  ${status}\n`
          if (latestNote) body += `${latestNote}\n`
          body += '\n'
        })
      }

      if (upcoming.length > 0) {
        body += `★★ Starting Within 2 Weeks (${upcoming.length}) ★★\n\n`
        upcoming.forEach(h => {
          const spoStatus = h.hasSpo ? '✓ Issued' : h.hasCpo ? '⚡ Cut Now' : '🔴 Chase CPO'
          const steelNote = h.steelFrom === 'ITW'
            ? `ITW — confirm ITW delivery schedule`
            : h.steelFrom || '—'
          const latestNote = (noteHistory[h.hop] || []).length > 0
            ? `  💬 Latest Note: ${new Date(noteHistory[h.hop][0].logged_at).toLocaleDateString('en-US', {month:'numeric',day:'numeric'})} — ${noteHistory[h.hop][0].note}`
            : ''
          body += `★ ${h.hop} ★`
          if (h.pathId) body += `  |  Path ID: ${h.pathId}`
          if (fullWorkload) body += `  |  Nokia PM: ${h.nokiaPm || '—'}`
          body += '\n'
          body += `  SPO: ${spoStatus}\n`
          body += `  NTP: ${h.hasNtp ? '✓' : '✗'}`
          if (!h.hasNtp && h.ntpWaitingOn) body += `  |  Waiting On: ${h.ntpWaitingOn}`
          body += '\n'
          body += `  Mat: ${h.hasMat ? '✓' : '✗'}  |  Steel From: ${steelNote}  |  GC Pickup F: ${h.gcPickupF || '—'}  |  GC Pickup A: ${h.gcPickupA || '✗'}\n`
          body += `  Vendor: ${h.vendorWindow.includes('🔴') ? h.vendorWindow : '✅ Clear'}\n`
          if (h.blockers.length > 0) {
            const blockerText = h.blockers.map(b => {
              if (b.includes('NTP') && h.ntpWaitingOn) return `${b} — Waiting On: ${h.ntpWaitingOn}`
              return b
            }).join(' | ')
            body += `  Blockers: ${blockerText}\n`
          }
          body += `  FC Start: ${h.ms15f}  |  Days Out: ${h.daysOut}d\n`
          if (latestNote) body += `${latestNote}\n`
          body += '\n'
        })
      }

      const actionNotes = Object.entries(sessionNotes).filter(([hop, note]) => note.trim() && cmHops.some(h => h.hop === hop))
      if (actionNotes.length > 0) {
        body += `★★ Action Items ★★\n\n`
        actionNotes.forEach(([hop, note], i) => {
          body += `${i + 1}. ${hop} — ${note}\n`
        })
        body += '\n'
      }

      body += `${div}\n\n`
    })

    const ccList = emailSettings.ccList.join(',')
    window.open(`mailto:?cc=${encodeURIComponent(ccList)}&subject=${encodeURIComponent(subj)}&body=${encodeURIComponent(body)}`)
  }

  return (
    <div className="min-h-screen bg-gray-950 text-white p-6">
      <div className="max-w-full mx-auto">

        <BackToDashboard />

        {/* Header */}
        <div className="flex items-center justify-between mb-6">
          <div>
            <h1 className="text-3xl font-bold">CM View</h1>
            <p className="text-gray-400 mt-1">Select a Site CM to view their pipeline, log field updates, and generate a follow-up email.</p>
          </div>
          <div className="flex gap-2">
            {pmUpdates.length > 0 && (
              <button onClick={() => setShowPmUpdates(!showPmUpdates)}
                className="bg-amber-500 hover:bg-amber-600 text-white px-4 py-2 rounded-lg text-sm font-semibold">
                📋 Pending Updates ({pmUpdates.length})
              </button>
            )}
          </div>
        </div>

        {/* PM Daily Updates Panel — consolidated list shared with Tracker
            and GC Call View (app/lib/pendingUpdates.ts): unified table for
            milestone edits AND CX note comments, tagged with which view each
            entry came from, no separate copy/clear-for-comments flow. */}
        {showPmUpdates && pmUpdates.length > 0 && (
          <div className="mb-4 bg-amber-50 border border-amber-300 rounded-xl p-4">
            <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
              <h2 className="text-amber-900 font-bold text-base">📋 PM Daily Updates — All Views Combined</h2>
              <div className="flex gap-2 items-center flex-wrap">
                <input
                  type="text"
                  placeholder="Search HOP or field..."
                  value={pmSearch}
                  onChange={(e) => setPmSearch(e.target.value)}
                  className="bg-white border border-amber-300 text-amber-900 text-xs rounded px-2 py-1 w-44 focus:outline-none focus:border-amber-500"
                />
                <button onClick={() => setPmSortField(prev => prev === 'field' ? 'hop' : 'field')}
                  className="bg-amber-200 hover:bg-amber-300 text-amber-900 text-xs px-3 py-1 rounded font-semibold">
                  Sort by {pmSortField === 'field' ? 'Field ↑' : 'HOP ↑'}
                </button>
              </div>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-amber-700 text-xs">
                    <th className="text-left p-2">Done</th>
                    <th className="text-left p-2">From</th>
                    <th className="text-left p-2">HOP</th>
                    <th className="text-left p-2">Field</th>
                    <th className="text-left p-2">Old Value</th>
                    <th className="text-left p-2">New Value</th>
                    <th className="text-left p-2">Logged At</th>
                  </tr>
                </thead>
                <tbody>
                  {[...pmUpdates]
                    .filter(u => {
                      if (!pmSearch) return true
                      const q = pmSearch.toLowerCase()
                      return u.hop.toLowerCase().includes(q) || u.field.toLowerCase().includes(q)
                    })
                    .sort((a, b) => pmSortField === 'field'
                      ? a.field.localeCompare(b.field)
                      : a.hop.localeCompare(b.hop))
                    .map((u) => (
                    <tr key={`${u.source}-${u.hop}-${u.field}-${u.timestamp}`} className={`border-t border-amber-200 ${u.completed ? 'opacity-40' : ''}`}>
                      <td className="p-2">
                        <input type="checkbox" checked={u.completed || false}
                          onChange={() => toggleUpdateCompleted(u.hop, u.field, u.timestamp)}
                          className="w-4 h-4 cursor-pointer accent-green-600" />
                      </td>
                      <td className="p-2">
                        <span className={`text-xs font-semibold px-2 py-0.5 rounded-full whitespace-nowrap ${SOURCE_BADGE_CLASSES[u.source]}`}>
                          {SOURCE_LABELS[u.source]}
                        </span>
                      </td>
                      <td className={`p-2 font-semibold ${u.completed ? 'line-through text-gray-500' : 'text-gray-900'}`}>{u.hop}</td>
                      <td className={`p-2 ${u.completed ? 'line-through text-gray-500' : 'text-amber-800'}`}>{u.field}</td>
                      <td className="p-2 text-gray-500">{u.oldValue || '—'}</td>
                      <td className={`p-2 font-bold ${u.completed ? 'text-gray-500' : 'text-green-700'}`}>{u.newValue}</td>
                      <td className="p-2 text-gray-500">{new Date(u.timestamp).toLocaleString()}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div className="mt-4 flex gap-3">
              <button onClick={clearCompletedUpdates} className="bg-green-700 hover:bg-green-600 text-white px-4 py-2 rounded text-sm font-semibold">
                ✅ Clear Completed
              </button>
              <button onClick={clearAllUpdates} className="bg-gray-600 hover:bg-gray-500 text-white px-4 py-2 rounded text-sm">
                🗑 Clear All
              </button>
            </div>
          </div>
        )}

        {snapshotTime && (
          <div className="mb-4 bg-gray-900 border border-gray-700 rounded-lg px-4 py-2 flex items-center justify-between">
            <p className="text-green-400 text-sm font-semibold">📡 Live data from {snapshotTime}</p>
            <p className="text-gray-500 text-xs">{fileName} — {cjHops.length} HOPs · Upload new tracker on Dashboard to refresh</p>
          </div>
        )}
        {!snapshotTime && (
          <div className="mb-4 bg-gray-900 border border-gray-700 rounded-lg px-4 py-8 text-center">
            <p className="text-gray-400">No tracker data found — go to Dashboard to upload your tracker</p>
          </div>
        )}

        {/* Workload Toggle */}
        <div className="flex gap-2 mb-4">
          <button onClick={() => setWorkloadMode('mine')}
            className={`px-4 py-2 rounded-lg text-sm font-semibold transition-all ${workloadMode === 'mine' ? 'bg-blue-600 text-white shadow-lg' : 'bg-gray-800 text-gray-300 hover:bg-gray-700'}`}>
            My HOPs Only
          </button>
          <button onClick={() => setWorkloadMode('full')}
            className={`px-4 py-2 rounded-lg text-sm font-semibold transition-all ${workloadMode === 'full' ? 'bg-blue-600 text-white shadow-lg' : 'bg-gray-800 text-gray-300 hover:bg-gray-700'}`}>
            Full CM Workload
          </button>
        </div>

        {/* CM Selector */}
        <div className="flex gap-3 mb-6 flex-wrap">
          {cmList.map((cm) => {
            const isSelected = selectedCM?.trim().toLowerCase() === cm?.trim().toLowerCase()
            const activeCount = cmActiveCounts.get(cm.toLowerCase()) ?? 0
            return (
              <button key={cm} onClick={() => setSelectedCM(cm)}
                title={`${activeCount} active site${activeCount === 1 ? '' : 's'}`}
                className={`flex items-center gap-2 px-6 py-3 rounded-lg font-semibold text-sm transition-all ${isSelected ? 'bg-blue-600 text-white shadow-lg scale-105' : 'bg-gray-800 text-gray-300 hover:bg-gray-700'}`}>
                {cm}
                <span className={`text-xs font-bold px-2 py-0.5 rounded-full ${isSelected ? 'bg-white/20 text-white' : 'bg-gray-700 text-gray-300'}`}>
                  {activeCount}
                </span>
              </button>
            )
          })}
        </div>

        {/* CM Panel */}
        {selectedCM && (
          <div className="bg-gray-900 rounded-xl border border-gray-700 p-6">
            <div className="flex items-center justify-between mb-6">
              <div>
                <h2 className="text-2xl font-bold">{selectedCM}</h2>
                {loaded && (
                  <p className="text-gray-400 mt-1">
                    <span className="text-gray-500">
                      {cmHops.length} HOPs · {active.length} active · {thisWeek.length} this week · {next2Wks.length} next 2 wks · {thisMonth.length} this month · {pipeline.length} pipeline
                    </span>
                  </p>
                )}
              </div>
              <div className="flex gap-2 flex-wrap">
                <button onClick={generateEmail}
                  className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-semibold">
                  ✉️ CM Email
                </button>
                <button onClick={generateAllCMsEmail}
                  className="bg-purple-600 hover:bg-purple-700 text-white px-4 py-2 rounded-lg text-sm font-semibold">
                  ✉️ All CMs Email
                </button>
                <button onClick={downloadAllCMs}
                  className="bg-green-700 hover:bg-green-600 text-white px-4 py-2 rounded-lg text-sm font-semibold">
                  📥 Download All CMs
                </button>
              </div>
            </div>

            {loaded && (
              <div className="space-y-8">

                {/* Active Sites */}
                <div>
                  <h3 className="text-lg font-semibold text-white mb-3">🔨 Active Sites ({active.length})</h3>
                  {active.length === 0
                    ? <p className="text-gray-500 text-sm">No active sites</p>
                    : (
                      <div className="overflow-x-auto">
                        <table className="w-full text-xs">
                          <thead>
                            <tr className="bg-gray-800 text-gray-400">
                              <th className="text-left p-2">HOP</th>
                              <th className="text-left p-2">Path ID</th>
                              {workloadMode === 'full' && <th className="text-left p-2">Nokia PM</th>}
                              <th className="text-left p-2">GC</th>
                              <th className="text-left p-2">Crew</th>
                              <th className="text-left p-2">Started</th>
                              <th className="text-left p-2">FC End</th>
                              <th className="text-left p-2">Days Elapsed</th>
                              <th className="text-left p-2">Status</th>
                              <th className="text-left p-2">MSS</th>
                              <th className="text-left p-2">Power-Up</th>
                              <th className="text-left p-2">Main Cutover</th>
                              <th className="text-left p-2">Diversity Cutover</th>
                              <th className="text-left p-2">MS16 Fc</th>
                              <th className="text-left p-2">Edit MS16 Fc</th>
                              <th className="text-left p-2">MS16 Act</th>
                              <th className="text-left p-2">Decom</th>
                              <th className="text-left p-2">Call Notes</th>
                              <th className="text-left p-2">Notes History</th>
                              <th className="text-left p-2">CX Notes</th>
                            </tr>
                          </thead>
                          <tbody>
                            {active.map(h => (
                              <tr key={h.hop} className={`border-t border-gray-800 ${h.statuses.some(s => s.includes('⚠️')) ? 'bg-red-950' : 'bg-gray-900'}`}>
                                <td className="p-2 font-semibold text-white whitespace-nowrap">{h.hop}</td>
                                <td className="p-2 text-gray-400 text-xs whitespace-nowrap">{h.pathId || '—'}</td>
                                {workloadMode === 'full' && <td className="p-2 text-gray-300 whitespace-nowrap">{h.nokiaPm || '—'}</td>}
                                <td className="p-2 text-gray-300 whitespace-nowrap">{h.gc}</td>
                                <td className="p-2 text-gray-300 whitespace-nowrap">{crewAssignments[h.hop] || '—'}</td>
                                <td className="p-2 text-gray-300 whitespace-nowrap">{h.ms15a}</td>
                                <td className="p-2 text-gray-300 whitespace-nowrap">{h.ms16f}</td>
                                <td className={`p-2 font-bold ${(h.daysElapsed ?? 0) > thresholds.durationAlertDays ? 'text-red-400' : 'text-green-400'}`}>
                                  {h.daysElapsed}d
                                </td>
                                <td className="p-2">
                                  <div className="flex flex-col gap-1">
                                    {h.statuses.map((s, i) => (
                                      <span key={i} className={`text-xs ${s.includes('⚠️') ? 'text-red-400' : s.includes('📡') || s.includes('🔗') || s.includes('♻️') ? 'text-yellow-400' : 'text-green-400'}`}>
                                        {s}
                                      </span>
                                    ))}
                                  </div>
                                </td>
                                <td className="p-2">
                                  <EditableDate hop={h.hop} field="MSS Completed NMS Ready " value={h.mss} editedDates={editedDates} logDateEdit={logDateEdit} />
                                </td>
                                <td className="p-2">
                                  <EditableDate hop={h.hop} field="Power-Up Completion" value={h.powerUp} editedDates={editedDates} logDateEdit={logDateEdit} />
                                </td>
                                <td className="p-2">
                                  <EditableDate hop={h.hop} field="Main Path Cutover Completed" value={h.mainCutover} editedDates={editedDates} logDateEdit={logDateEdit} />
                                </td>
                                <td className="p-2">
                                  <EditableDate hop={h.hop} field="Diversity Cutover Completed" value={h.divCutover} editedDates={editedDates} logDateEdit={logDateEdit} />
                                </td>
                                <td className="p-2 text-gray-300 whitespace-nowrap">{h.ms16f || '—'}</td>
                                <td className="p-2">
                                  <EditableDate hop={h.hop} field="MS16 Implementation Ends F" value={h.ms16f} alwaysEditable={true} editedDates={editedDates} logDateEdit={logDateEdit} />
                                </td>
                                <td className="p-2">
                                  <EditableDate hop={h.hop} field="MS16 Implementation Ends A" value={h.ms16a} editedDates={editedDates} logDateEdit={logDateEdit} />
                                </td>
                                <td className="p-2">
                                  <EditableDate hop={h.hop} field="Decom Complete" value={h.decom} editedDates={editedDates} logDateEdit={logDateEdit} />
                                </td>
                                <NoteCell
                                  hop={h.hop}
                                  noteValue={sessionNotes[h.hop] || ''}
                                  onNoteChange={(hop, val) => setSessionNotes(n => ({ ...n, [hop]: val }))}
                                  onSave={saveCallNote}
                                />
                                <HistoryCell hop={h.hop} noteHistory={noteHistory} />
                                <td className="p-2">
                                  {h.cxNotes ? (
                                    <button
                                      onClick={() => setCxNotesModal({ hop: h.hop, notes: h.cxNotes })}
                                      className="bg-gray-700 hover:bg-gray-600 text-white text-xs px-2 py-1 rounded flex items-center gap-1 whitespace-nowrap">
                                      📝 {h.cxNotes.split('\n').filter(Boolean).length || 1}
                                    </button>
                                  ) : (
                                    <span className="text-gray-600 text-xs">—</span>
                                  )}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    )
                  }
                </div>

                {/* Pipeline Sections */}
                <PipelineSection title="⚡ This Week (0–7 days)" rows={thisWeek} sessionNotes={sessionNotes} setSessionNotes={setSessionNotes} saveCallNote={saveCallNote} noteHistory={noteHistory} editedDates={editedDates} logDateEdit={logDateEdit} setCxNotesModal={setCxNotesModal} showNokiaPm={workloadMode === 'full'} crewAssignments={crewAssignments} />
                <PipelineSection title="🟠 Next 2 Weeks (8–14 days)" rows={next2Wks} sessionNotes={sessionNotes} setSessionNotes={setSessionNotes} saveCallNote={saveCallNote} noteHistory={noteHistory} editedDates={editedDates} logDateEdit={logDateEdit} setCxNotesModal={setCxNotesModal} showNokiaPm={workloadMode === 'full'} crewAssignments={crewAssignments} />
                <PipelineSection title="🟡 This Month (15–30 days)" rows={thisMonth} sessionNotes={sessionNotes} setSessionNotes={setSessionNotes} saveCallNote={saveCallNote} noteHistory={noteHistory} editedDates={editedDates} logDateEdit={logDateEdit} setCxNotesModal={setCxNotesModal} showNokiaPm={workloadMode === 'full'} crewAssignments={crewAssignments} />
                <PipelineSection title="🔵 Full Pipeline (30d+)" rows={pipeline} sessionNotes={sessionNotes} setSessionNotes={setSessionNotes} saveCallNote={saveCallNote} noteHistory={noteHistory} editedDates={editedDates} logDateEdit={logDateEdit} setCxNotesModal={setCxNotesModal} showNokiaPm={workloadMode === 'full'} crewAssignments={crewAssignments} />

              </div>
            )}

            {!loaded && (
              <div className="bg-gray-800 rounded-lg p-8 text-center border border-dashed border-gray-600">
                <p className="text-gray-400">📂 Upload your tracker above to load {selectedCM} pipeline</p>
              </div>
            )}
          </div>
        )}

        {!selectedCM && (
          <div className="bg-gray-900 rounded-xl border border-gray-700 p-12 text-center">
            <p className="text-gray-400 text-xl">👆 Select a CM above to begin</p>
          </div>
        )}

      </div>

      {cxNotesModal && (
        <div className="fixed inset-0 bg-black bg-opacity-80 z-50 flex items-start justify-center pt-20 px-4"
          onClick={() => setCxNotesModal(null)}>
          <div className="bg-gray-900 rounded-xl border border-gray-700 w-full max-w-2xl max-h-96 overflow-hidden"
            onClick={(e) => e.stopPropagation()}>
            <div className="flex items-center justify-between p-4 border-b border-gray-700">
              <h2 className="text-base font-bold text-white">📝 CX Notes — {cxNotesModal.hop}</h2>
              <button onClick={() => setCxNotesModal(null)} className="text-gray-400 hover:text-white text-xl font-bold">✕</button>
            </div>
            <div className="p-4 overflow-y-auto max-h-72">
              {cxNotesModal.notes.split('\n').filter(Boolean).map((line, i) => (
                <div key={i} className={`text-sm py-2 ${i > 0 ? 'border-t border-gray-800' : ''}`}>
                  <span className="text-gray-300">{line.trim()}</span>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
