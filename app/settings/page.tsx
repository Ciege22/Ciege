'use client'

export const dynamic = 'force-dynamic'

import { useState, useEffect } from 'react'
import { supabase, loadTrackerSnapshot } from '../lib/supabase'
import { GC_CONFIG } from '../lib/gcConfig'
import { GR_SORT_OPTIONS } from '../lib/grTracker'
import BackToDashboard from '../components/BackToDashboard'
import {
  ProgramSettings, DEFAULT_PROGRAM, loadProgramSettings, saveProgramSettings,
  ThresholdSettings, DEFAULT_THRESHOLDS, loadThresholdSettings, saveThresholdSettings,
  EmailSettings, loadEmailSettings, saveEmailSettings,
  DisplaySettings, DEFAULT_DISPLAY, loadDisplaySettings, saveDisplaySettings,
} from '../lib/settings'

interface SnapshotInfo {
  filename: string
  uploaded_at: string
}

function SectionCard({ title, subtitle, children }: { title: string; subtitle: string; children: React.ReactNode }) {
  return (
    <div className="bg-gray-900 rounded-xl border border-gray-700 p-6 mb-6">
      <h2 className="text-lg font-bold text-white">{title}</h2>
      <p className="text-gray-500 text-xs mt-1 mb-4">{subtitle}</p>
      {children}
    </div>
  )
}

function SaveButton({ onClick, saved }: { onClick: () => void; saved: boolean }) {
  return (
    <button
      onClick={onClick}
      className="mt-4 bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-semibold transition-colors">
      {saved ? '✓ Saved' : 'Save'}
    </button>
  )
}

function fmtTimestamp(iso: string) {
  const d = new Date(iso)
  return `${d.toLocaleDateString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric' })} at ${d.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })}`
}

export default function SettingsPage() {
  const [loaded, setLoaded] = useState(false)

  const [program, setProgram] = useState<ProgramSettings>(DEFAULT_PROGRAM)
  const [programSaved, setProgramSaved] = useState(false)

  const [thresholds, setThresholds] = useState<ThresholdSettings>(DEFAULT_THRESHOLDS)
  const [thresholdsSaved, setThresholdsSaved] = useState(false)

  const [ccText, setCcText] = useState('')
  const [financeText, setFinanceText] = useState('')
  const [gcContactRows, setGcContactRows] = useState<{ gc: string; email: string }[]>([])
  const [cmContactRows, setCmContactRows] = useState<{ cm: string; email: string }[]>([])
  const [emailSaved, setEmailSaved] = useState(false)

  const [display, setDisplay] = useState<DisplaySettings>(DEFAULT_DISPLAY)
  const [displaySaved, setDisplaySaved] = useState(false)

  const [trackerInfo, setTrackerInfo] = useState<SnapshotInfo | null>(null)
  const [spoInfo, setSpoInfo] = useState<SnapshotInfo | null>(null)
  const [crInfo, setCrInfo] = useState<SnapshotInfo | null>(null)

  useEffect(() => {
    const load = async () => {
      const [p, t, e, d] = await Promise.all([
        loadProgramSettings(), loadThresholdSettings(), loadEmailSettings(), loadDisplaySettings(),
      ])

      // First-run convenience: seed the GC list from the known GC config
      // instead of starting from a blank table.
      setProgram(p.gcs.length === 0 ? { ...p, gcs: GC_CONFIG.map(c => ({ gc: c.gc, cm: '', crewCount: 0 })) } : p)
      setThresholds(t)
      setCcText(e.ccList.join('\n'))
      setFinanceText(e.financeEmails.join('\n'))
      setGcContactRows(Object.entries(e.gcContactEmails).map(([gc, email]) => ({ gc, email })))
      setCmContactRows(Object.entries(e.cmContactEmails || {}).map(([cm, email]) => ({ cm, email })))
      setDisplay(d)

      const snap = await loadTrackerSnapshot()
      if (snap) setTrackerInfo({ filename: snap.filename, uploaded_at: snap.uploaded_at })

      const { data: spoSnap } = await supabase.from('report_snapshots').select('filename, uploaded_at').eq('id', 'spo').single()
      if (spoSnap) setSpoInfo(spoSnap)

      const { data: crSnap } = await supabase.from('report_snapshots').select('filename, uploaded_at').eq('id', 'cr').single()
      if (crSnap) setCrInfo(crSnap)

      setLoaded(true)
    }
    load()
  }, [])

  const flash = (setter: (v: boolean) => void) => {
    setter(true)
    setTimeout(() => setter(false), 2000)
  }

  const saveProgram = async () => {
    await saveProgramSettings(program)
    flash(setProgramSaved)
  }

  const saveThresholds = async () => {
    await saveThresholdSettings(thresholds)
    flash(setThresholdsSaved)
  }

  const saveEmail = async () => {
    const parsed: EmailSettings = {
      ccList: ccText.split('\n').map(s => s.trim()).filter(Boolean),
      financeEmails: financeText.split('\n').map(s => s.trim()).filter(Boolean),
      gcContactEmails: Object.fromEntries(
        gcContactRows.filter(r => r.gc.trim() && r.email.trim()).map(r => [r.gc.trim(), r.email.trim()])
      ),
      cmContactEmails: Object.fromEntries(
        cmContactRows.filter(r => r.cm.trim() && r.email.trim()).map(r => [r.cm.trim(), r.email.trim()])
      ),
    }
    await saveEmailSettings(parsed)
    flash(setEmailSaved)
  }

  const saveDisplay = async () => {
    await saveDisplaySettings(display)
    flash(setDisplaySaved)
  }

  // Program Settings helpers
  const addPm = () => setProgram(p => ({ ...p, nokiaPMs: [...p.nokiaPMs, ''] }))
  const updatePm = (i: number, val: string) => setProgram(p => ({ ...p, nokiaPMs: p.nokiaPMs.map((v, idx) => idx === i ? val : v) }))
  const removePm = (i: number) => setProgram(p => ({ ...p, nokiaPMs: p.nokiaPMs.filter((_, idx) => idx !== i) }))

  const addGc = () => setProgram(p => ({ ...p, gcs: [...p.gcs, { gc: '', cm: '', crewCount: 0 }] }))
  const updateGc = (i: number, field: 'gc' | 'cm' | 'crewCount', val: string) =>
    setProgram(p => ({
      ...p,
      gcs: p.gcs.map((g, idx) => idx === i ? { ...g, [field]: field === 'crewCount' ? Number(val) || 0 : val } : g),
    }))
  const removeGc = (i: number) => setProgram(p => ({ ...p, gcs: p.gcs.filter((_, idx) => idx !== i) }))

  // Email Settings — GC contact rows
  const addGcContact = () => setGcContactRows(r => [...r, { gc: '', email: '' }])
  const updateGcContact = (i: number, field: 'gc' | 'email', val: string) =>
    setGcContactRows(r => r.map((row, idx) => idx === i ? { ...row, [field]: val } : row))
  const removeGcContact = (i: number) => setGcContactRows(r => r.filter((_, idx) => idx !== i))

  // Email Settings — CM contact rows
  const addCmContact = () => setCmContactRows(r => [...r, { cm: '', email: '' }])
  const updateCmContact = (i: number, field: 'cm' | 'email', val: string) =>
    setCmContactRows(r => r.map((row, idx) => idx === i ? { ...row, [field]: val } : row))
  const removeCmContact = (i: number) => setCmContactRows(r => r.filter((_, idx) => idx !== i))

  return (
    <div className="min-h-screen bg-gray-950 text-white p-6">
      <div className="max-w-4xl mx-auto">

        <BackToDashboard />

        <div className="mb-6">
          <h1 className="text-3xl font-bold">Settings</h1>
          <p className="text-gray-400 mt-1">Program configuration, alert thresholds, email recipients, and display defaults</p>
        </div>

        {!loaded && <p className="text-gray-400">Loading settings...</p>}

        {loaded && (
          <>
            {/* Program Settings */}
            <SectionCard title="Program Settings" subtitle="Nokia PMs and GC roster used across the program">
              <div className="mb-6">
                <p className="text-sm font-semibold text-gray-300 mb-2">Nokia PM Names</p>
                <div className="flex flex-col gap-2">
                  {program.nokiaPMs.map((pm, i) => (
                    <div key={i} className="flex gap-2">
                      <input value={pm} onChange={(e) => updatePm(i, e.target.value)}
                        placeholder="e.g. CJ"
                        className="flex-1 bg-gray-800 text-white text-sm rounded px-3 py-2 border border-gray-600 focus:outline-none focus:border-blue-500" />
                      <button onClick={() => removePm(i)}
                        className="text-red-400 hover:text-red-300 text-sm px-3">✕</button>
                    </div>
                  ))}
                  <button onClick={addPm}
                    className="self-start text-blue-400 hover:text-blue-300 text-xs font-semibold mt-1">+ Add PM</button>
                </div>
              </div>

              <div>
                <p className="text-sm font-semibold text-gray-300 mb-2">GC Roster (CM &amp; crew count)</p>
                <div className="flex flex-col gap-2">
                  <div className="flex gap-2 text-xs text-gray-500 px-1">
                    <span className="flex-1">GC Name</span>
                    <span className="flex-1">Site CM</span>
                    <span className="w-24">Crew Count</span>
                    <span className="w-6"></span>
                  </div>
                  {program.gcs.map((g, i) => (
                    <div key={i} className="flex gap-2">
                      <input value={g.gc} onChange={(e) => updateGc(i, 'gc', e.target.value)}
                        placeholder="GC name"
                        className="flex-1 bg-gray-800 text-white text-sm rounded px-3 py-2 border border-gray-600 focus:outline-none focus:border-blue-500" />
                      <input value={g.cm} onChange={(e) => updateGc(i, 'cm', e.target.value)}
                        placeholder="CM name"
                        className="flex-1 bg-gray-800 text-white text-sm rounded px-3 py-2 border border-gray-600 focus:outline-none focus:border-blue-500" />
                      <input type="number" value={g.crewCount} onChange={(e) => updateGc(i, 'crewCount', e.target.value)}
                        className="w-24 bg-gray-800 text-white text-sm rounded px-3 py-2 border border-gray-600 focus:outline-none focus:border-blue-500" />
                      <button onClick={() => removeGc(i)}
                        className="w-6 text-red-400 hover:text-red-300 text-sm">✕</button>
                    </div>
                  ))}
                  <button onClick={addGc}
                    className="self-start text-blue-400 hover:text-blue-300 text-xs font-semibold mt-1">+ Add GC</button>
                </div>
              </div>

              <SaveButton onClick={saveProgram} saved={programSaved} />
            </SectionCard>

            {/* Alert Thresholds */}
            <SectionCard title="Alert Thresholds" subtitle="Day-count windows that drive alerts and status flags across the app">
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="text-sm text-gray-300 block mb-1">NTP Urgent Window (days)</label>
                  <input type="number" value={thresholds.ntpUrgentDays}
                    onChange={(e) => setThresholds(t => ({ ...t, ntpUrgentDays: Number(e.target.value) || 0 }))}
                    className="w-full bg-gray-800 text-white text-sm rounded px-3 py-2 border border-gray-600 focus:outline-none focus:border-blue-500" />
                </div>
                <div>
                  <label className="text-sm text-gray-300 block mb-1">Material Watch Window (days)</label>
                  <input type="number" value={thresholds.materialWatchDays}
                    onChange={(e) => setThresholds(t => ({ ...t, materialWatchDays: Number(e.target.value) || 0 }))}
                    className="w-full bg-gray-800 text-white text-sm rounded px-3 py-2 border border-gray-600 focus:outline-none focus:border-blue-500" />
                </div>
                <div>
                  <label className="text-sm text-gray-300 block mb-1">Duration Alert (days elapsed)</label>
                  <input type="number" value={thresholds.durationAlertDays}
                    onChange={(e) => setThresholds(t => ({ ...t, durationAlertDays: Number(e.target.value) || 0 }))}
                    className="w-full bg-gray-800 text-white text-sm rounded px-3 py-2 border border-gray-600 focus:outline-none focus:border-blue-500" />
                </div>
                <div>
                  <label className="text-sm text-gray-300 block mb-1">Pull-In Opportunity Buffer (days)</label>
                  <input type="number" value={thresholds.pullInBufferDays}
                    onChange={(e) => setThresholds(t => ({ ...t, pullInBufferDays: Number(e.target.value) || 0 }))}
                    className="w-full bg-gray-800 text-white text-sm rounded px-3 py-2 border border-gray-600 focus:outline-none focus:border-blue-500" />
                </div>
                <div>
                  <label className="text-sm text-gray-300 block mb-1">Schedule: Push Window (days)</label>
                  <input type="number" value={thresholds.pushWindow}
                    onChange={(e) => setThresholds(t => ({ ...t, pushWindow: Number(e.target.value) || 0 }))}
                    className="w-full bg-gray-800 text-white text-sm rounded px-3 py-2 border border-gray-600 focus:outline-none focus:border-blue-500" />
                </div>
                <div>
                  <label className="text-sm text-gray-300 block mb-1">Schedule: Push Amount (days)</label>
                  <input type="number" value={thresholds.pushAmount}
                    onChange={(e) => setThresholds(t => ({ ...t, pushAmount: Number(e.target.value) || 0 }))}
                    className="w-full bg-gray-800 text-white text-sm rounded px-3 py-2 border border-gray-600 focus:outline-none focus:border-blue-500" />
                </div>
                <div>
                  <label className="text-sm text-gray-300 block mb-1">Schedule: HOP Duration (days)</label>
                  <input type="number" value={thresholds.hopDuration}
                    onChange={(e) => setThresholds(t => ({ ...t, hopDuration: Number(e.target.value) || 0 }))}
                    className="w-full bg-gray-800 text-white text-sm rounded px-3 py-2 border border-gray-600 focus:outline-none focus:border-blue-500" />
                </div>
                <div>
                  <label className="text-sm text-gray-300 block mb-1">Schedule: Ramp-Up Threshold (HOPs/crew)</label>
                  <input type="number" value={thresholds.rampUpThreshold}
                    onChange={(e) => setThresholds(t => ({ ...t, rampUpThreshold: Number(e.target.value) || 0 }))}
                    className="w-full bg-gray-800 text-white text-sm rounded px-3 py-2 border border-gray-600 focus:outline-none focus:border-blue-500" />
                </div>
                <div>
                  <label className="text-sm text-gray-300 block mb-1">Schedule: Ramp-Up Window (days)</label>
                  <input type="number" value={thresholds.rampUpWindow}
                    onChange={(e) => setThresholds(t => ({ ...t, rampUpWindow: Number(e.target.value) || 0 }))}
                    className="w-full bg-gray-800 text-white text-sm rounded px-3 py-2 border border-gray-600 focus:outline-none focus:border-blue-500" />
                </div>
              </div>
              <SaveButton onClick={saveThresholds} saved={thresholdsSaved} />
            </SectionCard>

            {/* Email Settings */}
            <SectionCard title="Email Settings" subtitle="Recipients used when generating follow-up and GR release emails">
              <div className="grid grid-cols-2 gap-4 mb-6">
                <div>
                  <label className="text-sm text-gray-300 block mb-1">CC List (one email per line)</label>
                  <textarea value={ccText} onChange={(e) => setCcText(e.target.value)} rows={6}
                    className="w-full bg-gray-800 text-white text-sm rounded px-3 py-2 border border-gray-600 focus:outline-none focus:border-blue-500 font-mono" />
                </div>
                <div>
                  <label className="text-sm text-gray-300 block mb-1">Finance Team Emails — GR (one per line)</label>
                  <textarea value={financeText} onChange={(e) => setFinanceText(e.target.value)} rows={6}
                    className="w-full bg-gray-800 text-white text-sm rounded px-3 py-2 border border-gray-600 focus:outline-none focus:border-blue-500 font-mono" />
                </div>
              </div>

              <div>
                <p className="text-sm font-semibold text-gray-300 mb-2">GC Contact Emails</p>
                <div className="flex flex-col gap-2">
                  {gcContactRows.map((row, i) => (
                    <div key={i} className="flex gap-2">
                      <input value={row.gc} onChange={(e) => updateGcContact(i, 'gc', e.target.value)}
                        placeholder="GC name"
                        className="flex-1 bg-gray-800 text-white text-sm rounded px-3 py-2 border border-gray-600 focus:outline-none focus:border-blue-500" />
                      <input value={row.email} onChange={(e) => updateGcContact(i, 'email', e.target.value)}
                        placeholder="contact@example.com"
                        className="flex-1 bg-gray-800 text-white text-sm rounded px-3 py-2 border border-gray-600 focus:outline-none focus:border-blue-500" />
                      <button onClick={() => removeGcContact(i)}
                        className="text-red-400 hover:text-red-300 text-sm px-3">✕</button>
                    </div>
                  ))}
                  <button onClick={addGcContact}
                    className="self-start text-blue-400 hover:text-blue-300 text-xs font-semibold mt-1">+ Add GC Contact</button>
                </div>
              </div>

              <div className="mt-6">
                <p className="text-sm font-semibold text-gray-300 mb-2">CM Contact Emails</p>
                <div className="flex flex-col gap-2">
                  {cmContactRows.map((row, i) => (
                    <div key={i} className="flex gap-2">
                      <input value={row.cm} onChange={(e) => updateCmContact(i, 'cm', e.target.value)}
                        placeholder="CM name"
                        className="flex-1 bg-gray-800 text-white text-sm rounded px-3 py-2 border border-gray-600 focus:outline-none focus:border-blue-500" />
                      <input value={row.email} onChange={(e) => updateCmContact(i, 'email', e.target.value)}
                        placeholder="contact@example.com"
                        className="flex-1 bg-gray-800 text-white text-sm rounded px-3 py-2 border border-gray-600 focus:outline-none focus:border-blue-500" />
                      <button onClick={() => removeCmContact(i)}
                        className="text-red-400 hover:text-red-300 text-sm px-3">✕</button>
                    </div>
                  ))}
                  <button onClick={addCmContact}
                    className="self-start text-blue-400 hover:text-blue-300 text-xs font-semibold mt-1">+ Add CM Contact</button>
                </div>
              </div>

              <SaveButton onClick={saveEmail} saved={emailSaved} />
            </SectionCard>

            {/* Display Preferences */}
            <SectionCard title="Display Preferences" subtitle="Defaults applied when the dashboard loads">
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="text-sm text-gray-300 block mb-1">Default PM Filter</label>
                  <input value={display.defaultPmFilter}
                    onChange={(e) => setDisplay(d => ({ ...d, defaultPmFilter: e.target.value }))}
                    placeholder="ALL"
                    className="w-full bg-gray-800 text-white text-sm rounded px-3 py-2 border border-gray-600 focus:outline-none focus:border-blue-500" />
                  <p className="text-gray-600 text-xs mt-1">Use &quot;ALL&quot; or a Nokia PM initials (e.g. CJ)</p>
                </div>
                <div>
                  <label className="text-sm text-gray-300 block mb-1">Default GR Sort Order</label>
                  <select value={display.defaultSortOrder}
                    onChange={(e) => setDisplay(d => ({ ...d, defaultSortOrder: e.target.value }))}
                    className="w-full bg-gray-800 text-white text-sm rounded px-3 py-2 border border-gray-600 focus:outline-none focus:border-blue-500">
                    {GR_SORT_OPTIONS.map(opt => (
                      <option key={opt.value} value={opt.value}>{opt.label}</option>
                    ))}
                  </select>
                </div>
              </div>
              <SaveButton onClick={saveDisplay} saved={displaySaved} />
            </SectionCard>

            {/* Data Management */}
            <SectionCard title="Data Management" subtitle="Most recent uploads on file — read-only">
              <div className="grid grid-cols-3 gap-4 text-sm">
                <div className="bg-gray-800 rounded-lg p-4">
                  <p className="text-gray-500 text-xs uppercase font-semibold mb-1">Tracker</p>
                  {trackerInfo ? (
                    <>
                      <p className="text-white">{trackerInfo.filename}</p>
                      <p className="text-gray-500 text-xs mt-1">{fmtTimestamp(trackerInfo.uploaded_at)}</p>
                    </>
                  ) : <p className="text-gray-600">No upload found</p>}
                </div>
                <div className="bg-gray-800 rounded-lg p-4">
                  <p className="text-gray-500 text-xs uppercase font-semibold mb-1">SPO Report</p>
                  {spoInfo ? (
                    <>
                      <p className="text-white">{spoInfo.filename}</p>
                      <p className="text-gray-500 text-xs mt-1">{fmtTimestamp(spoInfo.uploaded_at)}</p>
                    </>
                  ) : <p className="text-gray-600">No upload found</p>}
                </div>
                <div className="bg-gray-800 rounded-lg p-4">
                  <p className="text-gray-500 text-xs uppercase font-semibold mb-1">CR Tracker</p>
                  {crInfo ? (
                    <>
                      <p className="text-white">{crInfo.filename}</p>
                      <p className="text-gray-500 text-xs mt-1">{fmtTimestamp(crInfo.uploaded_at)}</p>
                    </>
                  ) : <p className="text-gray-600">No upload found</p>}
                </div>
              </div>
            </SectionCard>
          </>
        )}
      </div>
    </div>
  )
}
