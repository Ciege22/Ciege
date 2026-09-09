'use client'

import { useState, useEffect, useCallback, useMemo } from 'react'
import * as XLSX from 'xlsx'
import JSZip from 'jszip'
import { saveChunkedReport, loadChunkedReport } from '../lib/reportChunks'
import {
  ScopSettings, DEFAULT_SCOP, loadScopSettings,
  EmailSettings, DEFAULT_EMAIL, loadEmailSettings,
} from '../lib/settings'
import {
  ScopRow, resolveScopCalcSettings, keyScopRows, detectScopHeaderRow,
  buildScopDataset, computeQuickBaseView, computePathwaveView, computeOverallView,
  buildItemizedMissingItems, buildAllGCReports, buildScopGcEmailMailto,
  pathwaveGcItems, QUICKBASE_ALL_ITEMS, agingColor, ScopGcReport,
} from '../lib/scop'
import { buildScopMasterWorkbook, buildScopGcReportWorkbook, downloadWorkbook } from '../lib/scopReport'
import { buildScopDeck } from '../lib/scopDeck'

// Deck design tokens (spec §4) — no red in card/section colors; red is only
// the aging number in tables.
const T = {
  navy: '#124191', teal: '#00A0B0', green: '#1E8449', amber: '#B7791F',
  grey: '#6B7280', greyBar: '#D0D5DD',
}

interface ScopInfo { filename: string; uploaded_at: string; row_count: number }

function fmtDate(d: Date | null): string {
  return d && !isNaN(d.getTime()) ? d.toLocaleDateString('en-US') : '—'
}

// ── small building blocks ───────────────────────────────────────────────────

function KpiCard({ label, value, color, sub }: { label: string; value: number | string; color: string; sub?: string }) {
  return (
    <div className="flex-1 rounded-xl border border-gray-700 bg-gray-900 p-4 text-center min-w-[150px]">
      <div className="text-3xl font-bold" style={{ color }}>{value}</div>
      <div className="text-xs font-semibold text-gray-200 mt-1">{label}</div>
      {sub && <div className="text-[11px] text-gray-500 mt-0.5">{sub}</div>}
    </div>
  )
}

function ProportionalBar({ segments, caption }: { segments: { value: number; color: string; label: string }[]; caption: string }) {
  const total = segments.reduce((s, x) => s + x.value, 0) || 1
  return (
    <div>
      <div className="flex h-7 w-full overflow-hidden rounded-lg border border-gray-700">
        {segments.map((s, i) => (
          <div key={i} title={`${s.label}: ${s.value}`}
            style={{ width: `${(s.value / total) * 100}%`, backgroundColor: s.color }}
            className="flex items-center justify-center text-[10px] font-bold text-white/90">
            {s.value / total > 0.06 ? s.value : ''}
          </div>
        ))}
      </div>
      <div className="mt-1.5 flex flex-wrap gap-x-4 gap-y-1">
        {segments.map((s, i) => (
          <span key={i} className="flex items-center gap-1.5 text-[11px] text-gray-400">
            <span className="inline-block h-2.5 w-2.5 rounded-sm" style={{ backgroundColor: s.color }} />
            {s.label} ({s.value})
          </span>
        ))}
      </div>
      <p className="mt-1 text-xs italic text-gray-500">{caption}</p>
    </div>
  )
}

function AgingCell({ days, thresholds }: { days: number | null; thresholds: { green: number; amber: number } }) {
  const c = agingColor(days, thresholds)
  const cls = c === 'red' ? 'text-red-400' : c === 'amber' ? 'text-amber-400' : 'text-green-400'
  return <span className={`font-bold ${cls}`}>{days ?? '—'}</span>
}

// ── the three dashboard views ───────────────────────────────────────────────

function View1QuickBase({ dataset }: { dataset: ScopRow[] }) {
  const v = computeQuickBaseView(dataset)
  return (
    <div className="space-y-5">
      <div>
        <h3 className="text-lg font-bold text-white">QuickBase Status</h3>
        <p className="text-sm text-gray-400">
          {v.percentComplete}% complete — {v.complete} of {v.trackedTotal} construction-complete HOPs
        </p>
      </div>
      <div className="flex flex-wrap gap-3">
        <KpiCard label="Complete" value={v.complete} color={T.green} />
        <KpiCard label="In Progress" value={v.pending} color={T.amber} />
        <KpiCard label="Pending HOP Completion" value={v.pendingHopCompletion} color={T.grey} />
      </div>
      <ProportionalBar
        segments={[
          { value: v.complete, color: T.green, label: 'Complete' },
          { value: v.pending, color: T.amber, label: 'In Progress' },
          { value: v.pendingHopCompletion, color: T.greyBar, label: 'Pending HOP Completion' },
        ]}
        caption={`Full program view — all ${v.total} HOPs.`}
      />
    </div>
  )
}

function View2Pathwave({ dataset }: { dataset: ScopRow[] }) {
  const v = computePathwaveView(dataset)
  const byGcRows = Object.entries(v.byGC)
    .map(([gc, b]) => ({ gc, ...b }))
    .sort((a, b) => b.inProgress - a.inProgress)

  return (
    <div className="space-y-5">
      <h3 className="text-lg font-bold text-white">Pathwave SCOP — Contractor Action Items</h3>
      <div className="flex flex-wrap gap-3">
        <KpiCard label="Complete & Approved" value={v.complete} color={T.green} />
        <KpiCard label="In Progress — GC Action" value={v.inProgress} color={T.amber} />
        <KpiCard label="OAD — Tracked Separately" value={v.oad} color={T.navy} />
        <KpiCard label="Pending HOP Completion" value={v.pendingHopCompletion} color={T.grey} />
      </div>
      <ProportionalBar
        segments={[
          { value: v.complete, color: T.green, label: 'Complete & Approved' },
          { value: v.inProgress, color: T.amber, label: 'In Progress — GC Action' },
          { value: v.oad, color: T.navy, label: 'OAD — Tracked Separately' },
          { value: v.pendingHopCompletion, color: T.greyBar, label: 'Pending HOP Completion' },
        ]}
        caption={`Full program view — all ${v.total} HOPs.`}
      />

      <div className="grid gap-5 lg:grid-cols-2">
        {/* Left: Pathwave Items by GC */}
        <div>
          <h4 className="mb-2 text-sm font-semibold text-gray-300">Pathwave Items by GC</h4>
          <div className="overflow-x-auto rounded-xl border border-gray-700 bg-gray-900">
            <table className="w-full text-xs">
              <thead>
                <tr className="bg-gray-800 text-gray-400">
                  <th className="p-2 text-left">GC</th>
                  <th className="p-2 text-left">In Progress</th>
                  <th className="p-2 text-left">Complete</th>
                </tr>
              </thead>
              <tbody>
                {byGcRows.map(r => (
                  <tr key={r.gc} className="border-t border-gray-800">
                    <td className="p-2 font-semibold text-white whitespace-nowrap">{r.gc}</td>
                    <td className="p-2 text-amber-400 font-bold">{r.inProgress}</td>
                    <td className="p-2 text-green-400">{r.complete}</td>
                  </tr>
                ))}
                {byGcRows.length === 0 && <tr><td colSpan={3} className="p-4 text-center text-gray-500">No construction-complete Pathwave rows</td></tr>}
              </tbody>
            </table>
          </div>
        </div>

        {/* Right: OAD Sites — deck layout is exactly HOP / GC / verbatim comment */}
        <div>
          <h4 className="mb-2 text-sm font-semibold text-gray-300">OAD Sites (Awaiting OAD, Not a GC Item)</h4>
          <div className="overflow-x-auto rounded-xl border border-gray-700 bg-gray-900">
            <table className="w-full text-xs">
              <thead>
                <tr className="bg-gray-800 text-gray-400">
                  <th className="p-2 text-left">HOP</th>
                  <th className="p-2 text-left">GC</th>
                  <th className="p-2 text-left">One and Done (verbatim)</th>
                </tr>
              </thead>
              <tbody>
                {v.oadSites.map(r => (
                  <tr key={r.hop} className="border-t border-gray-800">
                    <td className="p-2 font-semibold text-white whitespace-nowrap">{r.hop}</td>
                    <td className="p-2 text-gray-300 whitespace-nowrap">{r.gc}</td>
                    <td className="p-2 text-gray-400">{r.note || '—'}</td>
                  </tr>
                ))}
                {v.oadSites.length === 0 && <tr><td colSpan={3} className="p-4 text-center text-gray-500">No OAD sites</td></tr>}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  )
}

function View3Overall({ dataset }: { dataset: ScopRow[] }) {
  const v = computeOverallView(dataset)
  return (
    <div className="space-y-5">
      <h3 className="text-lg font-bold text-white">Overall SCOP Status — Program Summary</h3>
      <p className="text-sm text-gray-400">Scoped to the {v.trackedTotal} construction-complete HOPs.</p>

      <div className="grid gap-5 md:grid-cols-2">
        <div>
          <h4 className="mb-2 text-sm font-semibold text-gray-300">PATHWAVE</h4>
          <div className="flex gap-3">
            <KpiCard label="Complete" value={v.pathwaveComplete} color={T.green} />
            <KpiCard label="Pending" value={v.pathwavePending} color={T.amber} />
          </div>
        </div>
        <div>
          <h4 className="mb-2 text-sm font-semibold text-gray-300">QUICKBASE</h4>
          <div className="flex gap-3">
            <KpiCard label="Complete" value={v.quickbaseComplete} color={T.green} />
            <KpiCard label="Pending" value={v.quickbasePending} color={T.amber} />
          </div>
        </div>
      </div>

      <div>
        <h4 className="mb-2 text-sm font-semibold text-gray-300">Both Complete — Full SCOP Close-Out</h4>
        <div className="flex gap-3 md:max-w-md">
          <KpiCard label="100% Complete" value={v.fullyComplete} color={T.green} />
          <KpiCard label="Not Yet Fully Complete" value={v.notFullyComplete} color={T.grey} />
        </div>
      </div>
    </div>
  )
}

// ── Master Report (filterable table) ────────────────────────────────────────

type MasterSegment =
  | 'summary' | 'outstanding' | 'pathwave-pending' | 'pathwave-complete'
  | 'oad' | 'qb-pending' | 'qb-complete' | 'fully-complete'

const SEGMENTS: { id: MasterSegment; label: string }[] = [
  { id: 'summary', label: 'Summary' },
  { id: 'outstanding', label: 'Outstanding List' },
  { id: 'pathwave-pending', label: 'Pathwave Pending' },
  { id: 'pathwave-complete', label: 'Pathwave Complete' },
  { id: 'oad', label: 'Pathwave OAD Sites' },
  { id: 'qb-pending', label: 'QuickBase Pending' },
  { id: 'qb-complete', label: 'QuickBase Complete' },
  { id: 'fully-complete', label: 'Fully Complete' },
]

function SimpleTable({ rows }: { rows: ScopRow[] }) {
  return (
    <div className="overflow-x-auto rounded-xl border border-gray-700 bg-gray-900">
      <table className="w-full text-xs">
        <thead>
          <tr className="bg-gray-800 text-gray-400">
            {['HOP', 'Path ID', 'GC', 'CM', 'Near Site A', 'Far Site B'].map(h => <th key={h} className="p-2 text-left">{h}</th>)}
          </tr>
        </thead>
        <tbody>
          {rows.map(r => (
            <tr key={`${r.hop}-${r.pathId}`} className="border-t border-gray-800">
              <td className="p-2 font-semibold text-white whitespace-nowrap">{r.hop}</td>
              <td className="p-2 text-gray-400 whitespace-nowrap">{r.pathId || '—'}</td>
              <td className="p-2 text-gray-300 whitespace-nowrap">{r.gc || '—'}</td>
              <td className="p-2 text-gray-300 whitespace-nowrap">{r.cm || '—'}</td>
              <td className="p-2 text-gray-300 whitespace-nowrap">{r.nearSiteA || '—'}</td>
              <td className="p-2 text-gray-300 whitespace-nowrap">{r.farSiteB || '—'}</td>
            </tr>
          ))}
          {rows.length === 0 && <tr><td colSpan={6} className="p-4 text-center text-gray-500">No rows</td></tr>}
        </tbody>
      </table>
    </div>
  )
}

// ── main component ─────────────────────────────────────────────────────────

export default function ScopTab() {
  const [storeRows, setStoreRows] = useState<unknown[][]>([])
  const [info, setInfo] = useState<ScopInfo | null>(null)
  const [uploading, setUploading] = useState(false)
  const [scopSettings, setScopSettings] = useState<ScopSettings>(DEFAULT_SCOP)
  const [emailSettings, setEmailSettings] = useState<EmailSettings>(DEFAULT_EMAIL)
  const [subTab, setSubTab] = useState<'views' | 'master' | 'gc'>('views')
  const [viewIdx, setViewIdx] = useState<1 | 2 | 3>(1)
  const [segment, setSegment] = useState<MasterSegment>('summary')
  const [gcSearch, setGcSearch] = useState('')

  useEffect(() => {
    loadScopSettings().then(setScopSettings)
    loadEmailSettings().then(setEmailSettings)
    loadChunkedReport('scop').then(r => {
      if (r) {
        setStoreRows(r.rows)
        setInfo({ filename: r.filename, uploaded_at: r.uploaded_at, row_count: Math.max(0, r.rows.length - 1) })
      }
    })
  }, [])

  const calc = useMemo(() => resolveScopCalcSettings(scopSettings), [scopSettings])
  const dataset = useMemo<ScopRow[]>(() => {
    if (storeRows.length < 2) return []
    return buildScopDataset(keyScopRows(storeRows), calc)
  }, [storeRows, calc])

  const gcReports = useMemo(() => (dataset.length ? buildAllGCReports(dataset, calc) : []), [dataset, calc])
  const thresholds = scopSettings.agingColorThresholds

  const handleUpload = useCallback(async (file: File) => {
    setUploading(true)
    const reader = new FileReader()
    reader.onload = async (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer)
        const wb = XLSX.read(data, { type: 'array', cellDates: true })
        const ws = wb.Sheets['HOPs'] || wb.Sheets[wb.SheetNames.find(n => n.trim().toLowerCase() === 'hops') || '']
        if (!ws) { alert('Could not find a "HOPs" tab in this file.'); setUploading(false); return }

        const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as unknown[][]
        const override = scopSettings.headerRowOverride
        const headerIdx = override && override > 0 ? override - 1 : detectScopHeaderRow(rawRows)
        if (headerIdx < 0) {
          alert('Could not auto-detect the header row (looking for a row with both "HOP" and "General Contractor"). Set an explicit header row in Settings → SCOP and retry.')
          setUploading(false); return
        }
        const dataRows = rawRows.slice(headerIdx + 1).filter(r => (r || []).some(v => v !== null && v !== undefined && String(v).trim() !== ''))
        const toStore = [rawRows[headerIdx], ...dataRows]

        const { error } = await saveChunkedReport('scop', file.name, toStore)
        if (error) { console.error('[scop-upload] save failed', error); alert('SCOP upload failed to save — check console.'); setUploading(false); return }

        setStoreRows(toStore)
        setInfo({ filename: file.name, uploaded_at: new Date().toISOString(), row_count: dataRows.length })
      } catch (err) {
        console.error('[scop-upload] error', err)
        alert('SCOP upload failed — check the file format.')
      }
      setUploading(false)
    }
    reader.readAsArrayBuffer(file)
  }, [scopSettings.headerRowOverride])

  const today = fmtDate(calc.asOfDate).replace(/\//g, '-')

  // Master Report segment content
  const outstanding = useMemo(() => dataset.filter(r => !r.isConstructionComplete), [dataset])
  const pathwavePendingItemized = useMemo(
    () => buildItemizedMissingItems(dataset.filter(r => r.pathwaveStatus === 'IN_PROGRESS'), pathwaveGcItems(calc)),
    [dataset, calc])
  const qbPendingItemized = useMemo(
    () => buildItemizedMissingItems(dataset.filter(r => r.quickbaseStatus === 'PENDING'), QUICKBASE_ALL_ITEMS),
    [dataset])
  const pw = useMemo(() => computePathwaveView(dataset), [dataset])
  const qbv = useMemo(() => computeQuickBaseView(dataset), [dataset])
  const ov = useMemo(() => computeOverallView(dataset), [dataset])

  return (
    <div className="space-y-6">
      {/* Upload */}
      <div className="max-w-md">
        <p className="mb-2 text-sm font-semibold text-gray-400">SCOP TRACKER (COP tracker, HOPs tab)</p>
        <div
          className="cursor-pointer rounded-xl border-2 border-dashed border-gray-600 p-5 transition-colors hover:border-blue-500"
          onDragOver={(e) => e.preventDefault()}
          onDrop={(e) => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) handleUpload(f) }}
          onClick={() => document.getElementById('scop-upload')?.click()}
        >
          <input id="scop-upload" type="file" accept=".xlsx" className="hidden"
            onChange={(e) => { const f = e.target.files?.[0]; if (f) handleUpload(f) }} />
          {uploading
            ? <p className="text-center text-sm text-blue-400">⏳ Processing…</p>
            : info
            ? <div>
                <p className="text-sm font-semibold text-green-400">✅ SCOP tracker loaded</p>
                <p className="mt-1 text-xs text-gray-500">
                  {info.filename} · {info.row_count} rows · {new Date(info.uploaded_at).toLocaleDateString('en-US')} at {new Date(info.uploaded_at).toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })}
                </p>
              </div>
            : <p className="text-center text-sm text-gray-400">📂 Upload SCOP tracker</p>}
        </div>
      </div>

      {dataset.length === 0 && !uploading && (
        <div className="rounded-xl border border-gray-700 bg-gray-900 p-12 text-center">
          <p className="text-xl text-gray-400">📂 Upload the COP tracker above to build the SCOP views</p>
        </div>
      )}

      {dataset.length > 0 && (
        <>
          {/* scope banner */}
          <div className="flex flex-wrap items-center justify-between gap-2 rounded-lg border border-gray-700 bg-gray-900 px-4 py-2">
            <p className="text-sm font-semibold text-green-400">
              {dataset.length} Nokia HOPs in scope · {ov.trackedTotal} construction-complete
              {calc.asOfDate.toISOString().slice(0, 10) !== new Date().toISOString().slice(0, 10) && (
                <span className="ml-2 text-amber-400">· as-of {fmtDate(calc.asOfDate)} (override)</span>
              )}
            </p>
            <div className="flex flex-wrap gap-2">
              <button
                onClick={() => downloadWorkbook(buildScopMasterWorkbook(dataset, calc), `SCOP_Master_Report_${today}.xlsx`)}
                className="rounded-lg bg-green-700 px-3 py-1.5 text-xs font-semibold hover:bg-green-600">
                ⬇️ Export Master Report (.xlsx)
              </button>
              <button
                onClick={() => buildScopDeck({
                  totalNokiaHops: dataset.length,
                  constructionComplete: ov.trackedTotal,
                  notYetConstructed: dataset.length - ov.trackedTotal,
                  qbView: qbv, pwView: pw, overallView: ov,
                  generatedDate: calc.asOfDate,
                }, `SCOP_Status_Deck_${today}.pptx`)}
                className="rounded-lg bg-purple-700 px-3 py-1.5 text-xs font-semibold hover:bg-purple-600">
                ⬇️ Export Status Deck (.pptx)
              </button>
            </div>
          </div>

          {/* sub-tabs */}
          <div className="flex flex-wrap gap-2">
            {([['views', '3 Views'], ['master', 'Master Report'], ['gc', 'GC Reports & Emails']] as const).map(([id, label]) => (
              <button key={id} onClick={() => setSubTab(id)}
                className={`rounded-lg px-4 py-2 text-sm font-semibold transition-all ${subTab === id ? 'bg-blue-600 text-white' : 'bg-gray-800 text-gray-300 hover:bg-gray-700'}`}>
                {label}
              </button>
            ))}
          </div>

          {subTab === 'views' && (
            <div className="space-y-5">
              <div className="flex gap-2">
                {([1, 2, 3] as const).map(n => (
                  <button key={n} onClick={() => setViewIdx(n)}
                    className={`rounded-lg px-3 py-1.5 text-xs font-semibold ${viewIdx === n ? 'bg-blue-600 text-white' : 'bg-gray-800 text-gray-300 hover:bg-gray-700'}`}>
                    View {n}
                  </button>
                ))}
              </div>
              {viewIdx === 1 && <View1QuickBase dataset={dataset} />}
              {viewIdx === 2 && <View2Pathwave dataset={dataset} />}
              {viewIdx === 3 && <View3Overall dataset={dataset} />}
            </div>
          )}

          {subTab === 'master' && (
            <div className="space-y-4">
              <div className="flex flex-wrap gap-2">
                {SEGMENTS.map(s => (
                  <button key={s.id} onClick={() => setSegment(s.id)}
                    className={`rounded-lg px-3 py-1.5 text-xs font-semibold ${segment === s.id ? 'bg-blue-600 text-white' : 'bg-gray-800 text-gray-300 hover:bg-gray-700'}`}>
                    {s.label}
                  </button>
                ))}
              </div>

              {segment === 'summary' && (
                <div className="overflow-x-auto rounded-xl border border-gray-700 bg-gray-900">
                  <table className="w-full text-sm">
                    <tbody>
                      {[
                        ['Total Nokia HOPs', dataset.length],
                        ['Construction Complete (tracked)', ov.trackedTotal],
                        ['Pending HOP Completion', pw.pendingHopCompletion],
                        ['— Pathwave: Complete & Approved', pw.complete],
                        ['— Pathwave: In Progress — GC Action', pw.inProgress],
                        ['— Pathwave: OAD — Tracked Separately', pw.oad],
                        ['— Pathwave: Pending (In Progress + OAD, View 3)', ov.pathwavePending],
                        ['— QuickBase: Complete', qbv.complete],
                        ['— QuickBase: Pending', qbv.pending],
                        ['— QuickBase: % Complete (of tracked)', `${qbv.percentComplete}%`],
                        ['Fully Complete (both, within tracked)', ov.fullyComplete],
                        ['Not Yet Fully Complete (within tracked)', ov.notFullyComplete],
                      ].map(([k, val]) => (
                        <tr key={String(k)} className="border-t border-gray-800">
                          <td className="p-2 text-gray-300">{k}</td>
                          <td className="p-2 text-right font-bold text-white">{val}</td>
                        </tr>
                      ))}
                      <tr className="border-t border-gray-700 bg-gray-800">
                        <td className="p-2 font-semibold text-gray-300">Cross-foot: Pathwave 4-way = Total</td>
                        <td className={`p-2 text-right font-bold ${pw.complete + pw.inProgress + pw.oad + pw.pendingHopCompletion === dataset.length ? 'text-green-400' : 'text-red-400'}`}>
                          {pw.complete + pw.inProgress + pw.oad + pw.pendingHopCompletion === dataset.length ? 'OK' : 'MISMATCH'}
                        </td>
                      </tr>
                      <tr className="bg-gray-800">
                        <td className="p-2 font-semibold text-gray-300">Cross-foot: QuickBase C+P = Tracked</td>
                        <td className={`p-2 text-right font-bold ${qbv.complete + qbv.pending === ov.trackedTotal ? 'text-green-400' : 'text-red-400'}`}>
                          {qbv.complete + qbv.pending === ov.trackedTotal ? 'OK' : 'MISMATCH'}
                        </td>
                      </tr>
                      <tr className="bg-gray-800">
                        <td className="p-2 font-semibold text-gray-300">Cross-foot: Pathwave C+Pending = Tracked</td>
                        <td className={`p-2 text-right font-bold ${ov.pathwaveComplete + ov.pathwavePending === ov.trackedTotal ? 'text-green-400' : 'text-red-400'}`}>
                          {ov.pathwaveComplete + ov.pathwavePending === ov.trackedTotal ? 'OK' : 'MISMATCH'}
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              )}

              {segment === 'outstanding' && <SimpleTable rows={outstanding} />}
              {segment === 'pathwave-complete' && <SimpleTable rows={dataset.filter(r => r.pathwaveStatus === 'COMPLETE')} />}
              {segment === 'qb-complete' && <SimpleTable rows={dataset.filter(r => r.quickbaseStatus === 'COMPLETE')} />}
              {segment === 'fully-complete' && <SimpleTable rows={dataset.filter(r => r.fullyComplete)} />}

              {segment === 'oad' && (
                <div className="overflow-x-auto rounded-xl border border-gray-700 bg-gray-900">
                  <table className="w-full text-xs">
                    <thead>
                      <tr className="bg-gray-800 text-gray-400">
                        {['HOP', 'Path ID', 'GC', 'CM', 'Days Since Complete', 'One and Done (verbatim)'].map(h => <th key={h} className="p-2 text-left">{h}</th>)}
                      </tr>
                    </thead>
                    <tbody>
                      {pw.oadSites.map(r => (
                        <tr key={r.hop} className="border-t border-gray-800">
                          <td className="p-2 font-semibold text-white whitespace-nowrap">{r.hop}</td>
                          <td className="p-2 text-gray-400 whitespace-nowrap">{r.pathId || '—'}</td>
                          <td className="p-2 text-gray-300 whitespace-nowrap">{r.gc || '—'}</td>
                          <td className="p-2 text-gray-300 whitespace-nowrap">{r.cm || '—'}</td>
                          <td className="p-2"><AgingCell days={r.agingDays} thresholds={thresholds} /></td>
                          <td className="p-2 text-gray-400">{r.note || '—'}</td>
                        </tr>
                      ))}
                      {pw.oadSites.length === 0 && <tr><td colSpan={6} className="p-4 text-center text-gray-500">No OAD sites</td></tr>}
                    </tbody>
                  </table>
                </div>
              )}

              {(segment === 'pathwave-pending' || segment === 'qb-pending') && (
                <div>
                  {segment === 'qb-pending' && (
                    <p className="mb-2 text-xs italic text-amber-400">
                      Nokia-internal — NOT a GC action item. Do not send this list to a GC.
                    </p>
                  )}
                  <div className="overflow-x-auto rounded-xl border border-gray-700 bg-gray-900">
                    <table className="w-full text-xs">
                      <thead>
                        <tr className="bg-gray-800 text-gray-400">
                          {['HOP', 'Path ID', 'GC', 'CM', 'Near Site A', 'Far Site B', 'Site', 'Days Since Complete', 'Missing Items'].map(h => <th key={h} className="p-2 text-left">{h}</th>)}
                        </tr>
                      </thead>
                      <tbody>
                        {(segment === 'pathwave-pending' ? pathwavePendingItemized : qbPendingItemized).map((r, i) => (
                          <tr key={`${r.hop}-${r.site}-${i}`} className="border-t border-gray-800">
                            <td className="p-2 font-semibold text-white whitespace-nowrap">{r.hop}</td>
                            <td className="p-2 text-gray-400 whitespace-nowrap">{r.pathId || '—'}</td>
                            <td className="p-2 text-gray-300 whitespace-nowrap">{r.gc || '—'}</td>
                            <td className="p-2 text-gray-300 whitespace-nowrap">{r.cm || '—'}</td>
                            <td className="p-2 text-gray-300 whitespace-nowrap">{r.nearSiteA || '—'}</td>
                            <td className="p-2 text-gray-300 whitespace-nowrap">{r.farSiteB || '—'}</td>
                            <td className="p-2 text-gray-300 whitespace-nowrap">{r.site}</td>
                            <td className="p-2"><AgingCell days={r.agingDays} thresholds={thresholds} /></td>
                            <td className="p-2 text-gray-400">{r.missingItems}</td>
                          </tr>
                        ))}
                        {(segment === 'pathwave-pending' ? pathwavePendingItemized : qbPendingItemized).length === 0 && (
                          <tr><td colSpan={9} className="p-4 text-center text-gray-500">No rows</td></tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>
          )}

          {subTab === 'gc' && (
            <ScopGcReports
              gcReports={gcReports} dataset={dataset} calc={calc}
              emailSettings={emailSettings} thresholds={thresholds}
              today={today} search={gcSearch} setSearch={setGcSearch}
            />
          )}
        </>
      )}
    </div>
  )
}

// ── GC Reports & Emails sub-tab ─────────────────────────────────────────────

function ScopGcReports({
  gcReports, dataset, calc, emailSettings, thresholds, today, search, setSearch,
}: {
  gcReports: ScopGcReport[]
  dataset: ScopRow[]
  calc: ReturnType<typeof resolveScopCalcSettings>
  emailSettings: EmailSettings
  thresholds: { green: number; amber: number }
  today: string
  search: string
  setSearch: (v: string) => void
}) {
  // Every GC in scope, so a clean GC still shows (greyed, no buttons) —
  // matches the Reports page's all-GC grid intent.
  const allGcs = useMemo(
    () => [...new Set(dataset.map(r => r.gc).filter((g): g is string => !!g))].sort(),
    [dataset])
  const reportByGc = useMemo(() => new Map(gcReports.map(r => [r.gc, r])), [gcReports])
  const visible = allGcs.filter(g => g.toLowerCase().includes(search.trim().toLowerCase()))

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center gap-3">
        <input
          value={search} onChange={(e) => setSearch(e.target.value)}
          placeholder="Filter GCs…"
          className="w-48 rounded border border-gray-600 bg-gray-800 px-2 py-1 text-xs text-white focus:border-blue-500 focus:outline-none"
        />
        <span className="text-xs text-gray-500">{gcReports.length} of {allGcs.length} GCs have outstanding work</span>
        <button
          onClick={async () => {
            if (gcReports.length === 0) { alert('No GC has outstanding SCOP work.'); return }
            const zip = new JSZip()
            gcReports.forEach(r => {
              const wb = buildScopGcReportWorkbook(r)
              zip.file(`SCOP_GC_Report_${r.gc.replace(/[/ ]/g, '_')}_${today}.xlsx`, XLSX.write(wb, { type: 'array', bookType: 'xlsx' }))
            })
            const blob = await zip.generateAsync({ type: 'blob' })
            const url = URL.createObjectURL(blob)
            const a = document.createElement('a')
            a.href = url
            a.download = `SCOP_All_GC_Reports_${today}.zip`
            document.body.appendChild(a); a.click(); document.body.removeChild(a)
            URL.revokeObjectURL(url)
          }}
          disabled={gcReports.length === 0}
          className="rounded-lg bg-amber-700 px-3 py-1.5 text-xs font-semibold hover:bg-amber-600 disabled:cursor-not-allowed disabled:opacity-40">
          🗜️ Download All GC Reports (ZIP)
        </button>
      </div>

      <div className="grid gap-4 md:grid-cols-2 lg:grid-cols-3">
        {visible.map(gc => {
          const report = reportByGc.get(gc) ?? null
          const email = report ? buildScopGcEmailMailto(report, calc, emailSettings) : null
          return (
            <div key={gc} className={`rounded-xl border p-4 ${report ? 'border-gray-700 bg-gray-900' : 'border-gray-800 bg-gray-900/50'}`}>
              <h3 className="text-base font-bold text-white">{gc}</h3>
              {report ? (
                <>
                  <div className="mb-3 mt-1 flex flex-wrap gap-x-3 gap-y-0.5 text-xs text-gray-500">
                    <span>{report.summary.totalHOPs} HOPs</span><span>·</span>
                    <span className="text-green-400">{report.summary.completed} done</span><span>·</span>
                    <span className="text-amber-400">{report.summary.outstanding} outstanding</span>
                    {report.summary.oldestOutstandingDays != null && (
                      <>
                        <span>·</span>
                        <span className={agingColor(report.summary.oldestOutstandingDays, thresholds) === 'red' ? 'text-red-400' : agingColor(report.summary.oldestOutstandingDays, thresholds) === 'amber' ? 'text-amber-400' : 'text-green-400'}>
                          oldest {report.summary.oldestOutstandingDays}d
                        </span>
                      </>
                    )}
                    {report.sections.oad.length > 0 && <><span>·</span><span className="text-blue-300">{report.sections.oad.length} OAD</span></>}
                  </div>
                  <div className="flex flex-col gap-2">
                    <button
                      onClick={() => downloadWorkbook(buildScopGcReportWorkbook(report), `SCOP_GC_Report_${gc.replace(/[/ ]/g, '_')}_${today}.xlsx`)}
                      className="rounded bg-purple-700 px-3 py-2 text-xs font-semibold text-white hover:bg-purple-600">
                      📥 Download SCOP Report
                    </button>
                    <button
                      onClick={() => { if (email) window.open(email) }}
                      className="rounded bg-blue-600 px-3 py-2 text-xs font-semibold text-white hover:bg-blue-700">
                      ✉️ Generate SCOP Email
                    </button>
                  </div>
                </>
              ) : (
                <p className="mt-2 text-xs text-gray-600">No outstanding Pathwave or OAD work — no report this cycle.</p>
              )}
            </div>
          )
        })}
      </div>
    </div>
  )
}
