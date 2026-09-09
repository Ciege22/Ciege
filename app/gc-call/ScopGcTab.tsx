'use client'

import { useMemo } from 'react'
import type { EmailSettings, ScopSettings } from '../lib/settings'
import {
  resolveScopCalcSettings, keyScopRows, buildScopDataset, buildGCReport,
  buildScopGcEmailMailto, agingColor,
} from '../lib/scop'
import { buildScopGcReportWorkbook, downloadWorkbook } from '../lib/scopReport'

// Per-GC SCOP view for the GC Call View — same Pathwave + OAD content as the
// Reports → SCOP tab's per-GC report, scoped to the selected GC, so it can be
// reviewed on a call without leaving this page. Reads the SCOP tracker the
// Reports page uploaded (chunked report_snapshots id "scop").
export default function ScopGcTab({
  selectedGC, storeRows, scopSettings, emailSettings,
}: {
  selectedGC: string
  storeRows: unknown[][]
  scopSettings: ScopSettings
  emailSettings: EmailSettings
}) {
  const calc = useMemo(() => resolveScopCalcSettings(scopSettings), [scopSettings])
  const dataset = useMemo(
    () => (storeRows.length >= 2 ? buildScopDataset(keyScopRows(storeRows), calc) : []),
    [storeRows, calc],
  )

  // The GC name in the SCOP dataset is alias-normalised; match selectedGC
  // case-insensitively against it.
  const matchedGc = useMemo(() => {
    const target = selectedGC.trim().toLowerCase()
    return [...new Set(dataset.map(r => r.gc).filter((g): g is string => !!g))]
      .find(g => g.toLowerCase() === target) ?? null
  }, [dataset, selectedGC])

  const report = useMemo(
    () => (matchedGc ? buildGCReport(dataset, matchedGc, calc) : null),
    [dataset, matchedGc, calc],
  )

  if (dataset.length === 0) {
    return <p className="text-gray-400 text-sm">Upload the COP tracker on the Reports page → SCOP tab to enable this view.</p>
  }

  const thresholds = scopSettings.agingColorThresholds
  const agingCls = (d: number | null) => {
    const c = agingColor(d, thresholds)
    return c === 'red' ? 'text-red-400' : c === 'amber' ? 'text-amber-400' : 'text-green-400'
  }
  const today = new Date().toLocaleDateString('en-US').replace(/\//g, '-')

  if (!report) {
    return (
      <p className="text-green-400 text-sm">
        ✅ No outstanding Pathwave or OAD items for {selectedGC}
        {matchedGc ? '' : ' — no SCOP rows found for this GC in the current tracker'}.
      </p>
    )
  }

  const email = buildScopGcEmailMailto(report, calc, emailSettings)

  return (
    <div className="space-y-8">
      {/* summary */}
      <div className="flex flex-wrap items-center gap-x-4 gap-y-1 rounded-lg border border-gray-700 bg-gray-900 px-4 py-2 text-sm">
        <span className="font-semibold text-gray-200">{report.gc}</span>
        <span className="text-gray-400">{report.summary.totalHOPs} HOPs</span>
        <span className="text-green-400">{report.summary.completed} complete</span>
        <span className="text-amber-400">{report.summary.outstanding} outstanding</span>
        {report.summary.oldestOutstandingDays != null && (
          <span className={agingCls(report.summary.oldestOutstandingDays)}>oldest {report.summary.oldestOutstandingDays}d</span>
        )}
        {report.sections.oad.length > 0 && <span className="text-blue-300">{report.sections.oad.length} OAD</span>}
      </div>

      {/* Section 1 — Pathwave GC-owned items */}
      <div>
        <h3 className="mb-3 text-lg font-semibold text-white">
          🔨 Pathwave — GC-Owned Items ({report.sections.pathwave.length})
        </h3>
        {report.sections.pathwave.length === 0
          ? <p className="text-sm text-gray-500">No outstanding GC-owned Pathwave items</p>
          : (
            <div className="overflow-x-auto">
              <table className="w-full text-xs">
                <thead>
                  <tr className="bg-gray-800 text-gray-400">
                    {['HOP', 'Path ID', 'Near Site (A)', 'Far Site (B)', 'Site CM', 'Site', 'Days Since Complete', 'Missing Items'].map(h => (
                      <th key={h} className="p-2 text-left">{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {report.sections.pathwave.map((r, i) => (
                    <tr key={`${r.hop}-${r.site}-${i}`} className="border-t border-gray-800 bg-gray-900">
                      <td className="p-2 font-semibold text-white whitespace-nowrap">{r.hop}</td>
                      <td className="p-2 text-gray-400 whitespace-nowrap">{r.pathId || '—'}</td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{r.nearSiteA || '—'}</td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{r.farSiteB || '—'}</td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{r.cm || '—'}</td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{r.site}</td>
                      <td className={`p-2 font-bold ${agingCls(r.agingDays)}`}>{r.agingDays ?? '—'}</td>
                      <td className="p-2 text-gray-400">{r.missingItems}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
      </div>

      {/* Section 2 — OAD (only when the GC has OAD HOPs) */}
      {report.sections.oad.length > 0 && (
        <div>
          <h3 className="mb-3 text-lg font-semibold text-blue-300">
            🔵 OAD — Awaiting OAD (Not a GC Checklist Item) ({report.sections.oad.length})
          </h3>
          <div className="overflow-x-auto">
            <table className="w-full text-xs">
              <thead>
                <tr className="bg-gray-800 text-gray-400">
                  {['HOP', 'Path ID', 'Near Site (A)', 'Far Site (B)', 'Site CM', 'Days Since Complete', 'One and Done (verbatim)'].map(h => (
                    <th key={h} className="p-2 text-left">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {report.sections.oad.map((r, i) => (
                  <tr key={`${r.hop}-${i}`} className="border-t border-gray-800 bg-blue-950/40">
                    <td className="p-2 font-semibold text-white whitespace-nowrap">{r.hop}</td>
                    <td className="p-2 text-gray-400 whitespace-nowrap">{r.pathId || '—'}</td>
                    <td className="p-2 text-gray-300 whitespace-nowrap">{r.nearSiteA || '—'}</td>
                    <td className="p-2 text-gray-300 whitespace-nowrap">{r.farSiteB || '—'}</td>
                    <td className="p-2 text-gray-300 whitespace-nowrap">{r.cm || '—'}</td>
                    <td className={`p-2 font-bold ${agingCls(r.agingDays)}`}>{r.agingDays ?? '—'}</td>
                    <td className="p-2 text-gray-400">{r.note || '—'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* actions — same pattern as the Decom / GR tabs */}
      <div className="flex flex-wrap gap-3">
        <button
          onClick={() => downloadWorkbook(buildScopGcReportWorkbook(report), `SCOP_GC_Report_${report.gc.replace(/[/ ]/g, '_')}_${today}.xlsx`)}
          className="rounded-lg bg-purple-700 px-4 py-2 text-sm font-semibold text-white hover:bg-purple-600">
          📥 Download SCOP Report
        </button>
        <button
          onClick={() => window.open(email)}
          className="rounded-lg bg-blue-600 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-700">
          ✉️ Generate SCOP Email — {selectedGC} ({report.summary.outstanding} outstanding)
        </button>
      </div>
    </div>
  )
}
