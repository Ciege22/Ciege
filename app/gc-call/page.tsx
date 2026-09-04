'use client'

export const dynamic = 'force-dynamic'

import React, { useState, useCallback, useEffect } from 'react'
import * as XLSX from 'xlsx'
import { supabase, loadTrackerSnapshot } from '../lib/supabase'
import { GC_CONFIG, matches, SPO_VENDOR_COL_IN_MASTER, CR_SUPPLIER_COL_IN_MASTER } from '../lib/gcConfig'
import BackToDashboard from '../components/BackToDashboard'
import { ThresholdSettings, DEFAULT_THRESHOLDS, loadThresholdSettings, EmailSettings, DEFAULT_EMAIL, loadEmailSettings, ProgramSettings, DEFAULT_PROGRAM, loadProgramSettings, crewCountForGc, lookupContactEmail } from '../lib/settings'
import { PendingUpdate, SOURCE_LABELS, SOURCE_BADGE_CLASSES, loadPendingUpdates, persistPendingUpdates, upsertPendingUpdate } from '../lib/pendingUpdates'
import { loadChunkedReport } from '../lib/reportChunks'
import { parseDecomRows, decomRowsForGc, buildDecomEmailMailto, fmtDecomDate, parseTrackerHopsForDecom, findMissingDecom, DecomRow } from '../lib/decom'
import {
  GrRow, GrTileFilter, loadGrRows, groupGrRows, sortGrRowsBy, computeGrBreakdown, rowsForTileFilter,
  buildGrEmailMailto, fmtMoney, fmtMoneyShort,
} from '../lib/grTracker'


// The tracker's "General Contractor" column isn't guaranteed to be typed
// with consistent casing across rows (e.g. "Vikor" vs "VIKOR"). Resolves a
// raw cell value to GC_CONFIG's canonical casing when it's a known GC, so
// the GC tab list — and everything keyed off selectedGC (GC_CM_MAP,
// gcContactEmails via lookupContactEmail, email subject lines) — shows and
// uses the same spelling a person would type into Settings, instead of a
// forced-lowercase string nothing else in the app recognizes.
function canonicalGcName(raw: string): string {
  const trimmed = raw.trim()
  if (!trimmed) return ''
  const match = GC_CONFIG.find(cfg => cfg.gc.toLowerCase() === trimmed.toLowerCase())
  return match ? match.gc : trimmed
}

const GC_CM_MAP: Record<string, string> = {
  'MZI': 'Steve',
  'NV Tel': 'Steve',
  'Mastec': 'Benny',
  'Vikor': 'Benny',
  'Tech CX': 'Hap',
}

// Same raw-row column layout as app/reports/page.tsx's SPO/CR download section.
const SPO_COL_IDX = [7, 8, 33, 40, 41, 43, 47, 48, 49, 50, 51]
const SPO_HEADERS = ['Customer Site ID', 'Name', 'SOG Name', 'SPO Number', 'SPO Creation Date', 'SPO Vendor', 'SPO Value', 'IA Date', 'IA User', 'GR Date', 'GR Number']
const CR_COL_IDX = [0, 1, 2, 4, 5, 6, 7, 8, 9, 10, 11, 16, 18, 19, 25, 26, 27, 28, 29, 30, 31]
const CR_HEADERS = ['Requestor', 'Supplier Name', 'Path ID', 'Site Name', 'Site #', 'Network Site Name', 'Risk Budget', 'Materials or Labor', 'Reason for CR Details', 'Viaero Operation CR Filed', 'CR Type', 'Sellable to Who', 'PM Status', 'PM Status Owner', 'SPO Cost', 'GC Quote Shared', 'CQT Package', 'SPO #', 'SPO Issued Date', 'SPO IA/GR', 'CQT ID']
const REPORT_NAVY = '124191'

function downloadGcFilteredReport(rows: unknown[][], colIdx: number[], headers: string[], vendorColInMaster: number, matchList: string[], filename: string) {
  const gcRows = rows.filter(row => matches(row[vendorColInMaster], matchList))

  const wb = XLSX.utils.book_new()
  const sheetData = [
    headers,
    ...gcRows.map(row => colIdx.map(i => {
      const val = row[i]
      if (val instanceof Date) return val.toLocaleDateString('en-US')
      return val ?? ''
    }))
  ]

  const ws = XLSX.utils.aoa_to_sheet(sheetData)

  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1')
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: 0, c })]
    if (cell) {
      cell.s = {
        font: { bold: true, color: { rgb: 'FFFFFF' } },
        fill: { fgColor: { rgb: REPORT_NAVY } },
        alignment: { horizontal: 'center', wrapText: true }
      }
    }
  }

  ws['!cols'] = headers.map((h, i) => {
    const maxLen = Math.max(h.length, ...sheetData.slice(1).map(r => String(r[i] || '').length))
    return { wch: Math.min(Math.max(maxLen + 2, 10), 40) }
  })

  XLSX.utils.book_append_sheet(wb, ws, headers[0].includes('SPO') || filename.includes('SPO') ? 'SPO Report' : 'CR Tracker')
  XLSX.writeFile(wb, filename)
}

function styleDecomHeaderRow(ws: XLSX.WorkSheet, headers: string[], sheetData: (string | number)[][]) {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1')
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: 0, c })]
    if (cell) {
      cell.s = {
        font: { bold: true, color: { rgb: 'FFFFFF' } },
        fill: { fgColor: { rgb: REPORT_NAVY } },
        alignment: { horizontal: 'center', wrapText: true }
      }
    }
  }
  ws['!cols'] = headers.map((h, i) => {
    const maxLen = Math.max(h.length, ...sheetData.slice(1).map(r => String(r[i] ?? '').length))
    return { wch: Math.min(Math.max(maxLen + 2, 10), 40) }
  })
}

function decomRowToShortSheetRow(r: DecomRow, includeAging: boolean): (string | number)[] {
  const row: (string | number)[] = [r.hop, r.pathId, r.siteName, r.cm, fmtDecomDate(r.cxComplete)]
  if (includeAging) row.push(r.aging ?? '')
  row.push(r.podPathwave ? 'Yes' : 'No', r.podQuickBase ? 'Yes' : 'No', r.comment || '')
  return row
}

function downloadGcDecomExcel(gcRows: DecomRow[], filename: string) {
  const pendingDropOff = gcRows.filter(r => !r.dropOffDate)
  const pendingPathwave = gcRows.filter(r => r.dropOffDate && !r.podPathwave)
  // Pending POD in QuickBase = Pathwave confirmed, QuickBase isn't yet.
  const pendingQuickBase = gcRows.filter(r => r.dropOffDate && r.podPathwave && !r.podQuickBase)

  const dropOffHeaders = ['HOP', 'Path ID', 'Site Name', 'CM', 'CX Complete', 'Aging (days)', 'POD Pathwave', 'POD QuickBase', 'Comments']
  const pathwaveHeaders = ['HOP', 'Path ID', 'Site Name', 'CM', 'CX Complete', 'POD Pathwave', 'POD QuickBase', 'Comments']
  const quickBaseHeaders = ['HOP', 'Path ID', 'Site Name', 'CM', 'CX Complete', 'POD Pathwave', 'POD QuickBase', 'Comments']

  const wb = XLSX.utils.book_new()

  const sheet1Data = [dropOffHeaders, ...pendingDropOff.map(r => decomRowToShortSheetRow(r, true))]
  const ws1 = XLSX.utils.aoa_to_sheet(sheet1Data)
  styleDecomHeaderRow(ws1, dropOffHeaders, sheet1Data)
  XLSX.utils.book_append_sheet(wb, ws1, 'Pending Decom Drop Off')

  const sheet2Data = [pathwaveHeaders, ...pendingPathwave.map(r => decomRowToShortSheetRow(r, false))]
  const ws2 = XLSX.utils.aoa_to_sheet(sheet2Data)
  styleDecomHeaderRow(ws2, pathwaveHeaders, sheet2Data)
  XLSX.utils.book_append_sheet(wb, ws2, 'Pending POD in Pathwave')

  const sheet3Data = [quickBaseHeaders, ...pendingQuickBase.map(r => decomRowToShortSheetRow(r, false))]
  const ws3 = XLSX.utils.aoa_to_sheet(sheet3Data)
  styleDecomHeaderRow(ws3, quickBaseHeaders, sheet3Data)
  XLSX.utils.book_append_sheet(wb, ws3, 'Pending POD in QuickBase')

  XLSX.writeFile(wb, filename)
}

interface GrInvoicingTabProps {
  selectedGC: string
  grRows: GrRow[]
  grLoaded: boolean
  emailSettings: EmailSettings
}

function GrTile({ emoji, label, value, sub, color, valueColor, active, onClick }: {
  emoji: string; label: string; value: string; sub: string; color: string; valueColor?: string; active?: boolean; onClick: () => void
}) {
  return (
    <div onClick={onClick}
      className={`bg-gray-900 rounded-xl border p-4 text-center cursor-pointer hover:border-blue-500 hover:bg-gray-800 transition-all ${active ? 'ring-2 ring-blue-500' : ''} ${color}`}>
      <p className="text-2xl mb-1">{emoji}</p>
      <p className={`text-2xl font-bold ${valueColor || 'text-white'}`}>{value}</p>
      <p className="text-gray-400 text-xs mt-1 font-semibold">{label}</p>
      <p className="text-gray-600 text-xs mt-0.5">{sub}</p>
    </div>
  )
}

function GrInvoicingTab({ selectedGC, grRows, grLoaded, emailSettings }: GrInvoicingTabProps) {
  const [tileFilter, setTileFilter] = useState<GrTileFilter>(null)

  if (!grLoaded) return <p className="text-gray-400 text-sm">Loading GR data...</p>

  // Scoped to the currently selected GC — recalculates immediately when selectedGC changes.
  const gcGrRows = grRows.filter(r => r.gc?.trim().toLowerCase() === selectedGC?.trim().toLowerCase())
  const breakdown = computeGrBreakdown(gcGrRows)
  const ready = groupGrRows(gcGrRows).ready
  const emailMailto = ready.length > 0
    ? buildGrEmailMailto(selectedGC, ready, { financeEmails: emailSettings.financeEmails, gcContactEmails: emailSettings.gcContactEmails })
    : null

  const displayRows = sortGrRowsBy(rowsForTileFilter(gcGrRows, tileFilter), 'trigger')

  const selectTile = (filter: GrTileFilter) => setTileFilter(prev => prev === filter ? null : filter)

  const statusChip = (status: GrRow['status']) => (
    <span className={`text-xs px-2 py-0.5 rounded-full whitespace-nowrap ${
      status === 'GR Done' ? 'bg-green-900 text-green-200'
      : status === 'Ready to Release' ? 'bg-orange-900 text-orange-200'
      : 'bg-gray-700 text-gray-300'
    }`}>
      {status === 'GR Done' ? '✓ GR Done' : status === 'Ready to Release' ? '⏳ Pending' : 'Not Yet'}
    </span>
  )

  return (
    <div className="space-y-4">
      {/* KPI Tiles — Row 1: Volume */}
      <div className="grid grid-cols-2 gap-3">
        <GrTile emoji="📄" label="Total Base POs" value={String(breakdown.totalBasePOs)} sub="valid SPO + SOG rows"
          color="border-gray-700" active={tileFilter === 'totalBasePOs'}
          onClick={() => selectTile('totalBasePOs')} />
        <GrTile emoji="🔧" label="Total CRs" value={String(breakdown.totalCRs)} sub="SPO issued, no SOG tier"
          color="border-gray-700" active={tileFilter === 'totalCRs'}
          onClick={() => selectTile('totalCRs')} />
      </div>

      {/* KPI Tiles — Row 2: Value */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
        <GrTile emoji="✅" label="Base PO Value GR'd" value={fmtMoneyShort(breakdown.basePOValueGRd)} sub="GR Date populated"
          color="border-green-700" valueColor="text-green-400" active={tileFilter === 'basePOGRd'}
          onClick={() => selectTile('basePOGRd')} />
        <GrTile emoji="⏳" label="Base PO Value Pending" value={fmtMoneyShort(breakdown.basePOValuePending)} sub="trigger met, GR Date blank"
          color="border-orange-700" valueColor="text-orange-400" active={tileFilter === 'basePOPending'}
          onClick={() => selectTile('basePOPending')} />
        <GrTile emoji="✅" label="CR Value GR'd" value={fmtMoneyShort(breakdown.crValueGRd)} sub="GR Date populated"
          color="border-green-700" valueColor="text-green-400" active={tileFilter === 'crGRd'}
          onClick={() => selectTile('crGRd')} />
        <GrTile emoji="⏳" label="CR Value Pending" value={fmtMoneyShort(breakdown.crValuePending)} sub="GR Date blank"
          color="border-orange-700" valueColor="text-orange-400" active={tileFilter === 'crPending'}
          onClick={() => selectTile('crPending')} />
      </div>

      <div className="flex items-center justify-between">
        <div className="flex items-center gap-3">
          {tileFilter && (
            <button onClick={() => setTileFilter(null)}
              className="text-gray-400 hover:text-white text-xs underline">✕ Clear tile filter</button>
          )}
          <span className="text-gray-500 text-xs">{displayRows.length} rows</span>
        </div>
        <button onClick={() => { if (emailMailto) window.open(emailMailto) }}
          disabled={!emailMailto}
          className="bg-blue-600 hover:bg-blue-700 disabled:opacity-40 disabled:cursor-not-allowed text-white px-4 py-2 rounded-lg text-sm font-semibold">
          ✉️ Generate GR Email — {selectedGC} ({ready.length} pending)
        </button>
      </div>

      {/* Main Table */}
      <div className="overflow-x-auto bg-gray-900 rounded-xl border border-gray-700">
        <table className="w-full text-xs">
          <thead>
            <tr className="bg-gray-800 text-gray-400">
              <th className="text-left p-2">HOP</th>
              <th className="text-left p-2">Path ID</th>
              <th className="text-left p-2">SOG Name</th>
              <th className="text-left p-2">SPO #</th>
              <th className="text-left p-2">SPO Value</th>
              <th className="text-left p-2">Trigger Date</th>
              <th className="text-left p-2">GR Status</th>
              <th className="text-left p-2">GR Date</th>
            </tr>
          </thead>
          <tbody>
            {displayRows.map(r => (
              <tr key={`${r.hop}-${r.sogName}-${r.spoNumber}`} className={`border-t border-gray-800 ${r.isDecomScop ? 'text-gray-600' : 'text-gray-200'}`}>
                <td className={`p-2 font-semibold whitespace-nowrap ${r.isDecomScop ? 'text-gray-500' : 'text-white'}`}>
                  {r.isDecomScop && <span className="mr-1">👁</span>}{r.hopDisplay}
                </td>
                <td className="p-2 whitespace-nowrap">{r.pathId || '—'}</td>
                <td className="p-2 whitespace-nowrap">{r.sogName || 'CR'}</td>
                <td className="p-2 whitespace-nowrap">{r.spoNumber || '—'}</td>
                <td className="p-2 whitespace-nowrap">{fmtMoney(r.spoValue)}</td>
                <td className="p-2 whitespace-nowrap">{r.triggerDate || '—'}</td>
                <td className="p-2">{statusChip(r.status)}</td>
                <td className="p-2 whitespace-nowrap">{r.grDate || '—'}</td>
              </tr>
            ))}
            {displayRows.length === 0 && (
              <tr><td colSpan={8} className="p-6 text-center text-gray-500">No rows match the current filter</td></tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  )
}

interface GcReportsTabProps {
  selectedGC: string
  spoRawRows: unknown[][]
  crRawRows: unknown[][]
  decomRawRows: unknown[][]
}

function GcReportsTab({ selectedGC, spoRawRows, crRawRows, decomRawRows }: GcReportsTabProps) {
  const cfg = GC_CONFIG.find(c => c.gc?.trim().toLowerCase() === selectedGC?.trim().toLowerCase())
  const today = new Date().toLocaleDateString('en-US').replace(/\//g, '-')

  if (!cfg) return <p className="text-gray-400 text-sm">No report configuration found for {selectedGC}.</p>

  const spoCount = spoRawRows.filter(row => matches(row[SPO_VENDOR_COL_IN_MASTER], cfg.spo_match)).length
  const crCount = crRawRows.filter(row => matches(row[CR_SUPPLIER_COL_IN_MASTER], cfg.cr_match)).length
  const gcDecomRows = decomRowsForGc(parseDecomRows(decomRawRows), selectedGC)

  return (
    <div className="bg-gray-800 rounded-xl border border-gray-700 p-5 max-w-md">
      <h3 className="font-bold text-white text-base mb-1">{selectedGC} — SPO / CR Reports</h3>
      <div className="flex gap-2 text-xs text-gray-500 mb-4">
        <span>{spoCount} SPO rows</span>
        <span>·</span>
        <span>{crCount} CR rows</span>
        <span>·</span>
        <span>{gcDecomRows.length} Decom rows</span>
      </div>
      <div className="flex flex-col gap-2">
        <button
          onClick={() => downloadGcFilteredReport(spoRawRows, SPO_COL_IDX, SPO_HEADERS, SPO_VENDOR_COL_IN_MASTER, cfg.spo_match, `${cfg.spo_label}_-_SPO_Report_-_${today}.xlsx`)}
          disabled={spoRawRows.length === 0}
          className="bg-blue-700 hover:bg-blue-600 disabled:opacity-40 disabled:cursor-not-allowed text-white text-xs px-3 py-2 rounded font-semibold">
          📥 Download SPO Report
        </button>
        <button
          onClick={() => downloadGcFilteredReport(crRawRows, CR_COL_IDX, CR_HEADERS, CR_SUPPLIER_COL_IN_MASTER, cfg.cr_match, `${cfg.cr_label}_-_CR_Tracker_-_${today}.xlsx`)}
          disabled={crRawRows.length === 0}
          className="bg-teal-700 hover:bg-teal-600 disabled:opacity-40 disabled:cursor-not-allowed text-white text-xs px-3 py-2 rounded font-semibold">
          📥 Download CR Tracker
        </button>
        <button
          onClick={() => downloadGcDecomExcel(gcDecomRows, `${cfg.gc}_-_Decom_Report_-_${today}.xlsx`)}
          disabled={decomRawRows.length === 0}
          title={decomRawRows.length === 0 ? 'Upload Decom Tracker on Reports page first' : undefined}
          className="bg-purple-700 hover:bg-purple-600 disabled:opacity-40 disabled:cursor-not-allowed text-white text-xs px-3 py-2 rounded font-semibold">
          ⬇️ {selectedGC} Decom Report
        </button>
      </div>
      {spoRawRows.length === 0 && crRawRows.length === 0 && decomRawRows.length === 0 && (
        <p className="text-gray-500 text-xs mt-3">Upload SPO/CR/Decom master reports on the Reports page to enable downloads here.</p>
      )}
    </div>
  )
}

interface DecomTabProps {
  selectedGC: string
  decomRawRows: unknown[][]
  trackerRawRows: unknown[][]
  emailSettings: EmailSettings
}

function DecomTab({ selectedGC, decomRawRows, trackerRawRows, emailSettings }: DecomTabProps) {
  const [showComplete, setShowComplete] = useState(false)
  const [showMissingOverride, setShowMissingOverride] = useState<boolean | null>(null)
  const [showQuickBaseOverride, setShowQuickBaseOverride] = useState<boolean | null>(null)

  if (decomRawRows.length === 0) {
    return <p className="text-gray-400 text-sm">Upload the Decom Tracker on the Reports page to enable this view.</p>
  }

  const gcRows = decomRowsForGc(parseDecomRows(decomRawRows), selectedGC)
  const outstandingPending = gcRows
    .filter(r => r.status === 'outstanding' || r.status === 'pending')
    .sort((a, b) => (b.aging ?? -1) - (a.aging ?? -1))
  const podGap = gcRows.filter(r => r.status === 'pod_gap')
  // podGap splits into two mutually-exclusive buckets — same partition decom.ts's
  // summarizeDecomByGc uses for the Reports page GC breakdown — so this tab's
  // per-row lists line up with those counts instead of double-showing a row
  // that's actually just pending QuickBase under the "pending Pathwave" section.
  const pendingPathwaveRows = gcRows.filter(r => r.dropOffDate && !r.podPathwave)
  const pendingQuickBaseRows = gcRows.filter(r => r.dropOffDate && r.podPathwave && !r.podQuickBase)
  const complete = gcRows.filter(r => r.status === 'complete')
  const outstandingCount = gcRows.filter(r => r.status === 'outstanding').length
  const showQuickBase = showQuickBaseOverride !== null ? showQuickBaseOverride : pendingQuickBaseRows.length <= 5

  const gcTrackerHops = parseTrackerHopsForDecom(trackerRawRows)
    .filter(t => t.gc?.trim().toLowerCase() === selectedGC?.trim().toLowerCase())
  const missingSites = findMissingDecom(gcRows, gcTrackerHops)
    .sort((a, b) => b.daysElapsed - a.daysElapsed)
  const showMissing = showMissingOverride !== null ? showMissingOverride : missingSites.length <= 5

  const generateDecomEmail = () => {
    const mailto = buildDecomEmailMailto(selectedGC, gcRows, emailSettings)
    window.open(mailto)
  }

  return (
    <div className="space-y-8">
      {/* Section 0 — Not in Decom Tracker */}
      {missingSites.length > 0 && (
        <div>
          <div className="flex items-center gap-3 mb-3 cursor-pointer" onClick={() => setShowMissingOverride(!showMissing)}>
            <h3 className="text-lg font-semibold text-red-400"
              title="Construction complete but no entry exists in the decom tracker file. Chase your colleague to add these sites.">
              ⚠️ Not in Decom Tracker ({missingSites.length})
            </h3>
            <span className="text-gray-500 text-sm">{showMissing ? '▲' : '▼'}</span>
          </div>
          {showMissing && (
            <div className="overflow-x-auto">
              <table className="w-full text-xs">
                <thead>
                  <tr className="bg-gray-800 text-gray-400">
                    <th className="text-left p-2">HOP</th>
                    <th className="text-left p-2">Path ID</th>
                    <th className="text-left p-2">Site Name</th>
                    <th className="text-left p-2">Nokia PM</th>
                    <th className="text-left p-2">CX Complete</th>
                    <th className="text-left p-2">Days Since Complete</th>
                  </tr>
                </thead>
                <tbody>
                  {missingSites.map(m => (
                    <tr key={`${m.pathId || m.hop}-${m.siteName}`} className="border-t border-gray-800 bg-red-950">
                      <td className="p-2 font-semibold text-white whitespace-nowrap">{m.hop}</td>
                      <td className="p-2 text-gray-400 text-xs whitespace-nowrap">{m.pathId || '—'}</td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{m.siteName || '—'}</td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{m.nokiaPm || '—'}</td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{fmtDecomDate(m.ms16a) || '—'}</td>
                      <td className="p-2 text-red-400 font-bold">{m.daysElapsed}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}

      {/* Section 1 — Pending Decom Drop Off */}
      <div>
        <h3 className="text-lg font-semibold text-white mb-3"
          title="DECOM Drop Off date is blank — equipment has not been returned to the warehouse yet. Most urgent.">
          🔴 Pending Decom Drop Off ({outstandingPending.length})
        </h3>
        {outstandingPending.length === 0
          ? <p className="text-gray-500 text-sm">No outstanding or pending decom items</p>
          : (
            <div className="overflow-x-auto">
              <table className="w-full text-xs">
                <thead>
                  <tr className="bg-gray-800 text-gray-400">
                    <th className="text-left p-2">HOP</th>
                    <th className="text-left p-2">Path ID</th>
                    <th className="text-left p-2">Site Name</th>
                    <th className="text-left p-2">CM</th>
                    <th className="text-left p-2">CX Complete</th>
                    <th className="text-left p-2">Aging</th>
                    <th className="text-left p-2">POD Pathwave</th>
                    <th className="text-left p-2">POD QuickBase</th>
                    <th className="text-left p-2">Comments</th>
                  </tr>
                </thead>
                <tbody>
                  {outstandingPending.map(r => (
                    <tr key={r.rowKey} className={`border-t border-gray-800 ${(r.aging ?? 0) >= 7 ? 'bg-red-950' : 'bg-yellow-950'}`}>
                      <td className="p-2 font-semibold text-white whitespace-nowrap">{r.hop}</td>
                      <td className="p-2 text-gray-400 text-xs whitespace-nowrap">{r.pathId || '—'}</td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{r.siteName || '—'}</td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{r.cm || '—'}</td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{fmtDecomDate(r.cxComplete) || '—'}</td>
                      <td className="p-2">
                        <span className={`text-xs font-bold px-2 py-0.5 rounded-full whitespace-nowrap ${(r.aging ?? 0) >= 7 ? 'bg-red-800 text-red-200' : 'bg-yellow-800 text-yellow-200'}`}>
                          {(r.aging ?? 0) >= 7 ? '🔴' : '🟡'} {r.aging ?? '—'}d
                        </span>
                      </td>
                      <td className="p-2">{r.podPathwave ? <span className="text-green-400 font-bold">✓</span> : <span className="text-red-400 font-bold">✗</span>}</td>
                      <td className="p-2">{r.podQuickBase ? <span className="text-green-400 font-bold">✓</span> : <span className="text-red-400 font-bold">✗</span>}</td>
                      <td className="p-2 text-gray-400 text-xs max-w-48 truncate" title={r.comment}>{r.comment || '—'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )
        }
        <button onClick={generateDecomEmail}
          disabled={outstandingCount === 0 && podGap.length === 0}
          className="mt-4 bg-blue-600 hover:bg-blue-700 disabled:opacity-40 disabled:cursor-not-allowed text-white px-4 py-2 rounded-lg text-sm font-semibold">
          ✉️ Generate Decom Email — {selectedGC} ({outstandingPending.length} pending drop off · {podGap.length} pending POD)
        </button>
      </div>

      {/* Section 2 — Pending POD in Pathwave */}
      <div>
        <h3 className="text-lg font-semibold text-white mb-3"
          title="DECOM Drop Off confirmed but POD In Pathwave not yet confirmed. Equipment returned but paperwork incomplete.">
          ⚠️ Pending POD in Pathwave ({pendingPathwaveRows.length})
        </h3>
        {pendingPathwaveRows.length === 0
          ? <p className="text-gray-500 text-sm">No POD gaps</p>
          : (
            <div className="overflow-x-auto">
              <table className="w-full text-xs">
                <thead>
                  <tr className="bg-gray-800 text-gray-400">
                    <th className="text-left p-2">HOP</th>
                    <th className="text-left p-2">Path ID</th>
                    <th className="text-left p-2">Site Name</th>
                    <th className="text-left p-2">CM</th>
                    <th className="text-left p-2">CX Complete</th>
                    <th className="text-left p-2">POD Pathwave</th>
                    <th className="text-left p-2">POD QuickBase</th>
                    <th className="text-left p-2">Comments</th>
                  </tr>
                </thead>
                <tbody>
                  {pendingPathwaveRows.map(r => (
                    <tr key={r.rowKey} className="border-t border-gray-800 bg-yellow-950">
                      <td className="p-2 font-semibold text-white whitespace-nowrap">{r.hop}</td>
                      <td className="p-2 text-gray-400 text-xs whitespace-nowrap">{r.pathId || '—'}</td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{r.siteName || '—'}</td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{r.cm || '—'}</td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{fmtDecomDate(r.cxComplete) || '—'}</td>
                      <td className="p-2">{r.podPathwave ? <span className="text-green-400 font-bold">✓</span> : <span className="text-red-400 font-bold">✗</span>}</td>
                      <td className="p-2">{r.podQuickBase ? <span className="text-green-400 font-bold">✓</span> : <span className="text-red-400 font-bold">✗</span>}</td>
                      <td className="p-2 text-gray-400 text-xs max-w-48 truncate" title={r.comment}>{r.comment || '—'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )
        }
      </div>

      {/* Section 2.5 — Pending POD in QuickBase */}
      <div>
        <div className="flex items-center gap-3 mb-3 cursor-pointer" onClick={() => setShowQuickBaseOverride(!showQuickBase)}>
          <h3 className="text-lg font-semibold text-white"
            title="DECOM Drop Off confirmed and POD in Pathwave confirmed, but POD in QuickBase not yet submitted.">
            📋 Pending POD in QuickBase ({pendingQuickBaseRows.length})
          </h3>
          <span className="text-gray-500 text-sm">{showQuickBase ? '▲' : '▼'}</span>
        </div>
        {showQuickBase && (
          pendingQuickBaseRows.length === 0
            ? <p className="text-gray-500 text-sm">No QuickBase gaps</p>
            : (
              <div className="overflow-x-auto">
                <table className="w-full text-xs">
                  <thead>
                    <tr className="bg-gray-800 text-gray-400">
                      <th className="text-left p-2">HOP</th>
                      <th className="text-left p-2">Path ID</th>
                      <th className="text-left p-2">Site Name</th>
                      <th className="text-left p-2">CM</th>
                      <th className="text-left p-2">CX Complete</th>
                      <th className="text-left p-2">POD Pathwave</th>
                      <th className="text-left p-2">POD QuickBase</th>
                      <th className="text-left p-2">Comments</th>
                    </tr>
                  </thead>
                  <tbody>
                    {pendingQuickBaseRows.map(r => (
                      <tr key={r.rowKey} className="border-t border-gray-800 bg-orange-950">
                        <td className="p-2 font-semibold text-white whitespace-nowrap">{r.hop}</td>
                        <td className="p-2 text-gray-400 text-xs whitespace-nowrap">{r.pathId || '—'}</td>
                        <td className="p-2 text-gray-300 whitespace-nowrap">{r.siteName || '—'}</td>
                        <td className="p-2 text-gray-300 whitespace-nowrap">{r.cm || '—'}</td>
                        <td className="p-2 text-gray-300 whitespace-nowrap">{fmtDecomDate(r.cxComplete) || '—'}</td>
                        <td className="p-2">{r.podPathwave ? <span className="text-green-400 font-bold">✓</span> : <span className="text-red-400 font-bold">✗</span>}</td>
                        <td className="p-2">{r.podQuickBase ? <span className="text-green-400 font-bold">✓</span> : <span className="text-red-400 font-bold">✗</span>}</td>
                        <td className="p-2 text-gray-400 text-xs max-w-48 truncate" title={r.comment}>{r.comment || '—'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )
        )}
      </div>

      {/* Section 3 — Complete (collapsed by default) */}
      <div>
        <div className="flex items-center gap-3 mb-3 cursor-pointer" onClick={() => setShowComplete(s => !s)}>
          <h3 className="text-lg font-semibold text-white"
            title="DECOM Drop Off confirmed AND POD In Pathwave = Yes AND POD In QuickBase = Yes. Fully done.">
            ✅ Complete ({complete.length})
          </h3>
          <span className="text-gray-500 text-sm">{showComplete ? '▲' : '▼'}</span>
        </div>
        {showComplete && (
          complete.length === 0
            ? <p className="text-gray-500 text-sm">No completed decom items</p>
            : (
              <div className="overflow-x-auto">
                <table className="w-full text-xs">
                  <thead>
                    <tr className="bg-gray-800 text-gray-400">
                      <th className="text-left p-2">HOP</th>
                      <th className="text-left p-2">Path ID</th>
                      <th className="text-left p-2">Site Name</th>
                      <th className="text-left p-2">CM</th>
                      <th className="text-left p-2">CX Complete</th>
                      <th className="text-left p-2">POD Pathwave</th>
                      <th className="text-left p-2">POD QuickBase</th>
                      <th className="text-left p-2">Comments</th>
                    </tr>
                  </thead>
                  <tbody>
                    {complete.map(r => (
                      <tr key={r.rowKey} className="border-t border-gray-800 bg-green-950">
                        <td className="p-2 font-semibold text-white whitespace-nowrap">{r.hop}</td>
                        <td className="p-2 text-gray-400 text-xs whitespace-nowrap">{r.pathId || '—'}</td>
                        <td className="p-2 text-gray-300 whitespace-nowrap">{r.siteName || '—'}</td>
                        <td className="p-2 text-gray-300 whitespace-nowrap">{r.cm || '—'}</td>
                        <td className="p-2 text-gray-300 whitespace-nowrap">{fmtDecomDate(r.cxComplete) || '—'}</td>
                        <td className="p-2"><span className="text-green-400 font-bold">✓</span></td>
                        <td className="p-2"><span className="text-green-400 font-bold">✓</span></td>
                        <td className="p-2 text-gray-400 text-xs max-w-48 truncate" title={r.comment}>{r.comment || '—'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )
        )}
      </div>
    </div>
  )
}

interface HOP {
  hop: string
  pathId: string
  gc: string
  nokiaPm: string
  ops: string
  ms15f: string
  ms15a: string
  ms16f: string
  ms16a: string
  hasNtp: boolean
  hasMat: boolean
  wpApproved: boolean
  gcPickup: boolean
  ntpOwner: string
  ntpWaitingOn: string
  matForecast: string
  matReceived: string
  gcPickupDate: string
  hasSpo: boolean
  hasCpo: boolean
  spoIssued: string
  steelFrom: string
  itwStart: string
  itwEnd: string
  ssStart: string
  ssEnd: string
  daysOut: number | null
  daysElapsed: number | null
  inProgress: boolean
  complete: boolean
  vendorWindow: string
  blockers: string[]
  pullInReady: boolean
  pullInStatus: string
  internalConflict: string
  siteAConflict: string
  siteBConflict: string
  cm: string
  mss: string
  powerUp: string
  gcPickupF: string
  gcPickupA: string
  cxNotes: string
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
    const d = new Date((val - 25569) * 86400 * 1000)
    return isNaN(d.getTime()) ? null : d
  }
  const d = new Date(String(val))
  return isNaN(d.getTime()) ? null : d
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

function getVendorWindow(ms15f: Date | null, itwS: Date | null, itwE: Date | null, ssS: Date | null, ssE: Date | null): string {
  if (!ms15f) return ''
  const parts: string[] = []

  const checkVendor = (name: string, start: Date | null, end: Date | null) => {
    if (!start || !end) return
    const ms15fTime = ms15f!.getTime()
    const startTime = start.getTime()
    const endTime = end.getTime()

    if (startTime <= ms15fTime && ms15fTime <= endTime) {
      parts.push(`🔴 ${name} on site thru ${fmtDM(end)}`)
      return
    }

    if (endTime < ms15fTime) {
      const bufferDays = Math.round((ms15fTime - endTime) / (1000 * 60 * 60 * 24))
      if (bufferDays <= 5) {
        parts.push(`🔴 ${name} clears ${fmtDM(end)} — only ${bufferDays}d before our start`)
      } else if (bufferDays <= 10) {
        parts.push(`⚠️ ${name} clears ${fmtDM(end)} — ${bufferDays}d buffer, monitor`)
      } else {
        parts.push(`✅ ${name} clears ${fmtDM(end)} — ${bufferDays}d buffer`)
      }
      return
    }

    if (startTime > ms15fTime) {
      const bufferDays = Math.round((startTime - ms15fTime) / (1000 * 60 * 60 * 24))
      if (bufferDays <= 10) {
        parts.push(`⚠️ ${name} starts ${fmtDM(start)} — ${bufferDays}d after our start`)
      } else {
        parts.push(`✅ ${name} starts ${fmtDM(start)} — ${bufferDays}d after our start`)
      }
    }
  }

  checkVendor('ITW', itwS, itwE)
  checkVendor('Samsung', ssS, ssE)

  return parts.join(' | ') || '✅ No conflicts'
}

function getBlockers(h: HOP): string[] {
  const b: string[] = []
  if (!h.hasNtp) b.push('🔴 NTP pending')
  if (!h.hasMat) b.push('🔴 Material not received')
  if (h.hasMat && !h.gcPickup) b.push('🟠 Mat in warehouse — GC pickup needed')
  if (!h.wpApproved) b.push('🟡 WP not approved')
  if (h.vendorWindow.includes('🔴')) b.push(`🔴 Vendor conflict`)
  return b
}

function getPullInStatus(h: HOP): string {
  if (h.inProgress || h.complete) return '—'
  if (!h.hasNtp && !h.hasMat) return '🔴 Not ready — NTP + Mat missing'
  if (!h.hasNtp) return '🔴 Not ready — NTP missing'
  if (!h.hasMat) return '🔴 Not ready — Mat missing'
  if (h.blockers.some(b => b.includes('Vendor conflict'))) return '🔴 Cannot pull in — vendor conflict'
  if (h.vendorWindow.includes('🔴')) return '🔴 Cannot pull in — vendor conflict'
  if (h.vendorWindow.includes('⚠️')) return '⚠️ Risky — monitor vendor window'
  if (h.siteAConflict && h.siteBConflict) return '🔴 Cannot pull in — both sites occupied'
  if (h.siteAConflict) return '⚠️ Site B available — Site A has internal conflict'
  if (h.siteBConflict) return '⚠️ Site A available — Site B has internal conflict'
  if (!h.gcPickup) return '🟡 Ready — GC pickup needed first'
  return '✅ Ready to pull in'
}

interface GCEditableDateProps {
  hop: string
  field: string
  value: string
  editedDates: Record<string, Record<string, string>>
  logDateEdit: (hop: string, field: string, oldVal: string, newVal: string) => void
}

function GCEditableDate({ hop, field, value, editedDates, logDateEdit }: GCEditableDateProps) {
  const edited = editedDates[hop]?.[field]

  const toInputFormat = (dateStr: string) => {
    if (!dateStr) return ''
    const parts = dateStr.split('/')
    if (parts.length !== 3) return ''
    return `${parts[2]}-${parts[0].padStart(2,'0')}-${parts[1].padStart(2,'0')}`
  }

  const toDisplayFormat = (dateStr: string) => {
    if (!dateStr) return ''
    const parts = dateStr.split('-')
    if (parts.length !== 3) return ''
    return `${parseInt(parts[1])}/${parseInt(parts[2])}/${parts[0]}`
  }

  const currentValue = toInputFormat(edited || value)

  return (
    <div className="flex flex-col gap-1">
      <input
        type="date"
        value={currentValue}
        onChange={(e) => {
          const newDisplay = toDisplayFormat(e.target.value)
          logDateEdit(hop, field, value, newDisplay)
        }}
        className={`text-xs rounded px-2 py-1 border focus:outline-none focus:border-blue-500 cursor-pointer ${edited ? 'bg-yellow-900 border-yellow-500 text-yellow-200' : 'bg-gray-800 border-gray-600 text-gray-300'}`}
      />
      {edited && (
        <span className="text-yellow-400 text-xs">📝 {edited}</span>
      )}
    </div>
  )
}

interface PipelineTableProps {
  title: string
  rows: HOP[]
  sessionNotes: Record<string, string>
  setSessionNotes: React.Dispatch<React.SetStateAction<Record<string, string>>>
  saveCallNote: (hop: string) => void
  noteHistory: Record<string, { id: string; hop_name: string; note: string; logged_at: string }[]>
  editedDates: Record<string, Record<string, string>>
  logDateEdit: (hop: string, field: string, oldVal: string, newVal: string) => void
  setCxNotesModal: (val: { hop: string; notes: string } | null) => void
  crewAssignments: Record<string, string>
  maxCrews: number
  showCrewBadge: boolean
  onCrewChange: (hop: string, crew: string) => void
}

const CREW_BADGE_COLORS = [
  'bg-blue-800 text-blue-200',
  'bg-purple-800 text-purple-200',
  'bg-teal-800 text-teal-200',
  'bg-pink-800 text-pink-200',
]

function PipelineTable({ title, rows, sessionNotes, setSessionNotes, saveCallNote, noteHistory, editedDates, logDateEdit, setCxNotesModal, crewAssignments, maxCrews, showCrewBadge, onCrewChange }: PipelineTableProps) {
  return (
    <div className="mb-6">
      <h3 className="text-base font-semibold text-white mb-3">{title} ({rows.length})</h3>
      {rows.length === 0
        ? <p className="text-gray-500 text-sm">No sites in this window</p>
        : (
          <div className="overflow-x-auto">
            <table className="w-full text-xs">
              <thead>
                <tr className="bg-gray-800 text-gray-400">
                  <th className="text-left p-2">HOP</th>
                  <th className="text-left p-2">Crew</th>
                  <th className="text-left p-2">Path ID</th>
                  <th className="text-left p-2">FC Start</th>
                  <th className="text-left p-2">FC End</th>
                  <th className="text-left p-2">Days Out</th>
                  <th className="text-left p-2">SPO Issued</th>
                  <th className="text-left p-2">NTP</th>
                  <th className="text-left p-2">NTP Owner</th>
                  <th className="text-left p-2">NTP Waiting On</th>
                  <th className="text-left p-2">Mat</th>
                  <th className="text-left p-2">Mat Forecast</th>
                  <th className="text-left p-2">Mat Received</th>
                  <th className="text-left p-2">GC Pickup</th>
                  <th className="text-left p-2">Steel From</th>
                  <th className="text-left p-2">Pull-In Status</th>
                  <th className="text-left p-2">Vendor Window</th>
                  <th className="text-left p-2">Internal Conflict</th>
                  <th className="text-left p-2">Blockers</th>
                  <th className="text-left p-2">Edit MS15 Fc</th>
                  <th className="text-left p-2">Log MS15 Act</th>
                  <th className="text-left p-2">Call Notes (Today)</th>
                  <th className="text-left p-2">Notes History</th>
                  <th className="text-left p-2">CX Notes</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((h) => {
                  const hasConflict = h.vendorWindow.includes('🔴')
                  const isUrgent = h.blockers.length > 0 && (h.daysOut ?? 99) <= 7
                  const rowBg = hasConflict ? 'bg-red-950' : isUrgent ? 'bg-yellow-950' : h.blockers.length === 0 ? 'bg-green-950' : 'bg-gray-900'
                  return (
                    <tr key={h.hop} className={`border-t border-gray-800 ${rowBg}`}>
                      <td className="p-2 font-semibold text-white whitespace-nowrap">
                        {h.hop}
                        {showCrewBadge && crewAssignments[h.hop] && (
                          <span className={`ml-2 text-xs font-bold px-1.5 py-0.5 rounded ${CREW_BADGE_COLORS[(parseInt(crewAssignments[h.hop].replace('Crew ', ''), 10) - 1) % CREW_BADGE_COLORS.length] || 'bg-gray-700 text-gray-300'}`}>
                            C{crewAssignments[h.hop].replace('Crew ', '')}
                          </span>
                        )}
                      </td>
                      <td className="p-2">
                        <select value={crewAssignments[h.hop] || ''}
                          onChange={(e) => onCrewChange(h.hop, e.target.value)}
                          className="bg-gray-800 text-gray-300 text-xs rounded px-1 py-1 border border-gray-600 focus:outline-none focus:border-blue-500">
                          <option value="">--</option>
                          {Array.from({ length: maxCrews }, (_, i) => `Crew ${i + 1}`).map(c => (
                            <option key={c} value={c}>{c}</option>
                          ))}
                        </select>
                      </td>
                      <td className="p-2 text-gray-400 text-xs whitespace-nowrap">{h.pathId || '—'}</td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{h.ms15f}</td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{h.ms16f}</td>
                      <td className={`p-2 font-bold whitespace-nowrap ${(h.daysOut ?? 99) <= 7 ? 'text-red-400' : (h.daysOut ?? 99) <= 14 ? 'text-yellow-400' : 'text-gray-300'}`}>{h.daysOut}d</td>
                      <td className="p-2">
                        {h.hasSpo
                          ? <span className="text-green-400 font-bold text-sm" title={h.spoIssued}>✓</span>
                          : <span className="text-red-400 font-bold text-sm">✗</span>
                        }
                      </td>
                      <td className="p-2">{h.hasNtp ? <span className="text-green-400 font-bold text-sm">✓</span> : <span className="text-red-400 font-bold text-sm">✗</span>}</td>
                      <td className="p-2 text-gray-300 text-xs">
                        <div className="cursor-help" title={h.ntpOwner}>
                          {h.ntpOwner ? (h.ntpOwner.length > 12 ? h.ntpOwner.slice(0, 12) + '...' : h.ntpOwner) : '—'}
                        </div>
                      </td>
                      <td className="p-2 text-gray-300 text-xs">
                        <div className="cursor-help" title={h.ntpWaitingOn}>
                          {h.ntpWaitingOn ? (h.ntpWaitingOn.length > 12 ? h.ntpWaitingOn.slice(0, 12) + '...' : h.ntpWaitingOn) : '—'}
                        </div>
                      </td>
                      <td className="p-2">{h.hasMat ? <span className="text-green-400 font-bold text-sm">✓</span> : <span className="text-red-400 font-bold text-sm">✗</span>}</td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{h.matForecast || '—'}</td>
                      <td className="p-2 text-gray-300 whitespace-nowrap">{h.matReceived || '—'}</td>
                      <td className="p-2">
                        {h.gcPickupDate
                          ? <span className="text-green-400 text-xs">✓ {h.gcPickupDate}</span>
                          : <GCEditableDate hop={h.hop} field="GC Material Pick-up (A)" value="" editedDates={editedDates} logDateEdit={logDateEdit} />
                        }
                      </td>
                      <td className="p-2 text-gray-300 text-xs whitespace-nowrap">
                        {h.steelFrom || '—'}
                      </td>
                      <td className="p-2 text-xs whitespace-nowrap">
                        <span className={h.pullInStatus.includes('✅') ? 'text-green-400' : h.pullInStatus.includes('⚠️') ? 'text-yellow-400' : h.pullInStatus.includes('🔴') ? 'text-red-400' : h.pullInStatus.includes('🟡') ? 'text-yellow-300' : 'text-gray-500'}>
                          {h.pullInStatus || '—'}
                        </span>
                      </td>
                      <td className="p-2 text-xs max-w-36 truncate" title={h.vendorWindow}>
                        <span className={h.vendorWindow.includes('🔴') ? 'text-red-400' : h.vendorWindow.includes('⚠️') ? 'text-yellow-400' : 'text-green-400'}>
                          {h.vendorWindow || '—'}
                        </span>
                      </td>
                      <td className="p-2 text-xs">
                        {h.internalConflict
                          ? <span className={h.internalConflict.includes('🔴') ? 'text-red-400' : 'text-yellow-400'} title={h.internalConflict}>
                              {h.internalConflict.length > 30 ? h.internalConflict.slice(0, 30) + '...' : h.internalConflict}
                            </span>
                          : <span className="text-green-400">✅ Clear</span>
                        }
                      </td>
                      <td className="p-2 text-xs">
                        {h.blockers.length === 0
                          ? <span className="text-green-400">✅ Clear</span>
                          : <div className="flex flex-col gap-1">
                              {h.blockers.map((b, i) => (
                                <span key={i} className={`whitespace-nowrap ${b.includes('🔴') ? 'text-red-400' : b.includes('🟠') ? 'text-orange-400' : 'text-yellow-400'}`}>
                                  {b}
                                </span>
                              ))}
                            </div>
                        }
                      </td>
                      <td className="p-2">
                        <GCEditableDate hop={h.hop} field="MS15 Fc Start" value={h.ms15f} editedDates={editedDates} logDateEdit={logDateEdit} />
                      </td>
                      <td className="p-2">
                        <GCEditableDate hop={h.hop} field="MS15 Implementation Start A" value={h.ms15a} editedDates={editedDates} logDateEdit={logDateEdit} />
                      </td>
                      <td className="p-2">
                        <div className="flex gap-1">
                          <input type="text" placeholder="Note..." value={sessionNotes[h.hop] || ''}
                            onChange={(e) => setSessionNotes(s => ({ ...s, [h.hop]: e.target.value }))}
                            onKeyDown={(e) => { if (e.key === 'Enter') saveCallNote(h.hop) }}
                            className="w-36 bg-gray-800 text-white text-xs rounded px-2 py-1 border border-gray-600 focus:outline-none focus:border-blue-500" />
                          <button onClick={() => saveCallNote(h.hop)} className="text-xs bg-blue-700 hover:bg-blue-600 text-white px-2 py-1 rounded">💾</button>
                        </div>
                      </td>
                      <td className="p-2 max-w-48">
                        <div className="max-h-20 overflow-y-auto flex flex-col gap-1">
                          {(noteHistory[h.hop] || []).slice(0, 5).map((n, i) => (
                            <div key={i} className="text-xs text-gray-300 border-b border-gray-700 pb-1">
                              <span className="text-gray-500 text-xs">{new Date(n.logged_at).toLocaleDateString()}</span>
                              <span className="ml-1">{n.note}</span>
                            </div>
                          ))}
                          {!noteHistory[h.hop]?.length && <span className="text-gray-600 text-xs">No history</span>}
                        </div>
                      </td>
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
        )}
    </div>
  )
}

export default function GCCallPage() {
  const [selectedGC, setSelectedGC] = useState('')
  const [hops, setHops] = useState<HOP[]>([])
  const [trackerRawRows, setTrackerRawRows] = useState<unknown[][]>([])
  const [loaded, setLoaded] = useState(false)
  const [fileName, setFileName] = useState('')
  // Nokia PM filter — same pattern as the Dashboard's pmFilter/pmOptions.
  // gcList (below, computed at render time) and gcHops both narrow to this.
  const [pmFilter, setPmFilter] = useState<string>('ALL')
  const [pmOptions, setPmOptions] = useState<string[]>(['ALL'])
  const [noteHistory, setNoteHistory] = useState<Record<string, CallNote[]>>({})
  const [sessionNotes, setSessionNotes] = useState<Record<string, string>>({})
  const [editedDates, setEditedDates] = useState<Record<string, Record<string, string>>>({})
  const [pmUpdates, setPmUpdates] = useState<PendingUpdate[]>([])
  const [showPmUpdates, setShowPmUpdates] = useState(false)
  const [pmSortField, setPmSortField] = useState<'hop' | 'field'>('hop')
  const [pmSearch, setPmSearch] = useState('')
  const [snapshotTime, setSnapshotTime] = useState<string>('')
  const [cxNotesModal, setCxNotesModal] = useState<{ hop: string; notes: string } | null>(null)
  const [activeTab, setActiveTab] = useState<'pipeline' | 'gr' | 'decom' | 'reports'>('pipeline')
  const [grRows, setGrRows] = useState<GrRow[]>([])
  const [grLoaded, setGrLoaded] = useState(false)
  const [spoRawRows, setSpoRawRows] = useState<unknown[][]>([])
  const [crRawRows, setCrRawRows] = useState<unknown[][]>([])
  const [decomRawRows, setDecomRawRows] = useState<unknown[][]>([])
  const [thresholds, setThresholds] = useState<ThresholdSettings>(DEFAULT_THRESHOLDS)
  const [emailSettings, setEmailSettings] = useState<EmailSettings>(DEFAULT_EMAIL)
  const [program, setProgram] = useState<ProgramSettings>(DEFAULT_PROGRAM)
  // hop -> 'Crew 1' | 'Crew 2' | 'Crew 3' | 'Crew 4' — absent means unassigned
  const [crewAssignments, setCrewAssignments] = useState<Record<string, string>>({})
  // GC -> selected crew filter ('all' | 'unassigned' | 'Crew N') — remembered
  // per GC; a GC with no entry yet defaults to 'all'.
  const [crewFilterByGc, setCrewFilterByGc] = useState<Record<string, string>>({})

  useEffect(() => {
    const loadSettings = async () => {
      const [t, e, p] = await Promise.all([loadThresholdSettings(), loadEmailSettings(), loadProgramSettings()])
      setThresholds(t)
      setEmailSettings(e)
      setProgram(p)
    }
    loadSettings()
  }, [])

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

  const setCrewForHop = async (hop: string, crew: string) => {
    setCrewAssignments(prev => {
      const next = { ...prev }
      if (crew) next[hop] = crew
      else delete next[hop]
      return next
    })
    if (crew) {
      const { error } = await supabase.from('pm_updates_cache').upsert({
        id: `crew-assign-${hop}`,
        updates: JSON.stringify({ crew }),
        updated_at: new Date().toISOString()
      })
      if (error) console.error('[crew-assign] upsert failed:', error)
    } else {
      const { error } = await supabase.from('pm_updates_cache').delete().eq('id', `crew-assign-${hop}`)
      if (error) console.error('[crew-assign] delete failed:', error)
    }
  }

  useEffect(() => {
    if (selectedGC) console.log(`[gc-call] selectedGC: '${selectedGC}'`)
  }, [selectedGC])

  useEffect(() => {
    const loadGr = async () => {
      const rows = await loadGrRows()
      setGrRows(rows)
      setGrLoaded(true)
    }
    loadGr()
  }, [])

  useEffect(() => {
    const loadReports = async () => {
      const { data: spoSnap } = await supabase.from('report_snapshots').select('data').eq('id', 'spo').single()
      if (spoSnap?.data) setSpoRawRows(JSON.parse(spoSnap.data))

      const crReport = await loadChunkedReport('cr')
      if (crReport) setCrRawRows(crReport.rows)

      const decomReport = await loadChunkedReport('decom')
      if (decomReport) setDecomRawRows(decomReport.rows)
    }
    loadReports()
  }, [])

  const today = new Date()
  // Zero out the time-of-day so daysOut is a clean whole-day count — same fix
  // as the dashboard's computeKPIs (commit 8a32098).
  today.setHours(0, 0, 0, 0)

  useEffect(() => {
    const loadNoteHistory = async () => {
      const { data, error } = await supabase
        .from('hop_call_notes')
        .select('id, hop_name, note, logged_at')
        .order('logged_at', { ascending: false })
      if (error) { console.error('Error loading note history:', error); return }
      if (data) {
        const historyMap: Record<string, CallNote[]> = {}
        data.forEach((row: CallNote) => {
          if (!historyMap[row.hop_name]) historyMap[row.hop_name] = []
          historyMap[row.hop_name].push(row)
        })
        setNoteHistory(historyMap)
      }
    }
    loadNoteHistory()
  }, [])

  // Persisted to the shared pending-updates row (app/lib/pendingUpdates.ts) —
  // same list Tracker and CM View read/write, so an edit made here shows up
  // on their "Pending Updates" panel too, and vice versa.
  const persistPmUpdates = persistPendingUpdates

  useEffect(() => {
    loadPendingUpdates().then(setPmUpdates).catch(e => console.error('Error loading PM updates:', e))
  }, [])

  // Dedup on hop+field before appending (source: 'gc' tags where this entry
  // came from) — a second edit/comment for the same HOP+field today replaces
  // the pending entry instead of stacking a duplicate. The full note history
  // itself (hop_call_notes, below) is untouched by this — only what's queued
  // as "still needs to go in the tracker" collapses to the latest one.
  const upsertPmUpdate = (next: { hop: string; field: string; oldValue: string; newValue: string; timestamp: string }) => {
    setPmUpdates(prev => {
      const updated = upsertPendingUpdate(prev, { ...next, source: 'gc' })
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
  const saveCallNote = async (hop: string) => {
    const note = sessionNotes[hop]
    if (!note?.trim()) return
    const logged_at = new Date().toISOString()
    const { data, error } = await supabase
      .from('hop_call_notes')
      .insert({ hop_name: hop, note: note.trim(), logged_at })
      .select()
      .single()
    if (error) { console.error('Error saving note:', error); return }
    if (data) {
      setNoteHistory(h => ({ ...h, [hop]: [data as CallNote, ...(h[hop] || [])] }))
      setSessionNotes(s => ({ ...s, [hop]: '' }))
    }
    upsertPmUpdate({ hop, field: 'CX Notes', oldValue: '—', newValue: note.trim(), timestamp: logged_at })
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
      const row = rows[i] as unknown[]
      if (row && row.some(cell => String(cell).trim() === 'HOP')) {
        headerRow = i; break
      }
    }
    if (headerRow === -1) { alert('Could not find header row'); return }

    const headers = rows[headerRow] as string[]
    const col = (name: string) => headers.findIndex(h => String(h).trim() === name)

    const hopCol    = col('HOP')
    const gcCol     = col('General Contractor')
    const nokiaPmCol= col('Nokia PM')
    const opsCol    = col('Viaero Ops Field Ops')
    const ms15fCol  = col('MS15 Implementation Start F')
    const ms15aCol  = col('MS15 Implementation Start A')
    const ms16fCol  = col('MS16 Implementation Ends F')
    const ms16aCol  = col('MS16 Implementation Ends A')
    const ntpCol    = col('NTP A')
    const matCol    = headers.findIndex(h => String(h).trim() === 'Material Received A')
    const matFcCol  = col('Material Forecast +4ish')
    const wpCol     = col('Work Package Approved in QB')
    const pickupCol = col('GC Material Pick-up (A)')
    const spoCol    = headers.findIndex(h => String(h).trim().toLowerCase() === 'cx spo issued')
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
    const siteCmCol    = col('New CM')
    const mssCol       = col('MSS Completed NMS Ready ')
    const powerCol     = col('Power-Up Completion')
    const gcPickupFCol = col('GC Material Pick-up (F)')
    const gcPickupACol = col('GC Material Pick-up (A)')
    const cxNotesCol   = headers.findIndex(h => String(h).trim().replace(/^'+|'+$/g, '') === 'CX Notes:')
    const pathIdCol = headers.findIndex(h => String(h).trim().replace(/^'+|'+$/g, '') === 'Path ID')
    const ntpOwnCol = col('NTP Action Owner')
    const ntpWaitCol= col('NTP is waiting on')
    const don444Col = col('DON 444')
    const siteNameCol = col('Site Name')
    const siteNumCol  = col('Site Number')
    const itwSCol   = col('ITW Schedule Start')
    const itwECol   = col('ITW Schedule Complete')
    const ssSCol    = col('Samsung Schedule Start')
    const ssECol    = col('Samsung Schedule Complete')

    // First pass — collect all rows per HOP. Program-wide (every Nokia PM),
    // same as the Dashboard: the PM scoping used to be baked in here
    // (CJ-only), which meant no other PM's GCs/HOPs ever reached this page
    // at all. Now every PM's DON 444 rows survive parsing, and the PM
    // filter (pmFilter, set via the "All GCs" question, see the reactive
    // gcList below) narrows the view at render time instead — same
    // pattern the Dashboard's pmFilter/hopDetails already use.
    const hopRows = new Map<string, unknown[][]>()

    const TARGET_HOP = 'NE-SQUAW_MOUND-NE-CHADRON'
    let rawRowCount = 0
    let don444SurviveCount = 0
    let sawTargetRaw = false
    let sawTargetAfterDon444 = false

    for (let i = headerRow + 1; i < rows.length; i++) {
      const row = rows[i] as unknown[]
      rawRowCount++
      const hopRaw = String(row[hopCol] || '').trim()
      if (hopRaw === TARGET_HOP) sawTargetRaw = true

      const don = String(row[don444Col] || '').trim().toUpperCase()
      if (don !== 'DON 444') continue
      don444SurviveCount++
      if (hopRaw === TARGET_HOP) sawTargetAfterDon444 = true

      const hop = String(row[hopCol] || '').trim()
      if (!hop || hop === 'undefined') continue
      if (!hopRows.has(hop)) hopRows.set(hop, [])
      hopRows.get(hop)!.push(row)
    }

    console.log(`[gc-call] processRows: raw rows=${rawRowCount}, survived DON 444=${don444SurviveCount}, unique HOPs after dedup=${hopRows.size}`)
    console.log(`[gc-call] ${TARGET_HOP} present — raw: ${sawTargetRaw}, after DON 444: ${sawTargetAfterDon444}, after dedup: ${hopRows.has(TARGET_HOP)}`)

    // Build site occupancy map from all in-progress HOPs
    // Maps "SiteName|SiteNumber" -> { gc, hop, ms16f }
    const siteOccupancy = new Map<string, { gc: string, hop: string, ms16f: string }>()

    hopRows.forEach((rows2, hop) => {
      rows2.forEach(r => {
        const ms15a = parseDate(r[ms15aCol])
        const ms16a = parseDate(r[ms16aCol])
        const inProg = !!ms15a && !ms16a
        if (!inProg) return
        const siteName   = String(r[siteNameCol] || '').trim()
        const siteNumber = String(r[siteNumCol]  || '').trim()
        const gc         = String(r[gcCol]        || '').trim()
        const ms16f      = parseDateAny(r[ms16fCol])
        if (siteName && siteNumber) {
          const key = `${siteName}|${siteNumber}`.toLowerCase()
          siteOccupancy.set(key, { gc, hop, ms16f: fmtDate(ms16f) })
        }
      })
    })

    const parsed: HOP[] = []

    hopRows.forEach((rows2, hop) => {
      // Prefer the row where GC and Site CM are both populated — the other
      // row for this HOP may be a blank/partial duplicate.
      const row  = rows2.find(r => String(r[gcCol] || '').trim() && String(r[siteCmCol] || '').trim()) || rows2[0]
      const row2 = rows2.find(r => r !== row) || null

      const gc      = String(row[gcCol] || '').trim() || String(row2?.[gcCol] || '').trim()
      const nokiaPm = String(row[nokiaPmCol] || '').trim() || String(row2?.[nokiaPmCol] || '').trim()
      const ms15f   = parseDateAny(row[ms15fCol])
      const ms15a   = parseDate(row[ms15aCol])
      const ms16f   = parseDateAny(row[ms16fCol])
      const ms16a   = parseDate(row[ms16aCol])
      const ntpDate = parseDate(row[ntpCol])
      const matDate = parseDateAny(row[matCol])
      const matFc   = parseDateAny(row[matFcCol])
      const wpDate  = parseDateAny(row[wpCol])
      const pickupD = parseDateAny(row[pickupCol])
      const spoDate  = parseDateAny(row[spoCol]) || (row2 ? parseDateAny(row2[spoCol]) : null)
      const steelFrom = (() => {
        const idx = headers.findIndex(h => String(h).trim() === 'Steel From')
        if (idx === -1) return ''
        const v1 = String(row[idx] || '').trim().replace(/^'+|'+$/g, '').trim()
        const v2 = String(row2?.[idx] || '').trim().replace(/^'+|'+$/g, '').trim()
        return v1 || v2 || ''
      })()
      const itwS    = parseDateAny(row[itwSCol]) || (row2 ? parseDateAny(row2[itwSCol]) : null)
      const itwE    = parseDateAny(row[itwECol]) || (row2 ? parseDateAny(row2[itwECol]) : null)
      const ssS     = parseDateAny(row[ssSCol])  || (row2 ? parseDateAny(row2[ssSCol])  : null)
      const ssE     = parseDateAny(row[ssECol])  || (row2 ? parseDateAny(row2[ssECol])  : null)

      const hasNtp     = !!(ntpDate && ntpDate.getFullYear() >= 2025)
      const hasMat     = !!(matDate && matDate.getFullYear() >= 2020)
      const wpApproved = !!wpDate
      const gcPickup   = !!pickupD
      const started    = !!ms15a
      const complete   = !!ms16a
      const inProgress = started && !complete
      const daysOut    = ms15f ? daysBetween(today, ms15f) : null
      const daysElapsed = inProgress && ms15a ? daysBetween(ms15a, today) : null

      // Build vendor window by checking ALL site rows for worst case conflict
      const allVendorParts: string[] = []

      rows2.forEach(r => {
        const rItwS  = parseDateAny(r[itwSCol])
        const rItwE  = parseDateAny(r[itwECol])
        const rSsS   = parseDateAny(r[ssSCol])
        const rSsE   = parseDateAny(r[ssECol])
        const rMs15f = parseDateAny(r[ms15fCol]) || ms15f

        if (!rMs15f) return
        if (!rSsS && !rSsE && !rItwS && !rItwE) return

        const siteName = String(r[siteNameCol] || '').trim()
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
            if (buf <= 5)                            allVendorParts.push(`🔴 ${name} clears ${fmtDM(end)} — only ${buf}d before start${siteLabel}`)
            else if (buf <= thresholds.pullInBufferDays) allVendorParts.push(`⚠️ ${name} clears ${fmtDM(end)} — ${buf}d buffer${siteLabel}`)
            else                                      allVendorParts.push(`✅ ${name} clears ${fmtDM(end)} — ${buf}d buffer${siteLabel}`)
          } else {
            const buf = Math.round((startTime - ms15fTime) / (1000 * 60 * 60 * 24))
            if (buf <= thresholds.pullInBufferDays) allVendorParts.push(`⚠️ ${name} starts ${fmtDM(start)} — ${buf}d after start${siteLabel}`)
            else                                  allVendorParts.push(`✅ ${name} starts ${fmtDM(start)} — ${buf}d after start${siteLabel}`)
          }
        }
        checkV('ITW', rItwS, rItwE)
        checkV('Samsung', rSsS, rSsE)
      })

      const itwParts = Array.from(new Set(allVendorParts.filter(p => p.includes('ITW'))))
      const ssParts  = Array.from(new Set(allVendorParts.filter(p => p.includes('Samsung'))))

      const sortParts = (parts: string[]) => {
        const red    = parts.filter(p => p.includes('🔴'))
        const yellow = parts.filter(p => p.includes('⚠️'))
        const green  = parts.filter(p => p.includes('✅'))
        return [...red, ...yellow, ...green]
      }

      const itwSorted = sortParts(itwParts)
      const ssSorted  = sortParts(ssParts)

      const allParts = [...itwSorted, ...ssSorted].filter(Boolean)
      const vendorWindow = allParts.length > 0 ? allParts.join(' | ') : '✅ No conflicts'

      // Check internal conflicts at site level
      let siteAConflict = ''
      let siteBConflict = ''

      rows2.forEach((r, idx) => {
        const siteName   = String(r[siteNameCol] || '').trim()
        const siteNumber = String(r[siteNumCol]  || '').trim()
        const thisGc     = String(r[gcCol] || '').trim()
        if (!siteName || !siteNumber) return
        const key = `${siteName}|${siteNumber}`.toLowerCase()
        const occupant = siteOccupancy.get(key)
        if (occupant && occupant.gc !== thisGc && occupant.hop !== hop) {
          const conflictMsg = `⚠️ ${occupant.gc} on site — ${occupant.hop} completing ${occupant.ms16f}`
          if (idx === 0) siteAConflict = conflictMsg
          else siteBConflict = conflictMsg
        }
      })

      let internalConflict = ''
      if (siteAConflict && siteBConflict) {
        internalConflict = `🔴 Both sites occupied — Site A: ${siteAConflict} | Site B: ${siteBConflict}`
      } else if (siteAConflict) {
        internalConflict = `⚠️ Site A occupied — ${siteAConflict} | Site B available`
      } else if (siteBConflict) {
        internalConflict = `⚠️ Site B occupied — ${siteBConflict} | Site A available`
      }

      const hopObj: HOP = {
        hop, gc, nokiaPm,
        pathId:       String(row[pathIdCol] || '').trim().replace(/^'+|'+$/g, ''),
        ops:          String(row[opsCol] || '').trim(),
        ms15f:        fmtDate(ms15f),
        ms15a:        fmtDate(ms15a),
        ms16f:        fmtDate(ms16f),
        ms16a:        fmtDate(ms16a),
        hasNtp, hasMat, wpApproved, gcPickup,
        ntpOwner:     String(row[ntpOwnCol] || '').trim() || String(row2?.[ntpOwnCol] || '').trim(),
        ntpWaitingOn: String(row[ntpWaitCol] || '').trim() || String(row2?.[ntpWaitCol] || '').trim(),
        matForecast:  fmtDate(matFc),
        matReceived:  matDate ? fmtDate(matDate) : '',
        gcPickupDate: fmtDate(pickupD),
        hasSpo:       !!(spoDate && spoDate.getFullYear() >= 2020),
        hasCpo:       false,
        spoIssued:    spoDate ? fmtDate(spoDate) : '',
        steelFrom:    steelFrom,
        itwStart:     fmtDate(itwS),
        itwEnd:       fmtDate(itwE),
        ssStart:      fmtDate(ssS),
        ssEnd:        fmtDate(ssE),
        daysOut, daysElapsed, inProgress, complete,
        vendorWindow,
        internalConflict,
        siteAConflict,
        siteBConflict,
        blockers: [],
        pullInReady: hasNtp && hasMat && !vendorWindow.includes('🔴') && !vendorWindow.includes('⚠️') && !inProgress && !complete,
        pullInStatus: '',
        cm:        String(row[siteCmCol] || '').trim() || String(row2?.[siteCmCol] || '').trim(),
        mss:       fmtDate(parseDateAny(row[mssCol])),
        powerUp:   fmtDate(parseDateAny(row[powerCol])),
        gcPickupF: fmtDate(parseDateAny(row[gcPickupFCol])),
        gcPickupA: fmtDate(parseDateAny(row[gcPickupACol])),
        cxNotes:   String(row[cxNotesCol] || '').trim()
      }
      hopObj.blockers    = getBlockers(hopObj)
      hopObj.pullInStatus = getPullInStatus(hopObj)
      if (hop === 'NE-SQUAW_MOUND-NE-CHADRON') {
        console.log(`[gc-call] SQUAW_MOUND gc field: '${gc}'`)
      }
      parsed.push(hopObj)
    })

    const techCxHops = parsed.filter(h => h.gc?.trim().toLowerCase() === 'tech cx')
    console.log('[gc-call] tech cx hops after dedup:', techCxHops.map(h => h.hop).join(', '))

    setHops(parsed)

    // Nokia PM options for the filter row — 'ALL' plus every PM actually
    // present in this upload, same derivation as the Dashboard's pmOptions.
    // gcList itself (which GC tabs to show) is computed at render time
    // below, since it needs to react to pmFilter changing after upload.
    const pmSet = new Set<string>()
    parsed.forEach(h => { if (h.nokiaPm) pmSet.add(h.nokiaPm) })
    setPmOptions(['ALL', ...Array.from(pmSet).sort()])
    setLoaded(true)
  }, [thresholds])

  useEffect(() => {
    const loadFromSnapshot = async () => {
      const snap = await loadTrackerSnapshot()
      if (!snap) return
      console.log('[gc-call] fetched', snap.data.length, 'rows from Supabase')
      console.log('[gc-call] NE-SQUAW_MOUND-NE-CHADRON in fetched data:', snap.data.some(row => row.some(cell => String(cell).trim() === 'NE-SQUAW_MOUND-NE-CHADRON')))
      setSnapshotTime(new Date(snap.uploaded_at).toLocaleDateString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric' }) + ' at ' + new Date(snap.uploaded_at).toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' }))
      setFileName(snap.filename)
      setTrackerRawRows(snap.data)
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
      if (!ws) { alert('HOPs tab not found in tracker'); return }
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 }) as unknown[][]
      setSelectedGC('')
      processRows(rows, file.name)
    }
    reader.readAsArrayBuffer(file)
  }, [processRows])

  const logDateEdit = (hop: string, field: string, oldVal: string, newVal: string) => {
    if (!newVal || newVal === oldVal) return
    setEditedDates(prev => ({ ...prev, [hop]: { ...(prev[hop] || {}), [field]: newVal } }))
    upsertPmUpdate({ hop, field, oldValue: oldVal, newValue: newVal, timestamp: new Date().toISOString() })
  }

  // Matches the Nokia PM filter — 'ALL' passes everything.
  const matchesPmFilter = (h: HOP) => pmFilter === 'ALL' || h.nokiaPm?.trim().toUpperCase() === pmFilter.toUpperCase()

  // GC tab list (+ per-GC outstanding count for the tab badge) — dedup
  // case-insensitively (rows can spell the same GC differently) but keep
  // canonical display casing so the tab label, selectedGC, and every
  // gcContactEmails / GC_CM_MAP lookup keyed off it line up with what's
  // typed in Settings. Recomputed every render (not useState) so both react
  // immediately to pmFilter as well as to hops. Only lists a GC if they
  // have at least one non-complete HOP matching the current PM filter — a
  // GC with nothing outstanding for the selected PM has nothing to call
  // about, so their tab shouldn't show; gcOutstandingCounts is exactly that
  // same non-complete, PM-filtered count, per GC, for the "quick view" badge.
  const { gcList, gcOutstandingCounts } = (() => {
    const seenGc = new Map<string, string>() // lowercase key -> canonical display
    const counts = new Map<string, number>() // lowercase key -> outstanding count
    hops.forEach(h => {
      if (h.complete || !matchesPmFilter(h)) return
      const raw = h.gc?.trim()
      if (!raw) return
      const key = raw.toLowerCase()
      if (!seenGc.has(key)) seenGc.set(key, canonicalGcName(raw))
      counts.set(key, (counts.get(key) || 0) + 1)
    })
    return { gcList: Array.from(seenGc.values()).sort(), gcOutstandingCounts: counts }
  })()

  const gcHops      = hops.filter(h => h.gc?.trim().toLowerCase() === selectedGC?.trim().toLowerCase() && matchesPmFilter(h))
  const active      = gcHops.filter(h => h.inProgress).sort((a, b) => {
    const aTime = a.ms16f ? new Date(a.ms16f).getTime() : Infinity
    const bTime = b.ms16f ? new Date(b.ms16f).getTime() : Infinity
    return aTime - bTime
  })

  const maxCrews = crewCountForGc(program, selectedGC)
  const crewFilter = crewFilterByGc[selectedGC] || 'all'
  const setCrewFilter = (val: string) => setCrewFilterByGc(prev => ({ ...prev, [selectedGC]: val }))
  const showCrewBadge = crewFilter === 'all'

  const matchesCrewFilter = (h: HOP) => {
    if (crewFilter === 'all') return true
    if (crewFilter === 'unassigned') return !crewAssignments[h.hop]
    return crewAssignments[h.hop] === crewFilter
  }
  const crewSortKey = (h: HOP) => {
    const c = crewAssignments[h.hop]
    if (!c) return 999
    const n = parseInt(c.replace('Crew ', ''), 10)
    return isNaN(n) ? 999 : n
  }
  const sortByCrewThenDate = (a: HOP, b: HOP) => {
    if (crewFilter === 'all') {
      const ck = crewSortKey(a) - crewSortKey(b)
      if (ck !== 0) return ck
    }
    return (a.daysOut ?? 0) - (b.daysOut ?? 0)
  }

  const thisWeek    = gcHops.filter(h => !h.inProgress && !h.complete && h.daysOut !== null && h.daysOut >= 0 && h.daysOut <= 7 && matchesCrewFilter(h)).sort(sortByCrewThenDate)
  const next2Weeks  = gcHops.filter(h => !h.inProgress && !h.complete && h.daysOut !== null && h.daysOut > 7 && h.daysOut <= 14 && matchesCrewFilter(h)).sort(sortByCrewThenDate)
  const thisMonth   = gcHops.filter(h => !h.inProgress && !h.complete && h.daysOut !== null && h.daysOut > 14 && h.daysOut <= 30 && matchesCrewFilter(h)).sort(sortByCrewThenDate)
  const pullIns     = gcHops.filter(h => !h.inProgress && !h.complete && h.daysOut !== null && h.daysOut > 30 && matchesCrewFilter(h)).sort(sortByCrewThenDate)
  const pullInReady = pullIns.filter(h => h.pullInReady)

  const squawMound = gcHops.find(h => h.hop === 'NE-SQUAW_MOUND-NE-CHADRON')
  if (squawMound) {
    const bucket =
      squawMound.inProgress ? 'active' :
      squawMound.complete ? 'NONE — complete=true excludes it from every bucket' :
      squawMound.daysOut === null ? 'NONE — daysOut is null (no parseable MS15F), excludes it from every bucket' :
      squawMound.daysOut <= 7 ? 'thisWeek' :
      squawMound.daysOut <= 14 ? 'next2Weeks' :
      squawMound.daysOut <= 30 ? 'thisMonth' : 'pullIns'
    console.log('[gc-call] SQUAW_MOUND render-split check:', {
      ms15f: squawMound.ms15f, ms15a: squawMound.ms15a,
      ms16f: squawMound.ms16f, ms16a: squawMound.ms16a,
      daysOut: squawMound.daysOut, inProgress: squawMound.inProgress, complete: squawMound.complete,
      resolvedBucket: bucket
    })
  } else {
    console.log('[gc-call] SQUAW_MOUND render-split check: not present in gcHops for selectedGC =', selectedGC)
  }

  const generateEmail = () => {
    const cm   = GC_CM_MAP[selectedGC] || 'CM'
    const date = today.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })
    const subj = `Viaero MW Program — Weekly Site Update | ${selectedGC} | ${date}`
    const div  = '─'.repeat(60)
    const starDiv = '═'.repeat(60)

    let body = `Dear ${selectedGC} Team,\n\n`
    body += `Please find below and attached your weekly updates to include Cx Pipeline, SPOs, & CR report.\n`
    body += `${div}\n\n`

    if (active.length > 0) {
      body += `${starDiv}\n`
      body += `★★★  ACTIVE SITES (${active.length})  ★★★\n`
      body += `${starDiv}\n\n`
      active.forEach(h => {
        const status = (h.daysElapsed ?? 0) > thresholds.durationAlertDays
          ? `⚠️ OVER TARGET — ${h.daysElapsed}d elapsed — confirm completion date with crew`
          : `✅ On track — ${h.daysElapsed}d elapsed`
        const spoStatusActive = h.hasSpo ? '✓ Issued' : h.hasCpo ? '⚡ Cut Now' : '🔴 Chase CPO'
        body += `★ ${h.hop} ★`
        if (h.pathId) body += `  |  Path ID: ${h.pathId}`
        body += '\n'
        body += `  SPO: ${spoStatusActive}\n`
        body += `  AC Start: ${h.ms15a || '—'}  |  FC Complete: ${h.ms16f || '—'}\n`
        body += `  ${status}\n`
        const latestNote = (noteHistory[h.hop] || []).length > 0
          ? `  💬 Latest Note: ${new Date((noteHistory[h.hop][0].logged_at)).toLocaleDateString('en-US', {month:'numeric',day:'numeric'})} — ${noteHistory[h.hop][0].note}`
          : ''
        if (latestNote) body += `${latestNote}\n`
        if (sessionNotes[h.hop]) body += `  Note: ${sessionNotes[h.hop]}\n`
        body += '\n'
      })
    }

    const upcoming = [...thisWeek, ...next2Weeks]
    if (upcoming.length > 0) {
      body += `${starDiv}\n`
      body += `★★★  STARTING WITHIN 2 WEEKS (${upcoming.length})  ★★★\n`
      body += `${starDiv}\n\n`
      upcoming.forEach(h => {
        const spoStatus = h.hasSpo ? '✓ Issued' : h.hasCpo ? '⚡ Cut Now' : '🔴 Chase CPO'
        const steelNote = h.steelFrom === 'ITW'
          ? `ITW — confirm ITW delivery schedule`
          : h.steelFrom || '—'
        body += `★ ${h.hop} ★`
        if (h.pathId) body += `  |  Path ID: ${h.pathId}`
        body += '\n'
        body += `  SPO: ${spoStatus}\n`
        body += `  NTP: ${h.hasNtp ? '✓' : '✗'}`
        if (!h.hasNtp && h.ntpWaitingOn) body += `  |  Waiting On: ${h.ntpWaitingOn}`
        body += '\n'
        body += `  Mat: ${h.hasMat ? '✓' : '✗'}  |  Steel From: ${steelNote}  |  GC Pickup F: ${h.gcPickupDate || '—'}  |  GC Pickup A: ${h.gcPickupDate || '✗'}\n`
        body += `  Vendor: ${h.vendorWindow.includes('🔴') ? h.vendorWindow : '✅ Clear'}\n`
        if (h.blockers.length > 0) {
          const blockerText = h.blockers.map(b => {
            if (b.includes('NTP') && h.ntpWaitingOn) return `${b} — Waiting On: ${h.ntpWaitingOn}`
            return b
          }).join(' | ')
          body += `  Blockers: ${blockerText}\n`
        }
        if (h.internalConflict) body += `  Internal Conflict: ${h.internalConflict}\n`
        if (sessionNotes[h.hop]) body += `  Note: ${sessionNotes[h.hop]}\n`
        body += `  FC Start: ${h.ms15f}  |  Days Out: ${h.daysOut}d\n`
        const latestNote = (noteHistory[h.hop] || []).length > 0
          ? `  💬 Latest Note: ${new Date((noteHistory[h.hop][0].logged_at)).toLocaleDateString('en-US', {month:'numeric',day:'numeric'})} — ${noteHistory[h.hop][0].note}`
          : ''
        if (latestNote) body += `${latestNote}\n`
        body += '\n'
      })
    }

    if (thisMonth.length > 0) {
      body += `${starDiv}\n`
      body += `★★★  THIS MONTH — 15 TO 30 DAYS (${thisMonth.length})  ★★★\n`
      body += `${starDiv}\n\n`
      thisMonth.forEach(h => {
        const spoStatus = h.hasSpo ? '✓ Issued' : h.hasCpo ? '⚡ Cut Now' : '🔴 Chase CPO'
        const steelNote = h.steelFrom === 'ITW'
          ? `ITW — confirm ITW delivery schedule`
          : h.steelFrom || '—'
        body += `★ ${h.hop} ★`
        if (h.pathId) body += `  |  Path ID: ${h.pathId}`
        body += '\n'
        body += `  SPO: ${spoStatus}\n`
        body += `  NTP: ${h.hasNtp ? '✓' : '✗'}`
        if (!h.hasNtp && h.ntpWaitingOn) body += `  |  Waiting On: ${h.ntpWaitingOn}`
        body += '\n'
        body += `  Mat: ${h.hasMat ? '✓' : '✗'}  |  Steel From: ${steelNote}  |  GC Pickup F: ${h.gcPickupDate || '—'}  |  GC Pickup A: ${h.gcPickupDate || '✗'}\n`
        body += `  Vendor: ${h.vendorWindow.includes('🔴') ? h.vendorWindow : '✅ Clear'}\n`
        if (h.blockers.length > 0) body += `  Blockers: ${h.blockers.join(' | ')}\n`
        if (h.internalConflict) body += `  Internal Conflict: ${h.internalConflict}\n`
        if (sessionNotes[h.hop]) body += `  Note: ${sessionNotes[h.hop]}\n`
        body += `  FC Start: ${h.ms15f}  |  Days Out: ${h.daysOut}d\n`
        const latestNote = (noteHistory[h.hop] || []).length > 0
          ? `  💬 Latest Note: ${new Date((noteHistory[h.hop][0].logged_at)).toLocaleDateString('en-US', {month:'numeric',day:'numeric'})} — ${noteHistory[h.hop][0].note}`
          : ''
        if (latestNote) body += `${latestNote}\n`
        body += '\n'
      })
    }

    if (pullInReady.length > 0) {
      body += `${starDiv}\n`
      body += `★★★  PULL-IN OPPORTUNITIES (${pullInReady.length})  ★★★\n`
      body += `${starDiv}\n\n`
      body += `The following sites are ready to accelerate if schedule allows:\n\n`
      pullInReady.forEach(h => {
        body += `★ ${h.hop} ★  |  FC Start: ${h.ms15f}  |  ${h.daysOut}d out  |  NTP ✓  |  Mat ✓\n`
        if (sessionNotes[h.hop]) body += `  Note: ${sessionNotes[h.hop]}\n`
        body += '\n'
      })
    }

    const actionItems = [...thisWeek, ...next2Weeks].filter(h => h.blockers.length > 0)
    if (actionItems.length > 0) {
      body += `${starDiv}\n`
      body += `★★★  ACTION ITEMS REQUIRED  ★★★\n`
      body += `${starDiv}\n\n`
      actionItems.forEach((h, i) => {
        body += `${i + 1}. ★ ${h.hop} ★ (starts ${h.ms15f})\n`
        h.blockers.forEach(b => body += `   ${b}\n`)
        body += '\n'
      })
    }

    body += `${div}\n`
    body += `Please coordinate with your Site CM ${cm} for all field questions.\n`
    body += `For schedule, finance, or contract matters contact CJ directly.`

    const ccList = emailSettings.ccList.join(',')
    const to = lookupContactEmail(emailSettings.gcContactEmails, selectedGC)

    window.open(`mailto:${to}?cc=${encodeURIComponent(ccList)}&subject=${encodeURIComponent(subj)}&body=${encodeURIComponent(body)}`)
  }

  const downloadGCExcel = () => {
    try {
      const wb = XLSX.utils.book_new()
      const date = today.toLocaleDateString('en-US').replace(/\//g, '-')

      // Active Sites sheet
      const activeRows: (string | number)[][] = [
        ['HOP', 'CM', 'Days Elapsed', 'AC Start', 'FC End', 'AC End', 'MSS', 'Power-Up', 'SPO', 'Vendor Window', 'Notes']
      ]
      active.forEach(h => {
        const spoStatus = h.hasSpo ? '✓ Issued' : h.hasCpo ? '⚡ Cut Now' : '🔴 Chase CPO'
        const latestNote = (noteHistory[h.hop] || []).slice(0, 1).map(n => `${new Date(n.logged_at).toLocaleDateString()}: ${n.note}`).join('')
        activeRows.push([
          h.hop, h.cm || '—',
          h.daysElapsed !== null ? `${h.daysElapsed}d` : '—',
          h.ms15a || '—', h.ms16f || '—', h.ms16a || '—',
          h.mss || '—', h.powerUp || '—',
          spoStatus,
          h.vendorWindow.includes('🔴') ? h.vendorWindow : '✅ Clear',
          latestNote || sessionNotes[h.hop] || ''
        ])
      })
      const activeSheet = XLSX.utils.aoa_to_sheet(activeRows)
      activeSheet['!cols'] = [{ wch: 36 },{ wch: 12 },{ wch: 12 },{ wch: 12 },{ wch: 12 },{ wch: 12 },{ wch: 12 },{ wch: 12 },{ wch: 14 },{ wch: 30 },{ wch: 40 }]
      XLSX.utils.book_append_sheet(wb, activeSheet, '🔨 Active Sites')

      // Pipeline sheet — all upcoming HOPs
      const pipelineRows: (string | number)[][] = [
        ['HOP', 'CM', 'Days Out', 'NTP', 'NTP Waiting On', 'Mat', 'Steel From', 'GC Pickup F', 'GC Pickup A', 'SPO', 'FC Start', 'AC Start', 'Vendor Window', 'Pull-In Status', 'Notes']
      ]
      const allPipeline = [...thisWeek, ...next2Weeks, ...thisMonth, ...pullIns]
      allPipeline.forEach(h => {
        const spoStatus = h.hasSpo ? '✓ Issued' : h.hasCpo ? '⚡ Cut Now' : '🔴 Chase CPO'
        const pullIn = h.pullInReady ? '✅ Ready' : h.pullInStatus.includes('⚠️') ? '⚠️ Risky' : '🔴 Cannot'
        const latestNote = (noteHistory[h.hop] || []).slice(0, 1).map(n => `${new Date(n.logged_at).toLocaleDateString()}: ${n.note}`).join('')
        pipelineRows.push([
          h.hop, h.cm || '—',
          h.daysOut !== null ? `${h.daysOut}d` : '—',
          h.hasNtp ? '✓' : '✗',
          h.ntpWaitingOn || '—',
          h.hasMat ? '✓' : '✗',
          h.steelFrom || '—',
          h.gcPickupF || '—',
          h.gcPickupA || '—',
          spoStatus,
          h.ms15f || '—',
          h.ms15a || '—',
          h.vendorWindow.includes('🔴') ? h.vendorWindow : '✅ Clear',
          pullIn,
          latestNote || sessionNotes[h.hop] || ''
        ])
      })
      const pipelineSheet = XLSX.utils.aoa_to_sheet(pipelineRows)
      pipelineSheet['!cols'] = [{ wch: 36 },{ wch: 12 },{ wch: 10 },{ wch: 6 },{ wch: 40 },{ wch: 6 },{ wch: 12 },{ wch: 14 },{ wch: 14 },{ wch: 14 },{ wch: 12 },{ wch: 12 },{ wch: 30 },{ wch: 14 },{ wch: 40 }]
      XLSX.utils.book_append_sheet(wb, pipelineSheet, '📋 Pipeline')

      XLSX.writeFile(wb, `${selectedGC}_Pipeline_${date}.xlsx`)
    } catch (err) {
      console.error('Download error:', err)
      alert('Download failed — please try again')
    }
  }

  return (
    <div className="min-h-screen bg-gray-950 text-white p-6">
      <div className="max-w-full mx-auto">

        <BackToDashboard />

        {/* Header */}
        <div className="flex items-center justify-between mb-6">
          <div>
            <h1 className="text-3xl font-bold">GC Call View</h1>
            <p className="text-gray-400 mt-1">Select a contractor to view their full pipeline, blockers, and generate a follow-up email.</p>
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
            and CM View (app/lib/pendingUpdates.ts): unified table for
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
                        <input
                          type="checkbox"
                          checked={u.completed || false}
                          onChange={() => toggleUpdateCompleted(u.hop, u.field, u.timestamp)}
                          className="w-4 h-4 cursor-pointer accent-green-600"
                        />
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
            <p className="text-gray-500 text-xs">{fileName} — {hops.length} HOPs · Upload new tracker on Dashboard to refresh</p>
          </div>
        )}
        {!snapshotTime && (
          <div className="mb-4 bg-gray-900 border border-gray-700 rounded-lg px-4 py-8 text-center">
            <p className="text-gray-400">No tracker data found — go to Dashboard to upload your tracker</p>
          </div>
        )}

        {/* Nokia PM filter — same pattern as the Dashboard's pmFilter row */}
        {pmOptions.length > 1 && (
          <div className="flex gap-2 mb-4 flex-wrap items-center">
            <span className="text-gray-500 text-xs font-semibold">Nokia PM:</span>
            {pmOptions.map(pm => (
              <button key={pm} onClick={() => setPmFilter(pm)}
                className={`px-4 py-1.5 rounded-lg text-sm font-semibold transition-all ${pmFilter === pm ? 'bg-blue-600 text-white' : 'bg-gray-800 text-gray-300 hover:bg-gray-700'}`}>
                {pm}
              </button>
            ))}
          </div>
        )}

        {/* GC Selector */}
        <div className="flex gap-3 mb-6 flex-wrap">
          {gcList.map((gc) => {
            const isSelected = selectedGC?.trim().toLowerCase() === gc?.trim().toLowerCase()
            return (
              <button key={gc} onClick={() => setSelectedGC(gc)}
                title={`${gcOutstandingCounts.get(gc.toLowerCase()) ?? 0} outstanding HOP${(gcOutstandingCounts.get(gc.toLowerCase()) ?? 0) === 1 ? '' : 's'}`}
                className={`flex items-center gap-2 px-6 py-3 rounded-lg font-semibold text-sm transition-all ${isSelected ? 'bg-blue-600 text-white shadow-lg scale-105' : 'bg-gray-800 text-gray-300 hover:bg-gray-700'}`}>
                {gc}
                <span className={`text-xs font-bold px-2 py-0.5 rounded-full ${isSelected ? 'bg-white/20 text-white' : 'bg-gray-700 text-gray-300'}`}>
                  {gcOutstandingCounts.get(gc.toLowerCase()) ?? 0}
                </span>
              </button>
            )
          })}
        </div>

        {/* GC Panel */}
        {selectedGC && (
          <div className="bg-gray-900 rounded-xl border border-gray-700 p-6">
            <div className="flex items-center justify-between mb-6">
              <div>
                <h2 className="text-2xl font-bold">{selectedGC}</h2>
                <p className="text-gray-400 mt-1">
                  Site CM: <span className="text-blue-400 font-semibold">{GC_CM_MAP[selectedGC] || 'See Contacts'}</span>
                  {loaded && (
                    <span className="text-gray-500">
                      {' · '}{gcHops.length} HOPs{' · '}
                      {active.length} active{' · '}
                      {thisWeek.length} this week{' · '}
                      {next2Weeks.length} next 2 wks{' · '}
                      {thisMonth.length} this month{' · '}
                      <span className="text-green-400">{pullInReady.length} pull-in ready</span>
                    </span>
                  )}
                </p>
              </div>
              <div className="flex gap-2">
                <button onClick={generateEmail}
                  className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-semibold">
                  ✉️ Email
                </button>
                <button onClick={downloadGCExcel}
                  className="bg-green-700 hover:bg-green-600 text-white px-4 py-2 rounded-lg text-sm font-semibold">
                  📥 Download Excel
                </button>
              </div>
            </div>

            {/* Tab Bar */}
            <div className="flex gap-2 mb-6 border-b border-gray-800 pb-3">
              {([
                { key: 'pipeline', label: 'Pipeline' },
                { key: 'gr', label: '💰 GR / Invoicing' },
                { key: 'decom', label: 'Decom' },
                { key: 'reports', label: 'Reports' },
              ] as const).map(t => (
                <button key={t.key} onClick={() => setActiveTab(t.key)}
                  className={`px-4 py-2 rounded-lg text-sm font-semibold transition-all ${activeTab === t.key ? 'bg-blue-600 text-white' : 'bg-gray-800 text-gray-300 hover:bg-gray-700'}`}>
                  {t.label}
                </button>
              ))}
            </div>

            {activeTab === 'pipeline' && (
            <div className="space-y-8">

                {/* Crew Filter */}
                <div className="flex gap-2 flex-wrap">
                  <button onClick={() => setCrewFilter('all')}
                    className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all ${crewFilter === 'all' ? 'bg-blue-600 text-white' : 'bg-gray-800 text-gray-300 hover:bg-gray-700'}`}>
                    All Crews
                  </button>
                  {Array.from({ length: maxCrews }, (_, i) => `Crew ${i + 1}`).map(c => (
                    <button key={c} onClick={() => setCrewFilter(c)}
                      className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all ${crewFilter === c ? 'bg-blue-600 text-white' : 'bg-gray-800 text-gray-300 hover:bg-gray-700'}`}>
                      {c}
                    </button>
                  ))}
                  <button onClick={() => setCrewFilter('unassigned')}
                    className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all ${crewFilter === 'unassigned' ? 'bg-blue-600 text-white' : 'bg-gray-800 text-gray-300 hover:bg-gray-700'}`}>
                    Unassigned
                  </button>
                </div>

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
                              <th className="text-left p-2">Crew</th>
                              <th className="text-left p-2">Path ID</th>
                              <th className="text-left p-2">Started</th>
                              <th className="text-left p-2">FC End</th>
                              <th className="text-left p-2">Days Elapsed</th>
                              <th className="text-left p-2">Status</th>
                              <th className="text-left p-2">SPO Issued</th>
                              <th className="text-left p-2">MS16 Fc</th>
                              <th className="text-left p-2">Edit MS16 Fc</th>
                              <th className="text-left p-2">MS16 Act</th>
                              <th className="text-left p-2">Call Notes (Today)</th>
                              <th className="text-left p-2">Notes History</th>
                              <th className="text-left p-2">CX Notes</th>
                            </tr>
                          </thead>
                          <tbody>
                            {active.map((h) => (
                              <tr key={h.hop} className={`border-t border-gray-800 ${(h.daysElapsed ?? 0) > thresholds.durationAlertDays ? 'bg-red-950' : 'bg-gray-900'}`}>
                                <td className="p-2 font-semibold text-white whitespace-nowrap">
                                  {h.hop}
                                  {showCrewBadge && crewAssignments[h.hop] && (
                                    <span className={`ml-2 text-xs font-bold px-1.5 py-0.5 rounded ${CREW_BADGE_COLORS[(parseInt(crewAssignments[h.hop].replace('Crew ', ''), 10) - 1) % CREW_BADGE_COLORS.length] || 'bg-gray-700 text-gray-300'}`}>
                                      C{crewAssignments[h.hop].replace('Crew ', '')}
                                    </span>
                                  )}
                                </td>
                                <td className="p-2">
                                  <select value={crewAssignments[h.hop] || ''}
                                    onChange={(e) => setCrewForHop(h.hop, e.target.value)}
                                    className="bg-gray-800 text-gray-300 text-xs rounded px-1 py-1 border border-gray-600 focus:outline-none focus:border-blue-500">
                                    <option value="">--</option>
                                    {Array.from({ length: maxCrews }, (_, i) => `Crew ${i + 1}`).map(c => (
                                      <option key={c} value={c}>{c}</option>
                                    ))}
                                  </select>
                                </td>
                                <td className="p-2 text-gray-400 text-xs whitespace-nowrap">{h.pathId || '—'}</td>
                                <td className="p-2 text-gray-300 text-xs whitespace-nowrap">{h.ms15a || '—'}</td>
                                <td className="p-2 text-gray-300 text-xs whitespace-nowrap">{h.ms16f || '—'}</td>
                                <td className={`p-2 font-bold ${(h.daysElapsed ?? 0) > thresholds.durationAlertDays ? 'text-red-400' : 'text-green-400'}`}>{h.daysElapsed}d</td>
                                <td className="p-2">{(h.daysElapsed ?? 0) > thresholds.durationAlertDays ? <span className="text-red-400">⚠️ Over {thresholds.durationAlertDays}d</span> : <span className="text-green-400">On track</span>}</td>
                                <td className="p-2">
                                  {h.hasSpo
                                    ? <span className="text-green-400 font-bold text-sm" title={h.spoIssued}>✓</span>
                                    : <span className="text-red-400 font-bold text-sm">✗</span>
                                  }
                                </td>
                                <td className="p-2 text-gray-300 text-xs whitespace-nowrap">{h.ms16f || '—'}</td>
                                <td className="p-2">
                                  <GCEditableDate hop={h.hop} field="MS16 Implementation Ends F" value={h.ms16f} editedDates={editedDates} logDateEdit={logDateEdit} />
                                </td>
                                <td className="p-2">
                                  <GCEditableDate hop={h.hop} field="MS16 Implementation Ends A" value={h.ms16a} editedDates={editedDates} logDateEdit={logDateEdit} />
                                </td>
                                <td className="p-2">
                                  <div className="flex gap-1">
                                    <input type="text" placeholder="Note..." value={sessionNotes[h.hop] || ''}
                                      onChange={(e) => setSessionNotes(s => ({ ...s, [h.hop]: e.target.value }))}
                                      onKeyDown={(e) => { if (e.key === 'Enter') saveCallNote(h.hop) }}
                                      className="w-36 bg-gray-800 text-white text-xs rounded px-2 py-1 border border-gray-600 focus:outline-none focus:border-blue-500" />
                                    <button onClick={() => saveCallNote(h.hop)} className="text-xs bg-blue-700 hover:bg-blue-600 text-white px-2 py-1 rounded">💾</button>
                                  </div>
                                </td>
                                <td className="p-2 max-w-48">
                                  <div className="max-h-20 overflow-y-auto flex flex-col gap-1">
                                    {(noteHistory[h.hop] || []).slice(0, 5).map((n, i) => (
                                      <div key={i} className="text-xs text-gray-300 border-b border-gray-700 pb-1">
                                        <span className="text-gray-500 text-xs">{new Date(n.logged_at).toLocaleDateString()}</span>
                                        <span className="ml-1">{n.note}</span>
                                      </div>
                                    ))}
                                    {!noteHistory[h.hop]?.length && <span className="text-gray-600 text-xs">No history</span>}
                                  </div>
                                </td>
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
                    )}
                </div>

                {/* Pipeline Sections */}
                <PipelineTable title="⚡ This Week (0–7 days)" rows={thisWeek} sessionNotes={sessionNotes} setSessionNotes={setSessionNotes} saveCallNote={saveCallNote} noteHistory={noteHistory} editedDates={editedDates} logDateEdit={logDateEdit} setCxNotesModal={setCxNotesModal} crewAssignments={crewAssignments} maxCrews={maxCrews} showCrewBadge={showCrewBadge} onCrewChange={setCrewForHop} />
                <PipelineTable title="🟠 Next 2 Weeks (8–14 days)" rows={next2Weeks} sessionNotes={sessionNotes} setSessionNotes={setSessionNotes} saveCallNote={saveCallNote} noteHistory={noteHistory} editedDates={editedDates} logDateEdit={logDateEdit} setCxNotesModal={setCxNotesModal} crewAssignments={crewAssignments} maxCrews={maxCrews} showCrewBadge={showCrewBadge} onCrewChange={setCrewForHop} />
                <PipelineTable title="🟡 This Month (15–30 days)" rows={thisMonth} sessionNotes={sessionNotes} setSessionNotes={setSessionNotes} saveCallNote={saveCallNote} noteHistory={noteHistory} editedDates={editedDates} logDateEdit={logDateEdit} setCxNotesModal={setCxNotesModal} crewAssignments={crewAssignments} maxCrews={maxCrews} showCrewBadge={showCrewBadge} onCrewChange={setCrewForHop} />

                {/* Pull-In Opportunities */}
                <div>
                  <div className="flex items-center gap-3 mb-3">
                    <h3 className="text-lg font-semibold text-white">🚀 Full Pipeline + Pull-In Opportunities (30d+)</h3>
                    {pullInReady.length > 0 && (
                      <span className="bg-green-800 text-green-200 text-xs px-3 py-1 rounded-full font-semibold">
                        {pullInReady.length} ready to pull in
                      </span>
                    )}
                  </div>
                  <PipelineTable title="" rows={pullIns} sessionNotes={sessionNotes} setSessionNotes={setSessionNotes} saveCallNote={saveCallNote} noteHistory={noteHistory} editedDates={editedDates} logDateEdit={logDateEdit} setCxNotesModal={setCxNotesModal} crewAssignments={crewAssignments} maxCrews={maxCrews} showCrewBadge={showCrewBadge} onCrewChange={setCrewForHop} />
                </div>

              </div>
            )}

            {activeTab === 'gr' && (
              <GrInvoicingTab
                selectedGC={selectedGC}
                grRows={grRows}
                grLoaded={grLoaded}
                emailSettings={emailSettings}
              />
            )}

            {activeTab === 'decom' && (
              <DecomTab
                selectedGC={selectedGC}
                decomRawRows={decomRawRows}
                trackerRawRows={trackerRawRows}
                emailSettings={emailSettings}
              />
            )}

            {activeTab === 'reports' && (
              <GcReportsTab
                selectedGC={selectedGC}
                spoRawRows={spoRawRows}
                crRawRows={crRawRows}
                decomRawRows={decomRawRows}
              />
            )}

          </div>
        )}

        {!selectedGC && (
          <div className="bg-gray-900 rounded-xl border border-gray-700 p-12 text-center">
            <p className="text-gray-400 text-xl">👆 Select a GC above to begin</p>
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
