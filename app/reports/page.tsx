'use client'

export const dynamic = 'force-dynamic'

import { useState, useCallback, useEffect } from 'react'
import { useRouter } from 'next/navigation'
import * as XLSX from 'xlsx'
import { supabase } from '../lib/supabase'
import { GC_CONFIG, matches, SPO_VENDOR_COL_IN_MASTER, CR_SUPPLIER_COL_IN_MASTER } from '../lib/gcConfig'
import BackToDashboard from '../components/BackToDashboard'
import { saveChunkedReport, loadChunkedReport } from '../lib/reportChunks'
import { parseDecomRows, summarizeDecomByGc, decomRowsForGc, DecomRow } from '../lib/decom'

interface ReportSnapshot {
  filename: string
  uploaded_at: string
  row_count: number
}

const SPO_COL_IDX = [7, 8, 33, 40, 41, 43, 47, 48, 49, 50, 51]
const SPO_HEADERS = ['Customer Site ID', 'Name', 'SOG Name', 'SPO Number', 'SPO Creation Date', 'SPO Vendor', 'SPO Value', 'IA Date', 'IA User', 'GR Date', 'GR Number']

const CR_COL_IDX = [0, 1, 2, 4, 5, 6, 7, 8, 9, 10, 11, 16, 18, 19, 25, 26, 27, 28, 29, 30, 31]
const CR_HEADERS = ['Requestor', 'Supplier Name', 'Path ID', 'Site Name', 'Site #', 'Network Site Name', 'Risk Budget', 'Materials or Labor', 'Reason for CR Details', 'Viaero Operation CR Filed', 'CR Type', 'Sellable to Who', 'PM Status', 'PM Status Owner', 'SPO Cost', 'GC Quote Shared', 'CQT Package', 'SPO #', 'SPO Issued Date', 'SPO IA/GR', 'CQT ID']

const NAVY = '124191'
const TEAL = '00A0B0'

// Store only the columns Ciege actually reads (SPO_COL_IDX / CR_COL_IDX above —
// confirmed to be the full set referenced anywhere in the app or backend). Raw
// tracker exports run 50-90+ columns wide, and storing every column was pushing
// the Supabase upsert payload large enough to time out. Rows are truncated to a
// sparse array (nulled-out gaps, original index positions preserved, everything
// past the last needed column dropped entirely) rather than compacted or turned
// into an object — every existing row[N] lookup throughout the app and the
// Python backend (which indexes with a plain int, not a string key) keeps
// working unchanged, since this stays a real array both in JS and after
// json.loads() in Python.
function stripRow(row: unknown[], keepCols: number[]): unknown[] {
  const maxIdx = Math.max(...keepCols)
  const stripped: unknown[] = new Array(maxIdx + 1).fill(null)
  keepCols.forEach(i => { stripped[i] = row[i] ?? null })
  return stripped
}

function fmtDate(val: unknown): string {
  if (!val) return ''
  if (val instanceof Date) return val.toLocaleDateString('en-US')
  const d = new Date(String(val))
  return isNaN(d.getTime()) ? String(val) : d.toLocaleDateString('en-US')
}

function downloadGCReport(rows: unknown[][], colIdx: number[], headers: string[], vendorColInMaster: number, matchList: string[], filename: string) {
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

  // Style header row
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1')
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: 0, c })]
    if (cell) {
      cell.s = {
        font: { bold: true, color: { rgb: 'FFFFFF' } },
        fill: { fgColor: { rgb: NAVY } },
        alignment: { horizontal: 'center', wrapText: true }
      }
    }
  }

  // Auto column widths
  ws['!cols'] = headers.map((h, i) => {
    const maxLen = Math.max(h.length, ...sheetData.slice(1).map(r => String(r[i] || '').length))
    return { wch: Math.min(Math.max(maxLen + 2, 10), 40) }
  })

  XLSX.utils.book_append_sheet(wb, ws, headers[0].includes('SPO') || filename.includes('SPO') ? 'SPO Report' : 'CR Tracker')
  XLSX.writeFile(wb, filename)
}

const DECOM_HEADERS = ['HOP', 'Path ID', 'Site Name', 'Site Number', 'CM', 'CX Start', 'CX Complete', 'Aging (days)', 'DECOM Drop Off Date', 'DECOM Comments', 'POD Pathwave', 'POD QuickBase']

function decomRowToSheetRow(r: DecomRow): (string | number)[] {
  return [
    r.hop, r.pathId, r.siteName, r.siteNumber, r.cm,
    fmtDate(r.cxStart), fmtDate(r.cxComplete),
    r.aging ?? '',
    fmtDate(r.dropOffDate), r.comment,
    r.podPathwave ? 'Yes' : 'No', r.podQuickBase ? 'Yes' : 'No',
  ]
}

function styleDecomSheet(ws: XLSX.WorkSheet) {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1')
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: 0, c })]
    if (cell) {
      cell.s = {
        font: { bold: true, color: { rgb: 'FFFFFF' } },
        fill: { fgColor: { rgb: NAVY } },
        alignment: { horizontal: 'center', wrapText: true }
      }
    }
  }
  ws['!cols'] = DECOM_HEADERS.map(h => ({ wch: Math.max(h.length + 2, 12) }))
}

function downloadGcDecomReport(gcRows: DecomRow[], filename: string) {
  const outstandingPending = gcRows
    .filter(r => r.status === 'outstanding' || r.status === 'pending')
    .sort((a, b) => (b.aging ?? -1) - (a.aging ?? -1))
  const podGap = gcRows.filter(r => r.status === 'pod_gap')

  const wb = XLSX.utils.book_new()

  const ws1 = XLSX.utils.aoa_to_sheet([DECOM_HEADERS, ...outstandingPending.map(decomRowToSheetRow)])
  styleDecomSheet(ws1)
  XLSX.utils.book_append_sheet(wb, ws1, 'Outstanding & Pending')

  const ws2 = XLSX.utils.aoa_to_sheet([DECOM_HEADERS, ...podGap.map(decomRowToSheetRow)])
  styleDecomSheet(ws2)
  XLSX.utils.book_append_sheet(wb, ws2, 'POD Gap')

  XLSX.writeFile(wb, filename)
}

interface UploadBoxProps {
  type: 'spo' | 'cr' | 'decom'
  info: ReportSnapshot | null
  label: string
  uploading: string | null
  onUpload: (file: File, type: 'spo' | 'cr' | 'decom') => void
}

function UploadBox({ type, info, label, uploading, onUpload }: UploadBoxProps) {
  return (
    <div
      className="border-2 border-dashed border-gray-600 rounded-xl p-5 cursor-pointer hover:border-blue-500 transition-colors"
      onDragOver={(e) => e.preventDefault()}
      onDrop={(e) => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) onUpload(f, type) }}
      onClick={() => document.getElementById(`${type}-upload`)?.click()}
    >
      <input id={`${type}-upload`} type="file" accept=".xlsx" className="hidden"
        onChange={(e) => { const f = e.target.files?.[0]; if (f) onUpload(f, type) }} />
      {uploading === type
        ? <p className="text-blue-400 text-sm text-center">⏳ Processing...</p>
        : info
        ? <div>
            <p className="text-green-400 text-sm font-semibold">✅ {label} loaded</p>
            <p className="text-gray-500 text-xs mt-1">{info.filename} · {info.row_count} rows · {new Date(info.uploaded_at).toLocaleDateString('en-US')} at {new Date(info.uploaded_at).toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })}</p>
          </div>
        : <p className="text-gray-400 text-sm text-center">📂 Upload {label}</p>
      }
    </div>
  )
}

export default function ReportsPage() {
  const router = useRouter()
  const [spoRows, setSpoRows] = useState<unknown[][]>([])
  const [crRows, setCrRows] = useState<unknown[][]>([])
  const [decomRawRows, setDecomRawRows] = useState<unknown[][]>([])
  const [spoInfo, setSpoInfo] = useState<ReportSnapshot | null>(null)
  const [crInfo, setCrInfo] = useState<ReportSnapshot | null>(null)
  const [decomInfo, setDecomInfo] = useState<ReportSnapshot | null>(null)
  const [uploading, setUploading] = useState<string | null>(null)
  const today = new Date().toLocaleDateString('en-US').replace(/\//g, '-')

  const decomRows = parseDecomRows(decomRawRows)
  console.log('[decom-render] decomRawRows:', decomRawRows.length, 'rows (incl. header) → decomRows parsed:', decomRows.length)

  useEffect(() => {
    const load = async () => {
      const { data: spoSnap } = await supabase.from('report_snapshots').select('*').eq('id', 'spo').single()
      if (spoSnap) { setSpoRows(JSON.parse(spoSnap.data)); setSpoInfo({ filename: spoSnap.filename, uploaded_at: spoSnap.uploaded_at, row_count: JSON.parse(spoSnap.data).length }) }

      const crReport = await loadChunkedReport('cr')
      if (crReport) { setCrRows(crReport.rows); setCrInfo({ filename: crReport.filename, uploaded_at: crReport.uploaded_at, row_count: crReport.rows.length }) }

      const decomReport = await loadChunkedReport('decom')
      console.log('[decom-load] loadChunkedReport("decom") returned', decomReport ? decomReport.rows.length : 0, 'rows')
      if (decomReport) { setDecomRawRows(decomReport.rows); setDecomInfo({ filename: decomReport.filename, uploaded_at: decomReport.uploaded_at, row_count: decomReport.rows.length }) }
    }
    load()
  }, [])

  const handleUpload = useCallback(async (file: File, type: 'spo' | 'cr' | 'decom') => {
    setUploading(type)
    const reader = new FileReader()
    reader.onload = async (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer)
        const wb = XLSX.read(data, { type: 'array', cellDates: true })

        let ws, headerRowIndex
        if (type === 'spo') {
          ws = wb.Sheets['NDPd_NAM_003_Viaero_SPO_Report'] || wb.Sheets[wb.SheetNames[0]]
          headerRowIndex = 0
        } else if (type === 'cr') {
          ws = wb.Sheets['CR Tracker'] || wb.Sheets[wb.SheetNames[0]]
          headerRowIndex = 1
        } else {
          const DECOM_SHEET_NAME = 'R&R - Link Drop Off Status'
          const matchedSheetName = wb.SheetNames.find(n => n.trim() === DECOM_SHEET_NAME)
          ws = matchedSheetName ? wb.Sheets[matchedSheetName] : undefined
          if (!ws) {
            console.error(`[report-upload] Decom upload failed: sheet "${DECOM_SHEET_NAME}" not found. Available sheets:`, wb.SheetNames)
            alert(`Could not find the "${DECOM_SHEET_NAME}" tab in this file — check the tab name and try again.`)
            setUploading(null)
            return
          }
          headerRowIndex = 0
        }

        const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as unknown[][]

        let allRows = rawRows
        if (type === 'cr' || type === 'decom') {
          let lastDataRow = rawRows.length - 1
          while (lastDataRow >= 0 && !(rawRows[lastDataRow] || []).some(c => c !== null && c !== undefined && String(c).trim() !== '')) {
            lastDataRow--
          }
          const rows = rawRows.slice(0, lastDataRow + 1)
          console.log(`[report-upload] ${type.toUpperCase()} raw rows:`, rawRows.length, '→ trimmed:', rows.length)
          allRows = rows
        }

        const dataRows = allRows.slice(headerRowIndex + 1).filter(row => row.some(v => v !== null))

        if (type === 'spo' || type === 'cr') {
          const keepCols = type === 'spo' ? SPO_COL_IDX : CR_COL_IDX
          const totalCols = dataRows.reduce((max, row) => Math.max(max, row.length), 0)
          const strippedRows = dataRows.map(row => stripRow(row, keepCols))
          console.log(`[report-upload] ${type.toUpperCase()} columns: kept ${keepCols.length} of ${totalCols} (dropped ${totalCols - keepCols.length})`)

          if (type === 'spo') {
            console.log('[report-upload] upserting id="spo"', 'filename=', file.name)
            const { error } = await supabase.from('report_snapshots').upsert({
              id: 'spo',
              filename: file.name,
              uploaded_at: new Date().toISOString(),
              data: JSON.stringify(strippedRows)
            })
            if (error) {
              console.error('[report-upload] SPO upsert failed:', error)
              alert('SPO upload failed to save — check console for details')
              setUploading(null)
              return
            }
            setSpoRows(strippedRows)
            setSpoInfo({ filename: file.name, uploaded_at: new Date().toISOString(), row_count: strippedRows.length })
          } else {
            console.log('[report-upload] chunking id="cr"', 'filename=', file.name, 'rows=', strippedRows.length)
            const { error } = await saveChunkedReport('cr', file.name, strippedRows)
            if (error) {
              console.error('[report-upload] CR chunked save failed:', error)
              alert('CR upload failed to save — check console for details')
              setUploading(null)
              return
            }
            setCrRows(strippedRows)
            setCrInfo({ filename: file.name, uploaded_at: new Date().toISOString(), row_count: strippedRows.length })
          }
        } else {
          // Decom isn't column-stripped — the parser resolves columns by
          // header name at load time, so the header row travels with the
          // stored data instead of being dropped like SPO/CR.
          const decomStoreRows = [allRows[headerRowIndex], ...dataRows]
          console.log('[report-upload] chunking id="decom"', 'filename=', file.name, 'rows=', decomStoreRows.length)
          const { error } = await saveChunkedReport('decom', file.name, decomStoreRows)
          if (error) {
            console.error('[report-upload] Decom chunked save failed:', error)
            alert('Decom upload failed to save — check console for details')
            setUploading(null)
            return
          }
          setDecomRawRows(decomStoreRows)
          setDecomInfo({ filename: file.name, uploaded_at: new Date().toISOString(), row_count: dataRows.length })
        }
      } catch (err) {
        console.error('Upload error:', err)
        alert('Upload failed — check file format')
      }
      setUploading(null)
    }
    reader.readAsArrayBuffer(file)
  }, [])

  return (
    <div className="min-h-screen bg-gray-950 text-white p-6">
      <div className="max-w-6xl mx-auto">

        <BackToDashboard />

        <div className="mb-6">
          <h1 className="text-3xl font-bold">Reports</h1>
          <p className="text-gray-400 mt-1">Upload SPO and CR master reports — download GC-specific filtered Excel files on demand</p>
        </div>

        {/* Tab Bar */}
        <div className="flex gap-2 mb-6">
          <button
            className="px-5 py-2 rounded-lg text-sm font-semibold bg-blue-600 text-white shadow-lg">
            SPO / CR Reports
          </button>
          <button
            onClick={() => router.push('/gr-tracker')}
            className="px-5 py-2 rounded-lg text-sm font-semibold bg-gray-800 text-gray-300 hover:bg-gray-700">
            💰 GR Tracker
          </button>
        </div>

        {/* Upload Section */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
          <div>
            <p className="text-sm font-semibold text-gray-400 mb-2">SPO MASTER REPORT</p>
            <UploadBox type="spo" info={spoInfo} label="SPO Master Report" uploading={uploading} onUpload={handleUpload} />
          </div>
          <div>
            <p className="text-sm font-semibold text-gray-400 mb-2">CR CLAIMS TRACKER</p>
            <UploadBox type="cr" info={crInfo} label="CR Claims Tracker" uploading={uploading} onUpload={handleUpload} />
          </div>
        </div>
        <div className="mb-8">
          <p className="text-sm font-semibold text-gray-400 mb-2">DECOM TRACKER</p>
          <UploadBox type="decom" info={decomInfo} label="Decom Tracker" uploading={uploading} onUpload={handleUpload} />
        </div>

        {/* Decom Summary */}
        {decomRows.length > 0 && (
          <div className="mb-8">
            <h2 className="text-lg font-bold mb-4">Decom Summary by GC</h2>
            <div className="overflow-x-auto bg-gray-900 rounded-xl border border-gray-700">
              <table className="w-full text-xs">
                <thead>
                  <tr className="bg-gray-800 text-gray-400">
                    <th className="text-left p-2">GC</th>
                    <th className="text-left p-2">Total Sites</th>
                    <th className="text-left p-2">Pending Decom Drop Off</th>
                    <th className="text-left p-2">Pending POD in Pathwave</th>
                    <th className="text-left p-2">POD in QuickBase</th>
                    <th className="text-left p-2">Complete</th>
                    <th className="text-left p-2">Avg Aging (days)</th>
                  </tr>
                </thead>
                <tbody>
                  {(() => {
                    const summary = summarizeDecomByGc(decomRows, GC_CONFIG.map(c => c.gc)).sort((a, b) => (b.outstanding + b.pending) - (a.outstanding + a.pending))
                    const totals = summary.reduce((t, s) => ({
                      total: t.total + s.total,
                      complete: t.complete + s.complete,
                      podGap: t.podGap + s.podGap,
                      outstanding: t.outstanding + s.outstanding,
                      pending: t.pending + s.pending,
                      podQuickBaseCount: t.podQuickBaseCount + s.podQuickBaseCount,
                    }), { total: 0, complete: 0, podGap: 0, outstanding: 0, pending: 0, podQuickBaseCount: 0 })
                    const allAging = decomRows.filter(r => r.status === 'outstanding' || r.status === 'pending').map(r => r.aging).filter((a): a is number => a !== null)
                    const totalAvgAging = allAging.length > 0 ? Math.round(allAging.reduce((s, a) => s + a, 0) / allAging.length) : null
                    return (
                      <>
                        {summary.map(s => (
                          <tr key={s.gc} className="border-t border-gray-800">
                            <td className="p-2 font-semibold text-white whitespace-nowrap">{s.gc}</td>
                            <td className="p-2 text-gray-300">{s.total}</td>
                            <td className="p-2 text-red-400 font-bold">{s.outstanding + s.pending}</td>
                            <td className="p-2 text-orange-400">{s.podGap}</td>
                            <td className="p-2 text-blue-400">{s.podQuickBaseCount}</td>
                            <td className="p-2 text-green-400">{s.complete}</td>
                            <td className="p-2 text-gray-300">{s.avgAging ?? '—'}</td>
                          </tr>
                        ))}
                        <tr className="border-t border-gray-700 bg-gray-800 font-bold">
                          <td className="p-2 text-white">Total</td>
                          <td className="p-2 text-gray-200">{totals.total}</td>
                          <td className="p-2 text-red-400">{totals.outstanding + totals.pending}</td>
                          <td className="p-2 text-orange-400">{totals.podGap}</td>
                          <td className="p-2 text-blue-400">{totals.podQuickBaseCount}</td>
                          <td className="p-2 text-green-400">{totals.complete}</td>
                          <td className="p-2 text-gray-200">{totalAvgAging ?? '—'}</td>
                        </tr>
                      </>
                    )
                  })()}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* GC Download Grid */}
        {(spoRows.length > 0 || crRows.length > 0 || decomRows.length > 0) && (
          <div>
            <h2 className="text-lg font-bold mb-4">GC-Specific Reports — Download on Demand</h2>
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
              {GC_CONFIG.map(cfg => {
                const spoCount = spoRows.filter(row => matches(row[SPO_VENDOR_COL_IN_MASTER], cfg.spo_match)).length
                const crCount = crRows.filter(row => matches(row[CR_SUPPLIER_COL_IN_MASTER], cfg.cr_match)).length
                const gcDecomRows = decomRowsForGc(decomRows, cfg.gc)

                return (
                  <div key={cfg.gc} className="bg-gray-900 rounded-xl border border-gray-700 p-4">
                    <h3 className="font-bold text-white text-base mb-1">{cfg.gc}</h3>
                    <div className="flex gap-2 text-xs text-gray-500 mb-3">
                      <span>{spoCount} SPO rows</span>
                      <span>·</span>
                      <span>{crCount} CR rows</span>
                      <span>·</span>
                      <span>{gcDecomRows.length} Decom rows</span>
                    </div>
                    <div className="flex flex-col gap-2">
                      <button
                        onClick={() => downloadGCReport(spoRows, SPO_COL_IDX, SPO_HEADERS, SPO_VENDOR_COL_IN_MASTER, cfg.spo_match, `${cfg.spo_label}_-_SPO_Report_-_${today}.xlsx`)}
                        disabled={spoRows.length === 0}
                        className="bg-blue-700 hover:bg-blue-600 disabled:opacity-40 disabled:cursor-not-allowed text-white text-xs px-3 py-2 rounded font-semibold">
                        📥 Download SPO Report
                      </button>
                      <button
                        onClick={() => downloadGCReport(crRows, CR_COL_IDX, CR_HEADERS, CR_SUPPLIER_COL_IN_MASTER, cfg.cr_match, `${cfg.cr_label}_-_CR_Tracker_-_${today}.xlsx`)}
                        disabled={crRows.length === 0}
                        className="bg-teal-700 hover:bg-teal-600 disabled:opacity-40 disabled:cursor-not-allowed text-white text-xs px-3 py-2 rounded font-semibold">
                        📥 Download CR Tracker
                      </button>
                      <button
                        onClick={() => downloadGcDecomReport(gcDecomRows, `${cfg.gc}_-_Decom_Report_-_${today}.xlsx`)}
                        disabled={gcDecomRows.length === 0}
                        className="bg-purple-700 hover:bg-purple-600 disabled:opacity-40 disabled:cursor-not-allowed text-white text-xs px-3 py-2 rounded font-semibold">
                        📥 Download Decom Report
                      </button>
                    </div>
                  </div>
                )
              })}
            </div>
          </div>
        )}

        {spoRows.length === 0 && crRows.length === 0 && decomRows.length === 0 && (
          <div className="bg-gray-900 rounded-xl border border-gray-700 p-12 text-center">
            <p className="text-gray-400 text-xl">📂 Upload your master reports above to get started</p>
          </div>
        )}

      </div>
    </div>
  )
}
