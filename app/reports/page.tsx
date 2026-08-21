'use client'

export const dynamic = 'force-dynamic'

import { useState, useCallback, useEffect } from 'react'
import { useRouter } from 'next/navigation'
import * as XLSX from 'xlsx'
import { supabase } from '../lib/supabase'
import { GC_CONFIG, matches, SPO_VENDOR_COL_IN_MASTER, CR_SUPPLIER_COL_IN_MASTER } from '../lib/gcConfig'
import BackToDashboard from '../components/BackToDashboard'

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

export default function ReportsPage() {
  const router = useRouter()
  const [spoRows, setSpoRows] = useState<unknown[][]>([])
  const [crRows, setCrRows] = useState<unknown[][]>([])
  const [spoInfo, setSpoInfo] = useState<ReportSnapshot | null>(null)
  const [crInfo, setCrInfo] = useState<ReportSnapshot | null>(null)
  const [uploading, setUploading] = useState<string | null>(null)
  const today = new Date().toLocaleDateString('en-US').replace(/\//g, '-')

  useEffect(() => {
    const load = async () => {
      const { data: spoSnap } = await supabase.from('report_snapshots').select('*').eq('id', 'spo').single()
      const { data: crSnap } = await supabase.from('report_snapshots').select('*').eq('id', 'cr').single()
      if (spoSnap) { setSpoRows(JSON.parse(spoSnap.data)); setSpoInfo({ filename: spoSnap.filename, uploaded_at: spoSnap.uploaded_at, row_count: JSON.parse(spoSnap.data).length }) }
      if (crSnap) { setCrRows(JSON.parse(crSnap.data)); setCrInfo({ filename: crSnap.filename, uploaded_at: crSnap.uploaded_at, row_count: JSON.parse(crSnap.data).length }) }
    }
    load()
  }, [])

  const handleUpload = useCallback(async (file: File, type: 'spo' | 'cr') => {
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
        } else {
          ws = wb.Sheets['CR Tracker'] || wb.Sheets[wb.SheetNames[0]]
          headerRowIndex = 1
        }

        const allRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as unknown[][]
        const dataRows = allRows.slice(headerRowIndex + 1).filter(row => row.some(v => v !== null))

        if (type === 'spo') {
          console.log('[report-upload] upserting id="spo"', 'filename=', file.name)
          const { error } = await supabase.from('report_snapshots').upsert({
            id: 'spo',
            filename: file.name,
            uploaded_at: new Date().toISOString(),
            data: JSON.stringify(dataRows)
          })
          if (error) {
            console.error('[report-upload] SPO upsert failed:', error)
            alert('SPO upload failed to save — check console for details')
            setUploading(null)
            return
          }
        } else {
          console.log('[report-upload] upserting id="cr"', 'filename=', file.name)
          const { error } = await supabase.from('report_snapshots').upsert({
            id: 'cr',
            filename: file.name,
            uploaded_at: new Date().toISOString(),
            data: JSON.stringify(dataRows)
          })
          if (error) {
            console.error('[report-upload] CR upsert failed:', error)
            alert('CR upload failed to save — check console for details')
            setUploading(null)
            return
          }
        }

        if (type === 'spo') {
          setSpoRows(dataRows)
          setSpoInfo({ filename: file.name, uploaded_at: new Date().toISOString(), row_count: dataRows.length })
        } else {
          setCrRows(dataRows)
          setCrInfo({ filename: file.name, uploaded_at: new Date().toISOString(), row_count: dataRows.length })
        }
      } catch (err) {
        console.error('Upload error:', err)
        alert('Upload failed — check file format')
      }
      setUploading(null)
    }
    reader.readAsArrayBuffer(file)
  }, [])

  const UploadBox = ({ type, info, label }: { type: 'spo' | 'cr', info: ReportSnapshot | null, label: string }) => (
    <div
      className="border-2 border-dashed border-gray-600 rounded-xl p-5 cursor-pointer hover:border-blue-500 transition-colors"
      onDragOver={(e) => e.preventDefault()}
      onDrop={(e) => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) handleUpload(f, type) }}
      onClick={() => document.getElementById(`${type}-upload`)?.click()}
    >
      <input id={`${type}-upload`} type="file" accept=".xlsx" className="hidden"
        onChange={(e) => { const f = e.target.files?.[0]; if (f) handleUpload(f, type) }} />
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
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-8">
          <div>
            <p className="text-sm font-semibold text-gray-400 mb-2">SPO MASTER REPORT</p>
            <UploadBox type="spo" info={spoInfo} label="SPO Master Report" />
          </div>
          <div>
            <p className="text-sm font-semibold text-gray-400 mb-2">CR CLAIMS TRACKER</p>
            <UploadBox type="cr" info={crInfo} label="CR Claims Tracker" />
          </div>
        </div>

        {/* GC Download Grid */}
        {(spoRows.length > 0 || crRows.length > 0) && (
          <div>
            <h2 className="text-lg font-bold mb-4">GC-Specific Reports — Download on Demand</h2>
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
              {GC_CONFIG.map(cfg => {
                const spoCount = spoRows.filter(row => matches(row[SPO_VENDOR_COL_IN_MASTER], cfg.spo_match)).length
                const crCount = crRows.filter(row => matches(row[CR_SUPPLIER_COL_IN_MASTER], cfg.cr_match)).length

                return (
                  <div key={cfg.gc} className="bg-gray-900 rounded-xl border border-gray-700 p-4">
                    <h3 className="font-bold text-white text-base mb-1">{cfg.gc}</h3>
                    <div className="flex gap-2 text-xs text-gray-500 mb-3">
                      <span>{spoCount} SPO rows</span>
                      <span>·</span>
                      <span>{crCount} CR rows</span>
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
                    </div>
                  </div>
                )
              })}
            </div>
          </div>
        )}

        {spoRows.length === 0 && crRows.length === 0 && (
          <div className="bg-gray-900 rounded-xl border border-gray-700 p-12 text-center">
            <p className="text-gray-400 text-xl">📂 Upload your master reports above to get started</p>
          </div>
        )}

      </div>
    </div>
  )
}
