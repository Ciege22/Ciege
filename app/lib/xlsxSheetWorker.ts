// Dedicated Web Worker that decodes ONE XLSX sheet off the main thread —
// only dispatched for large buffers (see WORKER_THRESHOLD_BYTES in
// sharepointParsers.ts). CR (~30MB) is the one Nokia file big enough for
// this to matter; tracker/decom/scop/spo are all well under the threshold
// and never engage it. Plain worker script, not a React/'use client' file.
import * as XLSX from 'xlsx'
import { trimSheetRowRange } from './xlsxTrim'

export interface DecodeSheetRequest {
  id: number
  buffer: ArrayBuffer
  sheetName: string
}
export interface DecodeSheetResponse {
  id: number
  rows?: unknown[][]
  error?: string
}

self.onmessage = (e: MessageEvent<DecodeSheetRequest>) => {
  const { id, buffer, sheetName } = e.data
  try {
    const wb = XLSX.read(buffer, {
      type: 'array',
      sheets: [sheetName],
      cellDates: true,
      cellStyles: false,
      cellFormula: false,
      cellHTML: false,
      cellNF: false,
    })
    const ws = wb.Sheets[sheetName]
    if (!ws) throw new Error(`Sheet "${sheetName}" not found.`)
    trimSheetRowRange(ws)
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as unknown[][]
    const response: DecodeSheetResponse = { id, rows }
    postMessage(response)
  } catch (err) {
    const response: DecodeSheetResponse = { id, error: err instanceof Error ? err.message : 'Worker decode failed.' }
    postMessage(response)
  }
}
