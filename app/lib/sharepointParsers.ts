// Parses each Nokia file's raw SharePoint bytes into EXACTLY the shape its
// now-removed Supabase-backed equivalent (loadTrackerSnapshot()/the old
// chunked report_snapshots reads) used to produce, so every consumer
// (MapView, the Tracker grid, grTracker.ts, decom.ts, scop.ts...) kept
// working against this data unchanged through the migration.
//
// Verified against the real upload handlers this mirrors:
//   tracker -> app/page.tsx handleTrackerUpload
//   decom/cr/spo -> app/reports/page.tsx handleUpload
//   scop -> app/reports/ScopTab.tsx handleUpload
// Column-position arrays (SPO_COL_IDX/CR_COL_IDX) are duplicated from
// app/reports/page.tsx rather than imported — that page wasn't touched for
// this migration, so these have to be kept in sync by hand if that page's
// column list ever changes.
//
// Perf fix, 11/2026: each parser used to call XLSX.read() with no `sheets`
// filter, which decodes EVERY sheet in the workbook (styles/formulas/HTML/
// number-formats included) just to then pick one. Now: a cheap bookSheets-
// only pass resolves the real sheet name, then a second pass decodes ONLY
// that sheet with styles/formulas/HTML/NF parsing turned off, after
// trimming its declared range to the real last data row. Large buffers
// (CR, ~30MB) decode in a Web Worker instead, off the main thread; every
// other file here is well under the worker threshold and decodes inline.
import * as XLSX from 'xlsx'
import { detectScopHeaderRow } from './scop'
import { trimSheetRowRange } from './xlsxTrim'
import type { DecodeSheetRequest, DecodeSheetResponse } from './xlsxSheetWorker'
import { timeoutMsFor, type SharePointFileKey } from './sharepoint'

function trimTrailingBlankRows(rows: unknown[][]): unknown[][] {
  let last = rows.length - 1
  while (last >= 0 && !(rows[last] || []).some(c => c !== null && c !== undefined && String(c).trim() !== '')) {
    last--
  }
  return rows.slice(0, last + 1)
}

// Builds a sparse array of length max(keepCols)+1, keeping only the given
// column indices and nulling everything else, since downstream SPO/CR
// consumers (grTracker.ts, app/reports/page.tsx) index into these rows by
// the SAME fixed original positions, with no header row to resolve by name.
function stripRow(row: unknown[], keepCols: number[]): unknown[] {
  const maxCol = Math.max(...keepCols)
  const out: unknown[] = new Array(maxCol + 1).fill(null)
  keepCols.forEach(i => { out[i] = row[i] ?? null })
  return out
}

const SPO_COL_IDX = [7, 8, 33, 40, 41, 43, 47, 48, 49, 50, 51]
const CR_COL_IDX = [0, 1, 2, 4, 5, 6, 7, 8, 9, 10, 11, 16, 18, 19, 25, 26, 27, 28, 29, 30, 31]

// Above this, decode in a Web Worker instead of the main thread. Tracker/
// decom/scop/spo are all well under this in practice; only CR (~30MB)
// crosses it.
const WORKER_THRESHOLD_BYTES = 4 * 1024 * 1024

const FAST_READ_OPTS = {
  type: 'array' as const,
  cellDates: true,
  cellStyles: false,
  cellFormula: false,
  cellHTML: false,
  cellNF: false,
}

/** Cheap metadata-only pass — gets real sheet names without decoding any
 * cell data, so each parser's sheet-selection fallback (exact name /
 * trim-tolerant / case-insensitive / first-sheet) runs against them before
 * paying for a full decode. */
function listSheetNames(buffer: ArrayBuffer): string[] {
  const wb = XLSX.read(buffer, { type: 'array', bookSheets: true })
  return wb.SheetNames
}

let worker: Worker | null = null
let workerBroken = false
let reqId = 0
const pending = new Map<number, { resolve: (rows: unknown[][]) => void; reject: (err: Error) => void }>()

function getWorker(): Worker | null {
  if (workerBroken || typeof Worker === 'undefined') return null
  if (!worker) {
    try {
      worker = new Worker(new URL('./xlsxSheetWorker.ts', import.meta.url))
      worker.onmessage = (e: MessageEvent<DecodeSheetResponse>) => {
        const p = pending.get(e.data.id)
        if (!p) return
        pending.delete(e.data.id)
        if (e.data.error) p.reject(new Error(e.data.error))
        else p.resolve(e.data.rows ?? [])
      }
      worker.onerror = () => {
        // Something's wrong with the worker itself (not a per-message
        // decode error, which comes back as a normal response) — stop
        // trying to use it for the rest of this session; every caller
        // falls back to main-thread decoding regardless.
        workerBroken = true
        worker = null
      }
    } catch {
      workerBroken = true
      return null
    }
  }
  return worker
}

// Bug fix, 11/2026 ("SharePoint refresh gets stuck"): this had no timeout at
// all — if the worker never posted back (message silently dropped by a
// locked-down browser's worker sandboxing, or the worker thread wedged),
// the Promise here just never settled, and everything awaiting it (up
// through SharePointDataProvider's in-flight dedupe) hung forever right
// along with it. `timeoutMs` matches the same per-file budget the network
// fetch itself uses (timeoutMsFor) — CR gets 180s since it's the only file
// that actually crosses the worker threshold. On timeout, the pending
// entry is dropped (a late response is simply ignored) AND the worker
// itself is torn down and rebuilt fresh next time — a job that didn't
// finish in budget means the worker may be wedged, not just slow, so
// reusing it for the next request isn't safe.
function decodeSheetInWorker(buffer: ArrayBuffer, sheetName: string, timeoutMs: number): Promise<unknown[][]> {
  return new Promise((resolve, reject) => {
    const w = getWorker()
    if (!w) { reject(new Error('Worker unavailable')); return }
    const id = ++reqId
    const timer = setTimeout(() => {
      pending.delete(id)
      reject(new Error(`Worker did not respond in ${Math.round(timeoutMs / 1000)}s decoding "${sheetName}".`))
      try { worker?.terminate() } catch { /* best effort */ }
      worker = null
    }, timeoutMs)
    pending.set(id, {
      resolve: rows => { clearTimeout(timer); resolve(rows) },
      reject: err => { clearTimeout(timer); reject(err) },
    })
    // Transfer a COPY, never the caller's own buffer — if the worker call
    // fails, decodeTargetSheet falls back to re-reading the ORIGINAL buffer
    // synchronously, which would silently see 0 bytes if that one had been
    // the transferred (now-detached) object instead.
    const transferCopy = buffer.slice(0)
    const request: DecodeSheetRequest = { id, buffer: transferCopy, sheetName }
    w.postMessage(request, [transferCopy])
  })
}

// Diagnostics side-channel — Settings' Diagnostics panel wants "was the
// worker used" and "how long did decoding take" per file, but changing
// every parseXWorkbook's return type to thread that through would ripple
// into the provider's PARSERS signature and every caller. A plain module-
// level record, written right before decodeTargetSheet returns and read
// by the provider right after calling the parser, is lower-risk and keeps
// this strictly informational — nothing reads it for correctness.
export interface LastDecodeMeta {
  usedWorker: boolean
  decodeMs: number
}
const lastDecodeMeta: Partial<Record<string, LastDecodeMeta>> = {}
export function getLastDecodeMeta(label: string): LastDecodeMeta | undefined {
  return lastDecodeMeta[label]
}

/** Decodes exactly one sheet — via the `sheets` filter, so the workbook's
 * other sheets are never parsed at all — with styles/formulas/HTML/number-
 * format parsing off, after trimming its declared range. `label` is also
 * the key getLastDecodeMeta() reports against. */
async function decodeTargetSheet(buffer: ArrayBuffer, sheetName: string, label: SharePointFileKey): Promise<unknown[][]> {
  const start = Date.now()
  let usedWorker = false
  console.time(`[sp-parse] ${label}: decode+toRows (${(buffer.byteLength / 1024).toFixed(0)}KB)`)
  try {
    if (buffer.byteLength > WORKER_THRESHOLD_BYTES) {
      try {
        const rows = await decodeSheetInWorker(buffer, sheetName, timeoutMsFor(label))
        usedWorker = true
        return rows
      } catch (err) {
        console.warn(`[sp-parse] ${label}: worker decode failed, falling back to main thread`, err)
      }
    }
    const wb = XLSX.read(buffer, { ...FAST_READ_OPTS, sheets: [sheetName] })
    const ws = wb.Sheets[sheetName]
    if (!ws) throw new Error(`Sheet "${sheetName}" not found after targeted read.`)
    trimSheetRowRange(ws)
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as unknown[][]
  } finally {
    lastDecodeMeta[label] = { usedWorker, decodeMs: Date.now() - start }
    console.timeEnd(`[sp-parse] ${label}: decode+toRows (${(buffer.byteLength / 1024).toFixed(0)}KB)`)
  }
}

/** Tracker — 'HOPs' sheet, header row is wherever the 'DON 444' cell is
 * (scanned, not assumed at row 0), trailing blank rows trimmed. Matches
 * handleTrackerUpload in app/page.tsx exactly. */
export async function parseTrackerWorkbook(buffer: ArrayBuffer): Promise<unknown[][]> {
  console.time('[sp-parse] tracker: total')
  const names = listSheetNames(buffer)
  if (!names.includes('HOPs')) throw new Error('"HOPs" tab not found in the tracker workbook.')
  const rawRows = await decodeTargetSheet(buffer, 'HOPs', 'tracker')
  console.time('[sp-parse] tracker: reshape')
  const trimmed = trimTrailingBlankRows(rawRows)
  let headerRowIdx = trimmed.findIndex(r => (r || []).some(c => String(c).trim() === 'DON 444'))
  if (headerRowIdx === -1) headerRowIdx = 0
  const result = trimmed.slice(headerRowIdx)
  console.timeEnd('[sp-parse] tracker: reshape')
  console.timeEnd('[sp-parse] tracker: total')
  return result
}

/** Decom — 'R&R - Link Drop Off Status' sheet (name matched trim-tolerant,
 * the real tab carries a trailing space), header fixed at row 0, trailing
 * blanks trimmed, header kept at stored index 0 (parseDecomRows trusts
 * this unconditionally). Matches handleUpload's decom branch exactly. */
export async function parseDecomWorkbook(buffer: ArrayBuffer): Promise<unknown[][]> {
  console.time('[sp-parse] decom: total')
  const names = listSheetNames(buffer)
  const matchedName = names.find(n => n.trim() === 'R&R - Link Drop Off Status')
  if (!matchedName) throw new Error('"R&R - Link Drop Off Status" tab not found in the decom workbook.')
  const rawRows = await decodeTargetSheet(buffer, matchedName, 'decom')
  console.time('[sp-parse] decom: reshape')
  const trimmed = trimTrailingBlankRows(rawRows)
  const dataRows = trimmed.slice(1).filter(row => row.some(v => v !== null))
  const result = [trimmed[0], ...dataRows]
  console.timeEnd('[sp-parse] decom: reshape')
  console.timeEnd('[sp-parse] decom: total')
  return result
}

/** SCOP — 'HOPs' sheet (case-insensitive fallback), header auto-detected
 * via the same detectScopHeaderRow() the real handler calls, header kept
 * at stored index 0 (keyScopRows trusts this unconditionally). Matches
 * ScopTab.tsx's handleUpload exactly (header-row Settings override isn't
 * replicated here — this always auto-detects). */
export async function parseScopWorkbook(buffer: ArrayBuffer): Promise<unknown[][]> {
  console.time('[sp-parse] scop: total')
  const names = listSheetNames(buffer)
  const sheetName = names.includes('HOPs') ? 'HOPs' : names.find(n => n.trim().toLowerCase() === 'hops')
  if (!sheetName) throw new Error('"HOPs" tab not found in the SCOP workbook.')
  const rawRows = await decodeTargetSheet(buffer, sheetName, 'scop')
  console.time('[sp-parse] scop: reshape')
  const headerIdx = detectScopHeaderRow(rawRows)
  if (headerIdx < 0) throw new Error('Could not auto-detect the SCOP header row (looking for "HOP" and "General Contractor").')
  const dataRows = rawRows.slice(headerIdx + 1).filter(r => (r || []).some(v => v !== null && v !== undefined && String(v).trim() !== ''))
  const result = [rawRows[headerIdx], ...dataRows]
  console.timeEnd('[sp-parse] scop: reshape')
  console.timeEnd('[sp-parse] scop: total')
  return result
}

/** CR — 'CR Tracker' sheet (fallback first sheet), header fixed at row 1,
 * trailing blanks trimmed, rows column-stripped to CR_COL_IDX with the
 * header row dropped entirely (no name-based resolution downstream —
 * grTracker.ts-style consumers would index by fixed position). Matches
 * handleUpload's cr branch exactly. */
export async function parseCrWorkbook(buffer: ArrayBuffer): Promise<unknown[][]> {
  console.time('[sp-parse] cr: total')
  const names = listSheetNames(buffer)
  const sheetName = names.includes('CR Tracker') ? 'CR Tracker' : names[0]
  const rawRows = await decodeTargetSheet(buffer, sheetName, 'cr')
  console.time('[sp-parse] cr: reshape')
  const trimmed = trimTrailingBlankRows(rawRows)
  const dataRows = trimmed.slice(2).filter(row => row.some(v => v !== null)) // headerRowIndex=1
  const result = dataRows.map(row => stripRow(row, CR_COL_IDX))
  console.timeEnd('[sp-parse] cr: reshape')
  console.timeEnd('[sp-parse] cr: total')
  return result
}

/** SPO — 'NDPd_NAM_003_Viaero_SPO_Report' sheet (fallback first sheet),
 * header fixed at row 0, NO trailing-blank-trim (the real handler doesn't
 * apply one for spo), rows column-stripped to SPO_COL_IDX with the header
 * row dropped. Matches handleUpload's spo branch exactly. */
export async function parseSpoWorkbook(buffer: ArrayBuffer): Promise<unknown[][]> {
  console.time('[sp-parse] spo: total')
  const names = listSheetNames(buffer)
  const sheetName = names.includes('NDPd_NAM_003_Viaero_SPO_Report') ? 'NDPd_NAM_003_Viaero_SPO_Report' : names[0]
  const rawRows = await decodeTargetSheet(buffer, sheetName, 'spo')
  console.time('[sp-parse] spo: reshape')
  const dataRows = rawRows.slice(1).filter(row => row.some(v => v !== null)) // headerRowIndex=0
  const result = dataRows.map(row => stripRow(row, SPO_COL_IDX))
  console.timeEnd('[sp-parse] spo: reshape')
  console.timeEnd('[sp-parse] spo: total')
  return result
}

// ─────────────────────────────────────────────
// DIAGNOSTICS — "newest meaningful date" per file, for Settings' Diagnostics
// panel only (not used anywhere for business logic). tracker/decom/scop
// still carry a header row at index 0, so their column is resolved by
// name; spo/cr are already column-stripped to fixed original positions
// with no header row, so theirs is resolved by that fixed index instead.
// ─────────────────────────────────────────────

function parseAnyDate(val: unknown): Date | null {
  if (val === null || val === undefined || val === '') return null
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val
  if (typeof val === 'number') {
    if (val > 25000 && val < 60000) {
      const d = new Date((val - 25569) * 86400 * 1000)
      return isNaN(d.getTime()) ? null : d
    }
    return null
  }
  const d = new Date(String(val))
  return isNaN(d.getTime()) || d.getFullYear() < 1990 ? null : d
}

function newestDateByHeader(rows: unknown[][], headerName: string): Date | null {
  if (rows.length === 0) return null
  const headers = (rows[0] as unknown[]).map(h => String(h ?? '').trim().toLowerCase())
  const col = headers.indexOf(headerName.toLowerCase())
  if (col === -1) return null
  let newest: Date | null = null
  for (let i = 1; i < rows.length; i++) {
    const d = parseAnyDate((rows[i] as unknown[])[col])
    if (d && (!newest || d > newest)) newest = d
  }
  return newest
}

function newestDateByIndex(rows: unknown[][], colIdx: number): Date | null {
  let newest: Date | null = null
  for (const row of rows) {
    const d = parseAnyDate((row as unknown[])[colIdx])
    if (d && (!newest || d > newest)) newest = d
  }
  return newest
}

export interface NewestDateResult {
  columnLabel: string
  date: Date | null
}

/** `rows` is each file's already-parsed output (same shape parseXWorkbook
 * returns) — never re-reads the source buffer. */
export function computeNewestMeaningfulDate(key: 'tracker' | 'decom' | 'scop' | 'spo' | 'cr', rows: unknown[][]): NewestDateResult {
  switch (key) {
    case 'tracker':
      return { columnLabel: 'MS16 Implementation Ends A', date: newestDateByHeader(rows, 'MS16 Implementation Ends A') }
    case 'decom':
      return { columnLabel: 'DECOM Drop Off', date: newestDateByHeader(rows, 'DECOM Drop Off') }
    case 'scop':
      return { columnLabel: 'Construction Complete Actual', date: newestDateByHeader(rows, 'Construction Complete Actual') }
    case 'spo':
      // SPO_COL_IDX[4] = 41 = 'SPO Creation Date' (SPO_HEADERS[4]) — the
      // stripped rows have no header row, so this is a fixed position.
      return { columnLabel: 'SPO Creation Date', date: newestDateByIndex(rows, 41) }
    case 'cr':
      // CR_COL_IDX[18] = 29 = 'SPO Issued Date' (CR_HEADERS[18]) — the
      // closest thing CR has to a claim/submission/updated date, already
      // used elsewhere in the CR report.
      return { columnLabel: 'SPO Issued Date', date: newestDateByIndex(rows, 29) }
  }
}
