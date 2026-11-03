// Parses each Nokia file's raw SharePoint bytes into EXACTLY the shape its
// Supabase-backed equivalent currently produces, so SharePointDataProvider
// is a true drop-in for loadTrackerSnapshot()/loadChunkedReport() — every
// existing consumer (MapView, the Tracker grid, grTracker.ts, decom.ts,
// scop.ts...) keeps working against this data completely unchanged.
//
// Verified against the real upload handlers this mirrors:
//   tracker -> app/page.tsx handleTrackerUpload
//   decom/cr/spo -> app/reports/page.tsx handleUpload
//   scop -> app/reports/ScopTab.tsx handleUpload
// Column-position arrays (SPO_COL_IDX/CR_COL_IDX) are duplicated from
// app/reports/page.tsx rather than imported — that page wasn't touched for
// this migration, so these have to be kept in sync by hand if that page's
// column list ever changes.
import * as XLSX from 'xlsx'
import { detectScopHeaderRow } from './scop'

function trimTrailingBlankRows(rows: unknown[][]): unknown[][] {
  let last = rows.length - 1
  while (last >= 0 && !(rows[last] || []).some(c => c !== null && c !== undefined && String(c).trim() !== '')) {
    last--
  }
  return rows.slice(0, last + 1)
}

// Builds a sparse array of length max(keepCols)+1, keeping only the given
// column indices and nulling everything else — matches app/reports/page.tsx's
// stripRow() exactly, since downstream SPO/CR consumers (grTracker.ts) index
// into these rows by the SAME fixed original positions, with no header row
// to resolve by name.
function stripRow(row: unknown[], keepCols: number[]): unknown[] {
  const maxCol = Math.max(...keepCols)
  const out: unknown[] = new Array(maxCol + 1).fill(null)
  keepCols.forEach(i => { out[i] = row[i] ?? null })
  return out
}

const SPO_COL_IDX = [7, 8, 33, 40, 41, 43, 47, 48, 49, 50, 51]
const CR_COL_IDX = [0, 1, 2, 4, 5, 6, 7, 8, 9, 10, 11, 16, 18, 19, 25, 26, 27, 28, 29, 30, 31]

function readSheet(buffer: ArrayBuffer): XLSX.WorkBook {
  return XLSX.read(buffer, { type: 'array', cellDates: true })
}

/** Tracker — 'HOPs' sheet, header row is wherever the 'DON 444' cell is
 * (scanned, not assumed at row 0), trailing blank rows trimmed. Matches
 * handleTrackerUpload in app/page.tsx exactly. */
export function parseTrackerWorkbook(buffer: ArrayBuffer): unknown[][] {
  const wb = readSheet(buffer)
  const ws = wb.Sheets['HOPs']
  if (!ws) throw new Error('"HOPs" tab not found in the tracker workbook.')
  const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1 }) as unknown[][]
  const trimmed = trimTrailingBlankRows(rawRows)
  let headerRowIdx = trimmed.findIndex(r => (r || []).some(c => String(c).trim() === 'DON 444'))
  if (headerRowIdx === -1) headerRowIdx = 0
  return trimmed.slice(headerRowIdx)
}

/** Decom — 'R&R - Link Drop Off Status' sheet (name matched trim-tolerant,
 * the real tab carries a trailing space), header fixed at row 0, trailing
 * blanks trimmed, header kept at stored index 0 (parseDecomRows trusts
 * this unconditionally). Matches handleUpload's decom branch exactly. */
export function parseDecomWorkbook(buffer: ArrayBuffer): unknown[][] {
  const wb = readSheet(buffer)
  const matchedName = wb.SheetNames.find(n => n.trim() === 'R&R - Link Drop Off Status')
  const ws = matchedName ? wb.Sheets[matchedName] : undefined
  if (!ws) throw new Error('"R&R - Link Drop Off Status" tab not found in the decom workbook.')
  const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as unknown[][]
  const trimmed = trimTrailingBlankRows(rawRows)
  const dataRows = trimmed.slice(1).filter(row => row.some(v => v !== null))
  return [trimmed[0], ...dataRows]
}

/** SCOP — 'HOPs' sheet (case-insensitive fallback), header auto-detected
 * via the same detectScopHeaderRow() the real handler calls, header kept
 * at stored index 0 (keyScopRows trusts this unconditionally). Matches
 * ScopTab.tsx's handleUpload exactly (header-row Settings override isn't
 * replicated here — this always auto-detects). */
export function parseScopWorkbook(buffer: ArrayBuffer): unknown[][] {
  const wb = readSheet(buffer)
  const sheetName = wb.Sheets['HOPs'] ? 'HOPs' : wb.SheetNames.find(n => n.trim().toLowerCase() === 'hops')
  if (!sheetName) throw new Error('"HOPs" tab not found in the SCOP workbook.')
  const ws = wb.Sheets[sheetName]
  const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as unknown[][]
  const headerIdx = detectScopHeaderRow(rawRows)
  if (headerIdx < 0) throw new Error('Could not auto-detect the SCOP header row (looking for "HOP" and "General Contractor").')
  const dataRows = rawRows.slice(headerIdx + 1).filter(r => (r || []).some(v => v !== null && v !== undefined && String(v).trim() !== ''))
  return [rawRows[headerIdx], ...dataRows]
}

/** CR — 'CR Tracker' sheet (fallback first sheet), header fixed at row 1,
 * trailing blanks trimmed, rows column-stripped to CR_COL_IDX with the
 * header row dropped entirely (no name-based resolution downstream —
 * grTracker.ts-style consumers would index by fixed position). Matches
 * handleUpload's cr branch exactly. */
export function parseCrWorkbook(buffer: ArrayBuffer): unknown[][] {
  const wb = readSheet(buffer)
  const sheetName = wb.Sheets['CR Tracker'] ? 'CR Tracker' : wb.SheetNames[0]
  const ws = wb.Sheets[sheetName]
  const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as unknown[][]
  const trimmed = trimTrailingBlankRows(rawRows)
  const dataRows = trimmed.slice(2).filter(row => row.some(v => v !== null)) // headerRowIndex=1
  return dataRows.map(row => stripRow(row, CR_COL_IDX))
}

/** SPO — 'NDPd_NAM_003_Viaero_SPO_Report' sheet (fallback first sheet),
 * header fixed at row 0, NO trailing-blank-trim (the real handler doesn't
 * apply one for spo), rows column-stripped to SPO_COL_IDX with the header
 * row dropped. Matches handleUpload's spo branch exactly. */
export function parseSpoWorkbook(buffer: ArrayBuffer): unknown[][] {
  const wb = readSheet(buffer)
  const sheetName = wb.Sheets['NDPd_NAM_003_Viaero_SPO_Report'] ? 'NDPd_NAM_003_Viaero_SPO_Report' : wb.SheetNames[0]
  const ws = wb.Sheets[sheetName]
  const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as unknown[][]
  const dataRows = rawRows.slice(1).filter(row => row.some(v => v !== null)) // headerRowIndex=0
  return dataRows.map(row => stripRow(row, SPO_COL_IDX))
}
