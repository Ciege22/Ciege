// Shared by sharepointParsers.ts (main thread) and xlsxSheetWorker.ts (Web
// Worker) — kept as its own tiny, dependency-free module so the worker file
// doesn't have to import anything that could drag in browser/React-only code.
import * as XLSX from 'xlsx'

// Trims a sheet's declared used-range (!ref) down to the real last
// non-empty row, scanning only the cells SheetJS actually populated
// (proportional to real data) rather than the declared range — formatting
// applied to a whole column/row can inflate that declared range to near
// Excel's max (1,048,576 rows), confirmed on the tracker workbook. Doing
// this BEFORE sheet_to_json means it never walks or allocates for
// thousands of phantom trailing rows.
export function trimSheetRowRange(ws: XLSX.WorkSheet): void {
  const ref = ws['!ref']
  if (!ref) return
  const range = XLSX.utils.decode_range(ref)
  let maxRow = range.s.r
  for (const key of Object.keys(ws)) {
    if (key.startsWith('!')) continue
    const cell = (ws as unknown as Record<string, XLSX.CellObject>)[key]
    if (cell?.v === undefined || cell.v === null || String(cell.v).trim() === '') continue
    const addr = XLSX.utils.decode_cell(key)
    if (addr.r > maxRow) maxRow = addr.r
  }
  if (maxRow < range.e.r) {
    range.e.r = maxRow
    ws['!ref'] = XLSX.utils.encode_range(range)
  }
}
