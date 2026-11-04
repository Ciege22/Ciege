// Shared by app/deck-builder-v2/page.tsx and NtpCommentsV2.tsx — both send
// the in-memory, already-parsed tracker rows to a Railway endpoint as a
// multipart .xlsx file, since that's the upload shape the backend already
// parses (and Stage 6 keeps it that way deliberately, rather than adding a
// second "raw JSON rows" input the backend would also have to trust).
import * as XLSX from 'xlsx'

// backend/build_deck.py's extract_data(), when given a FILE (not pre-parsed
// rows), hardcodes pd.read_excel(tracker_path, sheet_name='HOPs', header=1)
// — it assumes row 0 is always a description row and the real header is
// row 1, same as every real Nokia tracker export. SharePointDataProvider's
// parsed tracker.rows has already found and stripped that description row
// (row 0 IS the real header there), so building a file from it as-is would
// make the backend treat our actual header as the description row and our
// first data row as headers. A single blank row has to go back in at the
// top to keep the file-upload path's assumption true. (Confirmed via a
// throwaway round-trip test: an empty array `[]` for that row gets dropped
// entirely by aoa_to_sheet, collapsing row 0 away — it has to be a
// one-cell row, e.g. [''].)
export function buildTrackerXlsxBlob(rows: unknown[][]): Blob {
  const wb = XLSX.utils.book_new()
  const ws = XLSX.utils.aoa_to_sheet([[''], ...rows], { cellDates: true })
  XLSX.utils.book_append_sheet(wb, ws, 'HOPs')
  const buffer = XLSX.write(wb, { type: 'array', bookType: 'xlsx', cellDates: true }) as ArrayBuffer
  return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
}
