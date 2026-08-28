'use client'

export const dynamic = 'force-dynamic'

import { useState, useEffect, useRef, useMemo, useCallback } from 'react'
import { supabase, loadTrackerSnapshot } from '../lib/supabase'
import { parseDateAny } from '../lib/grTracker'
import BackToDashboard from '../components/BackToDashboard'

const NAVY = '#124191'
const AMBER = '#FFF3CD'
const ALT_ROW = '#F7F8FA'
const SELECTED_ROW = '#DCEAFB'

// Virtualization tuning — fixed row height lets us compute the visible window
// from scrollTop with simple arithmetic instead of measuring the DOM.
const ROW_HEIGHT = 36
const ROW_BUFFER = 10
const COL_BUFFER = 3

// First five columns are pinned in place while the rest scroll horizontally.
// Matched against normalized header names (see normHeader).
const FROZEN_COL_NAMES = ['HOP', 'Site Name', 'Path ID', 'New CM', 'General Contractor']
// Subtle right-edge shadow that separates the frozen columns from the
// scrolling ones underneath them.
const FROZEN_SHADOW = '2px 0 4px -1px rgba(0,0,0,0.18)'

interface TrackerRowData {
  // Unique per physical site row — a HOP can span two site rows, so HOP alone
  // can't identify one (same convention as app/lib/decom.ts's DecomRow.rowKey:
  // Path ID first, falling back to Site Name + Site Number, then a row index).
  rowKey: string
  hop: string
  cells: unknown[]
}

interface TrackerChange {
  rowKey: string
  hop: string
  field: string
  oldValue: string
  newValue: string
  timestamp: string
  user: string
  completed?: boolean
}

interface TrackerView {
  name: string
  hiddenColumns: string[]
}

// Per-column ("Excel style") filter state. `sort` drives the whole grid's row
// order (only one column may hold a sort at a time). `selectedValues === null`
// means "all values pass" — i.e. no value filtering on this column.
type ColSort = 'asc' | 'desc' | null
interface ColumnFilterState {
  sort: ColSort
  selectedValues: string[] | null
}

interface GridColumn {
  index: number
  name: string
  isHop: boolean
  isDate: boolean
  width: number
}

// Three starter views, auto-created on first load if no saved views exist yet.
// "keep" lists are matched against normalized header names — anything not in
// the list gets hidden.
const STARTER_VIEW_DEFS: { name: string; keep: string[] }[] = [
  {
    name: 'Decom',
    keep: ['HOP', 'Path ID', 'Site Name', 'General Contractor', 'New CM', 'MS16 Implementation Ends A', 'NTP A', 'Material Received A'],
  },
  {
    name: 'Material Watch',
    keep: ['HOP', 'General Contractor', 'Material Forecast +4ish', 'Material Received A', 'NTP A', 'MS15 Implementation Start F'],
  },
  {
    name: 'Monthly Starts',
    keep: ['HOP', 'General Contractor', 'New CM', 'Nokia PM', 'POR Start', 'MS15 Implementation Start F', 'MS15 Implementation Start A', 'NTP A', 'Material Received A'],
  },
]

// Some headers carry a leading/trailing apostrophe (an Excel "force text"
// formatting artifact — see Path ID / CX Notes: elsewhere in the app) — strip
// it uniformly so column matching (views, dedup) never trips on it.
function normHeader(h: unknown): string {
  return String(h ?? '').trim().replace(/^'+|'+$/g, '')
}

const DATE_COL_REGEX = /start|end|complete|date|ntp|material|mss|power|forecast|actual/i
function isDateColumn(name: string): boolean {
  return DATE_COL_REGEX.test(name)
}

// Parses the full tracker snapshot with NO column subsetting — every column
// survives — and NO dedup-by-HOP either: a HOP can span two physical site
// rows with two different Site Names, and this grid shows both rather than
// silently picking one (unlike most other pages, which collapse to a single
// representative row per HOP). Deliberately does NOT filter by Nokia PM,
// since this grid is a program-wide utility, not scoped to one PM.
function parseAllTrackerRows(rows: unknown[][]): { headers: string[]; trackerRows: TrackerRowData[] } {
  let headerRowIdx = -1
  for (let i = 0; i < 10; i++) {
    const row = rows[i] as unknown[]
    if (row && row.some(c => normHeader(c) === 'HOP')) { headerRowIdx = i; break }
  }
  if (headerRowIdx === -1) return { headers: [], trackerRows: [] }

  const headers = (rows[headerRowIdx] as unknown[]).map(normHeader)
  const hopCol = headers.findIndex(h => h === 'HOP')
  const don444Col = headers.findIndex(h => h === 'DON 444')
  const pathIdCol = headers.findIndex(h => h === 'Path ID')
  const siteNameCol = headers.findIndex(h => h === 'Site Name')
  const siteNumberCol = headers.findIndex(h => h === 'Site Number')

  const trackerRows: TrackerRowData[] = []
  let rowCounter = 0
  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i] as unknown[]
    if (!row) continue
    const don = normHeader(row[don444Col]).toUpperCase()
    if (don !== 'DON 444') continue
    const hop = normHeader(row[hopCol])
    if (!hop || hop === 'undefined') continue

    const pathId = pathIdCol >= 0 ? normHeader(row[pathIdCol]) : ''
    const siteName = siteNameCol >= 0 ? normHeader(row[siteNameCol]) : ''
    const siteNumber = siteNumberCol >= 0 ? normHeader(row[siteNumberCol]) : ''
    const rowKey = pathId || (siteName || siteNumber ? `${siteName}|${siteNumber}` : `${hop}-row-${rowCounter}`)
    rowCounter++

    trackerRows.push({ rowKey, hop, cells: row })
  }
  trackerRows.sort((a, b) => a.hop.localeCompare(b.hop) || a.rowKey.localeCompare(b.rowKey))

  return { headers, trackerRows }
}

// A regex-matched "date" column can still hold free text on some rows (e.g.
// "NTP Action Owner" and "NTP is waiting on" both contain "ntp" but are
// people/text fields, not dates) — fall back to plain text display/editing
// per-cell rather than force an empty date picker over real data.
function cellDisplayValue(raw: unknown, isDateCol: boolean): { text: string; treatAsDate: boolean } {
  if (raw === null || raw === undefined || raw === '') return { text: '', treatAsDate: isDateCol }
  if (isDateCol) {
    const d = parseDateAny(raw)
    if (d) return { text: d.toLocaleDateString('en-US'), treatAsDate: true }
    return { text: String(raw).trim(), treatAsDate: false }
  }
  if (raw instanceof Date) return { text: raw.toLocaleDateString('en-US'), treatAsDate: false }
  return { text: String(raw).trim(), treatAsDate: false }
}

function toDateInputValue(display: string): string {
  if (!display) return ''
  const parts = display.split('/')
  if (parts.length !== 3) return ''
  return `${parts[2]}-${parts[0].padStart(2, '0')}-${parts[1].padStart(2, '0')}`
}
function fromDateInputValue(inputVal: string): string {
  if (!inputVal) return ''
  const parts = inputVal.split('-')
  if (parts.length !== 3) return ''
  return `${parseInt(parts[1], 10)}/${parseInt(parts[2], 10)}/${parts[0]}`
}

function formatChangeLine(c: TrackerChange): string {
  const dateStr = new Date(c.timestamp).toLocaleDateString('en-US', { month: '2-digit', day: '2-digit', year: 'numeric' })
  if (c.oldValue) {
    return `${c.hop} | ${dateStr}: (${c.user}) ${c.field} updated from '${c.oldValue}' to '${c.newValue}'`
  }
  return `${c.hop} | ${dateStr}: (${c.user}) ${c.field}: ${c.newValue}`
}

// EditableCell / DatePickerCell are defined at module scope (not inside the
// page component) so React never remounts them on re-render — a component
// defined inside another component's render body gets a new identity every
// render, which drops focus and local edit-buffer state on every keystroke.
interface CellProps {
  displayValue: string
  isChanged: boolean
  isEditing: boolean
  rowBg: string
  // When set, the cell is one of the frozen columns — pin it with the same
  // sticky offset used by its header so body and header stay aligned.
  stickyLeft?: number
  onStartEdit: () => void
  onCommit: (newValue: string) => void
  onCancel: () => void
}

function frozenTdStyle(stickyLeft: number | undefined, editing: boolean): React.CSSProperties {
  if (stickyLeft === undefined) return {}
  return {
    position: 'sticky',
    left: stickyLeft,
    zIndex: editing ? 40 : 10,
    boxShadow: FROZEN_SHADOW,
  }
}

function EditableCell({ displayValue, isChanged, isEditing, rowBg, stickyLeft, onStartEdit, onCommit, onCancel }: CellProps) {
  const [draft, setDraft] = useState(displayValue)
  // Reset the edit buffer exactly when entering edit mode — adjusted during
  // render (React's documented pattern for resetting state on a prop change)
  // rather than in an effect, which would cause an extra render pass.
  const [wasEditing, setWasEditing] = useState(isEditing)
  if (isEditing !== wasEditing) {
    setWasEditing(isEditing)
    if (isEditing) setDraft(displayValue)
  }

  if (!isEditing) {
    return (
      <td
        onDoubleClick={onStartEdit}
        className="px-2 py-1 text-xs whitespace-nowrap cursor-pointer border-b border-r border-gray-200 overflow-hidden text-ellipsis"
        style={{ backgroundColor: isChanged ? AMBER : rowBg, height: ROW_HEIGHT, ...frozenTdStyle(stickyLeft, false) }}
        title={displayValue}
      >
        {displayValue || <span className="text-gray-300">—</span>}
      </td>
    )
  }

  return (
    <td className="px-1 py-0.5 border-b border-r border-gray-200 bg-white relative" style={{ zIndex: 40, ...frozenTdStyle(stickyLeft, true), backgroundColor: '#fff' }}>
      <input
        autoFocus
        type="text"
        value={draft}
        onChange={(e) => setDraft(e.target.value)}
        onFocus={(e) => e.currentTarget.select()}
        onBlur={() => onCommit(draft)}
        onKeyDown={(e) => {
          if (e.key === 'Enter') { e.preventDefault(); e.currentTarget.blur() }
          else if (e.key === 'Escape') { e.preventDefault(); onCancel() }
        }}
        style={{ width: `${Math.max(90, draft.length * 7.5 + 28)}px` }}
        className="text-xs px-2 py-1 border-2 border-blue-500 rounded focus:outline-none"
      />
    </td>
  )
}

function DatePickerCell({ displayValue, isChanged, isEditing, rowBg, stickyLeft, onStartEdit, onCommit, onCancel }: CellProps) {
  const [draft, setDraft] = useState(displayValue)
  const [wasEditing, setWasEditing] = useState(isEditing)
  if (isEditing !== wasEditing) {
    setWasEditing(isEditing)
    if (isEditing) setDraft(displayValue)
  }

  if (!isEditing) {
    return (
      <td
        onDoubleClick={onStartEdit}
        className="px-2 py-1 text-xs whitespace-nowrap cursor-pointer border-b border-r border-gray-200"
        style={{ backgroundColor: isChanged ? AMBER : rowBg, height: ROW_HEIGHT, ...frozenTdStyle(stickyLeft, false) }}
      >
        {displayValue || <span className="text-gray-300">—</span>}
      </td>
    )
  }

  return (
    <td className="px-1 py-0.5 border-b border-r border-gray-200 bg-white relative" style={{ zIndex: 40, ...frozenTdStyle(stickyLeft, true), backgroundColor: '#fff' }}>
      <input
        autoFocus
        type="date"
        value={toDateInputValue(draft)}
        onChange={(e) => setDraft(fromDateInputValue(e.target.value))}
        onBlur={() => onCommit(draft)}
        onKeyDown={(e) => {
          if (e.key === 'Enter') { e.preventDefault(); e.currentTarget.blur() }
          else if (e.key === 'Escape') { e.preventDefault(); onCancel() }
        }}
        className="text-xs px-2 py-1 border-2 border-blue-500 rounded focus:outline-none"
      />
    </td>
  )
}

// Excel-style per-column filter dropdown. Module scope so it isn't remounted on
// every parent render. Holds its own working state; nothing is applied to the
// grid until "Apply" is pressed.
interface ColumnFilterPanelProps {
  panelRef: React.RefObject<HTMLDivElement | null>
  columnName: string
  isDateCol: boolean
  values: string[]
  current: ColumnFilterState | undefined
  pos: { top: number; left: number }
  onApply: (state: ColumnFilterState) => void
  onClose: () => void
}

const FILTER_CHECKLIST_LIMIT = 50
const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']

// Tri-state checkbox (checked / unchecked / indeterminate) — React has no
// JSX prop for `indeterminate`, it's only settable as a DOM property, so this
// sets it imperatively via a ref callback on mount/update.
function TriCheckbox({ state, onChange }: { state: 'all' | 'none' | 'some'; onChange: () => void }) {
  return (
    <input
      type="checkbox"
      checked={state === 'all'}
      ref={(el) => { if (el) el.indeterminate = state === 'some' }}
      onChange={onChange}
    />
  )
}

// year -> month name -> ordered list of underlying date values (the exact
// display strings used everywhere else for filtering/checked-state, e.g.
// "3/15/2026") that fall in that year+month.
type DateTree = Map<string, Map<string, string[]>>

// `blanks` is genuinely-empty cells only — always at most one entry, since
// `values` is a deduped set. `otherText` is anything non-empty that still
// failed to parse as a date (a "date"-detected column can still hold real
// text — e.g. NTP Action Owner / NTP is waiting on both match the date regex
// but hold names/comments) — those keep their own real label instead of all
// being mislabeled "(blank)" together, which is what made the checklist show
// several rows that all *looked* identical.
function buildDateTree(values: string[]): { tree: DateTree; blanks: string[]; otherText: string[] } {
  const tree: DateTree = new Map()
  const blanks: string[] = []
  const otherText: string[] = []
  for (const v of values) {
    if (v === '') { blanks.push(v); continue }
    const d = parseDateAny(v)
    if (!d) { otherText.push(v); continue }
    const year = String(d.getFullYear())
    const month = MONTH_NAMES[d.getMonth()]
    if (!tree.has(year)) tree.set(year, new Map())
    const months = tree.get(year)!
    if (!months.has(month)) months.set(month, [])
    months.get(month)!.push(v)
  }
  otherText.sort((a, b) => a.localeCompare(b))
  return { tree, blanks, otherText }
}

function groupState(vals: string[], checked: Set<string>): 'all' | 'none' | 'some' {
  const n = vals.filter(v => checked.has(v)).length
  if (n === 0) return 'none'
  if (n === vals.length) return 'all'
  return 'some'
}

function ColumnFilterPanel({ panelRef, columnName, isDateCol, values, current, pos, onApply, onClose }: ColumnFilterPanelProps) {
  const [sort, setSort] = useState<ColSort>(current?.sort ?? null)
  const [search, setSearch] = useState('')
  const [checked, setChecked] = useState<Set<string>>(
    () => (current?.selectedValues ? new Set(current.selectedValues) : new Set(values))
  )

  const { tree: dateTree, blanks: dateBlanks, otherText: dateOtherText } = useMemo(
    () => (isDateCol ? buildDateTree(values) : { tree: new Map<string, Map<string, string[]>>(), blanks: [], otherText: [] }),
    [isDateCol, values]
  )
  // Years start expanded (there are usually only a handful) — months start
  // collapsed, so the default view is a compact Year > Month list instead of
  // every individual day at once. Matches Excel's grouped date filter tree.
  const [expandedYears, setExpandedYears] = useState<Set<string>>(() => new Set(dateTree.keys()))
  const [expandedMonths, setExpandedMonths] = useState<Set<string>>(new Set())

  const q = search.trim().toLowerCase()
  const label = (v: string) => (v === '' ? '(blank)' : v)

  const toggle = (v: string) => setChecked(prev => {
    const n = new Set(prev)
    if (n.has(v)) n.delete(v)
    else n.add(v)
    return n
  })
  const toggleGroup = (vals: string[]) => setChecked(prev => {
    const n = new Set(prev)
    const allOn = vals.every(v => n.has(v))
    vals.forEach(v => (allOn ? n.delete(v) : n.add(v)))
    return n
  })

  const apply = () => {
    const allSelected = checked.size === values.length && values.every(v => checked.has(v))
    onApply({ sort, selectedValues: allSelected ? null : Array.from(checked) })
  }

  // Flat (non-date) checklist — unchanged behavior, capped at 50 shown.
  const shownFlat = (q ? values.filter(v => v.toLowerCase().includes(q)) : values).slice(0, FILTER_CHECKLIST_LIMIT)

  // Grouped date checklist — while searching, only show groups containing a
  // match, force-expanded so the match is visible without extra clicks.
  const yearEntries = Array.from(dateTree.entries()).map(([year, months]) => {
    const monthEntries = Array.from(months.entries()).map(([month, vals]) => {
      const filteredVals = q ? vals.filter(v => v.toLowerCase().includes(q)) : vals
      return { month, vals, filteredVals }
    }).filter(m => !q || m.filteredVals.length > 0)
    const yearVals = Array.from(months.values()).flat()
    return { year, monthEntries, yearVals }
  }).filter(y => !q || y.monthEntries.length > 0)
  const shownBlanks = q ? dateBlanks.filter(v => label(v).toLowerCase().includes(q)) : dateBlanks
  const shownOtherText = q ? dateOtherText.filter(v => v.toLowerCase().includes(q)) : dateOtherText

  return (
    <div
      ref={panelRef}
      data-col-filter
      style={{ position: 'fixed', top: pos.top, left: pos.left, zIndex: 60, width: 280 }}
      className="bg-white border border-gray-300 rounded-lg shadow-2xl p-3 text-xs text-gray-800"
    >
      <div className="font-bold text-gray-700 mb-2 truncate">{columnName}</div>

      <div className="flex flex-col gap-1 mb-2">
        <button
          onClick={() => setSort(sort === 'asc' ? null : 'asc')}
          className={`text-left px-2 py-1 rounded ${sort === 'asc' ? 'bg-blue-100 text-blue-800 font-semibold' : 'hover:bg-gray-100'}`}
        >
          ↑ {isDateCol ? 'Oldest to Newest' : 'A → Z'}
        </button>
        <button
          onClick={() => setSort(sort === 'desc' ? null : 'desc')}
          className={`text-left px-2 py-1 rounded ${sort === 'desc' ? 'bg-blue-100 text-blue-800 font-semibold' : 'hover:bg-gray-100'}`}
        >
          ↓ {isDateCol ? 'Newest to Oldest' : 'Z → A'}
        </button>
      </div>

      <input
        value={search}
        onChange={(e) => setSearch(e.target.value)}
        placeholder={`Search ${columnName}...`}
        className="w-full px-2 py-1 border border-gray-300 rounded mb-2 focus:outline-none focus:border-blue-500"
      />

      <div className="flex items-center gap-3 mb-1">
        <button onClick={() => setChecked(new Set(values))} className="text-blue-700 hover:underline font-semibold">Select All</button>
        <button onClick={() => setChecked(new Set())} className="text-blue-700 hover:underline font-semibold">Clear All</button>
      </div>

      {!isDateCol && (
        <div className="border border-gray-200 rounded max-h-48 overflow-y-auto mb-2">
          {shownFlat.map(v => (
            <label key={v} className="flex items-center gap-2 px-2 py-1 hover:bg-gray-50 cursor-pointer">
              <input type="checkbox" checked={checked.has(v)} onChange={() => toggle(v)} />
              <span className="truncate">{label(v)}</span>
            </label>
          ))}
          {shownFlat.length === 0 && <div className="px-2 py-2 text-gray-400">No matching values</div>}
          {!q && values.length > FILTER_CHECKLIST_LIMIT && (
            <div className="px-2 py-1 text-gray-400 italic">Showing first {FILTER_CHECKLIST_LIMIT} of {values.length}</div>
          )}
        </div>
      )}

      {isDateCol && (
        <div className="border border-gray-200 rounded max-h-56 overflow-y-auto mb-2">
          {shownBlanks.map(v => (
            <label key={v} className="flex items-center gap-2 px-2 py-1 hover:bg-gray-50 cursor-pointer">
              <input type="checkbox" checked={checked.has(v)} onChange={() => toggle(v)} />
              <span className="truncate">(blank)</span>
            </label>
          ))}
          {shownOtherText.map(v => (
            <label key={v} className="flex items-center gap-2 px-2 py-1 hover:bg-gray-50 cursor-pointer">
              <input type="checkbox" checked={checked.has(v)} onChange={() => toggle(v)} />
              <span className="truncate">{v}</span>
            </label>
          ))}
          {yearEntries.map(({ year, monthEntries, yearVals }) => {
            const yearOpen = q ? true : expandedYears.has(year)
            return (
              <div key={year}>
                <div className="flex items-center gap-1 px-2 py-1 hover:bg-gray-50">
                  <button
                    onClick={() => setExpandedYears(prev => {
                      const n = new Set(prev)
                      if (n.has(year)) n.delete(year); else n.add(year)
                      return n
                    })}
                    className="w-3 text-gray-500"
                  >
                    {yearOpen ? '▾' : '▸'}
                  </button>
                  <label className="flex items-center gap-2 cursor-pointer flex-1">
                    <TriCheckbox state={groupState(yearVals, checked)} onChange={() => toggleGroup(yearVals)} />
                    <span className="font-semibold">{year}</span>
                  </label>
                </div>
                {yearOpen && monthEntries.map(({ month, vals, filteredVals }) => {
                  const monthKey = `${year}|${month}`
                  const monthOpen = q ? true : expandedMonths.has(monthKey)
                  return (
                    <div key={month}>
                      <div className="flex items-center gap-1 pl-5 pr-2 py-1 hover:bg-gray-50">
                        <button
                          onClick={() => setExpandedMonths(prev => {
                            const n = new Set(prev)
                            if (n.has(monthKey)) n.delete(monthKey); else n.add(monthKey)
                            return n
                          })}
                          className="w-3 text-gray-500"
                        >
                          {monthOpen ? '▾' : '▸'}
                        </button>
                        <label className="flex items-center gap-2 cursor-pointer flex-1">
                          <TriCheckbox state={groupState(vals, checked)} onChange={() => toggleGroup(vals)} />
                          <span>{month}</span>
                        </label>
                      </div>
                      {monthOpen && filteredVals.map(v => {
                        const d = parseDateAny(v)
                        return (
                          <label key={v} className="flex items-center gap-2 pl-9 pr-2 py-1 hover:bg-gray-50 cursor-pointer">
                            <input type="checkbox" checked={checked.has(v)} onChange={() => toggle(v)} />
                            <span className="truncate">{d ? d.getDate() : v}</span>
                          </label>
                        )
                      })}
                    </div>
                  )
                })}
              </div>
            )
          })}
          {yearEntries.length === 0 && shownBlanks.length === 0 && shownOtherText.length === 0 && <div className="px-2 py-2 text-gray-400">No matching values</div>}
        </div>
      )}

      <div className="flex items-center justify-end gap-2">
        <button onClick={onClose} className="px-3 py-1 rounded bg-gray-200 hover:bg-gray-300 font-semibold">Cancel</button>
        <button onClick={apply} className="px-3 py-1 rounded text-white font-semibold" style={{ backgroundColor: NAVY }}>Apply</button>
      </div>
    </div>
  )
}

export default function TrackerGridPage() {
  const [headers, setHeaders] = useState<string[]>([])
  const [trackerRows, setTrackerRows] = useState<TrackerRowData[]>([])
  const [loaded, setLoaded] = useState(false)

  const [views, setViews] = useState<TrackerView[]>([])
  const [viewsLoaded, setViewsLoaded] = useState(false)
  const [activeViewName, setActiveViewName] = useState('Default')

  const [editorMode, setEditorMode] = useState(false)
  const [draftHidden, setDraftHidden] = useState<Set<string>>(new Set())
  const [savePromptOpen, setSavePromptOpen] = useState(false)
  const [viewNameDraft, setViewNameDraft] = useState('')
  // When set, the column editor is editing an existing saved view (rather than
  // creating a new one) — saving overwrites it, renaming drops the old row.
  const [editingViewName, setEditingViewName] = useState<string | null>(null)

  const [pendingChanges, setPendingChanges] = useState<TrackerChange[]>([])
  // Visible "PM Updates"-style panel — same pattern as GC Call View / CM View
  // (a toggleable table of every pending edit with a Done checkbox per row,
  // instead of edits only being visible as a badge count on Copy Updates).
  const [showPendingPanel, setShowPendingPanel] = useState(false)
  const [pendingSearch, setPendingSearch] = useState('')
  const [pendingSortField, setPendingSortField] = useState<'hop' | 'field'>('hop')
  const [editingCell, setEditingCell] = useState<{ rowKey: string; field: string } | null>(null)

  const longPressRef = useRef<ReturnType<typeof setTimeout> | null>(null)

  // --- Smart search bar. `searchInput` is the raw box value; `searchQuery` is
  // the debounced value that actually drives filtering (300ms) so we don't
  // re-filter on every keystroke.
  const [searchInput, setSearchInput] = useState('')
  const [searchQuery, setSearchQuery] = useState('')

  // --- Per-column ("Excel style") filters. Map of columnName -> filter state.
  // Only entries that actually constrain something are kept. Sort is tracked
  // separately (see sortOrder below) — this only ever holds value filters now.
  const [columnFilters, setColumnFilters] = useState<Record<string, ColumnFilterState>>({})
  // Excel-style stacked multi-column sort. Index 0 is the primary key; each
  // later entry is a tiebreaker applied only when everything before it is
  // equal. Sorting a *new* column appends it (lowest priority) without
  // disturbing columns already sorted; re-sorting an already-sorted column
  // just flips its direction in place.
  const [sortOrder, setSortOrder] = useState<{ name: string; dir: ColSort }[]>([])
  // Which column's filter dropdown is open, and where to anchor it (viewport
  // coords — the panel is position:fixed, outside the table flow). Only one at
  // a time.
  const [filterPanel, setFilterPanel] = useState<{ col: string; top: number; left: number } | null>(null)
  const filterPanelRef = useRef<HTMLDivElement>(null)

  // --- Rendering-performance state (virtual scrolling, lazy columns, row
  // selection) — none of this touches data loading, view, or copy-updates logic.
  const scrollContainerRef = useRef<HTMLDivElement>(null)
  const scrollRafRef = useRef<number | null>(null)
  const clickDebounceRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  const [scrollTop, setScrollTop] = useState(0)
  const [scrollLeft, setScrollLeft] = useState(0)
  const [viewportSize, setViewportSize] = useState({ width: 900, height: 600 })
  const [selectedRowKey, setSelectedRowKey] = useState<string | null>(null)
  // User-resized column widths, keyed by column name — overrides the default
  // width heuristic once a column's been dragged. Session-only (not persisted).
  const [columnWidths, setColumnWidths] = useState<Record<string, number>>({})
  const colResizeRef = useRef<{ name: string; startX: number; startWidth: number } | null>(null)

  // Load tracker snapshot — all columns, no dedup (see parseAllTrackerRows).
  useEffect(() => {
    const load = async () => {
      const snap = await loadTrackerSnapshot()
      if (!snap) { setLoaded(true); return }
      const { headers: h, trackerRows: rows } = parseAllTrackerRows(snap.data)
      setHeaders(h)
      setTrackerRows(rows)
      setLoaded(true)
    }
    load()
  }, [])

  // Load saved views.
  useEffect(() => {
    const load = async () => {
      const { data } = await supabase.from('pm_updates_cache').select('id, updates').like('id', 'tracker-view-%')
      const loadedViews: TrackerView[] = []
      ;(data || []).forEach((row: { id: string; updates: string }) => {
        try {
          const parsed = JSON.parse(row.updates) as TrackerView
          if (parsed?.name) loadedViews.push(parsed)
        } catch {}
      })
      setViews(loadedViews)
      setViewsLoaded(true)
    }
    load()
  }, [])

  // Auto-create the three starter views on first load, once both the tracker
  // headers and the saved-views list have loaded and no views exist yet.
  useEffect(() => {
    if (!loaded || !viewsLoaded) return
    if (headers.length === 0) return
    if (views.length > 0) return
    const createStarters = async () => {
      const starters: TrackerView[] = STARTER_VIEW_DEFS.map(def => ({
        name: def.name,
        hiddenColumns: headers.filter(h => !def.keep.includes(h)),
      }))
      for (const v of starters) {
        await supabase.from('pm_updates_cache').upsert({
          id: `tracker-view-${v.name}`,
          updates: JSON.stringify(v),
          updated_at: new Date().toISOString(),
        })
      }
      setViews(starters)
    }
    createStarters()
    // Intentionally reacts only to load-completion + emptiness, not to the
    // full headers/views array identity — this should fire exactly once.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [loaded, viewsLoaded, headers.length, views.length])

  // Load today's pending changes.
  useEffect(() => {
    const load = async () => {
      const todayKey = new Date().toISOString().slice(0, 10)
      const { data } = await supabase.from('pm_updates_cache').select('updates').eq('id', `tracker-changes-${todayKey}`).single()
      if (data?.updates) {
        try { setPendingChanges(JSON.parse(data.updates)) } catch {}
      }
    }
    load()
  }, [])

  // Debounce the search box — 300ms after the last keystroke the query commits
  // and the memoized row filter recomputes.
  useEffect(() => {
    const t = setTimeout(() => setSearchQuery(searchInput), 300)
    return () => clearTimeout(t)
  }, [searchInput])

  // Close the column-filter dropdown on any click outside of it (and outside
  // the ▼ buttons, which toggle it themselves). Listener only mounted while a
  // panel is open; cleaned up on unmount / close.
  useEffect(() => {
    if (!filterPanel) return
    const onDown = (e: MouseEvent) => {
      const t = e.target as HTMLElement | null
      if (!t) return
      if (filterPanelRef.current?.contains(t)) return
      if (t.closest('[data-col-filter-btn]')) return
      setFilterPanel(null)
    }
    document.addEventListener('mousedown', onDown)
    return () => document.removeEventListener('mousedown', onDown)
  }, [filterPanel])

  // Measure the scroll container so the row/column virtualization windows
  // know how many cells actually fit in view. Re-measures on window resize.
  useEffect(() => {
    if (!loaded) return
    const measure = () => {
      const el = scrollContainerRef.current
      if (el) setViewportSize({ width: el.clientWidth, height: el.clientHeight })
    }
    measure()
    window.addEventListener('resize', measure)
    return () => window.removeEventListener('resize', measure)
  }, [loaded])

  // Scroll position drives which rows/columns are rendered — rAF-throttled so
  // a fast scroll gesture only triggers at most one state update per frame
  // instead of one per native scroll event.
  const handleGridScroll = useCallback(() => {
    if (scrollRafRef.current !== null) return
    scrollRafRef.current = requestAnimationFrame(() => {
      const el = scrollContainerRef.current
      if (el) {
        setScrollTop(el.scrollTop)
        setScrollLeft(el.scrollLeft)
      }
      scrollRafRef.current = null
    })
  }, [])

  // Snap back to top-left whenever anything that changes what rows or
  // columns are showing takes effect — search, a column filter, the sort, or
  // switching views/entering the column editor. The real scrollbar doesn't
  // move on its own when the row or column set shrinks or shuffles, so
  // without this the grid can sit scrolled past content that no longer
  // exists there and show nothing (or the wrong columns) even though it
  // filtered/switched correctly. Matches Excel, which also jumps to the top
  // of a freshly filtered/sorted range.
  const gridResetKey = `${activeViewName}|${editorMode}|${searchQuery}|${JSON.stringify(columnFilters)}|${JSON.stringify(sortOrder)}`
  const [prevGridResetKey, setPrevGridResetKey] = useState(gridResetKey)
  if (gridResetKey !== prevGridResetKey) {
    setPrevGridResetKey(gridResetKey)
    setScrollTop(0)
    setScrollLeft(0)
  }
  // The React-state half of the reset happens above (during render, per
  // React's own pattern for resetting state on a dependency change) — this
  // effect only syncs the real DOM scroll position to match, which is a
  // legitimate external-system update for an effect to make.
  useEffect(() => {
    const el = scrollContainerRef.current
    if (el) { el.scrollTop = 0; el.scrollLeft = 0 }
  }, [gridResetKey])

  // Debounced row selection (single click) — collapses rapid repeated clicks
  // into a single state update instead of one per click.
  const handleRowClick = useCallback((rowKey: string) => {
    if (clickDebounceRef.current) clearTimeout(clickDebounceRef.current)
    clickDebounceRef.current = setTimeout(() => setSelectedRowKey(rowKey), 150)
  }, [])

  // Drag-to-resize a column header's right edge. Global mousemove/mouseup
  // listeners (not React handlers) so the drag keeps tracking even if the
  // cursor leaves the narrow handle itself mid-drag.
  const startColumnResize = useCallback((name: string, startWidth: number, e: React.MouseEvent) => {
    e.preventDefault()
    e.stopPropagation()
    colResizeRef.current = { name, startX: e.clientX, startWidth }
    const handleMouseMove = (ev: MouseEvent) => {
      const state = colResizeRef.current
      if (!state) return
      const newWidth = Math.max(60, state.startWidth + (ev.clientX - state.startX))
      setColumnWidths(prev => ({ ...prev, [state.name]: newWidth }))
    }
    const handleMouseUp = () => {
      colResizeRef.current = null
      window.removeEventListener('mousemove', handleMouseMove)
      window.removeEventListener('mouseup', handleMouseUp)
    }
    window.addEventListener('mousemove', handleMouseMove)
    window.addEventListener('mouseup', handleMouseUp)
  }, [])

  const persistChanges = async (changes: TrackerChange[]) => {
    const todayKey = new Date().toISOString().slice(0, 10)
    await supabase.from('pm_updates_cache').upsert({
      id: `tracker-changes-${todayKey}`,
      updates: JSON.stringify(changes),
      updated_at: new Date().toISOString(),
    })
  }

  const saveEdit = (rowKey: string, hop: string, field: string, originalValue: string, newValue: string) => {
    setEditingCell(null)
    if (newValue === originalValue) return
    const change: TrackerChange = {
      rowKey, hop, field, oldValue: originalValue, newValue,
      timestamp: new Date().toISOString(), user: 'CJ',
    }
    setPendingChanges(prev => {
      const next = [...prev.filter(c => !(c.rowKey === rowKey && c.field === field)), change]
      persistChanges(next)
      return next
    })
  }

  const copyUpdates = () => {
    if (pendingChanges.length === 0) return
    const sorted = [...pendingChanges].sort((a, b) => a.hop.localeCompare(b.hop) || new Date(a.timestamp).getTime() - new Date(b.timestamp).getTime())
    const text = sorted.map(formatChangeLine).join('\n')
    navigator.clipboard.writeText(text)
      .then(() => alert('✅ Copied to clipboard!'))
      .catch(() => alert('Copy failed — please try manually'))
  }

  const clearChanges = () => {
    setPendingChanges([])
    persistChanges([])
    setShowPendingPanel(false)
  }

  const toggleChangeCompleted = (rowKey: string, field: string, timestamp: string) => {
    setPendingChanges(prev => {
      const next = prev.map(c => (c.rowKey === rowKey && c.field === field && c.timestamp === timestamp ? { ...c, completed: !c.completed } : c))
      persistChanges(next)
      return next
    })
  }

  const clearCompletedChanges = () => {
    setPendingChanges(prev => {
      const next = prev.filter(c => !c.completed)
      persistChanges(next)
      if (next.length === 0) setShowPendingPanel(false)
      return next
    })
  }

  // Switching views resets the search box and every column filter, but a
  // half-typed search survives ordinary re-renders (it lives in state).
  const selectView = (name: string) => {
    setActiveViewName(name)
    setEditorMode(false)
    setEditingViewName(null)
    setSearchInput('')
    setSearchQuery('')
    setColumnFilters({})
    setSortOrder([])
    setFilterPanel(null)
  }

  const enterEditorMode = () => {
    setEditingViewName(null)
    setDraftHidden(new Set())
    setSavePromptOpen(false)
    setViewNameDraft('')
    setFilterPanel(null)
    setEditorMode(true)
  }

  // Enter the column editor pre-populated with an existing view's hidden
  // columns — same UI as Create View, but Save overwrites the view.
  const editView = (v: TrackerView) => {
    setEditingViewName(v.name)
    setDraftHidden(new Set(v.hiddenColumns))
    setSavePromptOpen(false)
    setViewNameDraft(v.name)
    setFilterPanel(null)
    setEditorMode(true)
  }

  const cancelEditorMode = () => {
    setEditorMode(false)
    setSavePromptOpen(false)
    setViewNameDraft('')
    setEditingViewName(null)
    setDraftHidden(new Set())
  }

  const toggleDraftHidden = (name: string) => {
    setDraftHidden(prev => new Set(prev).add(name))
  }

  const restoreDraftHidden = (name: string) => {
    setDraftHidden(prev => {
      const next = new Set(prev)
      next.delete(name)
      return next
    })
  }

  const confirmSaveView = async () => {
    const name = viewNameDraft.trim()
    if (!name) return
    const view: TrackerView = { name, hiddenColumns: Array.from(draftHidden) }
    await supabase.from('pm_updates_cache').upsert({
      id: `tracker-view-${name}`,
      updates: JSON.stringify(view),
      updated_at: new Date().toISOString(),
    })
    // Editing an existing view under a new name — drop the old row so we don't
    // leave a stale duplicate behind.
    if (editingViewName && editingViewName !== name) {
      await supabase.from('pm_updates_cache').delete().eq('id', `tracker-view-${editingViewName}`)
    }
    setViews(prev => [...prev.filter(v => v.name !== name && v.name !== editingViewName), view])
    setActiveViewName(name)
    setColumnFilters({})
    setSortOrder([])
    setFilterPanel(null)
    cancelEditorMode()
  }

  const deleteView = async (name: string) => {
    if (!confirm(`Delete view "${name}"?`)) return
    await supabase.from('pm_updates_cache').delete().eq('id', `tracker-view-${name}`)
    setViews(prev => prev.filter(v => v.name !== name))
    if (activeViewName === name) {
      setActiveViewName('Default')
      setColumnFilters({})
      setSortOrder([])
      setFilterPanel(null)
      setSearchInput('')
      setSearchQuery('')
    }
  }

  const handleTouchStart = (name: string) => {
    longPressRef.current = setTimeout(() => deleteView(name), 600)
  }
  const handleTouchEnd = () => {
    if (longPressRef.current) { clearTimeout(longPressRef.current); longPressRef.current = null }
  }

  const activeView = views.find(v => v.name === activeViewName) || null
  const effectiveHidden: Set<string> = useMemo(
    () => (editorMode ? draftHidden : new Set(activeView?.hiddenColumns || [])),
    [editorMode, draftHidden, activeView]
  )

  // Full column list — memoized so it's only recomputed when the tracker's
  // header row actually changes, not on every render (e.g. every keystroke
  // while editing a cell).
  const allColumns = useMemo<GridColumn[]>(() => {
    const hopColIdx = headers.findIndex(h => h === 'HOP')
    const orderedIndexes = headers.length > 0
      ? [hopColIdx, ...headers.map((_, i) => i).filter(i => i !== hopColIdx)]
      : []
    return orderedIndexes.map(i => {
      const name = headers[i] || `Column ${i + 1}`
      const isHop = i === hopColIdx
      const isDate = isDateColumn(name)
      const defaultWidth = isHop ? 220 : (isDate ? 150 : 130)
      return { index: i, name, isHop, isDate, width: columnWidths[name] ?? defaultWidth }
    })
  }, [headers, columnWidths])

  const filteredColumns = useMemo(
    () => allColumns.filter(c => c.isHop || !effectiveHidden.has(c.name)),
    [allColumns, effectiveHidden]
  )

  // The first five columns (HOP, Site Name, Path ID, New CM, General
  // Contractor) are pinned. They're pulled out of horizontal virtualization
  // entirely and always rendered as position:sticky cells with incrementing
  // left offsets so they stack.
  const frozenColumns = useMemo<GridColumn[]>(() => {
    const out: GridColumn[] = []
    for (const name of FROZEN_COL_NAMES) {
      const c = filteredColumns.find(col => col.name === name)
      if (c) out.push(c)
    }
    return out
  }, [filteredColumns])

  const { frozenLeft, frozenTotalWidth } = useMemo(() => {
    const left: number[] = []
    let acc = 0
    for (const c of frozenColumns) { left.push(acc); acc += c.width }
    return { frozenLeft: left, frozenTotalWidth: acc }
  }, [frozenColumns])

  // Everything that isn't frozen — this is what gets lazy-loaded + horizontally
  // virtualized.
  const scrollableColumns = useMemo(() => {
    const fset = new Set(frozenColumns.map(c => c.name))
    return filteredColumns.filter(c => !fset.has(c.name))
  }, [filteredColumns, frozenColumns])

  // name -> column, for filter lookups (covers frozen + scrollable).
  const colMap = useMemo(() => new Map(filteredColumns.map(c => [c.name, c])), [filteredColumns])

  // Every scrollable column is always "loaded" — horizontal virtualization
  // below already caps what actually reaches the DOM to whatever's in the
  // scroll viewport, so there's no separate need to artificially truncate the
  // pool of columns you can scroll to. (A prior lazy-load cap did that
  // truncation and required a "Show More Columns" click to see the rest of a
  // view — removed since it just meant a saved view's columns didn't all show
  // up right away, without actually saving any rendering cost.)
  const loadedColumns = scrollableColumns

  // Horizontal virtualization over whatever's currently loaded — cumulative
  // left-edge offsets, then a scroll-position scan to find the visible slice.
  const colOffsets = useMemo(() => {
    const offsets: number[] = []
    let acc = 0
    loadedColumns.forEach(c => { offsets.push(acc); acc += c.width })
    return { offsets, total: acc }
  }, [loadedColumns])

  // Horizontal virtualization window, in the scrollable columns' own offset
  // space (0 = first scrollable column). The frozen block is sticky and always
  // covers the leftmost `frozenTotalWidth` px of the viewport, so the visible
  // slice of scrolling columns runs from `scrollLeft` to
  // `scrollLeft + (viewportWidth - frozenTotalWidth)`.
  const { renderedColumns, leftSpacerWidth, rightSpacerWidth } = useMemo(() => {
    const { offsets, total } = colOffsets
    const windowLeft = scrollLeft
    const windowRight = scrollLeft + Math.max(0, viewportSize.width - frozenTotalWidth)
    let start = 0
    while (start < loadedColumns.length && offsets[start] + loadedColumns[start].width < windowLeft) start++
    start = Math.max(0, start - COL_BUFFER)
    let end = start
    let accWidth = start > 0 ? offsets[start] : 0
    while (end < loadedColumns.length && accWidth < windowRight) {
      accWidth += loadedColumns[end].width
      end++
    }
    end = Math.min(loadedColumns.length, end + COL_BUFFER)
    const rendered = loadedColumns.slice(start, end)
    const left = start > 0 ? offsets[start] : 0
    const right = total - (end < loadedColumns.length ? offsets[end] : total)
    return { renderedColumns: rendered, leftSpacerWidth: left, rightSpacerWidth: right }
  }, [loadedColumns, colOffsets, scrollLeft, viewportSize.width, frozenTotalWidth])

  const changeMap = useMemo(() => {
    const map = new Map<string, TrackerChange>()
    pendingChanges.forEach(c => map.set(`${c.rowKey}|${c.field}`, c))
    return map
  }, [pendingChanges])

  // Display text for a cell, honoring an in-flight edit override. Shared by
  // search, the per-column checklists, and the column-filter pass so all three
  // agree on what a cell "contains".
  const cellText = useCallback((row: TrackerRowData, col: GridColumn): string => {
    const change = changeMap.get(`${row.rowKey}|${col.name}`)
    if (change) return change.newValue ?? ''
    if (col.isHop) return row.hop
    return cellDisplayValue(row.cells[col.index], col.isDate).text
  }, [changeMap])

  // --- Search filter. Recomputes only when the debounced query, the row set,
  // or the visible column set changes — never on unrelated re-renders.
  const searchedRows = useMemo(() => {
    const q = searchQuery.trim().toLowerCase()
    if (!q) return trackerRows
    return trackerRows.filter(row =>
      filteredColumns.some(col => cellText(row, col).toLowerCase().includes(q))
    )
  }, [trackerRows, searchQuery, filteredColumns, cellText])

  const activeFilterCol = filterPanel ? colMap.get(filterPanel.col) : undefined

  // Rows a column's own checklist should be built from: search applied, plus
  // every OTHER column's active value filter (never this column's own — you're
  // in the middle of changing that one). Without this, filtering column A
  // then opening column B's dropdown showed every value B ever has anywhere
  // in the sheet, including combinations that can't coexist with A's filter —
  // pick one of those and you'd get zero rows with no visible reason why,
  // which reads exactly like "filters don't line up with what's on screen."
  const rowsForActiveChecklist = useMemo(() => {
    if (!activeFilterCol) return searchedRows
    const otherSpecs = Object.entries(columnFilters)
      .filter(([name, f]) => name !== activeFilterCol.name && f.selectedValues !== null)
      .map(([name, f]) => ({ col: colMap.get(name), allowed: new Set(f.selectedValues as string[]) }))
      .filter((s): s is { col: GridColumn; allowed: Set<string> } => !!s.col)
    if (otherSpecs.length === 0) return searchedRows
    return searchedRows.filter(row => otherSpecs.every(s => s.allowed.has(cellText(row, s.col))))
  }, [activeFilterCol, searchedRows, columnFilters, colMap, cellText])

  // Unique values for the filter checklist — computed for only the column
  // whose dropdown is actually open, not every visible column. This used to
  // scan every column x every row (up to 100+ columns in the Default view)
  // on every render where the row or column set changed, including every
  // view switch — an O(columns x rows) pass that was the main source of the
  // sluggishness switching views, worse now that a HOP's two site rows both
  // show (roughly doubling row count). A filter dropdown only ever needs one
  // column's values at a time, so there's no reason to compute the rest.
  const activeColumnValues = useMemo(() => {
    if (!activeFilterCol) return []
    const set = new Set<string>()
    for (const row of rowsForActiveChecklist) set.add(cellText(row, activeFilterCol))
    const arr = Array.from(set)
    if (activeFilterCol.isDate) {
      arr.sort((a, b) => (parseDateAny(a)?.getTime() ?? -Infinity) - (parseDateAny(b)?.getTime() ?? -Infinity))
    } else {
      arr.sort((a, b) => a.localeCompare(b))
    }
    return arr
  }, [activeFilterCol, rowsForActiveChecklist, cellText])

  // --- Column filters applied on top of the search filter. Single pass across
  // every active column filter simultaneously (never a sequential loop), then
  // one optional sort by whichever column currently holds a sort.
  // Date sort compares the raw cell value's parsed epoch directly, not the
  // already-formatted display string re-parsed a second time — going
  // Date -> "3/15/2026" -> Date again is lossy (a fresh Date parsed from a
  // locale string picks it up at local midnight, which can land on a
  // different calendar day than the original UTC value depending on the
  // viewer's timezone) and was corrupting sort order for populated dates.
  const dateSortValue = useCallback((row: TrackerRowData, col: GridColumn): number | null => {
    const change = changeMap.get(`${row.rowKey}|${col.name}`)
    const raw = change ? change.newValue : row.cells[col.index]
    const d = parseDateAny(raw)
    return d ? d.getTime() : null
  }, [changeMap])

  const compareRowsOnColumn = useCallback((a: TrackerRowData, b: TrackerRowData, col: GridColumn): number => {
    if (col.isDate) {
      const da = dateSortValue(a, col)
      const db = dateSortValue(b, col)
      return (da ?? -Infinity) - (db ?? -Infinity)
    }
    return cellText(a, col).localeCompare(cellText(b, col))
  }, [dateSortValue, cellText])

  const displayRows = useMemo(() => {
    const valueSpecs = Object.entries(columnFilters)
      .filter(([, f]) => f.selectedValues !== null)
      .map(([name, f]) => ({ col: colMap.get(name), allowed: new Set(f.selectedValues as string[]) }))
      .filter((s): s is { col: GridColumn; allowed: Set<string> } => !!s.col)

    let rows = searchedRows
    if (valueSpecs.length > 0) {
      rows = rows.filter(row => valueSpecs.every(s => s.allowed.has(cellText(row, s.col))))
    }

    // Excel-style stacked sort — sortOrder[0] is the primary key, each
    // subsequent entry only breaks ties left by the ones before it.
    const levels = sortOrder
      .map(s => ({ col: colMap.get(s.name), dir: s.dir }))
      .filter((l): l is { col: GridColumn; dir: ColSort } => !!l.col)
    if (levels.length > 0) {
      rows = [...rows].sort((a, b) => {
        for (const { col, dir } of levels) {
          const cmp = compareRowsOnColumn(a, b, col)
          if (cmp !== 0) return dir === 'asc' ? cmp : -cmp
        }
        return 0
      })
    }
    return rows
  }, [searchedRows, columnFilters, sortOrder, colMap, cellText, compareRowsOnColumn])

  // Vertical virtualization — same idea, over the filtered/sorted HOP row list.
  const { visibleRows, topSpacerHeight, bottomSpacerHeight } = useMemo(() => {
    const totalRows = displayRows.length
    const visibleSlots = Math.ceil(viewportSize.height / ROW_HEIGHT) + ROW_BUFFER * 2
    // scrollTop reflects wherever the container was scrolled to under the
    // *previous* (possibly much longer) row set — a search/filter that
    // shrinks the result set doesn't move the real scrollbar, so without
    // clamping, `start` could land past the new, shorter `totalRows`. That
    // made `.slice(start, end)` come back empty (start > end) with a huge
    // top spacer — the grid would render nothing and just look frozen, even
    // though the "Showing N of M" count above it was already correct.
    const rawStart = Math.floor(scrollTop / ROW_HEIGHT) - ROW_BUFFER
    const start = Math.max(0, Math.min(rawStart, totalRows))
    const end = Math.min(totalRows, start + visibleSlots)
    return {
      visibleRows: displayRows.slice(start, end).map((row, i) => ({ row, rowIndex: start + i })),
      topSpacerHeight: start * ROW_HEIGHT,
      bottomSpacerHeight: (totalRows - end) * ROW_HEIGHT,
    }
  }, [displayRows, scrollTop, viewportSize.height])

  const spacerColCount = frozenColumns.length + 2 + renderedColumns.length
  const tableWidth = frozenTotalWidth + colOffsets.total

  const isColFilterActive = (name: string) => !!columnFilters[name] || sortOrder.some(s => s.name === name)
  // 1-based priority + direction arrow for a column currently part of the
  // stacked sort — e.g. "1↑" for the primary key, "2↓" for the next tiebreaker.
  const sortBadge = (name: string): string | null => {
    const idx = sortOrder.findIndex(s => s.name === name)
    if (idx === -1) return null
    return `${idx + 1}${sortOrder[idx].dir === 'asc' ? '↑' : '↓'}`
  }

  // Master "clear everything" — every value filter and every level of the
  // stacked sort, same as Excel's Data > Clear (search stays untouched, same
  // as Excel not treating its separate Find box as part of the filter state).
  const clearAllSortsAndFilters = () => {
    setColumnFilters({})
    setSortOrder([])
    setFilterPanel(null)
  }

  const openColumnFilter = (name: string, e: React.MouseEvent) => {
    e.stopPropagation()
    if (filterPanel?.col === name) { setFilterPanel(null); return }
    const r = (e.currentTarget as HTMLElement).getBoundingClientRect()
    setFilterPanel({
      col: name,
      top: Math.min(r.bottom + 4, window.innerHeight - 380),
      left: Math.max(8, Math.min(r.left, window.innerWidth - 272)),
    })
  }

  // Header contents (label + editor ✕ + filter ▼ + resize handle) — shared by
  // frozen and scrolling header cells so both render identically. The resize
  // handle is a thin absolutely-positioned strip on the right edge — the
  // parent <th> is position:sticky, which (like position:relative) is a valid
  // containing block for it.
  const headerContent = (col: GridColumn) => {
    const badge = sortBadge(col.name)
    return (
    <>
      <div className="flex items-center justify-between gap-1">
        <span className="truncate">{col.name}</span>
        <div className="flex items-center gap-1 flex-shrink-0">
          {badge && (
            <span
              title="Sort priority — click the ▼ menu to change or clear it"
              className="text-xs font-bold px-1 rounded"
              style={{ backgroundColor: '#FFD166', color: NAVY }}
            >
              {badge}
            </span>
          )}
          {editorMode && !col.isHop && (
            <button
              onClick={() => toggleDraftHidden(col.name)}
              className="text-red-300 hover:text-red-100 font-bold text-xs"
              title="Hide column"
            >
              ✕
            </button>
          )}
          <button
            data-col-filter-btn
            onClick={(e) => openColumnFilter(col.name, e)}
            title="Filter / sort column"
            className="leading-none"
            // Headers are navy, so an active filter reads as a bright filled
            // marker rather than "navy on navy".
            style={{ fontSize: 10, color: isColFilterActive(col.name) ? '#FFD166' : 'rgba(255,255,255,0.5)' }}
          >
            ▼
          </button>
        </div>
      </div>
      <div
        onMouseDown={(e) => startColumnResize(col.name, col.width, e)}
        onClick={(e) => e.stopPropagation()}
        title="Drag to resize column"
        style={{ position: 'absolute', top: 0, right: 0, bottom: 0, width: 6, cursor: 'col-resize', zIndex: 5 }}
        className="hover:bg-blue-300/60"
      />
    </>
    )
  }

  const viewBtnClass = (name: string) =>
    `px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors ${activeViewName === name && !editorMode ? 'text-white' : 'bg-gray-200 text-gray-700 hover:bg-gray-300'}`

  return (
    <div className="min-h-screen bg-gray-100 text-gray-900 p-4">
      <BackToDashboard />

      <style>{`
        .tracker-scroll::-webkit-scrollbar { width: 12px; height: 12px; }
        .tracker-scroll::-webkit-scrollbar-track { background: #f1f2f4; border-radius: 6px; }
        .tracker-scroll::-webkit-scrollbar-thumb { background: #b8bcc4; border-radius: 6px; border: 3px solid #f1f2f4; }
        .tracker-scroll::-webkit-scrollbar-thumb:hover { background: #8b9099; }
        .tracker-scroll::-webkit-scrollbar-corner { background: #f1f2f4; }
        .tracker-scroll { scrollbar-width: thin; scrollbar-color: #b8bcc4 #f1f2f4; }
      `}</style>

      <div className="flex items-center justify-between mb-3 flex-wrap gap-3">
        <h1 className="text-xl font-bold">📊 Tracker Grid</h1>

        <div className="flex items-center gap-2 flex-wrap">
          <button
            onClick={() => selectView('Default')}
            style={activeViewName === 'Default' && !editorMode ? { backgroundColor: NAVY } : undefined}
            className={viewBtnClass('Default')}
          >
            Default
          </button>
          {views.map(v => (
            <span key={v.name} className="inline-flex items-center gap-0.5">
              <button
                onClick={() => selectView(v.name)}
                onContextMenu={(e) => { e.preventDefault(); deleteView(v.name) }}
                onTouchStart={() => handleTouchStart(v.name)}
                onTouchEnd={handleTouchEnd}
                style={activeViewName === v.name && !editorMode ? { backgroundColor: NAVY } : undefined}
                className={viewBtnClass(v.name)}
                title="Right-click or long-press to delete"
              >
                {v.name}
              </button>
              <button
                onClick={() => editView(v)}
                className="px-1 text-xs text-gray-500 hover:text-gray-900"
                title={`Edit "${v.name}"`}
              >
                ✏️
              </button>
              <button
                onClick={() => deleteView(v.name)}
                className="px-1 text-xs text-gray-500 hover:text-red-600"
                title={`Delete "${v.name}"`}
              >
                🗑️
              </button>
            </span>
          ))}
          <button
            onClick={enterEditorMode}
            className="px-3 py-1.5 rounded-lg text-xs font-semibold border border-dashed border-gray-400 text-gray-600 hover:bg-gray-200"
          >
            + Create View
          </button>
        </div>

        <div className="flex items-center gap-2">
          {pendingChanges.length > 0 && (
            <button
              onClick={() => setShowPendingPanel(s => !s)}
              className="bg-amber-500 hover:bg-amber-600 text-white px-4 py-2 rounded-lg text-sm font-semibold"
            >
              📋 Pending Updates ({pendingChanges.length})
            </button>
          )}
          <button
            onClick={copyUpdates}
            disabled={pendingChanges.length === 0}
            className="bg-blue-600 hover:bg-blue-700 disabled:opacity-40 disabled:cursor-not-allowed text-white px-4 py-2 rounded-lg text-sm font-semibold"
          >
            📋 Copy Updates ({pendingChanges.length})
          </button>
          <button
            onClick={clearChanges}
            disabled={pendingChanges.length === 0}
            className="bg-gray-600 hover:bg-gray-700 disabled:opacity-40 disabled:cursor-not-allowed text-white px-4 py-2 rounded-lg text-sm font-semibold"
          >
            🗑 Clear Changes
          </button>
        </div>
      </div>

      {/* Pending Updates panel — same shape as GC/CM view's PM Updates panel:
          a searchable, sortable table with a Done checkbox per row, persisted
          to Supabase on every change (via toggleChangeCompleted/persistChanges). */}
      {showPendingPanel && pendingChanges.length > 0 && (
        <div className="mb-4 bg-amber-50 border border-amber-300 rounded-xl p-4">
          <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
            <h2 className="text-amber-900 font-bold text-base">📋 Pending Updates — Update These in Your Tracker</h2>
            <div className="flex gap-2 items-center flex-wrap">
              <input
                type="text"
                placeholder="Search HOP or field..."
                value={pendingSearch}
                onChange={(e) => setPendingSearch(e.target.value)}
                className="bg-white border border-amber-300 text-amber-900 text-xs rounded px-2 py-1 w-44 focus:outline-none focus:border-amber-500"
              />
              <button
                onClick={() => setPendingSortField(prev => prev === 'field' ? 'hop' : 'field')}
                className="bg-amber-200 hover:bg-amber-300 text-amber-900 text-xs px-3 py-1 rounded font-semibold"
              >
                Sort by {pendingSortField === 'field' ? 'Field ↑' : 'HOP ↑'}
              </button>
            </div>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="text-amber-700 text-xs">
                  <th className="text-left p-2">Done</th>
                  <th className="text-left p-2">HOP</th>
                  <th className="text-left p-2">Field</th>
                  <th className="text-left p-2">Old Value</th>
                  <th className="text-left p-2">New Value</th>
                  <th className="text-left p-2">Logged At</th>
                </tr>
              </thead>
              <tbody>
                {[...pendingChanges]
                  .filter(c => {
                    if (!pendingSearch) return true
                    const q = pendingSearch.toLowerCase()
                    return c.hop.toLowerCase().includes(q) || c.field.toLowerCase().includes(q)
                  })
                  .sort((a, b) => pendingSortField === 'field' ? a.field.localeCompare(b.field) : a.hop.localeCompare(b.hop))
                  .map(c => (
                    <tr key={`${c.rowKey}-${c.field}-${c.timestamp}`} className={`border-t border-amber-200 ${c.completed ? 'opacity-40' : ''}`}>
                      <td className="p-2">
                        <input
                          type="checkbox"
                          checked={c.completed || false}
                          onChange={() => toggleChangeCompleted(c.rowKey, c.field, c.timestamp)}
                          className="w-4 h-4 cursor-pointer accent-green-600"
                        />
                      </td>
                      <td className={`p-2 font-semibold ${c.completed ? 'line-through text-gray-500' : 'text-gray-900'}`}>{c.hop}</td>
                      <td className={`p-2 ${c.completed ? 'line-through text-gray-500' : 'text-amber-800'}`}>{c.field}</td>
                      <td className="p-2 text-gray-500">{c.oldValue || '—'}</td>
                      <td className={`p-2 font-bold ${c.completed ? 'text-gray-500' : 'text-green-700'}`}>{c.newValue}</td>
                      <td className="p-2 text-gray-500">{new Date(c.timestamp).toLocaleString()}</td>
                    </tr>
                  ))}
              </tbody>
            </table>
          </div>
          <div className="mt-4 flex gap-3">
            <button onClick={clearCompletedChanges} className="bg-green-700 hover:bg-green-600 text-white px-4 py-2 rounded text-sm font-semibold">
              ✅ Clear Completed
            </button>
            <button onClick={clearChanges} className="bg-gray-600 hover:bg-gray-500 text-white px-4 py-2 rounded text-sm">
              🗑 Clear All
            </button>
          </div>
        </div>
      )}

      {editorMode && (
        <div className="mb-3 bg-blue-50 border border-blue-300 rounded-lg p-3">
          <div className="flex items-center justify-between flex-wrap gap-2 mb-2">
            <p className="text-sm text-blue-900 font-semibold">
              {editingViewName
                ? `Editing view "${editingViewName}" — click ✕ on any column to hide it`
                : 'Column Editor Mode — click ✕ on any column to hide it from this view'}
            </p>
            <div className="flex items-center gap-2">
              {!savePromptOpen ? (
                <>
                  <button onClick={() => setSavePromptOpen(true)} className="bg-blue-600 hover:bg-blue-700 text-white text-xs px-3 py-1.5 rounded font-semibold">
                    Save View
                  </button>
                  <button onClick={cancelEditorMode} className="bg-gray-300 hover:bg-gray-400 text-gray-800 text-xs px-3 py-1.5 rounded font-semibold">
                    Cancel
                  </button>
                </>
              ) : (
                <>
                  <input
                    autoFocus
                    type="text"
                    value={viewNameDraft}
                    onChange={(e) => setViewNameDraft(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') confirmSaveView()
                      else if (e.key === 'Escape') setSavePromptOpen(false)
                    }}
                    placeholder="View name..."
                    className="text-xs px-2 py-1 border border-blue-400 rounded focus:outline-none"
                  />
                  <button onClick={confirmSaveView} className="bg-green-600 hover:bg-green-700 text-white text-xs px-3 py-1.5 rounded font-semibold">
                    ✓ Save
                  </button>
                  <button onClick={cancelEditorMode} className="bg-gray-300 hover:bg-gray-400 text-gray-800 text-xs px-3 py-1.5 rounded font-semibold">
                    Cancel
                  </button>
                </>
              )}
            </div>
          </div>
          {draftHidden.size > 0 && (
            <div>
              <p className="text-xs text-gray-500 mb-1">Hidden Columns</p>
              <div className="flex flex-wrap gap-1">
                {Array.from(draftHidden).map(name => (
                  <span key={name} className="inline-flex items-center gap-1 bg-gray-300 text-gray-700 text-xs px-2 py-1 rounded-full">
                    {name}
                    <button onClick={() => restoreDraftHidden(name)} className="font-bold text-green-700 hover:text-green-900" title="Restore column">+</button>
                  </span>
                ))}
              </div>
            </div>
          )}
        </div>
      )}

      {!loaded && (
        <div className="bg-white border border-gray-300 rounded-lg px-4 py-8 text-center text-gray-500">Loading tracker...</div>
      )}
      {loaded && headers.length === 0 && (
        <div className="bg-white border border-gray-300 rounded-lg px-4 py-8 text-center text-gray-500">
          No tracker data found — go to Dashboard to upload your tracker
        </div>
      )}

      {loaded && headers.length > 0 && (
        <div className="mb-3 flex items-center gap-3 flex-wrap">
          <div className="relative flex-1 min-w-[280px] max-w-2xl">
            <input
              type="text"
              value={searchInput}
              onChange={(e) => setSearchInput(e.target.value)}
              placeholder="Search all columns — HOP, GC, Path ID, PO number..."
              className="w-full pl-3 pr-8 py-2 text-sm border border-gray-300 rounded-lg focus:outline-none focus:border-blue-500"
            />
            {searchInput && (
              <button
                onClick={() => setSearchInput('')}
                className="absolute right-2 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-700 text-lg leading-none"
                title="Clear search"
              >
                ×
              </button>
            )}
          </div>
          <span className="text-xs font-semibold text-gray-600">
            Showing {displayRows.length} of {trackerRows.length} HOPs
          </span>
          {(Object.keys(columnFilters).length > 0 || sortOrder.length > 0) && (
            <button
              onClick={clearAllSortsAndFilters}
              className="text-xs font-semibold text-red-600 hover:text-red-800 underline"
              title="Remove every column filter and every level of the sort"
            >
              ✕ Clear All Filters & Sorts
            </button>
          )}
        </div>
      )}

      {loaded && headers.length > 0 && (
        <div
          ref={scrollContainerRef}
          onScroll={handleGridScroll}
          className="border border-gray-300 rounded-lg bg-white tracker-scroll"
          style={{ overflowX: 'scroll', overflowY: 'auto', height: 'calc(100vh - 300px)' }}
        >
          <table style={{ tableLayout: 'fixed', borderCollapse: 'collapse', width: tableWidth }}>
            <colgroup>
              {frozenColumns.map(c => <col key={c.name} style={{ width: c.width }} />)}
              <col style={{ width: leftSpacerWidth }} />
              {renderedColumns.map(c => <col key={c.name} style={{ width: c.width }} />)}
              <col style={{ width: rightSpacerWidth }} />
            </colgroup>
            <thead>
              <tr>
                {frozenColumns.map((col, i) => (
                  <th
                    key={col.name}
                    style={{
                      position: 'sticky',
                      top: 0,
                      left: frozenLeft[i],
                      // Above the scrolling headers (z 20) so the corner stays clean.
                      zIndex: 30,
                      backgroundColor: NAVY,
                      height: ROW_HEIGHT,
                      boxShadow: FROZEN_SHADOW,
                    }}
                    className="text-white text-xs font-bold px-2 py-2 text-left border-r border-b border-blue-900"
                  >
                    {headerContent(col)}
                  </th>
                ))}
                <th style={{ position: 'sticky', top: 0, zIndex: 20, backgroundColor: NAVY }} />
                {renderedColumns.map(col => (
                  <th
                    key={col.name}
                    style={{ position: 'sticky', top: 0, zIndex: 20, backgroundColor: NAVY, height: ROW_HEIGHT }}
                    className="text-white text-xs font-bold px-2 py-2 text-left border-r border-b border-blue-900"
                  >
                    {headerContent(col)}
                  </th>
                ))}
                <th style={{ position: 'sticky', top: 0, zIndex: 20, backgroundColor: NAVY }} />
              </tr>
            </thead>
            <tbody>
              {topSpacerHeight > 0 && (
                <tr style={{ height: topSpacerHeight }}>
                  <td colSpan={spacerColCount} style={{ padding: 0, border: 'none' }} />
                </tr>
              )}
              {visibleRows.map(({ row, rowIndex }) => {
                const isSelected = row.rowKey === selectedRowKey
                const rowBg = isSelected ? SELECTED_ROW : (rowIndex % 2 === 1 ? ALT_ROW : '#FFFFFF')
                // Frozen columns carry a solid (white / selected) background so
                // scrolling cells never show through underneath them.
                const frozenBg = isSelected ? SELECTED_ROW : '#FFFFFF'
                return (
                  <tr key={row.rowKey} onClick={() => handleRowClick(row.rowKey)} style={{ height: ROW_HEIGHT, cursor: 'pointer' }}>
                    {frozenColumns.map((col, i) => {
                      if (col.isHop) {
                        return (
                          <td
                            key={col.name}
                            style={{
                              position: 'sticky',
                              left: frozenLeft[i],
                              zIndex: 10,
                              backgroundColor: frozenBg,
                              color: NAVY,
                              height: ROW_HEIGHT,
                              boxShadow: FROZEN_SHADOW,
                            }}
                            className="px-2 py-1 text-xs font-bold whitespace-nowrap border-r border-b border-gray-200"
                          >
                            {row.hop}
                          </td>
                        )
                      }
                      const raw = row.cells[col.index]
                      const { text: displayValue, treatAsDate } = cellDisplayValue(raw, col.isDate)
                      const change = changeMap.get(`${row.rowKey}|${col.name}`)
                      const shown = change ? change.newValue : displayValue
                      const isEditing = editingCell?.rowKey === row.rowKey && editingCell?.field === col.name
                      const cellProps: CellProps = {
                        displayValue: shown,
                        isChanged: !!change,
                        isEditing,
                        rowBg: frozenBg,
                        stickyLeft: frozenLeft[i],
                        onStartEdit: () => setEditingCell({ rowKey: row.rowKey, field: col.name }),
                        onCommit: (newValue: string) => saveEdit(row.rowKey, row.hop, col.name, displayValue, newValue),
                        onCancel: () => setEditingCell(null),
                      }
                      return treatAsDate
                        ? <DatePickerCell key={col.name} {...cellProps} />
                        : <EditableCell key={col.name} {...cellProps} />
                    })}
                    <td style={{ backgroundColor: rowBg, height: ROW_HEIGHT }} />
                    {renderedColumns.map(col => {
                      if (col.isHop) {
                        return (
                          <td
                            key={col.name}
                            style={{ backgroundColor: rowBg, color: NAVY, height: ROW_HEIGHT }}
                            className="px-2 py-1 text-xs font-bold whitespace-nowrap border-r border-b border-gray-200"
                          >
                            {row.hop}
                          </td>
                        )
                      }
                      const raw = row.cells[col.index]
                      const { text: displayValue, treatAsDate } = cellDisplayValue(raw, col.isDate)
                      const change = changeMap.get(`${row.rowKey}|${col.name}`)
                      const shown = change ? change.newValue : displayValue
                      const isEditing = editingCell?.rowKey === row.rowKey && editingCell?.field === col.name
                      const cellProps: CellProps = {
                        displayValue: shown,
                        isChanged: !!change,
                        isEditing,
                        rowBg,
                        onStartEdit: () => setEditingCell({ rowKey: row.rowKey, field: col.name }),
                        onCommit: (newValue: string) => saveEdit(row.rowKey, row.hop, col.name, displayValue, newValue),
                        onCancel: () => setEditingCell(null),
                      }
                      return treatAsDate
                        ? <DatePickerCell key={col.name} {...cellProps} />
                        : <EditableCell key={col.name} {...cellProps} />
                    })}
                    <td style={{ backgroundColor: rowBg, height: ROW_HEIGHT }} />
                  </tr>
                )
              })}
              {bottomSpacerHeight > 0 && (
                <tr style={{ height: bottomSpacerHeight }}>
                  <td colSpan={spacerColCount} style={{ padding: 0, border: 'none' }} />
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {filterPanel && activeFilterCol && (
        <ColumnFilterPanel
          key={filterPanel.col}
          panelRef={filterPanelRef}
          columnName={activeFilterCol.name}
          isDateCol={activeFilterCol.isDate}
          values={activeColumnValues}
          current={{
            sort: sortOrder.find(s => s.name === activeFilterCol.name)?.dir ?? null,
            selectedValues: columnFilters[activeFilterCol.name]?.selectedValues ?? null,
          }}
          pos={{ top: filterPanel.top, left: filterPanel.left }}
          onClose={() => setFilterPanel(null)}
          onApply={(state) => {
            const name = activeFilterCol.name

            // Sort is tracked separately from value filters (see sortOrder) —
            // Excel-style stacking: sorting a column that isn't already part
            // of the active sort appends it as the lowest-priority tiebreaker
            // without disturbing columns sorted earlier; re-sorting a column
            // that's already active just flips its direction in place;
            // clearing a column's sort removes only that column's level.
            setSortOrder(prev => {
              const idx = prev.findIndex(s => s.name === name)
              if (!state.sort) return prev.filter(s => s.name !== name)
              if (idx === -1) return [...prev, { name, dir: state.sort }]
              return prev.map(s => (s.name === name ? { ...s, dir: state.sort } : s))
            })

            setColumnFilters(prev => {
              const next = { ...prev }
              if (state.selectedValues === null) delete next[name]
              else next[name] = { sort: null, selectedValues: state.selectedValues }
              return next
            })
            setFilterPanel(null)
          }}
        />
      )}
    </div>
  )
}
