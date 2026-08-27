'use client'

export const dynamic = 'force-dynamic'

import { useState, useEffect, useRef } from 'react'
import { supabase, loadTrackerSnapshot } from '../lib/supabase'
import { parseDateAny } from '../lib/grTracker'
import BackToDashboard from '../components/BackToDashboard'

const NAVY = '#124191'
const AMBER = '#FFF3CD'
const ALT_ROW = '#F7F8FA'

interface TrackerRowData {
  hop: string
  cells: unknown[]
}

interface TrackerChange {
  hop: string
  field: string
  oldValue: string
  newValue: string
  timestamp: string
  user: string
}

interface TrackerView {
  name: string
  hiddenColumns: string[]
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
// survives. Same DON 444 filter + dedup-by-HOP rule as every other page
// (prefer the row where GC and New CM are both populated, since a HOP's other
// row may be a blank/partial duplicate) — deliberately does NOT filter by
// Nokia PM, since this grid is a program-wide utility, not scoped to one PM.
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
  const gcCol = headers.findIndex(h => h === 'General Contractor')
  const cmCol = headers.findIndex(h => h === 'New CM')

  const hopMap = new Map<string, unknown[][]>()
  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i] as unknown[]
    if (!row) continue
    const don = normHeader(row[don444Col]).toUpperCase()
    if (don !== 'DON 444') continue
    const hop = normHeader(row[hopCol])
    if (!hop || hop === 'undefined') continue
    if (!hopMap.has(hop)) hopMap.set(hop, [])
    hopMap.get(hop)!.push(row)
  }

  const trackerRows: TrackerRowData[] = []
  hopMap.forEach((candidateRows, hop) => {
    const chosen = candidateRows.find(r => normHeader(r[gcCol]) && normHeader(r[cmCol])) || candidateRows[0]
    trackerRows.push({ hop, cells: chosen })
  })
  trackerRows.sort((a, b) => a.hop.localeCompare(b.hop))

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
  onStartEdit: () => void
  onCommit: (newValue: string) => void
  onCancel: () => void
}

function EditableCell({ displayValue, isChanged, isEditing, rowBg, onStartEdit, onCommit, onCancel }: CellProps) {
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
        onClick={onStartEdit}
        className="px-2 py-1 text-xs whitespace-nowrap cursor-pointer border-b border-r border-gray-200 overflow-hidden text-ellipsis"
        style={{ backgroundColor: isChanged ? AMBER : rowBg }}
        title={displayValue}
      >
        {displayValue || <span className="text-gray-300">—</span>}
      </td>
    )
  }

  return (
    <td className="px-1 py-0.5 border-b border-r border-gray-200 bg-white relative" style={{ zIndex: 40 }}>
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

function DatePickerCell({ displayValue, isChanged, isEditing, rowBg, onStartEdit, onCommit, onCancel }: CellProps) {
  const [draft, setDraft] = useState(displayValue)
  const [wasEditing, setWasEditing] = useState(isEditing)
  if (isEditing !== wasEditing) {
    setWasEditing(isEditing)
    if (isEditing) setDraft(displayValue)
  }

  if (!isEditing) {
    return (
      <td
        onClick={onStartEdit}
        className="px-2 py-1 text-xs whitespace-nowrap cursor-pointer border-b border-r border-gray-200"
        style={{ backgroundColor: isChanged ? AMBER : rowBg }}
      >
        {displayValue || <span className="text-gray-300">—</span>}
      </td>
    )
  }

  return (
    <td className="px-1 py-0.5 border-b border-r border-gray-200 bg-white relative" style={{ zIndex: 40 }}>
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

  const [pendingChanges, setPendingChanges] = useState<TrackerChange[]>([])
  const [editingCell, setEditingCell] = useState<{ hop: string; field: string } | null>(null)

  const longPressRef = useRef<ReturnType<typeof setTimeout> | null>(null)

  // Load tracker snapshot — all columns, DON 444 + dedup only.
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

  const persistChanges = async (changes: TrackerChange[]) => {
    const todayKey = new Date().toISOString().slice(0, 10)
    await supabase.from('pm_updates_cache').upsert({
      id: `tracker-changes-${todayKey}`,
      updates: JSON.stringify(changes),
      updated_at: new Date().toISOString(),
    })
  }

  const saveEdit = (hop: string, field: string, originalValue: string, newValue: string) => {
    setEditingCell(null)
    if (newValue === originalValue) return
    const change: TrackerChange = {
      hop, field, oldValue: originalValue, newValue,
      timestamp: new Date().toISOString(), user: 'CJ',
    }
    setPendingChanges(prev => {
      const next = [...prev.filter(c => !(c.hop === hop && c.field === field)), change]
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
  }

  const enterEditorMode = () => {
    setDraftHidden(new Set())
    setSavePromptOpen(false)
    setViewNameDraft('')
    setEditorMode(true)
  }

  const cancelEditorMode = () => {
    setEditorMode(false)
    setSavePromptOpen(false)
    setViewNameDraft('')
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
    setViews(prev => [...prev.filter(v => v.name !== name), view])
    setActiveViewName(name)
    cancelEditorMode()
  }

  const deleteView = async (name: string) => {
    if (!confirm(`Delete view "${name}"?`)) return
    await supabase.from('pm_updates_cache').delete().eq('id', `tracker-view-${name}`)
    setViews(prev => prev.filter(v => v.name !== name))
    if (activeViewName === name) setActiveViewName('Default')
  }

  const handleTouchStart = (name: string) => {
    longPressRef.current = setTimeout(() => deleteView(name), 600)
  }
  const handleTouchEnd = () => {
    if (longPressRef.current) { clearTimeout(longPressRef.current); longPressRef.current = null }
  }

  const activeView = views.find(v => v.name === activeViewName) || null
  const effectiveHidden: Set<string> = editorMode ? draftHidden : new Set(activeView?.hiddenColumns || [])

  const hopColIdx = headers.findIndex(h => h === 'HOP')
  const orderedIndexes = headers.length > 0
    ? [hopColIdx, ...headers.map((_, i) => i).filter(i => i !== hopColIdx)]
    : []
  const allColumns = orderedIndexes.map(i => {
    const name = headers[i] || `Column ${i + 1}`
    const isHop = i === hopColIdx
    const isDate = isDateColumn(name)
    return { index: i, name, isHop, isDate, width: isHop ? 220 : (isDate ? 150 : 130) }
  })
  const visibleColumns = allColumns.filter(c => c.isHop || !effectiveHidden.has(c.name))
  const totalWidth = visibleColumns.reduce((s, c) => s + c.width, 0)

  const changeMap = new Map<string, TrackerChange>()
  pendingChanges.forEach(c => changeMap.set(`${c.hop}|${c.field}`, c))

  const viewBtnClass = (name: string) =>
    `px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors ${activeViewName === name && !editorMode ? 'text-white' : 'bg-gray-200 text-gray-700 hover:bg-gray-300'}`

  return (
    <div className="min-h-screen bg-gray-100 text-gray-900 p-4">
      <BackToDashboard />

      <div className="flex items-center justify-between mb-3 flex-wrap gap-3">
        <h1 className="text-xl font-bold">📊 Tracker Grid</h1>

        <div className="flex items-center gap-2 flex-wrap">
          <button
            onClick={() => { setActiveViewName('Default'); setEditorMode(false) }}
            style={activeViewName === 'Default' && !editorMode ? { backgroundColor: NAVY } : undefined}
            className={viewBtnClass('Default')}
          >
            Default
          </button>
          {views.map(v => (
            <button
              key={v.name}
              onClick={() => { setActiveViewName(v.name); setEditorMode(false) }}
              onContextMenu={(e) => { e.preventDefault(); deleteView(v.name) }}
              onTouchStart={() => handleTouchStart(v.name)}
              onTouchEnd={handleTouchEnd}
              style={activeViewName === v.name && !editorMode ? { backgroundColor: NAVY } : undefined}
              className={viewBtnClass(v.name)}
              title="Right-click or long-press to delete"
            >
              {v.name}
            </button>
          ))}
          <button
            onClick={enterEditorMode}
            className="px-3 py-1.5 rounded-lg text-xs font-semibold border border-dashed border-gray-400 text-gray-600 hover:bg-gray-200"
          >
            + Create View
          </button>
        </div>

        <div className="flex items-center gap-2">
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

      {editorMode && (
        <div className="mb-3 bg-blue-50 border border-blue-300 rounded-lg p-3">
          <div className="flex items-center justify-between flex-wrap gap-2 mb-2">
            <p className="text-sm text-blue-900 font-semibold">Column Editor Mode — click ✕ on any column to hide it from this view</p>
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
        <div className="border border-gray-300 rounded-lg bg-white" style={{ overflow: 'auto', height: 'calc(100vh - 230px)' }}>
          <table style={{ tableLayout: 'fixed', borderCollapse: 'collapse', width: totalWidth }}>
            <colgroup>
              {visibleColumns.map(c => <col key={c.name} style={{ width: c.width }} />)}
            </colgroup>
            <thead>
              <tr>
                {visibleColumns.map(col => (
                  <th
                    key={col.name}
                    style={{
                      position: 'sticky',
                      top: 0,
                      left: col.isHop ? 0 : undefined,
                      zIndex: col.isHop ? 30 : 20,
                      backgroundColor: NAVY,
                    }}
                    className="text-white text-xs font-bold px-2 py-2 text-left border-r border-b border-blue-900"
                  >
                    <div className="flex items-center justify-between gap-1">
                      <span className="truncate">{col.name}</span>
                      {editorMode && !col.isHop && (
                        <button
                          onClick={() => toggleDraftHidden(col.name)}
                          className="text-red-300 hover:text-red-100 font-bold text-xs flex-shrink-0"
                          title="Hide column"
                        >
                          ✕
                        </button>
                      )}
                    </div>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {trackerRows.map((row, rowIdx) => {
                const rowBg = rowIdx % 2 === 1 ? ALT_ROW : '#FFFFFF'
                return (
                  <tr key={row.hop}>
                    {visibleColumns.map(col => {
                      if (col.isHop) {
                        return (
                          <td
                            key={col.name}
                            style={{ position: 'sticky', left: 0, zIndex: 10, backgroundColor: rowBg, color: NAVY }}
                            className="px-2 py-1 text-xs font-bold whitespace-nowrap border-r border-b border-gray-200"
                          >
                            {row.hop}
                          </td>
                        )
                      }
                      const raw = row.cells[col.index]
                      const { text: displayValue, treatAsDate } = cellDisplayValue(raw, col.isDate)
                      const change = changeMap.get(`${row.hop}|${col.name}`)
                      const shown = change ? change.newValue : displayValue
                      const isEditing = editingCell?.hop === row.hop && editingCell?.field === col.name
                      const cellProps: CellProps = {
                        displayValue: shown,
                        isChanged: !!change,
                        isEditing,
                        rowBg,
                        onStartEdit: () => setEditingCell({ hop: row.hop, field: col.name }),
                        onCommit: (newValue: string) => saveEdit(row.hop, col.name, displayValue, newValue),
                        onCancel: () => setEditingCell(null),
                      }
                      return treatAsDate
                        ? <DatePickerCell key={col.name} {...cellProps} />
                        : <EditableCell key={col.name} {...cellProps} />
                    })}
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
