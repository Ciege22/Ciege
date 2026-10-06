'use client'

import { useEffect, useMemo, useRef, useState } from 'react'

// NTP comments for Deck Builder V2. Nothing saves until you press Save (per
// row or Save all); the server reads back what it wrote before confirming.
// Every build reads the saved comments into the NTP slides, and the Excel
// export is only for the customer after the call.

const RAILWAY = 'https://ciege-production.up.railway.app'

type NtpRow = {
  hop: string
  path_id: string
  category: 'External' | 'Other' | 'Program Team'
  owner: string
  gc: string
  fc_start: string
  fc_end: string
  blocker: string
  comment: string
  status: string
}

type NtpMonth = { sheet: string; label: string; rows: NtpRow[] }

type NtpPayload = { deck_date: string; statuses: string[]; months: NtpMonth[] }

type Edit = { sheet: string; hop: string; comment?: string; status?: string }

type ColKey = 'hop' | 'category' | 'owner' | 'gc' | 'fc_start' | 'fc_end' | 'blocker' | 'status' | 'comment'

const COLUMNS: { key: ColKey; label: string; width: string }[] = [
  { key: 'hop', label: 'HOP', width: 'min-w-[18rem]' },
  { key: 'category', label: 'Category', width: 'min-w-[8rem]' },
  { key: 'owner', label: 'Owner', width: 'min-w-[7rem]' },
  { key: 'gc', label: 'GC', width: 'min-w-[9rem]' },
  { key: 'fc_start', label: 'FC Start', width: 'min-w-[6rem]' },
  { key: 'fc_end', label: 'FC End', width: 'min-w-[6rem]' },
  { key: 'blocker', label: 'Blocker / waiting on', width: 'min-w-[16rem]' },
  { key: 'status', label: 'Status', width: 'min-w-[9rem]' },
  { key: 'comment', label: 'Comment', width: 'min-w-[30rem]' },
]

const BLANK = '(Blanks)'

const CATEGORY_STYLES: Record<string, string> = {
  External: 'text-red-300',
  Other: 'text-orange-300',
  'Program Team': 'text-sky-300',
}

function cellValue(r: NtpRow, key: ColKey): string {
  const v = r[key] ?? ''
  return v.trim() === '' ? BLANK : v
}

export default function NtpCommentsV2() {
  const [data, setData] = useState<NtpPayload | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [loading, setLoading] = useState(true)
  const [activeSheet, setActiveSheet] = useState<string>('')
  const [search, setSearch] = useState('')
  const [filters, setFilters] = useState<Partial<Record<ColKey, Set<string>>>>({})
  const [sort, setSort] = useState<{ key: ColKey; dir: 'asc' | 'desc' } | null>(null)
  const [openCol, setOpenCol] = useState<ColKey | null>(null)
  // Edits not yet saved, keyed by sheet + HOP. Only these are sent on save.
  const [dirty, setDirty] = useState<Record<string, Edit>>({})
  const [saving, setSaving] = useState(false)
  const [saveError, setSaveError] = useState<string | null>(null)
  const [savedAt, setSavedAt] = useState<string | null>(null)

  const dirtyCount = Object.keys(dirty).length

  // Warn before leaving the page with unsaved comments.
  useEffect(() => {
    if (dirtyCount === 0) return
    const warn = (e: BeforeUnloadEvent) => { e.preventDefault(); e.returnValue = '' }
    window.addEventListener('beforeunload', warn)
    return () => window.removeEventListener('beforeunload', warn)
  }, [dirtyCount])

  useEffect(() => {
    let cancelled = false
    fetch(`${RAILWAY}/ntp_comments`)
      .then(async res => {
        const json = await res.json()
        if (!res.ok) throw new Error(json.error || `Server error: ${res.status}`)
        return json as NtpPayload
      })
      .then(json => {
        if (cancelled) return
        setData(json)
        const firstWithRows = json.months.find(m => m.rows.length > 0) ?? json.months[0]
        setActiveSheet(firstWithRows?.sheet ?? '')
      })
      .catch(err => { if (!cancelled) setError(err instanceof Error ? err.message : 'Could not load NTP comments.') })
      .finally(() => { if (!cancelled) setLoading(false) })
    return () => { cancelled = true }
  }, [])

  const month = data?.months.find(m => m.sheet === activeSheet)

  // Distinct values per column for this month — the checklist in each filter.
  const valuesByCol = useMemo(() => {
    const out = {} as Record<ColKey, string[]>
    COLUMNS.forEach(c => {
      const set = new Set<string>()
      month?.rows.forEach(r => set.add(cellValue(r, c.key)))
      out[c.key] = Array.from(set).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
    })
    return out
  }, [month])

  const visibleRows = useMemo(() => {
    if (!month) return []
    const q = search.trim().toLowerCase()
    const rows = month.rows.filter(r => {
      for (const key of Object.keys(filters) as ColKey[]) {
        const allowed = filters[key]
        if (allowed && !allowed.has(cellValue(r, key))) return false
      }
      return !q || [r.hop, r.path_id, r.owner, r.gc, r.blocker, r.comment].some(v => v.toLowerCase().includes(q))
    })
    if (sort) {
      const dir = sort.dir === 'asc' ? 1 : -1
      rows.sort((a, b) => cellValue(a, sort.key).localeCompare(cellValue(b, sort.key), undefined, { numeric: true }) * dir)
    }
    return rows
  }, [month, filters, search, sort])

  const activeFilterCount = Object.keys(filters).length

  // Edits show on screen and are marked unsaved until the server confirms them.
  function edit(sheet: string, hop: string, patch: { comment?: string; status?: string }) {
    setData(prev => prev && {
      ...prev,
      months: prev.months.map(m => m.sheet !== sheet ? m : {
        ...m, rows: m.rows.map(r => r.hop === hop ? { ...r, ...patch } : r),
      }),
    })
    setDirty(prev => ({ ...prev, [`${sheet}|${hop}`]: { ...(prev[`${sheet}|${hop}`] ?? { sheet, hop }), ...patch } }))
    setSavedAt(null)
  }

  // Saves the given unsaved edits (all of them by default). One request per month
  // sheet; a failed sheet stays marked unsaved so nothing is silently dropped.
  async function saveEdits(keys?: string[]) {
    const pending = Object.entries(dirty).filter(([k]) => !keys || keys.includes(k))
    if (pending.length === 0) return
    setSaving(true)
    setSaveError(null)
    const bySheet = new Map<string, { entries: Edit[]; keys: string[] }>()
    pending.forEach(([k, d]) => {
      const group = bySheet.get(d.sheet) ?? { entries: [], keys: [] }
      group.entries.push({ sheet: d.sheet, hop: d.hop, comment: d.comment, status: d.status })
      group.keys.push(k)
      bySheet.set(d.sheet, group)
    })
    const failures: string[] = []
    for (const [sheet, group] of bySheet) {
      try {
        const res = await fetch(`${RAILWAY}/ntp_comments/save`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ sheet, entries: group.entries.map(({ hop, comment, status }) => ({ hop, comment, status })) }),
        })
        const json = await res.json().catch(() => ({}))
        if (!res.ok) throw new Error(json.error || `Save failed: ${res.status}`)
        setDirty(prev => {
          const next = { ...prev }
          group.keys.forEach(k => delete next[k])
          return next
        })
      } catch (err) {
        failures.push(`${sheet}: ${err instanceof Error ? err.message : 'save failed'}`)
      }
    }
    setSaving(false)
    if (failures.length) setSaveError(`Not saved — ${failures.join(' · ')}. Your edits are still marked unsaved; press Save again.`)
    else setSavedAt(new Date().toLocaleTimeString([], { hour: 'numeric', minute: '2-digit', second: '2-digit' }))
  }

  function setColumnFilter(key: ColKey, selected: Set<string> | null) {
    setFilters(prev => {
      const next = { ...prev }
      if (selected === null || selected.size === valuesByCol[key].length) delete next[key]
      else next[key] = selected
      return next
    })
  }

  function exportUrl() {
    const date = data?.deck_date ?? ''
    return `${RAILWAY}/ntp_comments/export?deck_date=${encodeURIComponent(date)}`
  }

  function confirmExport(e: React.MouseEvent<HTMLAnchorElement>) {
    if (dirtyCount > 0 && !window.confirm(`${dirtyCount} comment(s) are not saved yet, so the export will leave them out. Export anyway?`)) {
      e.preventDefault()
    }
  }

  if (loading) return <p className="text-sm text-zinc-400">Loading NTP pending HOPs…</p>
  if (error) return <p className="text-sm text-red-300">{error}</p>
  if (!data) return null

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center justify-between gap-3">
        <div>
          <p className="text-xs uppercase tracking-[0.3em] text-zinc-400">NTP comments</p>
          <p className="text-sm text-zinc-500">Call date {data.deck_date} · press Save to store changes · shows on the NTP slides in the next build</p>
        </div>
        <a
          href={exportUrl()} onClick={confirmExport}
          className="inline-flex items-center rounded-2xl bg-zinc-100 px-4 py-2 text-sm font-semibold text-zinc-950 hover:bg-white"
        >
          Export for customer (.xlsx)
        </a>
      </div>

      <div className="flex flex-wrap gap-2">
        {data.months.map(m => (
          <button
            key={m.sheet}
            type="button"
            onClick={() => { setActiveSheet(m.sheet); setFilters({}); setSort(null); setOpenCol(null) }}
            className={`rounded-xl px-3 py-1.5 text-sm ${m.sheet === activeSheet ? 'bg-emerald-500 text-zinc-950 font-semibold' : 'bg-zinc-800 text-zinc-300 hover:bg-zinc-700'}`}
          >
            {m.label} <span className="opacity-70">({m.rows.length})</span>
          </button>
        ))}
      </div>

      {month && (
        <>
          <div className="flex flex-wrap items-center gap-3">
            <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search HOP, owner, GC, blocker, comment"
              className="min-w-[18rem] flex-1 rounded-xl border border-white/10 bg-zinc-900 px-3 py-2 text-sm text-zinc-100 placeholder:text-zinc-600" />
            <span className="text-sm text-zinc-400">{visibleRows.length} of {month.rows.length} HOPs</span>
            {(activeFilterCount > 0 || sort) && (
              <button type="button" onClick={() => { setFilters({}); setSort(null) }}
                className="rounded-xl bg-zinc-800 px-3 py-2 text-sm text-zinc-200 hover:bg-zinc-700">
                Clear filters & sort
              </button>
            )}
            <button type="button" onClick={() => saveEdits()} disabled={saving || dirtyCount === 0}
              className="rounded-xl bg-emerald-500 px-4 py-2 text-sm font-semibold text-zinc-950 hover:bg-emerald-400 disabled:cursor-not-allowed disabled:opacity-40">
              {saving ? 'Saving…' : `Save all${dirtyCount ? ` (${dirtyCount})` : ''}`}
            </button>
            <span className="text-sm">
              {dirtyCount > 0
                ? <span className="text-amber-300">{dirtyCount} unsaved</span>
                : savedAt ? <span className="text-emerald-300">All saved · {savedAt}</span> : null}
            </span>
          </div>

          {saveError && (
            <p className="rounded-xl border border-red-500/30 bg-red-500/10 px-3 py-2 text-sm text-red-300">{saveError}</p>
          )}

          {visibleRows.length === 0 ? (
            <p className="text-sm text-zinc-500">No HOPs match these filters.</p>
          ) : (
            <div className="overflow-x-auto rounded-2xl border border-white/10">
              <table className="w-full text-sm text-left">
                <thead className="sticky top-0 z-10 bg-zinc-900 text-zinc-300">
                  <tr>
                    {COLUMNS.map(c => {
                      const active = !!filters[c.key]
                      const sorted = sort?.key === c.key
                      return (
                        <th key={c.key} className={`relative px-3 py-2 align-top whitespace-nowrap ${c.width}`}>
                          <button type="button" onClick={() => setOpenCol(openCol === c.key ? null : c.key)}
                            className="inline-flex items-center gap-1.5 font-semibold hover:text-white">
                            {c.label}
                            <span className={active || sorted ? 'text-emerald-400' : 'text-zinc-500'}>
                              {sorted ? (sort?.dir === 'asc' ? '▲' : '▼') : ''}{active ? ' ⏷' : ' ▾'}
                            </span>
                          </button>
                          {openCol === c.key && (
                            <ColumnFilterMenu
                              values={valuesByCol[c.key]}
                              selected={filters[c.key] ?? null}
                              sort={sorted ? sort!.dir : null}
                              onSort={dir => setSort(dir ? { key: c.key, dir } : null)}
                              onApply={sel => setColumnFilter(c.key, sel)}
                              onClose={() => setOpenCol(null)}
                            />
                          )}
                        </th>
                      )
                    })}
                    <th className="px-3 py-2 align-top whitespace-nowrap">Save</th>
                  </tr>
                </thead>
                <tbody>
                  {visibleRows.map(r => {
                    const key = `${month.sheet}|${r.hop}`
                    const isDirty = !!dirty[key]
                    return (
                      <tr key={r.hop} className={`border-t border-white/5 align-top hover:bg-white/[0.02] ${isDirty ? 'bg-amber-500/[0.06]' : ''}`}>
                        <td className="px-3 py-2 text-zinc-100">
                          <div className="font-medium">{r.hop}</div>
                          <div className="text-xs text-zinc-500">{r.path_id}</div>
                        </td>
                        <td className={`px-3 py-2 ${CATEGORY_STYLES[r.category] ?? 'text-zinc-300'}`}>{r.category}</td>
                        <td className="px-3 py-2 text-zinc-300">{r.owner}</td>
                        <td className="px-3 py-2 text-zinc-300">{r.gc}</td>
                        <td className="px-3 py-2 text-zinc-300 whitespace-nowrap">{r.fc_start}</td>
                        <td className="px-3 py-2 text-zinc-300 whitespace-nowrap">{r.fc_end}</td>
                        <td className="px-3 py-2 text-zinc-400">{r.blocker}</td>
                        <td className="px-3 py-2">
                          <select
                            value={r.status}
                            onChange={e => edit(month.sheet, r.hop, { status: e.target.value })}
                            className="rounded-lg border border-white/10 bg-zinc-900 px-2 py-1.5 text-zinc-100"
                          >
                            {data.statuses.map(s => <option key={s}>{s}</option>)}
                          </select>
                        </td>
                        <td className="px-3 py-2">
                          <textarea
                            rows={3}
                            value={r.comment}
                            onChange={e => edit(month.sheet, r.hop, { comment: e.target.value })}
                            placeholder="Add the call update, e.g. 10/13/2026: …"
                            className="w-full min-w-[28rem] rounded-lg border border-white/10 bg-zinc-900 px-2 py-1.5 text-zinc-100 placeholder:text-zinc-600"
                          />
                        </td>
                        <td className="px-3 py-2 whitespace-nowrap">
                          <button type="button" onClick={() => saveEdits([key])} disabled={!isDirty || saving}
                            className="rounded-lg bg-zinc-100 px-3 py-1.5 text-xs font-semibold text-zinc-950 hover:bg-white disabled:cursor-not-allowed disabled:opacity-30">
                            Save
                          </button>
                          {isDirty && <div className="mt-1 text-[11px] text-amber-300">unsaved</div>}
                        </td>
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            </div>
          )}
        </>
      )}
    </div>
  )
}

// Excel-style filter for one column: sort, search, and a checkbox per value.
function ColumnFilterMenu({ values, selected, sort, onSort, onApply, onClose }: {
  values: string[]
  selected: Set<string> | null
  sort: 'asc' | 'desc' | null
  onSort: (dir: 'asc' | 'desc' | null) => void
  onApply: (sel: Set<string> | null) => void
  onClose: () => void
}) {
  const [query, setQuery] = useState('')
  const [checked, setChecked] = useState<Set<string>>(() => new Set(selected ?? values))
  const ref = useRef<HTMLDivElement>(null)

  useEffect(() => {
    function onDown(e: MouseEvent) {
      if (ref.current && !ref.current.contains(e.target as Node)) onClose()
    }
    document.addEventListener('mousedown', onDown)
    return () => document.removeEventListener('mousedown', onDown)
  }, [onClose])

  const shown = values.filter(v => v.toLowerCase().includes(query.trim().toLowerCase()))

  function toggle(v: string) {
    setChecked(prev => {
      const next = new Set(prev)
      if (next.has(v)) next.delete(v); else next.add(v)
      return next
    })
  }

  return (
    <div ref={ref} className="absolute left-0 top-full z-20 mt-1 w-72 rounded-xl border border-white/10 bg-zinc-900 p-3 text-left font-normal shadow-2xl">
      <div className="mb-2 flex gap-2">
        <button type="button" onClick={() => onSort(sort === 'asc' ? null : 'asc')}
          className={`flex-1 rounded-lg px-2 py-1 text-xs ${sort === 'asc' ? 'bg-emerald-500 text-zinc-950' : 'bg-zinc-800 text-zinc-200'}`}>A → Z</button>
        <button type="button" onClick={() => onSort(sort === 'desc' ? null : 'desc')}
          className={`flex-1 rounded-lg px-2 py-1 text-xs ${sort === 'desc' ? 'bg-emerald-500 text-zinc-950' : 'bg-zinc-800 text-zinc-200'}`}>Z → A</button>
      </div>
      <input value={query} onChange={e => setQuery(e.target.value)} placeholder="Search values"
        className="mb-2 w-full rounded-lg border border-white/10 bg-zinc-950 px-2 py-1 text-xs text-zinc-100 placeholder:text-zinc-600" />
      <div className="mb-2 flex gap-2 text-xs">
        <button type="button" onClick={() => setChecked(new Set(values))} className="text-emerald-300 hover:underline">Select all</button>
        <button type="button" onClick={() => setChecked(new Set())} className="text-emerald-300 hover:underline">Clear</button>
      </div>
      <div className="max-h-64 overflow-y-auto space-y-1 pr-1">
        {shown.map(v => (
          <label key={v} className="flex items-start gap-2 text-xs text-zinc-200">
            <input type="checkbox" className="mt-0.5" checked={checked.has(v)} onChange={() => toggle(v)} />
            <span className="break-words">{v}</span>
          </label>
        ))}
        {shown.length === 0 && <p className="text-xs text-zinc-500">No matching values.</p>}
      </div>
      <div className="mt-3 flex justify-end gap-2">
        <button type="button" onClick={onClose} className="rounded-lg px-3 py-1 text-xs text-zinc-300 hover:bg-zinc-800">Cancel</button>
        <button type="button" onClick={() => { onApply(checked.size === values.length ? null : checked); onClose() }}
          className="rounded-lg bg-emerald-500 px-3 py-1 text-xs font-semibold text-zinc-950 hover:bg-emerald-400">Apply</button>
      </div>
    </div>
  )
}
