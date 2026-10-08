'use client'

import { useEffect, useRef, useState } from 'react'

// Excel-style filter for one table column: sort, search, and a checkbox per
// distinct value. Shared by the NTP comments tab and GC Clean-Up — built
// first for NTP, extracted here so a second table didn't duplicate it.
export default function ColumnFilterMenu({ values, selected, sort, onSort, onApply, onClose }: {
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
