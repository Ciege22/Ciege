'use client'

import { useEffect, useMemo, useState } from 'react'

// NTP comments for Deck Builder V2. Comments save to Supabase as you type
// (on blur), every build reads them back into the NTP slides, and the Excel
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

const CATEGORY_STYLES: Record<string, string> = {
  External: 'text-red-300',
  Other: 'text-orange-300',
  'Program Team': 'text-sky-300',
}

export default function NtpCommentsV2() {
  const [data, setData] = useState<NtpPayload | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [loading, setLoading] = useState(true)
  const [activeSheet, setActiveSheet] = useState<string>('')
  const [category, setCategory] = useState<string>('All')
  const [gc, setGc] = useState<string>('All')
  const [search, setSearch] = useState('')
  const [saving, setSaving] = useState<string | null>(null)
  const [saveError, setSaveError] = useState<string | null>(null)

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

  const gcOptions = useMemo(() => {
    const set = new Set<string>()
    month?.rows.forEach(r => { if (r.gc) set.add(r.gc) })
    return Array.from(set).sort()
  }, [month])

  const visibleRows = useMemo(() => {
    if (!month) return []
    const q = search.trim().toLowerCase()
    return month.rows.filter(r =>
      (category === 'All' || r.category === category) &&
      (gc === 'All' || r.gc === gc) &&
      (!q || [r.hop, r.path_id, r.owner, r.blocker, r.comment].some(v => v.toLowerCase().includes(q)))
    )
  }, [month, category, gc, search])

  // Local edit first so typing never fights a slow save; the server copy wins on reload.
  function applyLocal(sheet: string, hop: string, patch: Partial<Pick<NtpRow, 'comment' | 'status'>>) {
    setData(prev => prev && {
      ...prev,
      months: prev.months.map(m => m.sheet !== sheet ? m : {
        ...m, rows: m.rows.map(r => r.hop === hop ? { ...r, ...patch } : r),
      }),
    })
  }

  async function save(sheet: string, hop: string, patch: Partial<Pick<NtpRow, 'comment' | 'status'>>) {
    setSaving(hop)
    setSaveError(null)
    try {
      const res = await fetch(`${RAILWAY}/ntp_comments`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ sheet, hop, ...patch }),
      })
      if (!res.ok) {
        const json = await res.json().catch(() => ({}))
        throw new Error(json.error || `Save failed: ${res.status}`)
      }
    } catch (err) {
      setSaveError(err instanceof Error ? err.message : 'Save failed.')
    } finally {
      setSaving(null)
    }
  }

  function exportUrl() {
    const date = data?.deck_date ?? ''
    return `${RAILWAY}/ntp_comments/export?deck_date=${encodeURIComponent(date)}`
  }

  if (loading) return <p className="text-sm text-zinc-400">Loading NTP pending HOPs…</p>
  if (error) return <p className="text-sm text-red-300">{error}</p>
  if (!data) return null

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center justify-between gap-3">
        <div>
          <p className="text-xs uppercase tracking-[0.3em] text-zinc-400">NTP comments</p>
          <p className="text-sm text-zinc-500">Call date {data.deck_date} · saved as you type · shows on the NTP slides in the next build</p>
        </div>
        <a
          href={exportUrl()}
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
            onClick={() => { setActiveSheet(m.sheet); setGc('All') }}
            className={`rounded-xl px-3 py-1.5 text-sm ${m.sheet === activeSheet ? 'bg-emerald-500 text-zinc-950 font-semibold' : 'bg-zinc-800 text-zinc-300 hover:bg-zinc-700'}`}
          >
            {m.label} <span className="opacity-70">({m.rows.length})</span>
          </button>
        ))}
      </div>

      {month && (
        <>
          <div className="flex flex-wrap gap-3">
            <select value={category} onChange={e => setCategory(e.target.value)}
              className="rounded-xl border border-white/10 bg-zinc-900 px-3 py-2 text-sm text-zinc-100">
              {['All', 'External', 'Other', 'Program Team'].map(c => <option key={c}>{c}</option>)}
            </select>
            <select value={gc} onChange={e => setGc(e.target.value)}
              className="rounded-xl border border-white/10 bg-zinc-900 px-3 py-2 text-sm text-zinc-100">
              <option>All</option>
              {gcOptions.map(g => <option key={g}>{g}</option>)}
            </select>
            <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search HOP, owner, blocker, comment"
              className="min-w-[16rem] flex-1 rounded-xl border border-white/10 bg-zinc-900 px-3 py-2 text-sm text-zinc-100 placeholder:text-zinc-600" />
          </div>

          {saveError && <p className="text-sm text-red-300">{saveError}</p>}

          {visibleRows.length === 0 ? (
            <p className="text-sm text-zinc-500">No HOPs match these filters.</p>
          ) : (
            <div className="overflow-x-auto rounded-2xl border border-white/10">
              <table className="w-full min-w-[1100px] text-left text-xs">
                <thead className="bg-zinc-900 text-zinc-400">
                  <tr>
                    <th className="px-3 py-2">HOP</th>
                    <th className="px-3 py-2">Category</th>
                    <th className="px-3 py-2">Owner</th>
                    <th className="px-3 py-2">GC</th>
                    <th className="px-3 py-2">FC Start</th>
                    <th className="px-3 py-2">FC End</th>
                    <th className="px-3 py-2">Blocker / waiting on</th>
                    <th className="px-3 py-2">Status</th>
                    <th className="px-3 py-2 w-[28rem]">Comment</th>
                  </tr>
                </thead>
                <tbody>
                  {visibleRows.map(r => (
                    <tr key={r.hop} className="border-t border-white/5 align-top">
                      <td className="px-3 py-2 text-zinc-100">
                        <div className="font-medium">{r.hop}</div>
                        <div className="text-zinc-500">{r.path_id}</div>
                      </td>
                      <td className={`px-3 py-2 ${CATEGORY_STYLES[r.category] ?? 'text-zinc-300'}`}>{r.category}</td>
                      <td className="px-3 py-2 text-zinc-300">{r.owner}</td>
                      <td className="px-3 py-2 text-zinc-300">{r.gc}</td>
                      <td className="px-3 py-2 text-zinc-300">{r.fc_start}</td>
                      <td className="px-3 py-2 text-zinc-300">{r.fc_end}</td>
                      <td className="px-3 py-2 text-zinc-400 max-w-[18rem]">{r.blocker}</td>
                      <td className="px-3 py-2">
                        <select
                          value={r.status}
                          disabled={saving === r.hop}
                          onChange={e => {
                            const status = e.target.value
                            applyLocal(month.sheet, r.hop, { status })
                            save(month.sheet, r.hop, { status })
                          }}
                          className="rounded-lg border border-white/10 bg-zinc-900 px-2 py-1 text-zinc-100"
                        >
                          {data.statuses.map(s => <option key={s}>{s}</option>)}
                        </select>
                      </td>
                      <td className="px-3 py-2">
                        <textarea
                          rows={3}
                          value={r.comment}
                          onChange={e => applyLocal(month.sheet, r.hop, { comment: e.target.value })}
                          onBlur={e => save(month.sheet, r.hop, { comment: e.target.value })}
                          placeholder="Add the call update, e.g. 10/13/2026: …"
                          className="w-full rounded-lg border border-white/10 bg-zinc-900 px-2 py-1 text-zinc-100 placeholder:text-zinc-600"
                        />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </>
      )}
    </div>
  )
}
