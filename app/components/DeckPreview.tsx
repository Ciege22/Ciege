'use client'

import { useEffect, useState } from 'react'
import { supabase } from '../lib/supabase'

// Deck Builder V2 preview — steps through the last two V2 builds slide by slide
// (structured view of each slide's text, tables, and chart data), with a
// comment box under every slide. Comments are saved to the database keyed by
// build, so they can be reviewed later or handed off for corrections.

const BACKEND = 'https://ciege-production.up.railway.app'
const COMMENTS_ID = 'deck-slide-comments'

interface Chart { categories: string[]; series: { name: string; values: (number | null)[] }[] }
interface Slide { index: number; texts: string[]; tables: string[][][]; charts: Chart[] }
interface Build { which: string; filename: string; uploaded_at: string; deck_date: string | null; slides: Slide[] }
type Comments = Record<string, Record<string, string>>

async function fetchBuild(which: 'latest' | 'prior'): Promise<Build | null> {
  try {
    const res = await fetch(`${BACKEND}/v2_slides?which=${which}`)
    if (!res.ok) return null
    return await res.json()
  } catch {
    return null
  }
}

function SlideView({ slide }: { slide: Slide }) {
  const [title, ...rest] = slide.texts
  return (
    <div className="rounded-xl border border-white/10 bg-white text-zinc-900 p-5 min-h-[260px]">
      <div className="text-lg font-bold text-blue-900">{title || `Slide ${slide.index + 1}`}</div>
      {rest.length > 0 && (
        <div className="mt-2 space-y-1 text-sm">
          {rest.map((t, i) => <div key={i} className="whitespace-pre-wrap">{t}</div>)}
        </div>
      )}
      {slide.tables.map((tbl, ti) => (
        <div key={ti} className="mt-3 overflow-x-auto">
          <table className="w-full text-xs border-collapse">
            <tbody>
              {tbl.map((row, ri) => (
                <tr key={ri} className={ri === 0 ? 'bg-blue-900 text-white font-semibold' : (ri % 2 ? 'bg-zinc-100' : '')}>
                  {row.map((cell, ci) => <td key={ci} className="px-2 py-1 border border-zinc-200">{cell}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ))}
      {slide.charts.map((ch, ci) => (
        <div key={ci} className="mt-3 text-xs">
          <div className="font-semibold text-zinc-600">Chart — {ch.categories.join(' · ')}</div>
          {ch.series.map((s, si) => (
            <div key={si} className="text-zinc-700">
              {s.name}: {s.values.map(v => (v === null ? '—' : v)).join(', ')}
            </div>
          ))}
        </div>
      ))}
    </div>
  )
}

export default function DeckPreview() {
  const [latest, setLatest] = useState<Build | null>(null)
  const [prior, setPrior] = useState<Build | null>(null)
  const [loading, setLoading] = useState(true)
  const [index, setIndex] = useState(0)
  const [compare, setCompare] = useState(false)
  const [comments, setComments] = useState<Comments>({})
  const [edits, setEdits] = useState<Record<number, string>>({})
  const [saved, setSaved] = useState<string | null>(null)

  useEffect(() => {
    Promise.all([fetchBuild('latest'), fetchBuild('prior')]).then(([l, p]) => {
      setLatest(l)
      setPrior(p)
      setLoading(false)
    })
    supabase.from('pm_updates_cache').select('updates').eq('id', COMMENTS_ID).single().then(({ data }) => {
      if (data?.updates) {
        try { setComments(JSON.parse(data.updates)) } catch { /* ignore */ }
      }
    })
  }, [])

  const slides = latest?.slides ?? []
  const current = slides[index]
  const buildKey = latest?.uploaded_at ?? ''
  const existing = comments[buildKey]?.[String(index)] ?? ''

  const draft = edits[index] ?? existing
  const setDraft = (text: string) => setEdits(prev => ({ ...prev, [index]: text }))

  const saveComment = async () => {
    if (!buildKey) return
    setEdits(prev => { const n = { ...prev }; delete n[index]; return n })
    const next: Comments = {
      ...comments,
      [buildKey]: { ...(comments[buildKey] ?? {}), [String(index)]: draft.trim() },
    }
    setComments(next)
    const { error } = await supabase.from('pm_updates_cache').upsert({
      id: COMMENTS_ID, updates: JSON.stringify(next), updated_at: new Date().toISOString(),
    })
    setSaved(error ? 'Save failed' : 'Saved')
  }

  const priorSlide = prior?.slides[index]

  return (
    <div className="min-h-screen bg-zinc-950 text-zinc-100 p-4">
      <h2 className="text-xl font-semibold text-white mb-1">Slide preview &amp; comments</h2>
      <p className="text-sm text-zinc-400 mb-4">
        {latest ? `Latest build: ${latest.deck_date ?? ''} (saved ${new Date(latest.uploaded_at).toLocaleString()})` : 'No V2 build saved yet — build one on Deck Builder V2.'}
        {prior && ` · Prior: ${prior.deck_date ?? ''}`}
      </p>

      {loading && <div className="text-zinc-400 text-sm">Loading builds…</div>}

      {!loading && slides.length > 0 && (
        <div className="grid gap-4 lg:grid-cols-[220px_minmax(0,1fr)]">
          <aside className="rounded-xl border border-white/10 bg-zinc-900 p-2 max-h-[80vh] overflow-y-auto">
            {slides.map(s => (
              <button key={s.index} onClick={() => setIndex(s.index)}
                className={`block w-full text-left text-xs px-2 py-1.5 rounded ${s.index === index ? 'bg-blue-600 text-white' : 'text-zinc-300 hover:bg-white/5'}`}>
                {s.index + 1}. {s.texts[0] || '(no title)'}
              </button>
            ))}
          </aside>

          <section className="space-y-3">
            <div className="flex flex-wrap items-center gap-2">
              <button onClick={() => setIndex(Math.max(0, index - 1))} disabled={index === 0}
                className="bg-zinc-800 disabled:opacity-40 text-sm rounded px-3 py-1.5">← Prev</button>
              <span className="text-sm text-zinc-400">Slide {index + 1} of {slides.length}</span>
              <button onClick={() => setIndex(Math.min(slides.length - 1, index + 1))} disabled={index === slides.length - 1}
                className="bg-zinc-800 disabled:opacity-40 text-sm rounded px-3 py-1.5">Next →</button>
              <label className="ml-auto flex items-center gap-2 text-sm text-zinc-300">
                <input type="checkbox" checked={compare} onChange={e => setCompare(e.target.checked)} disabled={!prior} />
                Compare with prior build
              </label>
            </div>

            <div className={compare && priorSlide ? 'grid gap-3 md:grid-cols-2' : ''}>
              {compare && priorSlide && (
                <div>
                  <div className="text-xs text-zinc-500 mb-1">Prior build {prior?.deck_date ?? ''}</div>
                  <SlideView slide={priorSlide} />
                </div>
              )}
              {current && (
                <div>
                  {compare && priorSlide && <div className="text-xs text-zinc-500 mb-1">Latest build {latest?.deck_date ?? ''}</div>}
                  <SlideView slide={current} />
                </div>
              )}
            </div>

            <div className="rounded-xl border border-white/10 bg-zinc-900 p-3">
              <div className="text-xs text-zinc-400 mb-2">Comments on this slide (saved with this build)</div>
              <textarea value={draft} onChange={e => setDraft(e.target.value)} rows={3}
                placeholder="What's wrong or needs correcting on this slide?"
                className="w-full bg-zinc-800 text-sm rounded px-3 py-2 border border-white/10" />
              <div className="mt-2 flex items-center gap-3">
                <button onClick={saveComment} disabled={draft.trim() === existing.trim()}
                  className="bg-blue-600 hover:bg-blue-500 disabled:opacity-40 text-white text-sm font-semibold rounded px-4 py-1.5">Save comment</button>
                {saved && <span className="text-xs text-zinc-400">{saved}</span>}
              </div>
            </div>
          </section>
        </div>
      )}
    </div>
  )
}
