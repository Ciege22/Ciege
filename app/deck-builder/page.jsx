'use client'

import { useRef, useState, useEffect } from 'react'
import { loadTrackerSnapshot, getPreviousSnapshot } from '../lib/supabase'

export default function DeckBuilderPage() {
  const [loading, setLoading] = useState(false)
  const [success, setSuccess] = useState(false)
  const [error, setError] = useState(null)
  const [fileNames, setFileNames] = useState({})

  const [trackerLoaded, setTrackerLoaded] = useState(false)
  const [trackerInfo, setTrackerInfo] = useState(null)
  const [prevSnapInfo, setPrevSnapInfo] = useState(null)
  const [trackerRows, setTrackerRows] = useState(null)
  const [prevSnapRows, setPrevSnapRows] = useState(null)
  const [prevSnapDate, setPrevSnapDate] = useState(null)

  const fileRefs = useRef({})
  const dateRef = useRef(null)

  useEffect(() => {
    const loadFromSupabase = async () => {
      try {
        const [latest, previous] = await Promise.all([
          loadTrackerSnapshot(),
          getPreviousSnapshot()
        ])
        if (latest) {
          // Filter to only DON 444 rows before storing to reduce payload size
          const rows = latest.data
          let headerIdx = -1
          for (let i = 0; i < 10; i++) {
            if (rows[i]?.some(c => String(c).trim() === 'HOP')) { headerIdx = i; break }
          }
          if (headerIdx >= 0) {
            const headers = rows[headerIdx]
            const don444Col = headers.findIndex(h => String(h).trim() === 'DON 444')
            const filtered = [
              rows[headerIdx],
              ...rows.slice(headerIdx + 1).filter(row =>
                String(row[don444Col] || '').trim().toUpperCase() === 'DON 444'
              )
            ]
            setTrackerRows(filtered)
          } else {
            setTrackerRows(latest.data)
          }
          setTrackerInfo({ filename: latest.filename, uploaded_at: latest.uploaded_at, hop_count: latest.hop_count })
          setTrackerLoaded(true)
        }
        if (previous) {
          setPrevSnapRows(previous.data)
          setPrevSnapDate(previous.uploaded_at)
          setPrevSnapInfo({ filename: previous.filename, uploaded_at: previous.uploaded_at })
        }
      } catch (err) {
        console.error('Error loading from Supabase:', err)
      }
    }
    loadFromSupabase()
  }, [])

  function handleFileChange(id, e) {
    const file = e.target.files?.[0]
    setFileNames((prev) => ({ ...prev, [id]: file ? file.name : null }))
    setSuccess(false)
    setError(null)
  }

  function fmtDate(iso) {
    if (!iso) return ''
    const d = new Date(iso)
    return d.toLocaleDateString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric' }) +
      ' at ' + d.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })
  }

  async function handleSubmit(e) {
    e.preventDefault()
    setError(null)
    setSuccess(false)

    const formData = new FormData()

    // Previous deck (still a file upload)
    const prevDeckInput = document.getElementById('previous_deck')
    if (prevDeckInput?.files?.[0]) {
      formData.append('previous_deck', prevDeckInput.files[0])
    } else {
      setError('Previous deck (.pptx) is required.')
      return
    }

    // NTP comments (optional file upload)
    const ntpInput = document.getElementById('ntp_comments')
    if (ntpInput?.files?.[0]) {
      formData.append('ntp_comments', ntpInput.files[0])
    }

    // Tracker and snapshot from Supabase as JSON
    if (trackerRows) {
      formData.append('tracker_json', JSON.stringify(trackerRows))
    }
    if (prevSnapRows) {
      formData.append('prev_snapshot_json', JSON.stringify({
        rows: prevSnapRows,
        uploaded_at: prevSnapDate
      }))
    }

    const deckDate = dateRef.current?.value
    if (deckDate) {
      const [year, month, day] = deckDate.split('-')
      formData.append('deck_date', `${parseInt(month, 10)}/${parseInt(day, 10)}/${year}`)
    }

    setLoading(true)
    try {
      let res
      try {
        res = await fetch('https://ciege-production.up.railway.app/build', {
          method: 'POST',
          body: formData,
        })
      } catch (err) {
        console.error('Railway fetch error:', err)
        console.error('Error type:', err.name)
        console.error('Error message:', err.message)
        throw err
      }

      if (!res.ok) {
        const text = await res.text()
        throw new Error(text || `Server error: ${res.status}`)
      }

      const blob = await res.blob()
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      const disposition = res.headers.get('content-disposition')
      const match = disposition?.match(/filename="?([^"]+)"?/)
      a.download = match?.[1] ?? 'deck.zip'
      document.body.appendChild(a)
      a.click()
      a.remove()
      URL.revokeObjectURL(url)

      setSuccess(true)
    } catch (err) {
      setError(err.message ?? 'Something went wrong.')
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="min-h-screen bg-zinc-950 text-zinc-100">
      <div className="mx-auto max-w-[1600px] px-6 py-8">
        <div className="grid gap-6 lg:grid-cols-[280px_minmax(0,1fr)]">

          {/* Sidebar */}
          <aside className="rounded-[32px] border border-white/10 bg-white/5 p-6 shadow-[0_24px_120px_-80px_rgba(0,0,0,0.55)] backdrop-blur-xl">
            <div className="mb-8">
              <p className="text-xs uppercase tracking-[0.4em] text-emerald-300/80">Ciege</p>
              <h1 className="mt-3 text-3xl font-semibold text-white">PM Dashboard</h1>
              <p className="mt-2 text-sm leading-6 text-zinc-400">
                Navigate your planning platform and keep every delivery aligned.
              </p>
            </div>

            <nav className="space-y-2">
              {[
                { label: 'Dashboard', href: '/' },
                { label: 'HOP Readiness', href: '/weekly-focus' },
                { label: 'Deck Builder', href: '/deck-builder', active: true },
                { label: 'GC Call View', href: '/gc-call' },
                { label: 'CM Call View', href: '/cm-view' },
                { label: 'NTP Tracker', href: '/ntp-tracker' },
                { label: 'Change Log', href: '/change-log' },
              ].map((item) => (
                <a
                  key={item.label}
                  href={item.href}
                  className={`block w-full rounded-2xl px-4 py-3 text-left text-sm font-medium transition hover:bg-white/10 ${
                    item.active ? 'bg-emerald-400/10 text-emerald-300' : 'text-zinc-300'
                  }`}
                >
                  {item.label}
                </a>
              ))}
            </nav>

            <div className="mt-10 rounded-3xl border border-white/10 bg-zinc-900/70 p-4 text-sm text-zinc-400">
              <p className="font-medium text-white">Platform insight</p>
              <p className="mt-3 leading-6">
                Keep your project teams aligned with a single source of truth for decks, actions, NTPs, and invoices.
              </p>
            </div>
          </aside>

          {/* Main */}
          <main className="space-y-6">
            <section className="rounded-[32px] border border-white/10 bg-white/5 p-8 shadow-[0_24px_120px_-80px_rgba(0,0,0,0.55)] backdrop-blur-xl">
              <p className="text-sm uppercase tracking-[0.4em] text-emerald-300/80">Tools</p>
              <h2 className="mt-3 text-3xl font-semibold text-white">Deck Builder</h2>
              <p className="mt-2 text-sm leading-6 text-zinc-400">
                Upload your previous deck and set a date. Tracker and snapshot load automatically from the Dashboard.
              </p>
            </section>

            <section className="rounded-[32px] border border-white/10 bg-white/5 p-8 shadow-[0_24px_120px_-80px_rgba(0,0,0,0.55)] backdrop-blur-xl">
              <form onSubmit={handleSubmit} className="space-y-6">

                {/* Auto-loaded status cards */}
                <div className="grid gap-4 sm:grid-cols-2">
                  {/* Tracker status */}
                  <div className="bg-gray-800 border border-gray-600 rounded-lg p-4">
                    {trackerLoaded ? (
                      <div>
                        <p className="text-green-400 text-sm font-semibold">✅ Tracker loaded from Dashboard</p>
                        <p className="text-gray-500 text-xs mt-1">
                          {trackerInfo?.filename} · {trackerInfo?.hop_count} HOPs · {fmtDate(trackerInfo?.uploaded_at)}
                        </p>
                      </div>
                    ) : (
                      <div>
                        <p className="text-red-400 text-sm font-semibold">⚠️ No tracker found — upload on Dashboard first</p>
                      </div>
                    )}
                  </div>

                  {/* Previous snapshot status */}
                  <div className="bg-gray-800 border border-gray-600 rounded-lg p-4">
                    {prevSnapInfo ? (
                      <div>
                        <p className="text-green-400 text-sm font-semibold">✅ Previous session loaded automatically</p>
                        <p className="text-gray-500 text-xs mt-1">
                          Comparing against: {fmtDate(prevSnapInfo.uploaded_at) || 'previous upload'}
                        </p>
                      </div>
                    ) : (
                      <div>
                        <p className="text-yellow-400 text-sm font-semibold">⚠️ No previous session found — first build will have no delta</p>
                      </div>
                    )}
                  </div>
                </div>

                {/* File uploads */}
                <div className="grid gap-5 sm:grid-cols-2">
                  {[
                    { id: 'previous_deck', label: 'Previous Deck', accept: '.pptx', hint: '.pptx' },
                    { id: 'ntp_comments', label: 'NTP Comments', accept: '.xlsx', hint: '.xlsx (optional)' },
                  ].map((field) => (
                    <div key={field.id}>
                      <label
                        htmlFor={field.id}
                        className="mb-2 block text-xs font-medium uppercase tracking-[0.3em] text-zinc-400"
                      >
                        {field.label}
                        <span className="ml-2 font-normal normal-case tracking-normal text-zinc-600">
                          {field.hint}
                        </span>
                      </label>
                      <label
                        htmlFor={field.id}
                        className="flex cursor-pointer items-center gap-3 rounded-2xl border border-white/10 bg-zinc-900/70 px-4 py-3 transition hover:border-emerald-400/30 hover:bg-zinc-800/60"
                      >
                        <span className="flex h-8 w-8 shrink-0 items-center justify-center rounded-xl bg-emerald-400/10 text-emerald-300">
                          <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
                            <path
                              d="M8 2v8m0-8L5.5 4.5M8 2l2.5 2.5M2.5 11.5h11"
                              stroke="currentColor"
                              strokeWidth="1.5"
                              strokeLinecap="round"
                              strokeLinejoin="round"
                            />
                          </svg>
                        </span>
                        <span className="truncate text-sm text-zinc-400">
                          {fileNames[field.id] ?? 'Choose file…'}
                        </span>
                        <input
                          id={field.id}
                          type="file"
                          accept={field.accept}
                          className="sr-only"
                          onChange={(e) => handleFileChange(field.id, e)}
                        />
                      </label>
                    </div>
                  ))}
                </div>

                {/* Deck date */}
                <div className="max-w-xs">
                  <label
                    htmlFor="deck_date"
                    className="mb-2 block text-xs font-medium uppercase tracking-[0.3em] text-zinc-400"
                  >
                    Deck Date
                  </label>
                  <input
                    id="deck_date"
                    type="date"
                    ref={dateRef}
                    className="w-full rounded-2xl border border-white/10 bg-zinc-900/70 px-4 py-3 text-sm text-zinc-100 outline-none transition focus:border-emerald-400/40 focus:ring-1 focus:ring-emerald-400/20 [color-scheme:dark]"
                  />
                </div>

                {/* Feedback */}
                {error && (
                  <div className="rounded-2xl border border-red-500/20 bg-red-500/10 px-4 py-3 text-sm text-red-300">
                    {error}
                  </div>
                )}
                {success && (
                  <div className="rounded-2xl border border-emerald-500/20 bg-emerald-500/10 px-4 py-3 text-sm text-emerald-300">
                    Deck built successfully — your download should have started.
                  </div>
                )}

                {/* Submit */}
                <div className="flex items-center gap-4">
                  <button
                    type="submit"
                    disabled={!trackerLoaded || loading}
                    className="inline-flex items-center gap-2.5 rounded-2xl bg-emerald-500 px-6 py-3 text-sm font-semibold text-zinc-950 shadow-lg shadow-emerald-900/30 transition hover:bg-emerald-400 disabled:cursor-not-allowed disabled:opacity-50"
                  >
                    {loading ? (
                      <>
                        <svg className="h-4 w-4 animate-spin" viewBox="0 0 24 24" fill="none" aria-hidden="true">
                          <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                          <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v4a4 4 0 00-4 4H4z" />
                        </svg>
                        Building deck…
                      </>
                    ) : (
                      'Build Deck'
                    )}
                  </button>
                  {loading && (
                    <p className="text-sm text-zinc-400">Processing — this may take a moment.</p>
                  )}
                  {!trackerLoaded && (
                    <p className="text-sm text-red-400">Upload a tracker on the Dashboard to enable builds.</p>
                  )}
                </div>
              </form>
            </section>
          </main>
        </div>
      </div>
    </div>
  )
}
