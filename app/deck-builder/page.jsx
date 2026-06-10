'use client'

import { useRef, useState } from 'react'

const navItems = [
  { label: 'Deck Builder', href: '/deck-builder', active: true },
  { label: 'Action Board', href: '#', active: false },
  { label: 'NTP Tracker', href: '#', active: false },
  { label: 'SCOP Invoice', href: '#', active: false },
]

const fileFields = [
  {
    id: 'tracker',
    label: 'Excel Tracker',
    accept: '.xlsx',
    hint: '.xlsx',
  },
  {
    id: 'previous_deck',
    label: 'Previous Deck',
    accept: '.pptx',
    hint: '.pptx',
  },
  {
    id: 'snapshot',
    label: 'Snapshot',
    accept: '.json',
    hint: '.json',
  },
  {
    id: 'ntp_comments',
    label: 'NTP Comments',
    accept: '.xlsx',
    hint: '.xlsx',
  },
]

export default function DeckBuilderPage() {
  const [loading, setLoading] = useState(false)
  const [success, setSuccess] = useState(false)
  const [error, setError] = useState(null)
  const [fileNames, setFileNames] = useState({})

  const fileRefs = useRef({})
  const dateRef = useRef(null)

  function getRef(id) {
    if (!fileRefs.current[id]) {
      fileRefs.current[id] = { current: null }
    }
    return fileRefs.current[id]
  }

  function handleFileChange(id, e) {
    const file = e.target.files?.[0]
    setFileNames((prev) => ({ ...prev, [id]: file ? file.name : null }))
    setSuccess(false)
    setError(null)
  }

  async function handleSubmit(e) {
    e.preventDefault()
    setError(null)
    setSuccess(false)

    const formData = new FormData()

    for (const field of fileFields) {
      const input = document.getElementById(field.id)
      if (input?.files?.[0]) {
        formData.append(field.id, input.files[0])
      }
    }

    const deckDate = dateRef.current?.value
    if (deckDate) {
      const [year, month, day] = deckDate.split('-')
      formData.append('deck_date', `${parseInt(month, 10)}/${parseInt(day, 10)}/${year}`)
    }

    setLoading(true)
    try {
      const res = await fetch('https://ciege-production.up.railway.app/build', {
        method: 'POST',
        body: formData,
      })

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
              {navItems.map((item) => (
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
                Upload your source files, set a deck date, and generate a packaged deck in one click.
              </p>
            </section>

            <section className="rounded-[32px] border border-white/10 bg-white/5 p-8 shadow-[0_24px_120px_-80px_rgba(0,0,0,0.55)] backdrop-blur-xl">
              <form onSubmit={handleSubmit} className="space-y-6">
                <div className="grid gap-5 sm:grid-cols-2">
                  {fileFields.map((field) => (
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
                    disabled={loading}
                    className="inline-flex items-center gap-2.5 rounded-2xl bg-emerald-500 px-6 py-3 text-sm font-semibold text-zinc-950 shadow-lg shadow-emerald-900/30 transition hover:bg-emerald-400 disabled:cursor-not-allowed disabled:opacity-50"
                  >
                    {loading ? (
                      <>
                        <svg
                          className="h-4 w-4 animate-spin"
                          viewBox="0 0 24 24"
                          fill="none"
                          aria-hidden="true"
                        >
                          <circle
                            className="opacity-25"
                            cx="12"
                            cy="12"
                            r="10"
                            stroke="currentColor"
                            strokeWidth="4"
                          />
                          <path
                            className="opacity-75"
                            fill="currentColor"
                            d="M4 12a8 8 0 018-8v4a4 4 0 00-4 4H4z"
                          />
                        </svg>
                        Building deck…
                      </>
                    ) : (
                      'Build Deck'
                    )}
                  </button>
                  {loading && (
                    <p className="text-sm text-zinc-400">
                      Processing — this may take a moment.
                    </p>
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
