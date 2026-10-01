'use client'

export const dynamic = 'force-dynamic'

import { useState, useEffect } from 'react'
import { loadTrackerSnapshot } from '../lib/supabase'
import { loadChunkedReport } from '../lib/reportChunks'
import BackToDashboard from '../components/BackToDashboard'
import {
  parseDecomRows, parseTrackerHopsForDecom, findMissingDecom,
  countDroppedOffWithoutCxComplete,
} from '../lib/decom'
import { buildDecomDeckBlob } from '../lib/decomDeck'
import {
  resolveScopCalcSettings, keyScopRows, buildScopDataset,
  computeQuickBaseView, computePathwaveView, computeOverallView,
} from '../lib/scop'
import { buildScopDeckBlob } from '../lib/scopDeck'
import { loadScopSettings } from '../lib/settings'

// Deck Builder V2 — a completely separate page and backend route from
// /deck-builder (see backend/build_deck_v2.py's module docstring for the
// full list of behavioral differences). Nothing here touches /deck-builder's
// page, its Supabase reads, or backend/build_deck.py — this is intentionally
// isolated so CJ can test it for a few weeks before anyone considers
// switching over.

export default function DeckBuilderV2Page() {
  const [loading, setLoading] = useState(false)
  const [success, setSuccess] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [statusMsg, setStatusMsg] = useState('')

  const [previousDeckFile, setPreviousDeckFile] = useState<File | null>(null)
  const [trackerOverrideFile, setTrackerOverrideFile] = useState<File | null>(null)
  const [deckDate, setDeckDate] = useState('')

  const [trackerLoaded, setTrackerLoaded] = useState(false)
  const [trackerInfo, setTrackerInfo] = useState<{ filename: string; uploaded_at: string; hop_count: number } | null>(null)

  useEffect(() => {
    loadTrackerSnapshot().then(snap => {
      if (snap) {
        setTrackerInfo({ filename: snap.filename, uploaded_at: snap.uploaded_at, hop_count: snap.hop_count })
        setTrackerLoaded(true)
      }
    })
  }, [])

  function fmtDate(iso: string) {
    if (!iso) return ''
    const d = new Date(iso)
    return d.toLocaleDateString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric' }) +
      ' at ' + d.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })
  }

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    setError(null)
    setSuccess(false)

    if (!previousDeckFile) {
      setError('Previous deck (.pptx) is required.')
      return
    }
    if (!deckDate) {
      setError('Deck date is required.')
      return
    }

    setLoading(true)
    try {
      setStatusMsg('Loading Decom tracker data…')
      const [decomReport, trackerSnap] = await Promise.all([
        loadChunkedReport('decom'),
        loadTrackerSnapshot(),
      ])
      const decomRawRows = decomReport?.rows ?? []
      const trackerRawRows = trackerSnap?.data ?? []
      const decomRows = parseDecomRows(decomRawRows)
      const trackerHops = parseTrackerHopsForDecom(trackerRawRows)
      const missingDecomSites = findMissingDecom(decomRows, trackerHops)
      const extraDroppedOff = countDroppedOffWithoutCxComplete(decomRawRows)

      setStatusMsg('Building Decom Dashboard slides…')
      const decomBlob = await buildDecomDeckBlob(decomRows, missingDecomSites, extraDroppedOff)

      setStatusMsg('Loading SCOP tracker data…')
      const scopReport = await loadChunkedReport('scop')
      const scopSettings = await loadScopSettings()
      const calc = resolveScopCalcSettings(scopSettings)
      const scopDataset = buildScopDataset(keyScopRows(scopReport?.rows ?? []), calc)
      const qbv = computeQuickBaseView(scopDataset)
      const pw = computePathwaveView(scopDataset)
      const ov = computeOverallView(scopDataset)
      const notYetConstructed = scopDataset.length - ov.trackedTotal

      setStatusMsg('Building SCOP Status Deck slides…')
      const scopBlob = await buildScopDeckBlob({
        totalNokiaHops: scopDataset.length,
        constructionComplete: ov.trackedTotal,
        notYetConstructed,
        qbView: qbv, pwView: pw, overallView: ov,
        generatedDate: calc.asOfDate,
      })

      setStatusMsg('Building the deck…')
      const formData = new FormData()
      formData.append('previous_deck', previousDeckFile)
      if (trackerOverrideFile) formData.append('tracker', trackerOverrideFile)
      const [year, month, day] = deckDate.split('-')
      formData.append('deck_date', `${parseInt(month, 10)}/${parseInt(day, 10)}/${year}`)
      formData.append('decom_pptx', decomBlob, 'decom.pptx')
      formData.append('scop_pptx', scopBlob, 'scop.pptx')

      const res = await fetch('https://ciege-production.up.railway.app/build_v2', {
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
      a.download = match?.[1] ?? 'deck_v2.pptx'
      document.body.appendChild(a)
      a.click()
      a.remove()
      URL.revokeObjectURL(url)

      setSuccess(true)
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : 'Something went wrong.')
    } finally {
      setLoading(false)
      setStatusMsg('')
    }
  }

  return (
    <div className="min-h-screen bg-zinc-950 text-zinc-100">
      <div className="mx-auto max-w-3xl px-6 py-8">
        <BackToDashboard />

        <section className="rounded-[32px] border border-amber-400/20 bg-amber-400/5 p-6 mb-6">
          <p className="text-xs uppercase tracking-[0.4em] text-amber-300/80">Testing</p>
          <h1 className="mt-2 text-2xl font-semibold text-white">Deck Builder — V2 (In Progress)</h1>
          <p className="mt-2 text-sm leading-6 text-zinc-400">
            Separate from the production Deck Builder — nothing here affects it. Auto-merges the Decom Dashboard
            and SCOP Status Deck slides, uses a rolling POR window, keeps Plan numbers exactly as you last set
            them on the Cx Start/Complete charts, and compares &quot;What Changed&quot; against this tool&apos;s own
            last build instead of an unrelated tracker upload.
          </p>
        </section>

        <section className="rounded-[32px] border border-white/10 bg-white/5 p-8">
          <form onSubmit={handleSubmit} className="space-y-6">
            <div className="bg-gray-800 border border-gray-600 rounded-lg p-4">
              {trackerLoaded ? (
                <div>
                  <p className="text-green-400 text-sm font-semibold">✅ Tracker loaded from Dashboard</p>
                  <p className="text-gray-500 text-xs mt-1">
                    {trackerInfo?.filename} · {trackerInfo?.hop_count} HOPs · {fmtDate(trackerInfo?.uploaded_at ?? '')}
                  </p>
                </div>
              ) : (
                <p className="text-red-400 text-sm font-semibold">⚠️ No tracker found — upload on Dashboard first, or override below</p>
              )}
            </div>

            <div>
              <label className="mb-2 block text-xs font-medium uppercase tracking-[0.3em] text-zinc-400">
                Previous Deck (this tool&apos;s own last output) <span className="text-amber-400">required</span>
              </label>
              <input
                type="file" accept=".pptx"
                onChange={e => setPreviousDeckFile(e.target.files?.[0] ?? null)}
                className="w-full text-sm text-zinc-300 file:mr-4 file:rounded-lg file:border-0 file:bg-emerald-600 file:px-4 file:py-2 file:text-white file:text-sm"
              />
              {previousDeckFile && <p className="text-xs text-zinc-500 mt-1">{previousDeckFile.name}</p>}
            </div>

            <div>
              <label className="mb-2 block text-xs font-medium uppercase tracking-[0.3em] text-zinc-400">
                Tracker Override <span className="text-zinc-600 normal-case tracking-normal">(.xlsx, optional — overrides Supabase)</span>
              </label>
              <input
                type="file" accept=".xlsx"
                onChange={e => setTrackerOverrideFile(e.target.files?.[0] ?? null)}
                className="w-full text-sm text-zinc-300 file:mr-4 file:rounded-lg file:border-0 file:bg-zinc-700 file:px-4 file:py-2 file:text-white file:text-sm"
              />
            </div>

            <div className="max-w-xs">
              <label className="mb-2 block text-xs font-medium uppercase tracking-[0.3em] text-zinc-400">
                Deck Date
              </label>
              <input
                type="date" value={deckDate} onChange={e => setDeckDate(e.target.value)}
                className="w-full rounded-2xl border border-white/10 bg-zinc-900/70 px-4 py-3 text-sm text-zinc-100 outline-none [color-scheme:dark]"
              />
            </div>

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

            <div className="flex items-center gap-4">
              <button
                type="submit" disabled={loading}
                className="inline-flex items-center gap-2.5 rounded-2xl bg-emerald-500 px-6 py-3 text-sm font-semibold text-zinc-950 shadow-lg transition hover:bg-emerald-400 disabled:cursor-not-allowed disabled:opacity-50"
              >
                {loading ? 'Building…' : 'Build Deck (V2)'}
              </button>
              {loading && statusMsg && <p className="text-sm text-zinc-400">{statusMsg}</p>}
            </div>
          </form>
        </section>
      </div>
    </div>
  )
}
