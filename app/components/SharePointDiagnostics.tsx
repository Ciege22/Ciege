'use client'

import { useEffect, useState } from 'react'
import { useMsal } from '@azure/msal-react'
import { useSharePointFile, type SharePointFileState } from './SharePointDataProvider'
import { peekNokiaCache, isMemoryOnlyFile } from '../lib/nokiaCache'
import { computeNewestMeaningfulDate } from '../lib/sharepointParsers'
import type { SharePointFileKey } from '../lib/sharepoint'

const FILE_KEYS: SharePointFileKey[] = ['tracker', 'decom', 'scop', 'spo', 'cr']

function fmt(iso: string | null): string {
  if (!iso) return '—'
  return new Date(iso).toLocaleString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit', second: '2-digit' })
}

function fmtMs(ms: number | null): string {
  return ms === null ? '—' : `${ms.toLocaleString()} ms`
}

function fmtMB(bytes: number | null): string {
  return bytes === null ? '—' : `${(bytes / 1024 / 1024).toFixed(2)} MB`
}

interface CachePeek {
  savedAt: string
  sourceLastModified: string | null
  expired: boolean
}

interface DiagnosticsRow {
  key: SharePointFileKey
  state: SharePointFileState
  cachePeek: CachePeek | null | 'loading'
  rowCount: number
  newestDateLabel: string
  newestDateValue: string
}

function DiagnosticsFileBlock({ fileKey, accountId }: { fileKey: SharePointFileKey; accountId: string | undefined }) {
  const state = useSharePointFile(fileKey)
  const [cachePeek, setCachePeek] = useState<CachePeek | null | 'loading'>('loading')

  useEffect(() => {
    let cancelled = false
    const run = async () => {
      if (!accountId) { if (!cancelled) setCachePeek(null); return }
      if (!cancelled) setCachePeek('loading')
      const entry = await peekNokiaCache(fileKey, accountId)
      if (!cancelled) setCachePeek(entry)
    }
    run()
    return () => { cancelled = true }
    // Re-peek whenever this file's refreshedAt/cacheWriteStatus changes
    // (e.g. after a write) — not the whole `state` object, which would
    // re-run this on every unrelated patch (loading/phase changes too).
  }, [fileKey, accountId, state.refreshedAt, state.cacheWriteStatus])

  const rowCount = state.rows?.length ?? 0
  const newest = state.rows ? computeNewestMeaningfulDate(fileKey, state.rows) : { columnLabel: '—', date: null }
  const memoryOnly = isMemoryOnlyFile(fileKey)

  const row: DiagnosticsRow = {
    key: fileKey, state, cachePeek, rowCount,
    newestDateLabel: newest.columnLabel,
    newestDateValue: newest.date ? newest.date.toLocaleDateString('en-US') : '—',
  }

  return <DiagnosticsRowView row={row} memoryOnly={memoryOnly} />
}

function DiagnosticsRowView({ row, memoryOnly }: { row: DiagnosticsRow; memoryOnly: boolean }) {
  const { state } = row
  const cacheLine = memoryOnly
    ? 'memory-only — never written to IndexedDB (largest file, kept in memory for this session only)'
    : row.cachePeek === 'loading'
    ? 'checking…'
    : row.cachePeek === null
    ? 'none'
    : `savedAt ${fmt(row.cachePeek.savedAt)} · ${row.cachePeek.expired ? 'EXPIRED' : 'not expired'} · source version at save: ${fmt(row.cachePeek.sourceLastModified)}`

  return (
    <div className="bg-gray-800 rounded-lg p-3 text-xs font-mono text-gray-300 space-y-1">
      <p className="text-white font-bold font-sans">{row.key}</p>
      <p>source: <span className="text-cyan-300">{state.source}</span>{state.fromCache ? ' (currently showing cached data)' : ''}</p>
      <p>cache: {cacheLine}</p>
      <p>SharePoint lastModified (of data shown): {fmt(state.lastModifiedDateTime)}</p>
      <p>SharePoint lastModified (latest known, from a freshness check): {fmt(state.liveLastModifiedDateTime)}{state.staleVsSharePoint ? ' — NEWER than what is shown' : ''}</p>
      <p>last successful network load: {fmt(state.source === 'network' ? state.refreshedAt : null)}</p>
      <p>row count: {row.rowCount} · newest {row.newestDateLabel}: {row.newestDateValue}</p>
      <p>download: {fmtMB(state.downloadBytes)} in {fmtMs(state.downloadMs)} · parse: {fmtMs(state.parseMs)} · worker used: {state.usedWorker === null ? '—' : state.usedWorker ? 'yes' : 'no'}</p>
      <p>cache write: {state.cacheWriteStatus ?? '—'}{state.cacheWriteError ? ` (${state.cacheWriteError})` : ''}</p>
      <p>last error: {state.lastErrorMessage ? `${state.lastErrorMessage} (at ${fmt(state.lastErrorAt)})` : 'none'}</p>
      <p>background revalidation: {state.revalidateStatus}</p>
      <p>loading: {state.loading ? `yes (${state.phase})` : 'no'}</p>
    </div>
  )
}

function buildDiagnosticsText(rows: { key: SharePointFileKey; state: SharePointFileState; cachePeek: CachePeek | null | 'loading'; rowCount: number; newestDateLabel: string; newestDateValue: string }[]): string {
  const lines: string[] = [`Ciege SharePoint diagnostics — ${new Date().toLocaleString('en-US')}`, '']
  rows.forEach(r => {
    const memoryOnly = isMemoryOnlyFile(r.key)
    lines.push(`[${r.key}]`)
    lines.push(`  source: ${r.state.source}${r.state.fromCache ? ' (currently showing cached data)' : ''}`)
    lines.push(`  cache: ${memoryOnly ? 'memory-only, never cached to disk' : r.cachePeek === 'loading' ? 'checking…' : r.cachePeek === null ? 'none' : `savedAt ${fmt(r.cachePeek.savedAt)}, ${r.cachePeek.expired ? 'EXPIRED' : 'not expired'}, source version at save ${fmt(r.cachePeek.sourceLastModified)}`}`)
    lines.push(`  SharePoint lastModified (shown data): ${fmt(r.state.lastModifiedDateTime)}`)
    lines.push(`  SharePoint lastModified (latest known): ${fmt(r.state.liveLastModifiedDateTime)}${r.state.staleVsSharePoint ? ' (NEWER than shown)' : ''}`)
    lines.push(`  last successful network load: ${fmt(r.state.source === 'network' ? r.state.refreshedAt : null)}`)
    lines.push(`  row count: ${r.rowCount}, newest ${r.newestDateLabel}: ${r.newestDateValue}`)
    lines.push(`  download: ${fmtMB(r.state.downloadBytes)} in ${fmtMs(r.state.downloadMs)}, parse: ${fmtMs(r.state.parseMs)}, worker used: ${r.state.usedWorker === null ? '—' : r.state.usedWorker ? 'yes' : 'no'}`)
    lines.push(`  cache write: ${r.state.cacheWriteStatus ?? '—'}${r.state.cacheWriteError ? ` (${r.state.cacheWriteError})` : ''}`)
    lines.push(`  last error: ${r.state.lastErrorMessage ? `${r.state.lastErrorMessage} (at ${fmt(r.state.lastErrorAt)})` : 'none'}`)
    lines.push(`  background revalidation: ${r.state.revalidateStatus}`)
    lines.push(`  loading: ${r.state.loading ? `yes (${r.state.phase})` : 'no'}`)
    lines.push('')
  })
  return lines.join('\n')
}

/** Collapsed-by-default Diagnostics section for Settings — counts,
 * timestamps, and sizes only, NEVER Nokia row contents, so it's safe to
 * read aloud or paste into a message. Built specifically so CJ (whose
 * Nokia browser blocks devtools/F12) can see what's actually happening
 * with the SharePoint data pipeline without needing the console. */
export default function SharePointDiagnostics() {
  const { accounts } = useMsal()
  const accountId = accounts[0]?.homeAccountId
  const [expanded, setExpanded] = useState(false)
  const [copied, setCopied] = useState(false)

  // Each DiagnosticsFileBlock calls useSharePointFile itself (triggering
  // that file's load the same as anywhere else it's used) — rendered only
  // once expanded, so opening Diagnostics doesn't eagerly load files a
  // page wouldn't otherwise need.
  return (
    <div className="bg-gray-900 rounded-xl border border-gray-700 p-4 mb-6">
      <button
        type="button"
        onClick={() => setExpanded(e => !e)}
        className="flex items-center justify-between w-full text-left"
      >
        <span className="text-sm font-bold text-white">Diagnostics {expanded ? '▲' : '▼'}</span>
        <span className="text-xs text-gray-500">Nokia data pipeline status — counts and timestamps only, no row contents</span>
      </button>
      {expanded && (
        <div className="mt-4">
          {!accountId && <p className="text-gray-500 text-sm mb-3">Sign in with your Nokia account to see diagnostics.</p>}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mb-4">
            {FILE_KEYS.map(key => <DiagnosticsFileBlock key={key} fileKey={key} accountId={accountId} />)}
          </div>
          <CopyDiagnosticsButton accountId={accountId} copied={copied} setCopied={setCopied} />
        </div>
      )}
    </div>
  )
}

// Separate component so it can call useSharePointFile for all five keys
// itself (hooks can't be called in a loop/conditionally in the parent) to
// build the same text buildDiagnosticsText produces for display.
function CopyDiagnosticsButton({ accountId, copied, setCopied }: { accountId: string | undefined; copied: boolean; setCopied: (v: boolean) => void }) {
  const trackerState = useSharePointFile('tracker')
  const decomState = useSharePointFile('decom')
  const scopState = useSharePointFile('scop')
  const spoState = useSharePointFile('spo')
  const crState = useSharePointFile('cr')
  const states: Record<SharePointFileKey, SharePointFileState> = {
    tracker: trackerState, decom: decomState, scop: scopState, spo: spoState, cr: crState,
  }

  const [peeks, setPeeks] = useState<Record<SharePointFileKey, CachePeek | null | 'loading'>>({
    tracker: 'loading', decom: 'loading', scop: 'loading', spo: 'loading', cr: 'loading',
  })
  useEffect(() => {
    if (!accountId) return
    let cancelled = false
    Promise.all(FILE_KEYS.map(k => peekNokiaCache(k, accountId).then(entry => [k, entry] as const))).then(results => {
      if (cancelled) return
      const next = { ...peeks }
      results.forEach(([k, entry]) => { next[k] = entry })
      setPeeks(next)
    })
    return () => { cancelled = true }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [accountId, trackerState.refreshedAt, decomState.refreshedAt, scopState.refreshedAt, spoState.refreshedAt, crState.refreshedAt])

  const handleCopy = async () => {
    const rows = FILE_KEYS.map(key => {
      const state = states[key]
      const newest = state.rows ? computeNewestMeaningfulDate(key, state.rows) : { columnLabel: '—', date: null }
      return {
        key, state, cachePeek: peeks[key], rowCount: state.rows?.length ?? 0,
        newestDateLabel: newest.columnLabel,
        newestDateValue: newest.date ? newest.date.toLocaleDateString('en-US') : '—',
      }
    })
    const text = buildDiagnosticsText(rows)
    try {
      await navigator.clipboard.writeText(text)
      setCopied(true)
      setTimeout(() => setCopied(false), 2000)
    } catch {
      // Clipboard API can be blocked by the same managed-browser policies
      // that block devtools — fall back to a prompt so the text is still
      // get-able by hand (select-all, copy).
      window.prompt('Copy failed — select and copy manually:', text)
    }
  }

  return (
    <button onClick={handleCopy} className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-semibold">
      {copied ? '✓ Copied' : 'Copy diagnostics'}
    </button>
  )
}
