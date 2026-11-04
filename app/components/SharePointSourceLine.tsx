'use client'

import { useEffect, useState } from 'react'
import { useSharePointFile } from './SharePointDataProvider'
import type { SharePointFileKey } from '../lib/sharepoint'

function fmt(iso: string | null): string {
  if (!iso) return '—'
  return new Date(iso).toLocaleString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit' })
}

const PHASE_LABEL: Record<string, string> = {
  downloading: 'Downloading from SharePoint…',
  parsing: 'Parsing…',
  building: 'Building view…',
}

/** "Last loaded from SharePoint: <time>" + a Refresh/Retry button — drop
 * this in on any page that reads Nokia data through SharePointDataProvider,
 * so it's always visible where that data came from and when it was last
 * pulled. While `fromCache` is true, `refreshedAt` is the on-device cache's
 * savedAt (not a live SharePoint response yet) and a small "refreshing…"
 * indicator shows next to it — the background fetch that will replace it
 * is already in flight (see SharePointDataProvider's stale-while-revalidate
 * load path), not something this button needs to trigger.
 *
 * Bug fix, 11/2026: before the FIRST successful load (no cached rows, no
 * refreshedAt yet), this used to just show "last modified — · Last loaded
 * from SharePoint: —" with a disabled "Refreshing…" button — for a large,
 * slow-loading file (CR, the only one routed through a Web Worker) that
 * read as "broken," not "working." Now shows the actual load phase
 * (Downloading/Parsing/Building) during that first load, and relabels the
 * button "Retry" (still the same refresh() call) when there's an error, so
 * the fix action is obvious instead of generic.
 *
 * Bug fix, 11/2026 ("CR still shows stale data"): staleVsSharePoint (set by
 * the provider's cheap metadata-version check — see
 * SharePointDataProvider's bug-fix note) now surfaces as an explicit amber
 * "Showing older data, refreshing…" line, instead of the page quietly
 * displaying an old version with nothing to indicate a newer one exists
 * and is on its way. If that background refetch then fails, the amber line
 * is replaced by the normal error + "Retry" path below — it never just
 * goes back to looking "done" with stale data still on screen.
 *
 * Bug fix, 11/2026 ("SharePoint refresh gets stuck"): the Refresh button
 * used to be `disabled={loading}` — if a load ever hung (nothing in the
 * chain had a timeout at the time), `loading` stayed true forever and the
 * button was permanently unclickable, exactly matching the reported
 * "Refresh does nothing." The button is never disabled now; clicking it
 * while something's already running aborts that attempt and starts a fresh
 * one (see SharePointDataProvider's refresh()). While loading, this also
 * shows elapsed seconds in the current phase and a Cancel link. */
function ElapsedSeconds({ since }: { since: string | null }) {
  // Date.now() belongs in the effect, not render — computing elapsed time
  // directly during render is an impure read (React's purity rule), even
  // though it's only cosmetic here.
  const [secs, setSecs] = useState<number | null>(null)
  useEffect(() => {
    if (!since) return
    const sinceMs = new Date(since).getTime()
    const tick = () => setSecs(Math.max(0, Math.floor((Date.now() - sinceMs) / 1000)))
    // Deferred, not called synchronously in the effect body — satisfies
    // React's "no setState directly in an effect" rule while still showing
    // a value on the first render after `since` appears, not a full second
    // later.
    const kick = setTimeout(tick, 0)
    const id = setInterval(tick, 1000)
    return () => { clearTimeout(kick); clearInterval(id) }
  }, [since])
  if (secs === null) return null
  return <span className="text-gray-500">({secs}s)</span>
}

export default function SharePointSourceLine({ fileKey }: { fileKey: SharePointFileKey }) {
  const { lastModifiedDateTime, refreshedAt, loading, phase, phaseStartedAt, fromCache, error, staleVsSharePoint, revalidateStatus, refresh, cancel } = useSharePointFile(fileKey)

  const firstLoadInProgress = loading && !refreshedAt
  const buttonLabel = error ? 'Retry' : loading && !fromCache ? 'Refreshing…' : 'Refresh'
  const showingOlderData = staleVsSharePoint && !error && (revalidateStatus === 'running' || loading)

  return (
    <div className="flex flex-wrap items-center gap-2 text-xs text-gray-500">
      {firstLoadInProgress ? (
        <span className="text-blue-400">{PHASE_LABEL[phase] ?? 'Loading from SharePoint…'}</span>
      ) : (
        <span>
          Source: SharePoint · last modified {fmt(lastModifiedDateTime)} · Last loaded from SharePoint: {fmt(refreshedAt)}
        </span>
      )}
      {fromCache && phase === 'refreshing' && !showingOlderData && (
        <span className="text-blue-400 italic">refreshing… <ElapsedSeconds since={phaseStartedAt} /></span>
      )}
      {showingOlderData && (
        <span className="text-amber-400 font-semibold">⚠ Showing older data, refreshing… <ElapsedSeconds since={phaseStartedAt} /></span>
      )}
      <button onClick={refresh}
        className={`underline ${error ? 'text-red-300 hover:text-red-200 font-semibold' : 'text-blue-400 hover:text-blue-300'}`}>
        {buttonLabel}
      </button>
      {loading && (
        <>
          <ElapsedSeconds since={phaseStartedAt} />
          <button onClick={cancel} className="underline text-gray-400 hover:text-gray-300">Cancel</button>
        </>
      )}
      {error && <span className="text-red-400">{error}</span>}
    </div>
  )
}
