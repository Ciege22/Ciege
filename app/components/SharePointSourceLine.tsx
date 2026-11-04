'use client'

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
 * the fix action is obvious instead of generic. */
export default function SharePointSourceLine({ fileKey }: { fileKey: SharePointFileKey }) {
  const { lastModifiedDateTime, refreshedAt, loading, phase, fromCache, error, refresh } = useSharePointFile(fileKey)

  const firstLoadInProgress = loading && !refreshedAt
  const buttonLabel = error ? 'Retry' : loading && !fromCache ? 'Refreshing…' : 'Refresh'

  return (
    <div className="flex flex-wrap items-center gap-2 text-xs text-gray-500">
      {firstLoadInProgress ? (
        <span className="text-blue-400">{PHASE_LABEL[phase] ?? 'Loading from SharePoint…'}</span>
      ) : (
        <span>
          Source: SharePoint · last modified {fmt(lastModifiedDateTime)} · Last loaded from SharePoint: {fmt(refreshedAt)}
        </span>
      )}
      {fromCache && phase === 'refreshing' && (
        <span className="text-blue-400 italic">refreshing…</span>
      )}
      <button onClick={refresh} disabled={loading}
        className={`underline disabled:no-underline disabled:opacity-40 ${error ? 'text-red-300 hover:text-red-200 font-semibold' : 'text-blue-400 hover:text-blue-300'}`}>
        {buttonLabel}
      </button>
      {error && <span className="text-red-400">{error}</span>}
    </div>
  )
}
