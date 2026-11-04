'use client'

import { useSharePointFile } from './SharePointDataProvider'
import type { SharePointFileKey } from '../lib/sharepoint'

function fmt(iso: string | null): string {
  if (!iso) return '—'
  return new Date(iso).toLocaleString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit' })
}

/** "Last loaded from SharePoint: <time>" + a Refresh button — drop this in
 * on any page that reads Nokia data through SharePointDataProvider, so
 * it's always visible where that data came from and when it was last
 * pulled. While `fromCache` is true, `refreshedAt` is the on-device cache's
 * savedAt (not a live SharePoint response yet) and a small "refreshing…"
 * indicator shows next to it — the background fetch that will replace it
 * is already in flight (see SharePointDataProvider's stale-while-revalidate
 * load path), not something this button needs to trigger. */
export default function SharePointSourceLine({ fileKey }: { fileKey: SharePointFileKey }) {
  const { lastModifiedDateTime, refreshedAt, loading, phase, fromCache, error, refresh } = useSharePointFile(fileKey)

  return (
    <div className="flex flex-wrap items-center gap-2 text-xs text-gray-500">
      <span>
        Source: SharePoint · last modified {fmt(lastModifiedDateTime)} · Last loaded from SharePoint: {fmt(refreshedAt)}
      </span>
      {fromCache && phase === 'refreshing' && (
        <span className="text-blue-400 italic">refreshing…</span>
      )}
      <button onClick={refresh} disabled={loading}
        className="text-blue-400 hover:text-blue-300 disabled:opacity-40 underline disabled:no-underline">
        {loading && !fromCache ? 'Refreshing…' : 'Refresh'}
      </button>
      {error && <span className="text-red-400">{error}</span>}
    </div>
  )
}
