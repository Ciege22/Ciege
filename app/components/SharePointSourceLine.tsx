'use client'

import { useSharePointFile } from './SharePointDataProvider'
import type { SharePointFileKey } from '../lib/sharepoint'

function fmt(iso: string | null): string {
  if (!iso) return '—'
  return new Date(iso).toLocaleString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit' })
}

/** "Source: SharePoint · last modified <time> · refreshed <time>" + a
 * Refresh button — drop this in on any page that reads Nokia data through
 * SharePointDataProvider, so it's always visible where that data came
 * from and when it was last pulled. */
export default function SharePointSourceLine({ fileKey }: { fileKey: SharePointFileKey }) {
  const { lastModifiedDateTime, refreshedAt, loading, error, refresh } = useSharePointFile(fileKey)

  return (
    <div className="flex flex-wrap items-center gap-2 text-xs text-gray-500">
      <span>
        Source: SharePoint · last modified {fmt(lastModifiedDateTime)} · refreshed {fmt(refreshedAt)}
      </span>
      <button onClick={refresh} disabled={loading}
        className="text-blue-400 hover:text-blue-300 disabled:opacity-40 underline disabled:no-underline">
        {loading ? 'Refreshing…' : 'Refresh'}
      </button>
      {error && <span className="text-red-400">{error}</span>}
    </div>
  )
}
