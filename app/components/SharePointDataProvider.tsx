'use client'

import { createContext, useCallback, useContext, useEffect, useMemo, useRef, useState, type ReactNode } from 'react'
import { useMsal } from '@azure/msal-react'
import { graphLoginRequest } from '../lib/msalConfig'
import {
  readNamedFile, getNamedFileMetadata, type SharePointFileKey,
  SharePointRateLimitError,
} from '../lib/sharepoint'
import {
  parseTrackerWorkbook, parseDecomWorkbook, parseCrWorkbook, parseSpoWorkbook, parseScopWorkbook,
} from '../lib/sharepointParsers'

// Stage 2 Part A — the shared, in-memory-only data layer every page will
// eventually read Nokia files through instead of Supabase. Deliberately
// NOT persisted anywhere (no localStorage/IndexedDB, no Supabase writes of
// Nokia data) — a page refresh means every file reloads from SharePoint
// again, by design. Lives here (not app/lib/) since it's a component
// (needs useMsal(), which only works inside React) despite otherwise
// reading like a data-layer module.
//
// Each file loads lazily — only once some page actually calls
// useSharePointFile(key) — and is shared across every page/navigation for
// the rest of the session once loaded, so switching between pages that
// both need the tracker doesn't re-download it. This matters most for CR
// (~30MB): a page that never touches CR data never triggers its load.

const PARSERS: Record<SharePointFileKey, (buf: ArrayBuffer) => unknown[][]> = {
  tracker: parseTrackerWorkbook,
  decom: parseDecomWorkbook,
  cr: parseCrWorkbook,
  spo: parseSpoWorkbook,
  scop: parseScopWorkbook,
}

const FILE_KEYS: SharePointFileKey[] = ['tracker', 'decom', 'cr', 'spo', 'scop']

export interface SharePointFileState {
  rows: unknown[][] | null
  lastModifiedDateTime: string | null
  refreshedAt: string | null
  loading: boolean
  error: string | null
}

const EMPTY_FILE_STATE: SharePointFileState = {
  rows: null, lastModifiedDateTime: null, refreshedAt: null, loading: false, error: null,
}

interface SharePointDataContextValue {
  files: Record<SharePointFileKey, SharePointFileState>
  ensureLoaded: (key: SharePointFileKey) => void
  refresh: (key: SharePointFileKey) => void
}

const SharePointDataContext = createContext<SharePointDataContextValue | null>(null)

export function SharePointDataProvider({ children }: { children: ReactNode }) {
  const { instance, accounts } = useMsal()
  const [files, setFiles] = useState<Record<SharePointFileKey, SharePointFileState>>(() => {
    const initial = {} as Record<SharePointFileKey, SharePointFileState>
    FILE_KEYS.forEach(k => { initial[k] = EMPTY_FILE_STATE })
    return initial
  })

  // Gating state lives in refs, not the `files` state above — setState
  // updaters must stay pure (no side effects), so "should a new load even
  // start" can't safely be decided from inside one. inFlight also de-dupes
  // concurrent callers (two components both mounting useSharePointFile for
  // the same key at once triggers exactly one read, not two).
  const inFlight = useRef<Partial<Record<SharePointFileKey, Promise<void>>>>({})
  const loadedKeys = useRef<Set<SharePointFileKey>>(new Set())

  const patchFile = useCallback((key: SharePointFileKey, patch: Partial<SharePointFileState>) => {
    setFiles(prev => ({ ...prev, [key]: { ...prev[key], ...patch } }))
  }, [])

  const getAccessToken = useCallback(async (): Promise<string> => {
    const account = accounts[0]
    if (!account) throw new Error('Not signed in with a Nokia account.')
    try {
      const result = await instance.acquireTokenSilent({ ...graphLoginRequest, account })
      return result.accessToken
    } catch {
      // Silent acquisition fails when the cached token's expired or needs
      // re-consent — fall back to an interactive popup once rather than
      // leaving the caller permanently stuck.
      const result = await instance.acquireTokenPopup(graphLoginRequest)
      return result.accessToken
    }
  }, [instance, accounts])

  const load = useCallback((key: SharePointFileKey, isRefresh: boolean): Promise<void> => {
    if (!isRefresh && loadedKeys.current.has(key)) return Promise.resolve()
    const existing = inFlight.current[key]
    if (existing) return existing

    const run = (async () => {
      patchFile(key, { loading: true, error: null })
      try {
        const accessToken = await getAccessToken()

        const readBoth = () => Promise.all([
          readNamedFile(accessToken, key),
          getNamedFileMetadata(accessToken, key),
        ] as const)

        let buffer: ArrayBuffer
        let metadata: { lastModifiedDateTime: string }
        try {
          ;[buffer, metadata] = await readBoth()
        } catch (err) {
          if (err instanceof SharePointRateLimitError) {
            // One retry after the server-specified delay — a second
            // failure surfaces as a real error instead of retrying forever.
            await new Promise(resolve => setTimeout(resolve, err.retryAfterSeconds * 1000))
            ;[buffer, metadata] = await readBoth()
          } else {
            throw err
          }
        }

        const rows = PARSERS[key](buffer)
        loadedKeys.current.add(key)
        patchFile(key, {
          rows,
          lastModifiedDateTime: metadata.lastModifiedDateTime,
          refreshedAt: new Date().toISOString(),
          loading: false,
          error: null,
        })
      } catch (err) {
        // Deliberately NOT added to loadedKeys on failure, so the next
        // ensureLoaded() call (e.g. the consuming page re-mounting) retries
        // instead of silently staying broken forever.
        patchFile(key, { loading: false, error: err instanceof Error ? err.message : 'Load failed.' })
      } finally {
        delete inFlight.current[key]
      }
    })()

    inFlight.current[key] = run
    return run
  }, [getAccessToken, patchFile])

  const ensureLoaded = useCallback((key: SharePointFileKey) => { void load(key, false) }, [load])
  const refresh = useCallback((key: SharePointFileKey) => { void load(key, true) }, [load])

  const value = useMemo<SharePointDataContextValue>(
    () => ({ files, ensureLoaded, refresh }),
    [files, ensureLoaded, refresh],
  )

  return <SharePointDataContext.Provider value={value}>{children}</SharePointDataContext.Provider>
}

/** Per-file Nokia data — rows, Graph's lastModifiedDateTime, when this
 * session last refreshed it, loading/error state, and a refresh() to force
 * a fresh read. Triggers the file's load on first use; afterwards every
 * caller (any page, any number of times) shares the same cached result for
 * the rest of the session — nothing is written to disk/Supabase. */
export function useSharePointFile(key: SharePointFileKey): SharePointFileState & { refresh: () => void } {
  const ctx = useContext(SharePointDataContext)
  if (!ctx) throw new Error('useSharePointFile must be used within SharePointDataProvider')
  const { files, ensureLoaded, refresh } = ctx

  // An effect, not a direct call during render — ensureLoaded ends up
  // calling setFiles (via patchFile) on the PROVIDER, and updating one
  // component's state synchronously from inside another component's
  // render is exactly what React's rules disallow. ensureLoaded is
  // ref-gated/idempotent regardless, so this only ever starts one load
  // per key no matter how many components call this hook for it.
  useEffect(() => { ensureLoaded(key) }, [key, ensureLoaded])

  return useMemo(() => ({ ...files[key], refresh: () => refresh(key) }), [files, key, refresh])
}
