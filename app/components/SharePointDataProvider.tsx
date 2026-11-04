'use client'

import { createContext, useCallback, useContext, useEffect, useMemo, useRef, useState, type ReactNode } from 'react'
import { useMsal } from '@azure/msal-react'
import { graphLoginRequest } from '../lib/msalConfig'
import {
  readNamedFileWithMetadata, type SharePointFileKey,
  SharePointRateLimitError,
} from '../lib/sharepoint'
import {
  parseTrackerWorkbook, parseDecomWorkbook, parseCrWorkbook, parseSpoWorkbook, parseScopWorkbook,
} from '../lib/sharepointParsers'
import { readNokiaCache, writeNokiaCache } from '../lib/nokiaCache'

// Stage 2 Part A — the shared, in-memory-only data layer every page reads
// Nokia files through instead of Supabase. NEVER persisted to
// localStorage/sessionStorage/Supabase — the only on-device persistence is
// app/lib/nokiaCache.ts's IndexedDB store (opt-in, account-scoped, 12h
// expiry), used purely for stale-while-revalidate so a hard reload with a
// warm cache can render before the network even responds.
//
// Each file loads lazily — only once some page actually calls
// useSharePointFile(key) — and is shared across every page/navigation for
// the rest of the session once loaded (in-memory `loadedKeys`), so
// switching between pages that both need the tracker doesn't re-download
// it. This matters most for CR (~30MB): a page that never touches CR data
// never triggers its load.
//
// Perf fix, 11/2026: this used to make two separate Graph calls per file
// (content + metadata) via Promise.all — readNamedFileWithMetadata now
// does it as one Graph call (the DriveItem, which carries a direct
// downloadUrl) plus one non-Graph content fetch. The real fix for "revisit
// feels slow," though, wasn't in here at all — the Dashboard's nav cards
// were plain <a href> tags, so every click into /map or /tracker was a
// full page reload that destroyed this provider (and its in-memory
// `loadedKeys` cache) and remounted it from scratch. See app/page.tsx.

const PARSERS: Record<SharePointFileKey, (buf: ArrayBuffer) => Promise<unknown[][]>> = {
  tracker: parseTrackerWorkbook,
  decom: parseDecomWorkbook,
  cr: parseCrWorkbook,
  spo: parseSpoWorkbook,
  scop: parseScopWorkbook,
}

const FILE_KEYS: SharePointFileKey[] = ['tracker', 'decom', 'cr', 'spo', 'scop']

export type SharePointLoadPhase = 'idle' | 'downloading' | 'parsing' | 'building' | 'refreshing' | 'done' | 'error'

export interface SharePointFileState {
  rows: unknown[][] | null
  lastModifiedDateTime: string | null
  refreshedAt: string | null
  loading: boolean
  // Granular progress for a "Downloading… Parsing… Building view…" UI
  // instead of one opaque spinner. 'refreshing' means a cached entry is
  // already rendering (rows is non-null) while a background fetch is in
  // flight to replace it.
  phase: SharePointLoadPhase
  // True while `rows` reflects the on-device cache rather than a live
  // SharePoint response — cleared the moment the real fetch lands.
  fromCache: boolean
  error: string | null
}

const EMPTY_FILE_STATE: SharePointFileState = {
  rows: null, lastModifiedDateTime: null, refreshedAt: null, loading: false, phase: 'idle', fromCache: false, error: null,
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
    console.time('[sp-timing] msal acquireTokenSilent')
    try {
      const result = await instance.acquireTokenSilent({ ...graphLoginRequest, account })
      console.timeEnd('[sp-timing] msal acquireTokenSilent')
      return result.accessToken
    } catch {
      console.timeEnd('[sp-timing] msal acquireTokenSilent')
      // Silent acquisition fails when the cached token's expired or needs
      // re-consent — fall back to an interactive popup once rather than
      // leaving the caller permanently stuck. Only hit if initialize()
      // couldn't restore a usable session at all (see MsalProviderInner).
      console.time('[sp-timing] msal acquireTokenPopup')
      const result = await instance.acquireTokenPopup(graphLoginRequest)
      console.timeEnd('[sp-timing] msal acquireTokenPopup')
      return result.accessToken
    }
  }, [instance, accounts])

  const load = useCallback((key: SharePointFileKey, isRefresh: boolean): Promise<void> => {
    if (!isRefresh && loadedKeys.current.has(key)) return Promise.resolve()
    const existing = inFlight.current[key]
    if (existing) return existing

    const run = (async () => {
      const accountId = accounts[0]?.homeAccountId

      // Stale-while-revalidate: a non-expired cached entry renders
      // immediately (phase 'refreshing', fromCache true) while the real
      // fetch below still runs in the background and replaces it —
      // never shown for an explicit Refresh click, only a fresh load.
      if (!isRefresh && accountId) {
        const cached = await readNokiaCache(key, accountId)
        if (cached) {
          patchFile(key, {
            rows: cached.rows, refreshedAt: cached.savedAt,
            loading: true, phase: 'refreshing', fromCache: true, error: null,
          })
          loadedKeys.current.add(key)
        }
      }

      patchFile(key, { loading: true, phase: 'downloading', error: null })
      try {
        const accessToken = await getAccessToken()

        const readOnce = () => readNamedFileWithMetadata(accessToken, key)
        let buffer: ArrayBuffer
        let metadata: { lastModifiedDateTime: string }
        try {
          ({ buffer, metadata } = await readOnce())
        } catch (err) {
          if (err instanceof SharePointRateLimitError) {
            // One retry after the server-specified delay — a second
            // failure surfaces as a real error instead of retrying forever.
            await new Promise(resolve => setTimeout(resolve, err.retryAfterSeconds * 1000))
            ;({ buffer, metadata } = await readOnce())
          } else {
            throw err
          }
        }

        patchFile(key, { phase: 'parsing' })
        const rows = await PARSERS[key](buffer)

        patchFile(key, { phase: 'building' })
        loadedKeys.current.add(key)
        const now = new Date().toISOString()
        patchFile(key, {
          rows,
          lastModifiedDateTime: metadata.lastModifiedDateTime,
          refreshedAt: now,
          loading: false,
          phase: 'done',
          fromCache: false,
          error: null,
        })
        if (accountId) void writeNokiaCache(key, accountId, rows)
      } catch (err) {
        // Deliberately NOT added to loadedKeys on failure, so the next
        // ensureLoaded() call (e.g. the consuming page re-mounting) retries
        // instead of silently staying broken forever.
        patchFile(key, { loading: false, phase: 'error', error: err instanceof Error ? err.message : 'Load failed.' })
      } finally {
        delete inFlight.current[key]
      }
    })()

    inFlight.current[key] = run
    return run
  }, [getAccessToken, patchFile, accounts])

  const ensureLoaded = useCallback((key: SharePointFileKey) => { void load(key, false) }, [load])
  const refresh = useCallback((key: SharePointFileKey) => { void load(key, true) }, [load])

  const value = useMemo<SharePointDataContextValue>(
    () => ({ files, ensureLoaded, refresh }),
    [files, ensureLoaded, refresh],
  )

  return <SharePointDataContext.Provider value={value}>{children}</SharePointDataContext.Provider>
}

/** Per-file Nokia data — rows, Graph's lastModifiedDateTime, when this
 * session last refreshed it, a granular load phase, whether `rows` is
 * currently the on-device cache vs. a live response, error state, and a
 * refresh() to force a fresh read. Triggers the file's load on first use;
 * afterwards every caller (any page, any number of times) shares the same
 * cached result for the rest of the session — nothing is written to
 * disk/Supabase except through app/lib/nokiaCache.ts's approved store. */
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
