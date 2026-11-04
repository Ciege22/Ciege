'use client'

import { createContext, useCallback, useContext, useEffect, useMemo, useRef, useState, type ReactNode } from 'react'
import { useMsal } from '@azure/msal-react'
import { graphLoginRequest } from '../lib/msalConfig'
import {
  readNamedFileWithMetadata, getNamedFileMetadata, type SharePointFileKey,
  SharePointRateLimitError,
} from '../lib/sharepoint'
import {
  parseTrackerWorkbook, parseDecomWorkbook, parseCrWorkbook, parseSpoWorkbook, parseScopWorkbook,
  getLastDecodeMeta,
} from '../lib/sharepointParsers'
import { readNokiaCache, writeNokiaCache, isMemoryOnlyFile } from '../lib/nokiaCache'

// Stage 2 Part A — the shared, in-memory-only data layer every page reads
// Nokia files through instead of Supabase. NEVER persisted to
// localStorage/sessionStorage/Supabase — the only on-device persistence is
// app/lib/nokiaCache.ts's IndexedDB store (opt-in, account-scoped, 12h
// expiry, CR excluded entirely — see that file), used purely for stale-
// while-revalidate so a hard reload with a warm cache can render before the
// network even responds.
//
// Bug fix, 11/2026 ("CR still shows stale data"): a key that had already
// loaded once THIS SESSION (loadedKeys) used to short-circuit every later
// ensureLoaded() call completely — no re-check of any kind, for the rest of
// the session, no matter how long ago that first load was or whether
// SharePoint's copy had changed since. Revisiting a page (unmount+remount)
// re-ran ensureLoaded() but got nothing from it. Now every ensureLoaded()
// call — first-ever or not — does at least a cheap Graph metadata-only
// check (getNamedFileMetadata, no content download) and compares its
// lastModifiedDateTime against what's currently shown; only a real mismatch
// triggers a full re-fetch. See verifyFreshnessThenFetch below. This also
// closes the cache side of the same bug: a cache entry now stores the
// source file's lastModifiedDateTime at write time, so a cache hit is
// version-checked the same way instead of only ever being time-checked
// (the 12h expiry alone couldn't tell a 10-minute-old cache entry from a
// 10-minute-old SharePoint edit).

const PARSERS: Record<SharePointFileKey, (buf: ArrayBuffer) => Promise<unknown[][]>> = {
  tracker: parseTrackerWorkbook,
  decom: parseDecomWorkbook,
  cr: parseCrWorkbook,
  spo: parseSpoWorkbook,
  scop: parseScopWorkbook,
}

const FILE_KEYS: SharePointFileKey[] = ['tracker', 'decom', 'cr', 'spo', 'scop']

export type SharePointLoadPhase = 'idle' | 'downloading' | 'parsing' | 'building' | 'refreshing' | 'done' | 'error'
export type SharePointDataSource = 'none' | 'memory' | 'cache' | 'network'
export type RevalidateStatus = 'idle' | 'running' | 'finished' | 'failed'

export interface SharePointFileState {
  rows: unknown[][] | null
  // Graph lastModifiedDateTime of the file version the CURRENTLY-SHOWN
  // rows came from (cache or network) — not necessarily SharePoint's
  // latest; see liveLastModifiedDateTime/staleVsSharePoint for that.
  lastModifiedDateTime: string | null
  refreshedAt: string | null
  loading: boolean
  // Granular progress for a "Downloading… Parsing… Building view…" UI
  // instead of one opaque spinner. 'refreshing' means a cached entry (or
  // already-in-memory data) is already rendering while a background fetch
  // is in flight to replace it.
  phase: SharePointLoadPhase
  // True while `rows` reflects the on-device cache rather than a live
  // SharePoint response — cleared the moment a real fetch lands.
  fromCache: boolean
  error: string | null

  // --- Diagnostics (Settings' Diagnostics panel; counts/timestamps/sizes
  // only, never row contents) ---
  source: SharePointDataSource
  cacheSavedAt: string | null
  cacheExpired: boolean
  cacheSourceLastModified: string | null
  cacheWriteStatus: 'written' | 'skipped' | 'failed' | null
  cacheWriteError: string | null
  memoryOnly: boolean
  downloadBytes: number | null
  downloadMs: number | null
  parseMs: number | null
  usedWorker: boolean | null
  lastErrorMessage: string | null
  lastErrorAt: string | null
  revalidateStatus: RevalidateStatus
  // The newest lastModifiedDateTime a metadata check has actually observed
  // from SharePoint — may be newer than `lastModifiedDateTime` above if a
  // background refetch triggered by that check hasn't landed yet.
  liveLastModifiedDateTime: string | null
  staleVsSharePoint: boolean
}

const EMPTY_FILE_STATE: SharePointFileState = {
  rows: null, lastModifiedDateTime: null, refreshedAt: null, loading: false, phase: 'idle', fromCache: false, error: null,
  source: 'none', cacheSavedAt: null, cacheExpired: false, cacheSourceLastModified: null,
  cacheWriteStatus: null, cacheWriteError: null, memoryOnly: false,
  downloadBytes: null, downloadMs: null, parseMs: null, usedWorker: null,
  lastErrorMessage: null, lastErrorAt: null, revalidateStatus: 'idle',
  liveLastModifiedDateTime: null, staleVsSharePoint: false,
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

  // Mirrors `files` for synchronous reads inside the async callbacks below
  // — load()/verifyFreshnessThenFetch() aren't in `files`'s dependency
  // chain (that would recreate them, and every consumer's effect, on
  // every single patch), so they read the current value through this ref
  // instead of closing over a `files` that may be stale by the time an
  // await resumes.
  const filesRef = useRef(files)
  useEffect(() => { filesRef.current = files }, [files])

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

  // The real content fetch + parse + cache write — always does a live
  // network read, regardless of why it was called (first load, explicit
  // refresh, or a freshness check finding SharePoint's version newer than
  // what's shown). `isBackground` only changes which phase/status labels
  // are used, not the work itself.
  const fetchAndStore = useCallback(async (key: SharePointFileKey, accountId: string | undefined, isBackground: boolean) => {
    patchFile(key, {
      loading: true, error: null,
      phase: isBackground ? 'refreshing' : 'downloading',
      revalidateStatus: isBackground ? 'running' : 'idle',
    })
    try {
      const accessToken = await getAccessToken()

      const readOnce = () => readNamedFileWithMetadata(accessToken, key)
      let buffer: ArrayBuffer
      let metadata: { lastModifiedDateTime: string }
      const downloadStart = Date.now()
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
      const downloadMs = Date.now() - downloadStart

      patchFile(key, { phase: 'parsing' })
      const parseStart = Date.now()
      const rows = await PARSERS[key](buffer)
      const parseMs = Date.now() - parseStart
      const decodeMeta = getLastDecodeMeta(key)

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
        source: 'network',
        downloadBytes: buffer.byteLength,
        downloadMs,
        parseMs,
        usedWorker: decodeMeta?.usedWorker ?? false,
        liveLastModifiedDateTime: metadata.lastModifiedDateTime,
        staleVsSharePoint: false,
        revalidateStatus: isBackground ? 'finished' : 'idle',
      })
      if (accountId) {
        const writeResult = await writeNokiaCache(key, accountId, rows, metadata.lastModifiedDateTime)
        patchFile(key, {
          cacheWriteStatus: writeResult.status,
          cacheWriteError: writeResult.error ?? null,
          memoryOnly: isMemoryOnlyFile(key),
        })
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Load failed.'
      const now = new Date().toISOString()
      // Deliberately does NOT clear `rows` — whatever was showing (cache
      // or a previous live load) stays visible rather than blanking the
      // page, but `error`/`revalidateStatus` make the failure visible
      // instead of leaving stale data looking like a normal success. Also
      // deliberately NOT added to loadedKeys on failure, so the next
      // ensureLoaded() call retries instead of staying broken forever.
      patchFile(key, {
        loading: false, phase: 'error', error: message,
        lastErrorMessage: message, lastErrorAt: now,
        revalidateStatus: isBackground ? 'failed' : 'idle',
      })
    } finally {
      delete inFlight.current[key]
    }
  }, [getAccessToken, patchFile])

  // Cheap metadata-only check (no content download) — compares SharePoint's
  // current lastModifiedDateTime against `compareAgainst` (whatever the
  // caller is currently showing, from cache or memory). Only triggers the
  // real, expensive fetchAndStore when they differ — a null compareAgainst
  // (unknown version, e.g. a cache entry saved before this field existed)
  // always counts as "differs," failing open to a real fetch rather than
  // trusting an unverifiable cache.
  const verifyFreshnessThenFetch = useCallback(async (
    key: SharePointFileKey, accountId: string | undefined, compareAgainst: string | null,
  ) => {
    patchFile(key, { revalidateStatus: 'running' })
    try {
      const accessToken = await getAccessToken()
      const meta = await getNamedFileMetadata(accessToken, key)
      if (compareAgainst && meta.lastModifiedDateTime === compareAgainst) {
        patchFile(key, {
          liveLastModifiedDateTime: meta.lastModifiedDateTime,
          staleVsSharePoint: false,
          revalidateStatus: 'finished',
        })
        return
      }
      // SharePoint's version differs (or we don't know the cached
      // version) — the data on screen is confirmed stale until the real
      // fetch below replaces it.
      patchFile(key, { liveLastModifiedDateTime: meta.lastModifiedDateTime, staleVsSharePoint: true })
      await fetchAndStore(key, accountId, true)
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Freshness check failed.'
      patchFile(key, {
        revalidateStatus: 'failed', error: message,
        lastErrorMessage: message, lastErrorAt: new Date().toISOString(),
      })
    } finally {
      delete inFlight.current[key]
    }
  }, [getAccessToken, patchFile, fetchAndStore])

  const load = useCallback((key: SharePointFileKey, isRefresh: boolean): Promise<void> => {
    const existing = inFlight.current[key]
    if (existing) return existing

    const alreadyLoadedThisSession = loadedKeys.current.has(key)
    const accountId = accounts[0]?.homeAccountId

    // Already have real data in memory this session and this isn't an
    // explicit refresh — don't reload the whole UI, just verify freshness
    // in the background (see the bug-fix note at the top of this file).
    if (!isRefresh && alreadyLoadedThisSession) {
      patchFile(key, { source: 'memory' })
      const current = filesRef.current[key]
      const run = verifyFreshnessThenFetch(key, accountId, current.lastModifiedDateTime)
      inFlight.current[key] = run
      return run
    }

    const run = (async () => {
      // Stale-while-revalidate: a non-expired cached entry renders
      // immediately (phase 'refreshing', fromCache true) while a freshness
      // check + (if needed) the real fetch runs in the background —
      // never shown for an explicit Refresh click, only a fresh load.
      if (!isRefresh && accountId) {
        const cached = await readNokiaCache(key, accountId)
        if (cached) {
          patchFile(key, {
            rows: cached.rows,
            lastModifiedDateTime: cached.sourceLastModified,
            refreshedAt: cached.savedAt,
            loading: true, phase: 'refreshing', fromCache: true, error: null,
            source: 'cache', cacheSavedAt: cached.savedAt, cacheExpired: false,
            cacheSourceLastModified: cached.sourceLastModified,
          })
          loadedKeys.current.add(key)
          await verifyFreshnessThenFetch(key, accountId, cached.sourceLastModified)
          return
        }
      }
      // No usable cache (or an explicit refresh) — go straight to a real fetch.
      await fetchAndStore(key, accountId, isRefresh)
    })()

    inFlight.current[key] = run
    run.finally(() => { delete inFlight.current[key] })
    return run
  }, [accounts, patchFile, verifyFreshnessThenFetch, fetchAndStore])

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
 * currently the on-device cache vs. a live response, error state, a full
 * set of diagnostics fields (Settings' Diagnostics panel), and a refresh()
 * to force a fresh read. Triggers the file's load on first use; every
 * later mount re-verifies freshness cheaply instead of trusting "already
 * loaded this session" forever (see the bug-fix note at the top of this
 * file) — nothing is written to disk/Supabase except through
 * app/lib/nokiaCache.ts's approved store. */
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

/** Like useSharePointFile, but does NOT trigger a load on mount — the
 * caller decides when, via the returned `load()`. For a file a page only
 * needs some of the time (e.g. GC Call View's Reports tab needing 'cr',
 * by far the largest file), so it isn't eagerly fetched just because the
 * page itself mounted — only once that tab is actually opened. */
export function useSharePointLazyFile(key: SharePointFileKey): SharePointFileState & { refresh: () => void; load: () => void } {
  const ctx = useContext(SharePointDataContext)
  if (!ctx) throw new Error('useSharePointLazyFile must be used within SharePointDataProvider')
  const { files, ensureLoaded, refresh } = ctx
  return useMemo(
    () => ({ ...files[key], refresh: () => refresh(key), load: () => ensureLoaded(key) }),
    [files, key, ensureLoaded, refresh],
  )
}
