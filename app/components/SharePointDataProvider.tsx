'use client'

import { createContext, useCallback, useContext, useEffect, useMemo, useRef, useState, type ReactNode } from 'react'
import { useMsal } from '@azure/msal-react'
import { graphLoginRequest } from '../lib/msalConfig'
import { useMsalReady } from '../lib/msalReady'
import {
  readNamedFileWithMetadata, getNamedFileMetadata, timeoutMsFor, type SharePointFileKey,
  SharePointRateLimitError, SharePointCancelledError,
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
//
// Bug fix, 11/2026 ("SharePoint refresh gets stuck"): the freshness check
// this added runs a real Graph call on every ensureLoaded() — and until
// this fix, NOTHING in the whole chain (the Graph fetch itself, the
// downloadUrl content fetch, the worker decode, acquireTokenSilent) had any
// timeout. A single hung network call on Nokia's locked-down connection
// left `inFlight` permanently populated for that key (its `finally` never
// ran, because the async function awaiting the hung call never resumed),
// which meant every later ensureLoaded()/refresh() call — including every
// click of the Refresh button — just rejoined that same dead promise
// forever. Every Graph/content/worker/token call below now has its own
// timeout (see sharepoint.ts's timeoutMsFor and getAccessToken below), each
// file's in-flight/abort state is cleaned up via a generation counter
// (`genRef`) that a superseded attempt can check before touching state a
// newer attempt already owns, and refresh()/cancel() can always abort
// whatever's currently running for a key, so the Refresh button is never
// stuck waiting on a hang it has no way to interrupt.

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
  // When the CURRENT phase started — lets the UI show "stuck for Ns" /
  // elapsed-seconds instead of a single opaque spinner, and lets a Refresh
  // click decide whether the in-flight attempt is worth aborting.
  phaseStartedAt: string | null
  // Set on every successful CONTENT fetch (not a cache hit, not a freshness
  // check alone) — independent of whatever `source`/`fromCache` currently
  // show, which can read 'cache' long after the underlying data was last
  // actually pulled live if every freshness check since has found no
  // change. Bug fix, 11/2026: Diagnostics used to derive this from
  // `source === 'network' ? refreshedAt : null`, which read as a permanent
  // dash on a page mostly served from a freshness-confirmed cache.
  lastSuccessfulNetworkAt: string | null
}

const EMPTY_FILE_STATE: SharePointFileState = {
  rows: null, lastModifiedDateTime: null, refreshedAt: null, loading: false, phase: 'idle', fromCache: false, error: null,
  source: 'none', cacheSavedAt: null, cacheExpired: false, cacheSourceLastModified: null,
  cacheWriteStatus: null, cacheWriteError: null, memoryOnly: false,
  downloadBytes: null, downloadMs: null, parseMs: null, usedWorker: null,
  lastErrorMessage: null, lastErrorAt: null, revalidateStatus: 'idle',
  liveLastModifiedDateTime: null, staleVsSharePoint: false,
  phaseStartedAt: null, lastSuccessfulNetworkAt: null,
}

interface SharePointDataContextValue {
  files: Record<SharePointFileKey, SharePointFileState>
  ensureLoaded: (key: SharePointFileKey) => void
  refresh: (key: SharePointFileKey) => void
  // Aborts whatever's currently running for this key WITHOUT starting a
  // new attempt (unlike refresh, which aborts-then-restarts) — the UI's
  // "Cancel" link while a load is in progress.
  cancel: (key: SharePointFileKey) => void
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
  // The AbortController actually driving whatever's currently running for a
  // key — refresh()/cancel() fire it to interrupt a hung or unwanted
  // attempt. Overwritten (not merged) as an attempt moves between phases
  // (e.g. a freshness check's metadata call handing off to the real
  // content fetch) — it always points at whichever call is actually live.
  const abortControllers = useRef<Partial<Record<SharePointFileKey, AbortController>>>({})
  // Bumped once per NEW load() call for a key. Every async function started
  // by that call captures the value at its start and checks it again before
  // touching `inFlight`/`abortControllers` in its own `finally` — if the
  // number's moved on (a newer attempt superseded this one, e.g. via
  // Refresh), a late-settling old attempt skips cleanup instead of deleting
  // state a newer attempt already owns.
  const genRef = useRef<Partial<Record<SharePointFileKey, number>>>({})
  const msalReady = useMsalReady()

  const patchFile = useCallback((key: SharePointFileKey, patch: Partial<SharePointFileState>) => {
    setFiles(prev => ({ ...prev, [key]: { ...prev[key], ...patch } }))
  }, [])

  const SILENT_TOKEN_TIMEOUT_MS = 15_000

  const getAccessToken = useCallback(async (): Promise<string> => {
    const account = accounts[0]
    if (!account) throw new Error('Not signed in with a Nokia account.')
    console.time('[sp-timing] msal acquireTokenSilent')
    let timedOut = false
    let timer: ReturnType<typeof setTimeout>
    const timeoutPromise = new Promise<never>((_, reject) => {
      timer = setTimeout(() => { timedOut = true; reject(new Error('silent-token-timeout')) }, SILENT_TOKEN_TIMEOUT_MS)
    })
    try {
      const result = await Promise.race([
        instance.acquireTokenSilent({ ...graphLoginRequest, account }),
        timeoutPromise,
      ])
      console.timeEnd('[sp-timing] msal acquireTokenSilent')
      return result.accessToken
    } catch {
      console.timeEnd('[sp-timing] msal acquireTokenSilent')
      if (timedOut) {
        // Bug fix, 11/2026: acquireTokenSilent had no timeout of its own —
        // MSAL doesn't accept an AbortSignal for it, so this can't be
        // cancelled outright, but giving up waiting after 15s and surfacing
        // a clear "sign in again" message beats leaving every caller
        // (and therefore inFlight) stuck on it indefinitely. Deliberately
        // does NOT fall through to the interactive popup fallback below —
        // a popup that never got a user gesture can itself hang/get
        // blocked, which would just trade one indefinite wait for another.
        throw new Error('Sign in again — SharePoint sign-in did not respond in 15s.')
      }
      // Silent acquisition failed outright (not a timeout) — expired/needs
      // re-consent — fall back to an interactive popup once rather than
      // leaving the caller permanently stuck. Only hit if initialize()
      // couldn't restore a usable session at all (see MsalProviderInner).
      console.time('[sp-timing] msal acquireTokenPopup')
      const result = await instance.acquireTokenPopup(graphLoginRequest)
      console.timeEnd('[sp-timing] msal acquireTokenPopup')
      return result.accessToken
    } finally {
      clearTimeout(timer!)
    }
  }, [instance, accounts])

  // The real content fetch + parse + cache write — always does a live
  // network read, regardless of why it was called (first load, explicit
  // refresh, or a freshness check finding SharePoint's version newer than
  // what's shown). `isBackground` only changes which phase/status labels
  // are used, not the work itself. `myGen` is this attempt's generation
  // number (captured by the caller from genRef at the moment it started) —
  // used ONLY to decide, in the finally block, whether this attempt is
  // still the current one for `key` before touching inFlight/abortControllers
  // (see genRef's doc comment above).
  const fetchAndStore = useCallback(async (
    key: SharePointFileKey, accountId: string | undefined, isBackground: boolean, myGen: number,
  ) => {
    const controller = new AbortController()
    abortControllers.current[key] = controller
    patchFile(key, {
      loading: true, error: null,
      phase: isBackground ? 'refreshing' : 'downloading',
      phaseStartedAt: new Date().toISOString(),
      revalidateStatus: isBackground ? 'running' : 'idle',
    })
    try {
      const accessToken = await getAccessToken()

      const control = { signal: controller.signal, timeoutMs: timeoutMsFor(key) }
      const readOnce = () => readNamedFileWithMetadata(accessToken, key, control)
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

      patchFile(key, { phase: 'parsing', phaseStartedAt: new Date().toISOString() })
      const parseStart = Date.now()
      const rows = await PARSERS[key](buffer)
      const parseMs = Date.now() - parseStart
      const decodeMeta = getLastDecodeMeta(key)

      patchFile(key, { phase: 'building', phaseStartedAt: new Date().toISOString() })
      loadedKeys.current.add(key)
      const now = new Date().toISOString()
      patchFile(key, {
        rows,
        lastModifiedDateTime: metadata.lastModifiedDateTime,
        refreshedAt: now,
        loading: false,
        phase: 'done',
        phaseStartedAt: now,
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
        lastSuccessfulNetworkAt: now,
        // A later success clears whatever error this file was last showing
        // — a stale "last error" sitting in Diagnostics next to data that's
        // now loading fine reads as an ongoing problem when it isn't one.
        lastErrorMessage: null, lastErrorAt: null,
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
      // A user-initiated Cancel isn't a failure worth alarming over — quietly
      // return to idle (whatever was showing stays showing) instead of
      // leaving a red error banner for something the user asked for.
      const cancelled = err instanceof SharePointCancelledError
      const message = err instanceof Error ? err.message : 'Load failed.'
      const now = new Date().toISOString()
      // Deliberately does NOT clear `rows` — whatever was showing (cache
      // or a previous live load) stays visible rather than blanking the
      // page, but `error`/`revalidateStatus` make the failure visible
      // instead of leaving stale data looking like a normal success. Also
      // deliberately NOT added to loadedKeys on failure, so the next
      // ensureLoaded() call retries instead of staying broken forever.
      patchFile(key, {
        loading: false, phase: cancelled ? 'idle' : 'error', phaseStartedAt: now,
        error: cancelled ? null : message,
        lastErrorMessage: cancelled ? null : message,
        lastErrorAt: cancelled ? null : now,
        revalidateStatus: isBackground ? (cancelled ? 'idle' : 'failed') : 'idle',
      })
    } finally {
      if (genRef.current[key] === myGen) {
        delete inFlight.current[key]
        delete abortControllers.current[key]
      }
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
    key: SharePointFileKey, accountId: string | undefined, compareAgainst: string | null, myGen: number,
  ) => {
    const controller = new AbortController()
    abortControllers.current[key] = controller
    patchFile(key, { revalidateStatus: 'running', phaseStartedAt: new Date().toISOString() })
    try {
      const accessToken = await getAccessToken()
      // Always the cheap 45s budget here regardless of file — this is a
      // tiny metadata-only call even for CR.
      const meta = await getNamedFileMetadata(accessToken, key, { signal: controller.signal, timeoutMs: 45_000 })
      if (compareAgainst && meta.lastModifiedDateTime === compareAgainst) {
        // Bug fix, 11/2026: this branch used to leave `loading`/`phase`/
        // `fromCache` exactly as the caller set them before calling here —
        // for the cache-first load path, that's loading:true,
        // phase:'refreshing', fromCache:true, and NOTHING ever cleared them
        // on this success path. A freshness check that found the cache
        // already current (the common, fast, everything-working case) left
        // the Dashboard tracker box (and every other SharePointSourceLine)
        // showing "Refreshing from SharePoint…"/"refreshing…" forever,
        // with no hang or error involved at all — confirmed as the likely
        // dominant cause of the reported stuck-UI, independent of the
        // no-timeout issue this same fix round also addresses.
        const now = new Date().toISOString()
        patchFile(key, {
          liveLastModifiedDateTime: meta.lastModifiedDateTime,
          staleVsSharePoint: false,
          revalidateStatus: 'finished',
          loading: false, phase: 'done', phaseStartedAt: now,
          error: null, lastErrorMessage: null, lastErrorAt: null,
        })
        return
      }
      // SharePoint's version differs (or we don't know the cached
      // version) — the data on screen is confirmed stale until the real
      // fetch below replaces it.
      patchFile(key, { liveLastModifiedDateTime: meta.lastModifiedDateTime, staleVsSharePoint: true })
      await fetchAndStore(key, accountId, true, myGen)
    } catch (err) {
      const cancelled = err instanceof SharePointCancelledError
      const message = err instanceof Error ? err.message : 'Freshness check failed.'
      const now = new Date().toISOString()
      patchFile(key, {
        loading: false, phase: cancelled ? 'idle' : 'error', phaseStartedAt: now,
        revalidateStatus: cancelled ? 'idle' : 'failed',
        error: cancelled ? null : message,
        lastErrorMessage: cancelled ? null : message,
        lastErrorAt: cancelled ? null : now,
      })
    } finally {
      if (genRef.current[key] === myGen) {
        delete inFlight.current[key]
        delete abortControllers.current[key]
      }
    }
  }, [getAccessToken, patchFile, fetchAndStore])

  const load = useCallback((key: SharePointFileKey, isRefresh: boolean): Promise<void> => {
    if (!msalReady) {
      // MSAL hasn't finished restoring its cached session yet (see
      // msalReady's doc comment) — starting a load now would just throw
      // "Not signed in" against the stub context. The effect that calls
      // ensureLoaded() re-runs once msalReady flips true (load's identity
      // changes because msalReady is in its own dependency array below),
      // so nothing needs to be queued here.
      return Promise.resolve()
    }

    if (isRefresh) {
      // Explicit Refresh always wins — abort whatever's currently running
      // for this key so a hung or merely-slow prior attempt can never make
      // Refresh a no-op. Safe even if nothing's running (abort on an
      // already-settled/no controller is a no-op).
      abortControllers.current[key]?.abort(new DOMException('Refresh requested a new attempt.', 'AbortError'))
    } else {
      const existing = inFlight.current[key]
      if (existing) return existing
    }

    const myGen = (genRef.current[key] ?? 0) + 1
    genRef.current[key] = myGen

    const alreadyLoadedThisSession = loadedKeys.current.has(key)
    const accountId = accounts[0]?.homeAccountId

    // Already have real data in memory this session and this isn't an
    // explicit refresh — don't reload the whole UI, just verify freshness
    // in the background (see the bug-fix note at the top of this file).
    if (!isRefresh && alreadyLoadedThisSession) {
      patchFile(key, { source: 'memory' })
      const current = filesRef.current[key]
      const run = verifyFreshnessThenFetch(key, accountId, current.lastModifiedDateTime, myGen)
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
            loading: true, phase: 'refreshing', phaseStartedAt: new Date().toISOString(), fromCache: true, error: null,
            source: 'cache', cacheSavedAt: cached.savedAt, cacheExpired: false,
            cacheSourceLastModified: cached.sourceLastModified,
          })
          loadedKeys.current.add(key)
          await verifyFreshnessThenFetch(key, accountId, cached.sourceLastModified, myGen)
          return
        }
      }
      // No usable cache (or an explicit refresh) — go straight to a real fetch.
      await fetchAndStore(key, accountId, isRefresh, myGen)
    })()

    inFlight.current[key] = run
    run.finally(() => {
      if (genRef.current[key] === myGen) delete inFlight.current[key]
    })
    return run
  }, [accounts, patchFile, verifyFreshnessThenFetch, fetchAndStore, msalReady])

  const ensureLoaded = useCallback((key: SharePointFileKey) => { void load(key, false) }, [load])
  const refresh = useCallback((key: SharePointFileKey) => { void load(key, true) }, [load])
  const cancel = useCallback((key: SharePointFileKey) => {
    abortControllers.current[key]?.abort(new DOMException('Cancelled by user.', 'AbortError'))
  }, [])

  const value = useMemo<SharePointDataContextValue>(
    () => ({ files, ensureLoaded, refresh, cancel }),
    [files, ensureLoaded, refresh, cancel],
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
export function useSharePointFile(key: SharePointFileKey): SharePointFileState & { refresh: () => void; cancel: () => void } {
  const ctx = useContext(SharePointDataContext)
  if (!ctx) throw new Error('useSharePointFile must be used within SharePointDataProvider')
  const { files, ensureLoaded, refresh, cancel } = ctx

  // An effect, not a direct call during render — ensureLoaded ends up
  // calling setFiles (via patchFile) on the PROVIDER, and updating one
  // component's state synchronously from inside another component's
  // render is exactly what React's rules disallow. ensureLoaded is
  // ref-gated/idempotent regardless, so this only ever starts one load
  // per key no matter how many components call this hook for it.
  useEffect(() => { ensureLoaded(key) }, [key, ensureLoaded])

  return useMemo(() => ({ ...files[key], refresh: () => refresh(key), cancel: () => cancel(key) }), [files, key, refresh, cancel])
}

/** Like useSharePointFile, but does NOT trigger a load on mount — the
 * caller decides when, via the returned `load()`. For a file a page only
 * needs some of the time (e.g. GC Call View's Reports tab needing 'cr',
 * by far the largest file), so it isn't eagerly fetched just because the
 * page itself mounted — only once that tab is actually opened. */
export function useSharePointLazyFile(key: SharePointFileKey): SharePointFileState & { refresh: () => void; cancel: () => void; load: () => void } {
  const ctx = useContext(SharePointDataContext)
  if (!ctx) throw new Error('useSharePointLazyFile must be used within SharePointDataProvider')
  const { files, ensureLoaded, refresh, cancel } = ctx
  return useMemo(
    () => ({ ...files[key], refresh: () => refresh(key), cancel: () => cancel(key), load: () => ensureLoaded(key) }),
    [files, key, ensureLoaded, refresh, cancel],
  )
}
