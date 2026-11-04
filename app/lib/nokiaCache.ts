// The ONE approved on-device cache for Nokia data (SharePoint migration perf
// work, 11/2026 spec). Nokia data may live in memory and in this IndexedDB
// store ONLY — never localStorage, sessionStorage, the Cache API, a service
// worker, or Supabase/Railway. `indexedDB` must appear ONLY in this file —
// every other module that wants cached Nokia rows goes through the
// functions below, never touches the browser API directly.
//
// Stale-while-revalidate: SharePointDataProvider reads a non-expired entry
// here immediately (instant render) while it kicks off a real SharePoint
// fetch in the background to replace it. Entries expire after 12 hours and
// are deleted — not just ignored — the moment something tries to read them
// past that age, so a stale entry can never render even briefly.
//
// Scoped per signed-in account (keyed by MSAL's homeAccountId, not just the
// file key) so one Nokia account's cached rows are never reachable by a
// different account signed into the same browser profile. Cleared entirely
// on sign-out (see MsalProviderInner's LOGOUT_SUCCESS handler) and via the
// "Clear local data" button in Settings.

const DB_NAME = 'ciege-nokia-cache'
const DB_VERSION = 1
const STORE_NAME = 'files'
const EXPIRY_MS = 12 * 60 * 60 * 1000 // 12 hours

// CR is never written to IndexedDB — it's by far the largest parsed set
// (the source file is ~30MB; a 30MB structured-clone write is the kind of
// thing a locked-down/managed corporate browser profile can throttle,
// quota-restrict, or silently fail in ways a normal desktop Chrome profile
// wouldn't — and there's no way to verify that from this environment.
// Rather than risk a write that succeeds unreliably (or "succeeds" but
// serves a corrupted/partial entry later), CR stays memory-only: refetched
// every session, never persisted. Settings' Diagnostics panel says so
// explicitly. Every other file is small enough that this isn't a concern.
const MEMORY_ONLY_KEYS = new Set(['cr'])

export interface NokiaCacheEntry {
  rows: unknown[][]
  savedAt: string // ISO timestamp — when this entry was written
  // The SharePoint file's Graph lastModifiedDateTime AT THE TIME this entry
  // was written — lets a reader tell "is this cache entry from an older
  // version of the source file than what's live now," independent of how
  // long ago it was saved. null only for entries written before this field
  // existed (treated as unknown-freshness, not implicitly stale).
  sourceLastModified: string | null
}

function killSwitchOn(): boolean {
  return process.env.NEXT_PUBLIC_NOKIA_CACHE === 'off'
}

function mobileOverrideOn(): boolean {
  return process.env.NEXT_PUBLIC_NOKIA_CACHE_MOBILE === 'on'
}

// Coarse pointer (touch) OR a mobile/tablet user-agent token — either signal
// alone is enough to call it a phone/tablet, since a real desktop/laptop
// trips neither under normal use.
function isMobileDevice(): boolean {
  if (typeof window === 'undefined') return false
  try {
    const coarsePointer = typeof window.matchMedia === 'function' && window.matchMedia('(pointer: coarse)').matches
    const ua = typeof navigator !== 'undefined' ? navigator.userAgent || '' : ''
    const mobileUa = /Android|iPhone|iPad|iPod|Mobile|Tablet/i.test(ua)
    return coarsePointer || mobileUa
  } catch {
    return false
  }
}

function cacheEnabled(fileKey: string): boolean {
  if (typeof window === 'undefined') return false
  if (killSwitchOn()) return false
  if (MEMORY_ONLY_KEYS.has(fileKey)) return false
  if (isMobileDevice() && !mobileOverrideOn()) return false
  if (!('indexedDB' in window) || !window.indexedDB) return false
  return true
}

/** Whether this file is deliberately never written to IndexedDB (see
 * MEMORY_ONLY_KEYS above) — exported so the Diagnostics panel can say so
 * explicitly instead of just showing "no cache entry" indistinguishably
 * from "cache write failed" or "never loaded yet." */
export function isMemoryOnlyFile(fileKey: string): boolean {
  return MEMORY_ONLY_KEYS.has(fileKey)
}

function keyFor(fileKey: string, accountId: string): string {
  return `${fileKey}::${accountId}`
}

let dbPromise: Promise<IDBDatabase> | null = null

function openDb(): Promise<IDBDatabase> {
  if (dbPromise) return dbPromise
  dbPromise = new Promise<IDBDatabase>((resolve, reject) => {
    try {
      const req = indexedDB.open(DB_NAME, DB_VERSION)
      req.onupgradeneeded = () => {
        const db = req.result
        if (!db.objectStoreNames.contains(STORE_NAME)) db.createObjectStore(STORE_NAME)
      }
      req.onsuccess = () => resolve(req.result)
      req.onerror = () => reject(req.error ?? new Error('IndexedDB open failed'))
    } catch (err) {
      reject(err instanceof Error ? err : new Error('IndexedDB unavailable'))
    }
  })
  // A failed open should never poison future attempts (e.g. a one-off
  // private-browsing restriction) — let the next call try again fresh.
  dbPromise.catch(() => { dbPromise = null })
  return dbPromise
}

/** Pure — exported separately so the 12-hour expiry rule itself can be unit
 * tested without touching IndexedDB at all. */
export function isCacheEntryExpired(savedAt: string, now: number = Date.now()): boolean {
  const savedMs = new Date(savedAt).getTime()
  if (!Number.isFinite(savedMs)) return true
  return now - savedMs > EXPIRY_MS
}

/** Returns the cached entry for this file+account, or null if there isn't
 * one, it's expired (and has just been deleted), the cache is disabled
 * (kill switch / mobile / IndexedDB unavailable / memory-only file), or
 * anything about the read failed — every failure path is silent, falling
 * back to "no cache," never a thrown error the caller has to handle. */
export async function readNokiaCache(fileKey: string, accountId: string): Promise<NokiaCacheEntry | null> {
  if (!accountId || !cacheEnabled(fileKey)) return null
  const key = keyFor(fileKey, accountId)
  try {
    const db = await openDb()
    const entry = await new Promise<NokiaCacheEntry | undefined>((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readonly')
      const req = tx.objectStore(STORE_NAME).get(key)
      req.onsuccess = () => resolve(req.result as NokiaCacheEntry | undefined)
      req.onerror = () => reject(req.error ?? new Error('IndexedDB read failed'))
    })
    if (!entry) return null
    if (isCacheEntryExpired(entry.savedAt)) {
      await deleteKey(key)
      return null
    }
    return entry
  } catch {
    return null
  }
}

/** Diagnostics-only — reads the raw stored entry (if any) WITHOUT applying
 * or acting on expiry, so Settings' Diagnostics panel can show "there's a
 * cache entry and here's whether it's expired" even when nothing has
 * actually tried to read (and thus auto-delete) it recently. Never used by
 * any real data-loading path — readNokiaCache above is what every page
 * actually gets served from. */
export async function peekNokiaCache(fileKey: string, accountId: string): Promise<{ savedAt: string; sourceLastModified: string | null; expired: boolean } | null> {
  if (!accountId) return null
  const key = keyFor(fileKey, accountId)
  try {
    const db = await openDb()
    const entry = await new Promise<NokiaCacheEntry | undefined>((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readonly')
      const req = tx.objectStore(STORE_NAME).get(key)
      req.onsuccess = () => resolve(req.result as NokiaCacheEntry | undefined)
      req.onerror = () => reject(req.error ?? new Error('IndexedDB read failed'))
    })
    if (!entry) return null
    return { savedAt: entry.savedAt, sourceLastModified: entry.sourceLastModified, expired: isCacheEntryExpired(entry.savedAt) }
  } catch {
    return null
  }
}

async function deleteKey(key: string): Promise<void> {
  try {
    const db = await openDb()
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readwrite')
      tx.objectStore(STORE_NAME).delete(key)
      tx.oncomplete = () => resolve()
      tx.onerror = () => reject(tx.error ?? new Error('IndexedDB delete failed'))
    })
  } catch {
    // best-effort
  }
}

export interface WriteCacheResult {
  // 'written': succeeded. 'skipped': disabled/memory-only/not signed in —
  // not an error, just never attempted. 'failed': actually attempted and
  // threw (e.g. IndexedDB quota, a managed-browser storage restriction) —
  // this is the case that used to be swallowed with no caller visibility
  // at all; now it's returned, not just caught, so Diagnostics can show it.
  status: 'written' | 'skipped' | 'failed'
  error?: string
}

/** Writes/replaces the cached entry for this file+account. Never throws —
 * the caller already has the freshly-fetched rows in memory regardless, so
 * a failed cache write never blocks or degrades the page itself — but
 * DOES report success/skip/failure back via the return value instead of
 * swallowing it silently, so Diagnostics can show real write status. */
export async function writeNokiaCache(
  fileKey: string, accountId: string, rows: unknown[][], sourceLastModified: string | null,
): Promise<WriteCacheResult> {
  if (!accountId || !cacheEnabled(fileKey)) return { status: 'skipped' }
  const key = keyFor(fileKey, accountId)
  const entry: NokiaCacheEntry = { rows, savedAt: new Date().toISOString(), sourceLastModified }
  try {
    const db = await openDb()
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readwrite')
      tx.objectStore(STORE_NAME).put(entry, key)
      tx.oncomplete = () => resolve()
      tx.onerror = () => reject(tx.error ?? new Error('IndexedDB write failed'))
    })
    return { status: 'written' }
  } catch (err) {
    return { status: 'failed', error: err instanceof Error ? err.message : 'IndexedDB write failed' }
  }
}

/** Deletes every cached entry, for every file and every account — called on
 * sign-out and from the "Clear local data" button in Settings. Safe to call
 * even if the cache was never enabled/populated. */
export async function clearNokiaCache(): Promise<void> {
  try {
    const db = await openDb()
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readwrite')
      tx.objectStore(STORE_NAME).clear()
      tx.oncomplete = () => resolve()
      tx.onerror = () => reject(tx.error ?? new Error('IndexedDB clear failed'))
    })
  } catch {
    // best-effort — nothing to clear if it never worked in the first place
  }
}
