// SharePoint read service — Stage 1 of the manual-upload-to-SharePoint
// migration. Reads a file straight out of Nokia's SharePoint via Microsoft
// Graph, using whichever Nokia account the signed-in user authenticated
// with (delegated access — this app never has its own standing access to
// a user's files, only what that user could already open themselves).
//
// Permission model: Sites.Selected (see app/lib/msalConfig.ts) — the app
// has NO site access by default, only whichever specific sites Nokia IT
// has separately granted via the Graph site-permissions API. Under
// Sites.Selected, Graph's combined "resolve-site-and-read-item-by-path-in-
// one-call" shortcut (/sites/{hostname}:{path}:/drive/root:{item}:/content)
// doesn't carry the per-site grant through correctly — reading only works
// as two explicit steps: resolve the site's real ID first (GET
// /sites/{hostname}:/{path}), then address the drive by that ID (GET
// /sites/{siteId}/drive/root:{item}:/content). A 403 at either step most
// likely means "this site hasn't been granted to the app yet," not
// necessarily a problem with the signed-in user — see
// SharePointAccessDeniedError.
//
// Every page now reads through one of the FILES entries below instead of
// the old loadTrackerSnapshot()/chunked report_snapshots calls.

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0'

// NAMDS removed, 11/3/2026 — spo moved to viaeroWireless (same site as
// everything else) and nothing else referenced NAMDS.
export const SITES = {
  viaeroWireless: 'nokianam.sharepoint.com/sites/viaero-wireless2',
} as const

// Paths are relative to the DRIVE (the document library itself), not the
// SharePoint site — "Shared Documents" as seen in the SharePoint UI/URL
// IS the drive, not a folder inside it, so it's never part of the
// drive-root-relative path Graph expects (GET /sites/{id}/drive/root:
// {path}:/content). Confirmed by CJ, 11/3/2026 — the original paths here
// included that leading segment and every read 404'd.
export const FILES = {
  tracker: {
    site: SITES.viaeroWireless,
    path: '/Team PM/CJ Ojeda/Viaero - Construction Schedule Tracker - WIP 2.0.xlsx',
  },
  spo: {
    site: SITES.viaeroWireless,
    path: '/Team PM/CJ Ojeda/SPO Report/NDPd_NAM_003_Viaero_SPO_Report.xlsx',
  },
  decom: {
    site: SITES.viaeroWireless,
    path: '/Team Decom/Decom Tracker Updates/Nokia Decom Status_Report_LIVE - NEW.xlsx',
  },
  cr: {
    site: SITES.viaeroWireless,
    path: '/Team CQT/Viaero - CR Claims Tracker.xlsx',
  },
  scop: {
    site: SITES.viaeroWireless,
    path: '/Team Close Out/Viaero COP Tracker_LIVE.xlsx',
  },
} as const

export type SharePointFileKey = keyof typeof FILES

// Graph's path-based site addressing takes "{hostname}:{server-relative-path}"
// as one segment — SITES values above are "{hostname}/{path}", so split on
// the first slash to build that segment correctly.
function siteSegment(site: string): string {
  const slash = site.indexOf('/')
  const hostname = site.slice(0, slash)
  const sitePath = site.slice(slash) // includes the leading '/'
  return `${hostname}:${sitePath}`
}

// Every thrown error below carries the raw Graph status/code/message, not
// just a generic "couldn't do it" string — so a 403 (not granted yet), a
// 404 (wrong path), and anything else Graph returns are distinguishable
// from the error text alone, not just from which error class it is.
export interface GraphErrorDetail {
  status: number
  code: string
  message: string
}

function formatGraphDetail(d: GraphErrorDetail): string {
  return `Graph ${d.status}${d.code ? ` ${d.code}` : ''}: ${d.message || '(no error message in response body)'}`
}

async function parseGraphError(res: Response): Promise<GraphErrorDetail> {
  try {
    const body = await res.json()
    return { status: res.status, code: body?.error?.code ?? '', message: body?.error?.message ?? '' }
  } catch {
    const text = await res.text().catch(() => '')
    return { status: res.status, code: '', message: text.slice(0, 300) }
  }
}

export class SharePointAuthError extends Error {
  constructor(public detail: GraphErrorDetail) {
    super(`SharePoint session expired — sign in again. (${formatGraphDetail(detail)})`)
    this.name = 'SharePointAuthError'
  }
}

export class SharePointNotFoundError extends Error {
  constructor(public fileKey: string, public filePath: string, public detail: GraphErrorDetail) {
    super(`Could not find "${fileKey}" on SharePoint at: ${filePath} (${formatGraphDetail(detail)})`)
    this.name = 'SharePointNotFoundError'
  }
}

// Under Sites.Selected, a 403 on an otherwise-valid, signed-in request
// almost always means the TARGET SITE hasn't been granted to this app yet
// (a separate IT action per site, not something the signed-in user's own
// permissions control) — distinct enough from a generic failure to call
// out explicitly rather than let it read as "something's broken."
export class SharePointAccessDeniedError extends Error {
  constructor(public siteUrl: string, public detail: GraphErrorDetail) {
    super(`Access denied for site "${siteUrl}" — this site may not be granted to the app yet (Sites.Selected requires each site to be explicitly approved). (${formatGraphDetail(detail)})`)
    this.name = 'SharePointAccessDeniedError'
  }
}

export class SharePointRateLimitError extends Error {
  constructor(public retryAfterSeconds: number, public detail: GraphErrorDetail) {
    super(`SharePoint/Graph API rate limited — retry after ${retryAfterSeconds}s. (${formatGraphDetail(detail)})`)
    this.name = 'SharePointRateLimitError'
  }
}

// Catch-all for any other non-2xx Graph response — still carries the full
// status/code/message rather than collapsing to a generic message.
export class GraphApiError extends Error {
  constructor(context: string, public detail: GraphErrorDetail) {
    super(`${context} (${formatGraphDetail(detail)})`)
    this.name = 'GraphApiError'
  }
}

// Bug fix, 11/2026 ("SharePoint refresh gets stuck"): every fetch() call in
// this file used to have no timeout and no AbortController at all — on
// Nokia's locked-down network, a single Graph call (or the direct
// downloadUrl content fetch) that just never resolves or rejects leaves
// whatever awaited it paused forever, with no way for the caller to ever
// know or recover. SharePointDataProvider's in-flight dedupe then keeps
// EVERY later ensureLoaded()/refresh() call joined to that same permanently-
// pending promise — which is exactly why clicking Refresh did nothing.
// Every network call below now goes through fetchWithTimeout: a timeout-
// driven AbortController (45s for everything except CR's own content
// fetch, which gets 180s — see timeoutMsFor) combined with an optional
// caller-supplied AbortSignal, so the provider can both time out AND let a
// user-initiated Cancel/Refresh abort an in-flight call outright.
export class SharePointTimeoutError extends Error {
  constructor(public timeoutMs: number, public context: string) {
    super(`SharePoint did not respond in ${Math.round(timeoutMs / 1000)}s (${context}).`)
    this.name = 'SharePointTimeoutError'
  }
}

export class SharePointCancelledError extends Error {
  constructor(public context: string) {
    super(`Cancelled (${context}).`)
    this.name = 'SharePointCancelledError'
  }
}

export interface FetchControl {
  // Caller-owned signal (e.g. a per-file AbortController the provider can
  // fire on Cancel/Refresh) — combined with this call's own timeout signal,
  // whichever fires first wins.
  signal?: AbortSignal
  timeoutMs: number
}

const DEFAULT_CONTROL: FetchControl = { timeoutMs: 45_000 }

/** 45s for every file's metadata/content calls except CR's own content
 * fetch, which is by far the largest (~30MB) and gets 180s. The cheap
 * metadata-only check (getNamedFileMetadata) always uses the 45s default
 * regardless of file — it's a tiny JSON response either way. */
export function timeoutMsFor(key: SharePointFileKey): number {
  return key === 'cr' ? 180_000 : 45_000
}

function combineSignals(signals: AbortSignal[]): AbortSignal {
  const live = signals.filter(Boolean)
  if (live.length === 1) return live[0]
  const controller = new AbortController()
  for (const s of live) {
    if (s.aborted) { controller.abort(s.reason); break }
    s.addEventListener('abort', () => controller.abort(s.reason), { once: true })
  }
  return controller.signal
}

async function fetchWithTimeout(url: string, init: RequestInit, control: FetchControl, context: string): Promise<Response> {
  const timeoutController = new AbortController()
  let timedOut = false
  const timer = setTimeout(() => { timedOut = true; timeoutController.abort() }, control.timeoutMs)
  const signal = control.signal ? combineSignals([timeoutController.signal, control.signal]) : timeoutController.signal
  try {
    return await fetch(url, { ...init, signal })
  } catch (err) {
    if (timedOut) throw new SharePointTimeoutError(control.timeoutMs, context)
    if (control.signal?.aborted) throw new SharePointCancelledError(context)
    throw err
  } finally {
    clearTimeout(timer)
  }
}

// Site IDs are permanent (don't expire with the token) and this read
// service is called once per file per page load at most — caching avoids
// a redundant resolve call every time the same site's file is read again.
const siteIdCache = new Map<string, string>()

/** Throws a typed, detail-carrying error for any non-ok Graph response. No-op if res.ok. */
async function assertGraphOk(res: Response, ctx: { siteUrl: string; fileKey?: string; filePath?: string; context: string }): Promise<void> {
  if (res.ok) return
  const detail = await parseGraphError(res)
  if (res.status === 401) throw new SharePointAuthError(detail)
  if (res.status === 403) throw new SharePointAccessDeniedError(ctx.siteUrl, detail)
  if (res.status === 404) throw new SharePointNotFoundError(ctx.fileKey ?? ctx.filePath ?? '', ctx.filePath ?? '', detail)
  if (res.status === 429) {
    const retryAfter = Number(res.headers.get('Retry-After')) || 30
    throw new SharePointRateLimitError(retryAfter, detail)
  }
  throw new GraphApiError(ctx.context, detail)
}

async function resolveSiteId(accessToken: string, siteUrl: string, control: FetchControl): Promise<string> {
  const cached = siteIdCache.get(siteUrl)
  if (cached) return cached

  console.time(`[sp-timing] resolveSiteId ${siteUrl}`)
  const url = `${GRAPH_BASE}/sites/${siteSegment(siteUrl)}`
  const res = await fetchWithTimeout(url, { headers: { Authorization: `Bearer ${accessToken}` } }, control, `resolving site ${siteUrl}`)
  await assertGraphOk(res, { siteUrl, context: `Could not resolve SharePoint site "${siteUrl}"` })

  const data = await res.json() as { id: string }
  siteIdCache.set(siteUrl, data.id)
  console.timeEnd(`[sp-timing] resolveSiteId ${siteUrl}`)
  return data.id
}

/** Reads a file from SharePoint and returns it as raw bytes. */
export async function readSharePointFile(
  accessToken: string,
  siteUrl: string,
  filePath: string,
  control: FetchControl = DEFAULT_CONTROL,
): Promise<ArrayBuffer> {
  const siteId = await resolveSiteId(accessToken, siteUrl, control)

  // encodeURI (not encodeURIComponent) — keeps the '/' path separators
  // intact while percent-encoding spaces/parens/etc. in each segment,
  // which these real file paths are full of.
  console.time(`[sp-timing] graph content fetch ${filePath}`)
  const url = `${GRAPH_BASE}/sites/${siteId}/drive/root:${encodeURI(filePath)}:/content`
  const res = await fetchWithTimeout(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  }, control, `reading ${filePath}`)
  await assertGraphOk(res, { siteUrl, filePath, context: `SharePoint read failed for ${filePath}` })

  const buffer = await res.arrayBuffer()
  console.timeEnd(`[sp-timing] graph content fetch ${filePath}`)
  return buffer
}

export interface SharePointFileMetadata {
  lastModifiedDateTime: string
  size: number
}

interface DriveItemResponse {
  lastModifiedDateTime: string
  size: number
  '@microsoft.graph.downloadUrl'?: string
}

async function fetchDriveItem(accessToken: string, siteUrl: string, filePath: string, control: FetchControl): Promise<DriveItemResponse> {
  const siteId = await resolveSiteId(accessToken, siteUrl, control)
  const url = `${GRAPH_BASE}/sites/${siteId}/drive/root:${encodeURI(filePath)}`
  const res = await fetchWithTimeout(url, { headers: { Authorization: `Bearer ${accessToken}` } }, control, `getting metadata for ${filePath}`)
  await assertGraphOk(res, { siteUrl, filePath, context: `Could not get metadata for ${filePath}` })
  return res.json() as Promise<DriveItemResponse>
}

/** Fetches a file's metadata (no content) — the DriveItem resource itself,
 * same addressing as readSharePointFile minus the trailing `:/content`. */
export async function getSharePointFileMetadata(
  accessToken: string,
  siteUrl: string,
  filePath: string,
  control: FetchControl = DEFAULT_CONTROL,
): Promise<SharePointFileMetadata> {
  const data = await fetchDriveItem(accessToken, siteUrl, filePath, control)
  return { lastModifiedDateTime: data.lastModifiedDateTime, size: data.size }
}

/** Convenience wrapper over getSharePointFileMetadata for a named FILES entry.
 * Always uses the 45s default budget — a metadata-only call is a tiny JSON
 * response regardless of which file it's for, so CR doesn't need (and
 * shouldn't get) the longer 180s content-fetch budget here. */
export async function getNamedFileMetadata(accessToken: string, key: SharePointFileKey, control: FetchControl = DEFAULT_CONTROL): Promise<SharePointFileMetadata> {
  const { site, path } = FILES[key]
  try {
    return await getSharePointFileMetadata(accessToken, site, path, control)
  } catch (err) {
    if (err instanceof SharePointNotFoundError) {
      throw new SharePointNotFoundError(key, path, err.detail)
    }
    throw err
  }
}

/** Convenience wrapper over readSharePointFile for one of the named FILES entries. */
export async function readNamedFile(accessToken: string, key: SharePointFileKey, control?: FetchControl): Promise<ArrayBuffer> {
  const { site, path } = FILES[key]
  try {
    return await readSharePointFile(accessToken, site, path, control ?? { timeoutMs: timeoutMsFor(key) })
  } catch (err) {
    if (err instanceof SharePointNotFoundError) {
      throw new SharePointNotFoundError(key, path, err.detail)
    }
    throw err
  }
}

/** Perf fix, 11/2026: one Graph API call (the DriveItem, which already
 * carries @microsoft.graph.downloadUrl) instead of two — the previous path
 * made a separate /content call AND a separate metadata call for the same
 * file every load. The downloadUrl is a pre-authenticated, time-limited
 * direct link straight to the file's bytes (no Authorization header, often
 * CDN-backed) — fetching it is a plain request, not a second Graph API
 * call. Falls back to the old two-call path only if a DriveItem somehow
 * lacks a downloadUrl (shouldn't happen for a real file, but never seen in
 * this account's tenant either way, so this is defensive, not load-bearing). */
export async function readNamedFileWithMetadata(
  accessToken: string,
  key: SharePointFileKey,
  control: FetchControl = { timeoutMs: timeoutMsFor(key) },
): Promise<{ buffer: ArrayBuffer; metadata: SharePointFileMetadata }> {
  const { site, path } = FILES[key]
  console.time(`[sp-timing] graph metadata+downloadUrl ${key}`)
  let item: DriveItemResponse
  try {
    // The metadata half of this combined call is cheap regardless of file
    // — always bounded by the 45s default, not CR's longer content budget,
    // same reasoning as getNamedFileMetadata above.
    item = await fetchDriveItem(accessToken, site, path, { ...DEFAULT_CONTROL, signal: control.signal })
  } catch (err) {
    if (err instanceof SharePointNotFoundError) throw new SharePointNotFoundError(key, path, err.detail)
    throw err
  }
  console.timeEnd(`[sp-timing] graph metadata+downloadUrl ${key}`)

  const metadata: SharePointFileMetadata = { lastModifiedDateTime: item.lastModifiedDateTime, size: item.size }
  const downloadUrl = item['@microsoft.graph.downloadUrl']
  if (!downloadUrl) {
    // Defensive fallback — see doc comment above.
    const buffer = await readNamedFile(accessToken, key, control)
    return { buffer, metadata }
  }

  console.time(`[sp-timing] downloadUrl content fetch ${key}`)
  const res = await fetchWithTimeout(downloadUrl, {}, control, `downloading ${key}`)
  if (!res.ok) {
    throw new GraphApiError(`Direct download failed for ${path}`, { status: res.status, code: '', message: res.statusText })
  }
  const buffer = await res.arrayBuffer()
  console.timeEnd(`[sp-timing] downloadUrl content fetch ${key}`)
  return { buffer, metadata }
}

export interface FolderChild {
  name: string
  isFolder: boolean
  childCount: number | null
}

/** Lists the children of a folder path (relative to the site's default
 * drive, same addressing as readSharePointFile) — '' or '/' lists the
 * drive root. Built so CJ can see exactly what the app can see, since
 * guessing at drive-relative paths is what caused the 404 this fixes. */
export async function listFolderChildren(
  accessToken: string,
  siteUrl: string,
  folderPath: string,
  control: FetchControl = DEFAULT_CONTROL,
): Promise<FolderChild[]> {
  const siteId = await resolveSiteId(accessToken, siteUrl, control)
  const trimmed = folderPath.trim()
  const isRoot = trimmed === '' || trimmed === '/'
  const url = isRoot
    ? `${GRAPH_BASE}/sites/${siteId}/drive/root/children`
    : `${GRAPH_BASE}/sites/${siteId}/drive/root:${encodeURI(trimmed.startsWith('/') ? trimmed : `/${trimmed}`)}:/children`
  const res = await fetchWithTimeout(url, { headers: { Authorization: `Bearer ${accessToken}` } }, control, `listing folder "${folderPath || '/'}"`)
  await assertGraphOk(res, { siteUrl, filePath: folderPath, context: `Could not list folder "${folderPath || '/'}"` })

  const data = await res.json() as { value: { name: string; folder?: { childCount?: number } }[] }
  return data.value.map(v => ({
    name: v.name,
    isFolder: !!v.folder,
    childCount: v.folder?.childCount ?? null,
  }))
}
