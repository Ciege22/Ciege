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
// Not wired into any existing page yet — this stage only proves the read
// path itself works. Later stages replace each page's loadTrackerSnapshot/
// loadChunkedReport call with one of the FILES entries below.

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

async function resolveSiteId(accessToken: string, siteUrl: string): Promise<string> {
  const cached = siteIdCache.get(siteUrl)
  if (cached) return cached

  const url = `${GRAPH_BASE}/sites/${siteSegment(siteUrl)}`
  const res = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } })
  await assertGraphOk(res, { siteUrl, context: `Could not resolve SharePoint site "${siteUrl}"` })

  const data = await res.json() as { id: string }
  siteIdCache.set(siteUrl, data.id)
  return data.id
}

/** Reads a file from SharePoint and returns it as raw bytes. */
export async function readSharePointFile(
  accessToken: string,
  siteUrl: string,
  filePath: string,
): Promise<ArrayBuffer> {
  const siteId = await resolveSiteId(accessToken, siteUrl)

  // encodeURI (not encodeURIComponent) — keeps the '/' path separators
  // intact while percent-encoding spaces/parens/etc. in each segment,
  // which these real file paths are full of.
  const url = `${GRAPH_BASE}/sites/${siteId}/drive/root:${encodeURI(filePath)}:/content`
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  })
  await assertGraphOk(res, { siteUrl, filePath, context: `SharePoint read failed for ${filePath}` })

  return res.arrayBuffer()
}

export interface SharePointFileMetadata {
  lastModifiedDateTime: string
  size: number
}

/** Fetches a file's metadata (no content) — the DriveItem resource itself,
 * same addressing as readSharePointFile minus the trailing `:/content`. */
export async function getSharePointFileMetadata(
  accessToken: string,
  siteUrl: string,
  filePath: string,
): Promise<SharePointFileMetadata> {
  const siteId = await resolveSiteId(accessToken, siteUrl)
  const url = `${GRAPH_BASE}/sites/${siteId}/drive/root:${encodeURI(filePath)}`
  const res = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } })
  await assertGraphOk(res, { siteUrl, filePath, context: `Could not get metadata for ${filePath}` })

  const data = await res.json() as { lastModifiedDateTime: string; size: number }
  return { lastModifiedDateTime: data.lastModifiedDateTime, size: data.size }
}

/** Convenience wrapper over getSharePointFileMetadata for a named FILES entry. */
export async function getNamedFileMetadata(accessToken: string, key: SharePointFileKey): Promise<SharePointFileMetadata> {
  const { site, path } = FILES[key]
  try {
    return await getSharePointFileMetadata(accessToken, site, path)
  } catch (err) {
    if (err instanceof SharePointNotFoundError) {
      throw new SharePointNotFoundError(key, path, err.detail)
    }
    throw err
  }
}

/** Convenience wrapper over readSharePointFile for one of the named FILES entries. */
export async function readNamedFile(accessToken: string, key: SharePointFileKey): Promise<ArrayBuffer> {
  const { site, path } = FILES[key]
  try {
    return await readSharePointFile(accessToken, site, path)
  } catch (err) {
    if (err instanceof SharePointNotFoundError) {
      throw new SharePointNotFoundError(key, path, err.detail)
    }
    throw err
  }
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
): Promise<FolderChild[]> {
  const siteId = await resolveSiteId(accessToken, siteUrl)
  const trimmed = folderPath.trim()
  const isRoot = trimmed === '' || trimmed === '/'
  const url = isRoot
    ? `${GRAPH_BASE}/sites/${siteId}/drive/root/children`
    : `${GRAPH_BASE}/sites/${siteId}/drive/root:${encodeURI(trimmed.startsWith('/') ? trimmed : `/${trimmed}`)}:/children`
  const res = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } })
  await assertGraphOk(res, { siteUrl, filePath: folderPath, context: `Could not list folder "${folderPath || '/'}"` })

  const data = await res.json() as { value: { name: string; folder?: { childCount?: number } }[] }
  return data.value.map(v => ({
    name: v.name,
    isFolder: !!v.folder,
    childCount: v.folder?.childCount ?? null,
  }))
}
