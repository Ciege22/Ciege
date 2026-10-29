// SharePoint read service — Stage 1 of the manual-upload-to-SharePoint
// migration. Reads a file straight out of Nokia's SharePoint via Microsoft
// Graph, using whichever Nokia account the signed-in user authenticated
// with (delegated access — this app never has its own standing access to
// a user's files, only what that user could already open themselves).
//
// Permission model: Sites.Selected (see app/lib/msalConfig.ts) — the app
// has NO site access by default, only whichever specific sites Nokia IT
// has separately granted via the Graph site-permissions API. The request
// shape below (reading by site+path) is unchanged from a Sites.Read.All
// setup; what differs is that a 403 here most likely means "this site
// hasn't been granted to the app yet" rather than a real permission
// problem with the signed-in user — see SharePointAccessDeniedError.
//
// Not wired into any existing page yet — this stage only proves the read
// path itself works. Later stages replace each page's loadTrackerSnapshot/
// loadChunkedReport call with one of the FILES entries below.

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0'

export const SITES = {
  viaeroWireless: 'nokianam.sharepoint.com/sites/viaero-wireless2',
  namds: 'nokia.sharepoint.com/sites/NAMDS',
} as const

export const FILES = {
  tracker: {
    site: SITES.viaeroWireless,
    path: '/Shared Documents/Team PM/CJ Ojeda/Viaero - Construction Schedule Tracker - WIP 2.0.xlsx',
  },
  spo: {
    site: SITES.namds,
    path: '/Shared Documents/NDPd (SiteForge)/US Majors/US Majors NDPd Reports/Viaero NDPd Reports/Current/NDPd_NAM_003_Viaero_SPO_Report.xlsx',
  },
  decom: {
    site: SITES.viaeroWireless,
    path: '/Shared Documents/Team Decom/Decom Tracker Updates/Nokia Decom Status_Report_LIVE - NEW.xlsx',
  },
  cr: {
    site: SITES.viaeroWireless,
    path: '/Shared Documents/Team CQT/Viaero - CR Claims Tracker.xlsx',
  },
  scop: {
    site: SITES.viaeroWireless,
    path: '/Shared Documents/Team Close Out/Viaero COP Tracker_LIVE.xlsx',
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

export class SharePointAuthError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'SharePointAuthError'
  }
}

export class SharePointNotFoundError extends Error {
  constructor(public fileKey: string, public filePath: string) {
    super(`Could not find "${fileKey}" on SharePoint at: ${filePath}`)
    this.name = 'SharePointNotFoundError'
  }
}

// Under Sites.Selected, a 403 on an otherwise-valid, signed-in request
// almost always means the TARGET SITE hasn't been granted to this app yet
// (a separate IT action per site, not something the signed-in user's own
// permissions control) — distinct enough from a generic failure to call
// out explicitly rather than let it read as "something's broken."
export class SharePointAccessDeniedError extends Error {
  constructor(public siteUrl: string) {
    super(`Access denied for site "${siteUrl}" — this site may not be granted to the app yet (Sites.Selected requires each site to be explicitly approved).`)
    this.name = 'SharePointAccessDeniedError'
  }
}

/** Reads a file from SharePoint and returns it as raw bytes. */
export async function readSharePointFile(
  accessToken: string,
  siteUrl: string,
  filePath: string,
): Promise<ArrayBuffer> {
  // encodeURI (not encodeURIComponent) — keeps the '/' path separators
  // intact while percent-encoding spaces/parens/etc. in each segment,
  // which these real file paths are full of.
  const url = `${GRAPH_BASE}/sites/${siteSegment(siteUrl)}:/drive/root:${encodeURI(filePath)}:/content`
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  })

  if (res.status === 401) {
    throw new SharePointAuthError('SharePoint session expired — sign in again.')
  }
  if (res.status === 403) {
    throw new SharePointAccessDeniedError(siteUrl)
  }
  if (res.status === 404) {
    throw new SharePointNotFoundError(filePath, filePath)
  }
  if (res.status === 429) {
    const retryAfter = Number(res.headers.get('Retry-After')) || 30
    throw new SharePointRateLimitError(retryAfter)
  }
  if (!res.ok) {
    const body = await res.text().catch(() => '')
    throw new Error(`SharePoint read failed (${res.status}): ${body.slice(0, 300)}`)
  }

  return res.arrayBuffer()
}

export class SharePointRateLimitError extends Error {
  constructor(public retryAfterSeconds: number) {
    super(`SharePoint/Graph API rate limited — retry after ${retryAfterSeconds}s.`)
    this.name = 'SharePointRateLimitError'
  }
}

/** Convenience wrapper over readSharePointFile for one of the named FILES entries. */
export async function readNamedFile(accessToken: string, key: SharePointFileKey): Promise<ArrayBuffer> {
  const { site, path } = FILES[key]
  try {
    return await readSharePointFile(accessToken, site, path)
  } catch (err) {
    if (err instanceof SharePointNotFoundError) {
      throw new SharePointNotFoundError(key, path)
    }
    throw err
  }
}
