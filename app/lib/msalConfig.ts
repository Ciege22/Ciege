// MSAL (Microsoft Authentication Library) config for Nokia SSO login —
// lets a signed-in Nokia user's browser session read SharePoint on their
// own behalf (delegated auth), separate from Railway's own app-only Graph
// access in backend/sharepoint.py. Client ID/tenant are public by design
// (MSAL is a browser-side OAuth public client, not a secret) — Vercel's
// NEXT_PUBLIC_ prefix makes that explicit.
import { Configuration, LogLevel } from '@azure/msal-browser'

const clientId = process.env.NEXT_PUBLIC_AZURE_CLIENT_ID ?? ''
const tenantId = process.env.NEXT_PUBLIC_AZURE_TENANT_ID ?? ''

export const msalConfig: Configuration = {
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri: typeof window !== 'undefined' ? window.location.origin : '',
  },
  cache: {
    // Session, not local — the access token lives in memory via MSAL's
    // in-memory cache on top of this; sessionStorage here only holds the
    // ID token / account info needed to avoid re-prompting on every page
    // nav within the same tab, and clears when the tab closes.
    cacheLocation: 'sessionStorage',
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) return
        if (level === LogLevel.Error) console.error('[msal]', message)
      },
    },
  },
}

// Sites.Read.All covers both SharePoint sites this app reads from
// (viaero-wireless2 and NAMDS) without needing a separate consent per site.
export const graphLoginRequest = {
  scopes: ['Sites.Read.All'],
}
