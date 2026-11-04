'use client'

import { ReactNode, useEffect, useMemo, useState } from 'react'
import { PublicClientApplication, EventType, type AuthenticationResult } from '@azure/msal-browser'
import { MsalProvider } from '@azure/msal-react'
import { msalConfig } from '../lib/msalConfig'
import { clearNokiaCache } from '../lib/nokiaCache'

// The actual MSAL instance/provider — split out from MsalProviderWrapper so
// that one can load it via next/dynamic(ssr:false). PublicClientApplication
// touches `window` during construction, which crashes/spams errors under
// Next's server-side prerender pass (same class of issue as Leaflet on the
// Map page — see app/map/page.tsx for the same pattern).
//
// Perf fix, 11/2026: this never called instance.initialize() or restored an
// already-signed-in account from cache — msal-browser v3+ requires
// initialize() before ANY other API call, and without it or an explicit
// getAllAccounts()/setActiveAccount() on mount, useMsal().accounts stays
// empty after every hard reload even though the account is still sitting in
// sessionStorage (cacheLocation, see msalConfig.ts) — so a returning user
// looked signed-out and had to click "Sign in" again every time. Rendering
// now waits for initialize() (near-instant — it's just MSAL reading its own
// cache, not a network call) before mounting MsalProvider, so
// acquireTokenSilent in SharePointDataProvider has a real account to use
// from the very first render that needs one.
export default function MsalProviderInner({ children }: { children: ReactNode }) {
  const [ready, setReady] = useState(false)
  const msalInstance = useMemo(() => new PublicClientApplication(msalConfig), [])

  useEffect(() => {
    let cancelled = false
    msalInstance.initialize().then(() => {
      const accounts = msalInstance.getAllAccounts()
      if (accounts.length > 0 && !msalInstance.getActiveAccount()) {
        msalInstance.setActiveAccount(accounts[0])
      }
      if (!cancelled) setReady(true)
    })

    // Keeps the active account set after a popup login completes, so
    // useMsal().accounts isn't empty on the very next render. Clears the
    // on-device Nokia cache on sign-out — nothing from a previous account
    // should outlive that account's own session.
    const callbackId = msalInstance.addEventCallback(event => {
      if (event.eventType === EventType.LOGIN_SUCCESS && event.payload) {
        const account = (event.payload as AuthenticationResult).account
        if (account) msalInstance.setActiveAccount(account)
      }
      if (event.eventType === EventType.LOGOUT_SUCCESS) {
        void clearNokiaCache()
      }
    })

    return () => {
      cancelled = true
      if (callbackId) msalInstance.removeEventCallback(callbackId)
    }
  }, [msalInstance])

  // Children render immediately either way (same as before) — only the MSAL
  // context itself waits, and that wait is a cache read, not a network
  // round trip, so it resolves within a render or two.
  if (!ready) return <>{children}</>
  return <MsalProvider instance={msalInstance}>{children}</MsalProvider>
}
