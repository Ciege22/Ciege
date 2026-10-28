'use client'

import { ReactNode, useMemo } from 'react'
import { PublicClientApplication, EventType, type AuthenticationResult } from '@azure/msal-browser'
import { MsalProvider } from '@azure/msal-react'
import { msalConfig } from '../lib/msalConfig'

// The actual MSAL instance/provider — split out from MsalProviderWrapper so
// that one can load it via next/dynamic(ssr:false). PublicClientApplication
// touches `window` during construction, which crashes/spams errors under
// Next's server-side prerender pass (same class of issue as Leaflet on the
// Map page — see app/map/page.tsx for the same pattern).
export default function MsalProviderInner({ children }: { children: ReactNode }) {
  const msalInstance = useMemo(() => {
    const instance = new PublicClientApplication(msalConfig)
    // Keeps the active account set after a redirect/popup login completes,
    // so useMsal().accounts isn't empty on the very next render.
    instance.addEventCallback(event => {
      if (event.eventType === EventType.LOGIN_SUCCESS && event.payload) {
        const account = (event.payload as AuthenticationResult).account
        if (account) instance.setActiveAccount(account)
      }
    })
    return instance
  }, [])

  return <MsalProvider instance={msalInstance}>{children}</MsalProvider>
}
