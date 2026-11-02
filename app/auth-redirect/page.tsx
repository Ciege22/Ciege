'use client'

export const dynamic = 'force-dynamic'

import { useEffect, useState } from 'react'

// Target of the loginPopup() popup window after Microsoft redirects back.
// msal-browser v5's popup flow doesn't work by the opener polling this
// page's URL — the popup itself must call broadcastResponseToMainFrame(),
// which reads the #code=... response from the URL, sends it to the opener
// over a BroadcastChannel (keyed by an id embedded in the request's state
// param), and closes this window. That's the library's own default
// waitForPopupResponse path (confirmed in PopupClient.ts — no
// popupRelayUri or custom hook configured here), so nothing else needs to
// change on the opener/login side.
//
// Imported from '@azure/msal-browser/redirect-bridge', not the main
// package — a standalone function, no PublicClientApplication/MsalProvider
// needed. MsalProviderWrapper also deliberately skips mounting MSAL on
// this path (see its own comment) so there's never a second, competing
// MSAL instance in this window.
export default function AuthRedirectPage() {
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    import('@azure/msal-browser/redirect-bridge')
      .then(({ broadcastResponseToMainFrame }) => broadcastResponseToMainFrame())
      .catch(err => {
        setError(err instanceof Error ? err.message : 'Sign-in could not be completed.')
      })
  }, [])

  return (
    <div style={{
      display: 'flex', alignItems: 'center', justifyContent: 'center',
      height: '100vh', background: '#09090b', color: '#a1a1aa', fontFamily: 'sans-serif', fontSize: 14,
      textAlign: 'center', padding: 16,
    }}>
      {error ? `${error} — you can close this window and try again.` : 'Signing in… this window should close automatically.'}
    </div>
  )
}
