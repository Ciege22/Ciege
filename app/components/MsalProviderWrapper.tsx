'use client'

import { ReactNode, useEffect, useState, type ComponentType } from 'react'
import { usePathname } from 'next/navigation'
import { AUTH_REDIRECT_PATH } from '../lib/msalConfig'

// Thin wrapper that keeps MsalProviderInner (and therefore MSAL's
// PublicClientApplication, which touches `window` on construction) out of
// Next's server-side render/prerender pass entirely — children render
// immediately either way, this only adds the MSAL context once mounted in
// the browser. Hand-rolled dynamic import (rather than next/dynamic) so
// `children` doesn't have to thread through a separate loading slot.
//
// Deliberately skips mounting MSAL at all on AUTH_REDIRECT_PATH — that
// page loads inside the loginPopup() popup window, and if it also
// constructs its own PublicClientApplication, that instance's own
// handleRedirectPromise() races the OPENER window's popup-monitoring code
// for the same #code=... hash, which is what left the popup rendering the
// full dashboard instead of closing itself. The opener's own (already
// live) MSAL instance is what completes the popup flow — nothing inside
// the popup needs to process it at all.
export default function MsalProviderWrapper({ children }: { children: ReactNode }) {
  const pathname = usePathname()
  const [Inner, setInner] = useState<ComponentType<{ children: ReactNode }> | null>(null)

  useEffect(() => {
    if (pathname === AUTH_REDIRECT_PATH) return
    let cancelled = false
    import('./MsalProviderInner').then(mod => {
      if (!cancelled) setInner(() => mod.default)
    })
    return () => { cancelled = true }
  }, [pathname])

  if (pathname === AUTH_REDIRECT_PATH || !Inner) return <>{children}</>
  const MsalInner = Inner
  return <MsalInner>{children}</MsalInner>
}
