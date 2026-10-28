'use client'

import { ReactNode, useEffect, useState, type ComponentType } from 'react'

// Thin wrapper that keeps MsalProviderInner (and therefore MSAL's
// PublicClientApplication, which touches `window` on construction) out of
// Next's server-side render/prerender pass entirely — children render
// immediately either way, this only adds the MSAL context once mounted in
// the browser. Hand-rolled dynamic import (rather than next/dynamic) so
// `children` doesn't have to thread through a separate loading slot.
export default function MsalProviderWrapper({ children }: { children: ReactNode }) {
  const [Inner, setInner] = useState<ComponentType<{ children: ReactNode }> | null>(null)

  useEffect(() => {
    let cancelled = false
    import('./MsalProviderInner').then(mod => {
      if (!cancelled) setInner(() => mod.default)
    })
    return () => { cancelled = true }
  }, [])

  if (!Inner) return <>{children}</>
  const MsalInner = Inner
  return <MsalInner>{children}</MsalInner>
}
