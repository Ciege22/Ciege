import { createContext, useContext } from 'react'

// Bug fix, 11/2026 ("SharePoint refresh gets stuck" Part A #5): useMsal()
// returns a STUBBED context (accounts: [] always, a no-op instance) for
// every render before the real MsalProvider has mounted — not just before
// the user is actually signed in. MsalProviderInner renders children
// without any provider at all until its own initialize() (a cache read,
// not a network call, but still async) resolves, and MsalProviderWrapper
// does the same during its one-tick dynamic import. SharePointDataProvider
// sits inside both, so every file's very first ensureLoaded() call used to
// fire against that stub — accounts[0] is always undefined there, so it
// immediately set a leftover "Not signed in with a Nokia account." error on
// every file, superseded a render or two later by the real attempt once the
// real account was available. Harmless in itself (the stub attempt settles
// immediately and clears its own in-flight entry — see
// SharePointDataProvider's generation-guarded cleanup), but needless and
// confusing in Diagnostics. This context lets the provider wait for the
// real thing before starting any load at all. Default `false` — only
// MsalProviderInner's real `ready` branch ever provides `true`.
export const MsalReadyContext = createContext(false)

export function useMsalReady(): boolean {
  return useContext(MsalReadyContext)
}
