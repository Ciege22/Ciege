'use client'

import { useState } from 'react'
import { useMsal, useIsAuthenticated } from '@azure/msal-react'
import { InteractionStatus } from '@azure/msal-browser'
import { graphLoginRequest } from '../lib/msalConfig'

// Nokia SSO login/logout — Stage 1 of the SharePoint migration. Purely
// additive: nothing existing reads auth state yet, this only proves the
// login flow itself works. Popup (not redirect) so it doesn't navigate away
// from whatever page it's rendered on.
export default function MicrosoftLogin() {
  const { instance, accounts, inProgress } = useMsal()
  const isAuthenticated = useIsAuthenticated()
  const [error, setError] = useState<string | null>(null)

  const account = accounts[0]
  const busy = inProgress !== InteractionStatus.None

  const signIn = async () => {
    setError(null)
    try {
      await instance.loginPopup(graphLoginRequest)
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Sign-in failed — try again.')
    }
  }

  const signOut = async () => {
    setError(null)
    try {
      await instance.logoutPopup({ account })
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Sign-out failed — try again.')
    }
  }

  if (inProgress === InteractionStatus.Startup) {
    return <span className="text-xs text-zinc-500">Checking sign-in…</span>
  }

  if (isAuthenticated && account) {
    return (
      <div className="flex items-center gap-2 text-xs">
        <span className="text-zinc-300">{account.name || account.username}</span>
        <button onClick={signOut} disabled={busy}
          className="text-zinc-500 hover:text-white underline disabled:opacity-50">
          Sign out
        </button>
        {error && <span className="text-red-400">{error}</span>}
      </div>
    )
  }

  return (
    <div className="flex items-center gap-2 text-xs">
      <button onClick={signIn} disabled={busy}
        className="bg-blue-700 hover:bg-blue-600 disabled:opacity-50 text-white font-semibold rounded px-3 py-1.5">
        {busy ? 'Signing in…' : 'Sign in with Microsoft'}
      </button>
      {error && <span className="text-red-400">{error}</span>}
    </div>
  )
}
