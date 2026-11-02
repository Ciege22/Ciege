// Deliberately not 'use client' with any MSAL/data logic — this page must
// stay inert. It's the target loginPopup() redirects the popup window to;
// MsalProviderWrapper skips mounting MSAL here specifically (see its own
// comment) so nothing inside the popup competes with the OPENER window's
// MSAL instance, which is what actually reads the #code=... hash off this
// page's URL and closes the popup. If this page ever starts doing real
// work, that's a sign the fix regressed.
export const dynamic = 'force-dynamic'

export default function AuthRedirectPage() {
  return (
    <div style={{
      display: 'flex', alignItems: 'center', justifyContent: 'center',
      height: '100vh', background: '#09090b', color: '#a1a1aa', fontFamily: 'sans-serif', fontSize: 14,
    }}>
      Signing in… this window should close automatically.
    </div>
  )
}
