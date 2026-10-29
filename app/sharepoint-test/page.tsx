'use client'

export const dynamic = 'force-dynamic'

import { useState } from 'react'
import * as XLSX from 'xlsx'
import { useMsal, useIsAuthenticated } from '@azure/msal-react'
import { InteractionStatus } from '@azure/msal-browser'
import BackToDashboard from '../components/BackToDashboard'
import MicrosoftLogin from '../components/MicrosoftLogin'
import { graphLoginRequest } from '../lib/msalConfig'
import {
  readNamedFile, FILES, type SharePointFileKey,
  SharePointAuthError, SharePointNotFoundError, SharePointRateLimitError, SharePointAccessDeniedError,
} from '../lib/sharepoint'

// Stage 1 verification page for the SharePoint migration — NOT linked from
// any nav, not used by any real feature. Proves end to end: sign in with a
// Nokia account, pull one file straight from SharePoint via Graph, and
// parse it with the same SheetJS path every existing upload handler uses,
// so the byte-for-byte result can be compared against today's manual
// upload. Delete this page once Stage 2+ migrate real pages onto the same
// read path and it's no longer needed as a standalone check.
const FILE_KEYS: SharePointFileKey[] = ['tracker', 'spo', 'decom', 'cr', 'scop']

export default function SharePointTestPage() {
  const { instance, accounts, inProgress } = useMsal()
  const isAuthenticated = useIsAuthenticated()
  const [fileKey, setFileKey] = useState<SharePointFileKey>('tracker')
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [result, setResult] = useState<{
    fileKey: string; byteLength: number; sheetNames: string[]
    firstSheetRowCount: number; firstSheetSample: unknown[][]
  } | null>(null)

  const busy = inProgress !== InteractionStatus.None

  async function runTest() {
    setError(null)
    setResult(null)
    setLoading(true)
    try {
      const account = accounts[0]
      if (!account) throw new Error('Not signed in.')

      let accessToken: string
      try {
        const tokenResult = await instance.acquireTokenSilent({ ...graphLoginRequest, account })
        accessToken = tokenResult.accessToken
      } catch {
        // Silent acquisition can fail (e.g. needs re-consent) — fall back
        // to an interactive popup rather than leaving the user stuck.
        const tokenResult = await instance.acquireTokenPopup(graphLoginRequest)
        accessToken = tokenResult.accessToken
      }

      const buffer = await readNamedFile(accessToken, fileKey)
      const wb = XLSX.read(buffer, { type: 'array' })
      const firstSheetName = wb.SheetNames[0]
      const firstSheet = wb.Sheets[firstSheetName]
      const rows = XLSX.utils.sheet_to_json<unknown[]>(firstSheet, { header: 1 })

      setResult({
        fileKey,
        byteLength: buffer.byteLength,
        sheetNames: wb.SheetNames,
        firstSheetRowCount: rows.length,
        firstSheetSample: rows.slice(0, 3),
      })
    } catch (err) {
      if (err instanceof SharePointAuthError) setError(`${err.message} (sign in again below)`)
      else if (err instanceof SharePointAccessDeniedError) setError(err.message)
      else if (err instanceof SharePointNotFoundError) setError(err.message)
      else if (err instanceof SharePointRateLimitError) setError(err.message)
      else setError(err instanceof Error ? err.message : 'Read failed.')
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="min-h-screen bg-gray-950 text-white p-6 max-w-2xl mx-auto">
      <BackToDashboard />
      <h1 className="text-2xl font-bold mb-1">🔧 SharePoint Read — Stage 1 Test</h1>
      <p className="text-gray-400 text-sm mb-6">
        Not part of any real feature yet. Signs in with your Nokia account, pulls one file
        straight from SharePoint via Graph, and parses it the same way every current upload
        handler does — just to prove the path works before any real page switches to it.
      </p>

      <div className="bg-gray-900 rounded-xl border border-gray-700 p-4 mb-4">
        <MicrosoftLogin />
      </div>

      {isAuthenticated && (
        <div className="bg-gray-900 rounded-xl border border-gray-700 p-4 space-y-4">
          <div>
            <label className="block text-xs text-gray-400 mb-1">File to read</label>
            <select value={fileKey} onChange={e => setFileKey(e.target.value as SharePointFileKey)}
              className="bg-gray-800 text-white text-sm rounded px-2 py-1.5 border border-gray-600">
              {FILE_KEYS.map(k => <option key={k} value={k}>{k} — {FILES[k].path}</option>)}
            </select>
          </div>
          <button onClick={runTest} disabled={loading || busy}
            className="bg-emerald-700 hover:bg-emerald-600 disabled:opacity-40 text-white text-sm font-semibold rounded px-4 py-2">
            {loading ? 'Reading…' : `Read "${fileKey}" from SharePoint`}
          </button>

          {error && (
            <p className="bg-red-950 border border-red-700 text-red-200 text-sm rounded-lg p-3">{error}</p>
          )}

          {result && (
            <div className="text-sm text-gray-300 space-y-1 border-t border-gray-700 pt-3">
              <p>✅ Read <span className="text-emerald-400">{result.fileKey}</span> — {result.byteLength.toLocaleString()} bytes</p>
              <p>Sheets: {result.sheetNames.join(', ')}</p>
              <p>First sheet: {result.firstSheetRowCount.toLocaleString()} rows</p>
              <pre className="bg-gray-950 border border-gray-800 rounded p-2 mt-2 overflow-x-auto text-xs text-gray-400">
                {JSON.stringify(result.firstSheetSample, null, 2)}
              </pre>
            </div>
          )}
        </div>
      )}
    </div>
  )
}
