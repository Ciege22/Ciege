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
  readNamedFile, listFolderChildren, FILES, SITES, type SharePointFileKey, type FolderChild,
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

const SITE_OPTIONS = Object.entries(SITES) as [string, string][]

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

  // List Folder — lets CJ browse exactly what the app can see instead of
  // guessing at drive-relative paths (what caused the original 404: the
  // FILES paths included "Shared Documents," which is the drive itself,
  // not a folder inside it).
  const [folderSite, setFolderSite] = useState<string>(SITES.viaeroWireless)
  const [folderPath, setFolderPath] = useState('')
  const [folderLoading, setFolderLoading] = useState(false)
  const [folderError, setFolderError] = useState<string | null>(null)
  const [folderResult, setFolderResult] = useState<FolderChild[] | null>(null)

  const busy = inProgress !== InteractionStatus.None

  async function getAccessToken(): Promise<string> {
    const account = accounts[0]
    if (!account) throw new Error('Not signed in.')
    try {
      const tokenResult = await instance.acquireTokenSilent({ ...graphLoginRequest, account })
      return tokenResult.accessToken
    } catch {
      // Silent acquisition can fail (e.g. needs re-consent) — fall back
      // to an interactive popup rather than leaving the user stuck.
      const tokenResult = await instance.acquireTokenPopup(graphLoginRequest)
      return tokenResult.accessToken
    }
  }

  async function runTest() {
    setError(null)
    setResult(null)
    setLoading(true)
    try {
      const accessToken = await getAccessToken()
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

  async function listFolder() {
    setFolderError(null)
    setFolderResult(null)
    setFolderLoading(true)
    try {
      const accessToken = await getAccessToken()
      const children = await listFolderChildren(accessToken, folderSite, folderPath)
      setFolderResult(children)
    } catch (err) {
      setFolderError(err instanceof Error ? err.message : 'List failed.')
    } finally {
      setFolderLoading(false)
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

      {isAuthenticated && (
        <div className="bg-gray-900 rounded-xl border border-gray-700 p-4 mt-4 space-y-4">
          <div>
            <h2 className="text-sm font-semibold text-white mb-1">List folder</h2>
            <p className="text-gray-500 text-xs">
              Browse exactly what the app can see, relative to the site&apos;s drive root — blank lists
              the root; try &quot;Team PM&quot; next.
            </p>
          </div>
          <div className="flex flex-wrap gap-2 items-end">
            <div>
              <label className="block text-xs text-gray-400 mb-1">Site</label>
              <select value={folderSite} onChange={e => setFolderSite(e.target.value)}
                className="bg-gray-800 text-white text-sm rounded px-2 py-1.5 border border-gray-600">
                {SITE_OPTIONS.map(([key, url]) => <option key={key} value={url}>{key} — {url}</option>)}
              </select>
            </div>
            <div className="flex-1 min-w-[180px]">
              <label className="block text-xs text-gray-400 mb-1">Folder path (blank = drive root)</label>
              <input value={folderPath} onChange={e => setFolderPath(e.target.value)} placeholder="Team PM"
                className="w-full bg-gray-800 text-white text-sm rounded px-2 py-1.5 border border-gray-600" />
            </div>
            <button onClick={listFolder} disabled={folderLoading || busy}
              className="bg-blue-700 hover:bg-blue-600 disabled:opacity-40 text-white text-sm font-semibold rounded px-4 py-2 h-[34px]">
              {folderLoading ? 'Listing…' : 'List Folder'}
            </button>
          </div>

          {folderError && (
            <p className="bg-red-950 border border-red-700 text-red-200 text-sm rounded-lg p-3">{folderError}</p>
          )}

          {folderResult && (
            <div className="text-sm text-gray-300 border-t border-gray-700 pt-3">
              <p className="text-gray-500 text-xs mb-2">{folderResult.length} item{folderResult.length === 1 ? '' : 's'}</p>
              <ul className="space-y-1">
                {folderResult.map(c => (
                  <li key={c.name} className="flex items-center gap-2 text-xs">
                    <span>{c.isFolder ? '📁' : '📄'}</span>
                    <span>{c.name}</span>
                    {c.isFolder && c.childCount != null && <span className="text-gray-600">({c.childCount} items)</span>}
                  </li>
                ))}
              </ul>
            </div>
          )}
        </div>
      )}
    </div>
  )
}
