'use client'

export const dynamic = 'force-dynamic'

import { useState } from 'react'
import * as XLSX from 'xlsx'
import { useMsal, useIsAuthenticated } from '@azure/msal-react'
import { InteractionStatus } from '@azure/msal-browser'
import BackToDashboard from '../components/BackToDashboard'
import MicrosoftLogin from '../components/MicrosoftLogin'
import { graphLoginRequest } from '../lib/msalConfig'
import { detectScopHeaderRow } from '../lib/scop'
import {
  readNamedFile, getNamedFileMetadata, listFolderChildren, FILES, SITES,
  type SharePointFileKey, type FolderChild,
  SharePointAuthError, SharePointNotFoundError, SharePointRateLimitError, SharePointAccessDeniedError,
} from '../lib/sharepoint'

// Stage 1 verification page for the SharePoint migration — NOT linked from
// any nav, not used by any real feature. Proves end to end: sign in with a
// Nokia account, pull one file straight from SharePoint via Graph, and
// parse it exactly the way its real upload handler does, so the result can
// be compared against today's manual upload. Delete this page once Stage
// 2+ migrate real pages onto the same read path and it's no longer needed
// as a standalone check.
const FILE_KEYS: SharePointFileKey[] = ['tracker', 'spo', 'decom', 'cr', 'scop']

const SITE_OPTIONS = Object.entries(SITES) as [string, string][]

function trimTrailingBlankRows(rows: unknown[][]): unknown[][] {
  let last = rows.length - 1
  while (last >= 0 && !(rows[last] || []).some(c => c !== null && c !== undefined && String(c).trim() !== '')) {
    last--
  }
  return rows.slice(0, last + 1)
}

interface UsedSheetCheck {
  usedSheetName: string | null
  rowCount: number | null
  note?: string
}

// Mirrors each real upload handler's own sheet-selection + header-row +
// trailing-blank-trim logic exactly, so the row count shown here is the
// same number that handler would actually store, not just "whatever's in
// the first tab." See each handler for the source of truth this copies:
//   tracker -> app/page.tsx handleTrackerUpload
//   decom/cr/spo -> app/reports/page.tsx handleUpload
//   scop -> app/reports/ScopTab.tsx handleUpload
function checkUsedSheet(fileKey: SharePointFileKey, wb: XLSX.WorkBook): UsedSheetCheck {
  if (fileKey === 'tracker') {
    const ws = wb.Sheets['HOPs']
    if (!ws) return { usedSheetName: null, rowCount: null, note: '"HOPs" tab not found' }
    const raw = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1 })
    const trimmed = trimTrailingBlankRows(raw)
    const headerIdx = trimmed.findIndex(r => (r || []).some(c => String(c).trim() === 'DON 444'))
    if (headerIdx === -1) return { usedSheetName: 'HOPs', rowCount: null, note: 'Could not find the "DON 444" header row' }
    return { usedSheetName: 'HOPs', rowCount: trimmed.length - headerIdx - 1 }
  }

  if (fileKey === 'decom') {
    // Real tab name may carry a trailing space — trim before comparing,
    // same as the real handler.
    const matchedName = wb.SheetNames.find(n => n.trim() === 'R&R - Link Drop Off Status')
    const ws = matchedName ? wb.Sheets[matchedName] : undefined
    if (!ws) return { usedSheetName: null, rowCount: null, note: '"R&R - Link Drop Off Status" tab not found' }
    const raw = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null })
    const trimmed = trimTrailingBlankRows(raw)
    const dataRows = trimmed.slice(1).filter(row => row.some(v => v !== null))
    return { usedSheetName: matchedName!, rowCount: dataRows.length }
  }

  if (fileKey === 'cr') {
    const usedSheetName = wb.Sheets['CR Tracker'] ? 'CR Tracker' : wb.SheetNames[0]
    const ws = wb.Sheets[usedSheetName]
    const raw = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null })
    const trimmed = trimTrailingBlankRows(raw)
    const dataRows = trimmed.slice(2).filter(row => row.some(v => v !== null)) // headerRowIndex=1
    return { usedSheetName, rowCount: dataRows.length }
  }

  if (fileKey === 'spo') {
    const usedSheetName = wb.Sheets['NDPd_NAM_003_Viaero_SPO_Report'] ? 'NDPd_NAM_003_Viaero_SPO_Report' : wb.SheetNames[0]
    const ws = wb.Sheets[usedSheetName]
    const raw = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null })
    // No trailing-blank-trim for spo — the real handler doesn't apply one.
    const dataRows = raw.slice(1).filter(row => row.some(v => v !== null)) // headerRowIndex=0
    return { usedSheetName, rowCount: dataRows.length }
  }

  // scop
  const usedSheetName = wb.Sheets['HOPs'] ? 'HOPs' : wb.SheetNames.find(n => n.trim().toLowerCase() === 'hops')
  if (!usedSheetName) return { usedSheetName: null, rowCount: null, note: '"HOPs" tab not found' }
  const ws = wb.Sheets[usedSheetName]
  const raw = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null })
  const headerIdx = detectScopHeaderRow(raw)
  if (headerIdx < 0) return { usedSheetName, rowCount: null, note: 'Could not auto-detect the header row' }
  const dataRows = raw.slice(headerIdx + 1).filter(r => (r || []).some(v => v !== null && v !== undefined && String(v).trim() !== ''))
  return { usedSheetName, rowCount: dataRows.length }
}

export default function SharePointTestPage() {
  const { instance, accounts, inProgress } = useMsal()
  const isAuthenticated = useIsAuthenticated()
  const [fileKey, setFileKey] = useState<SharePointFileKey>('tracker')
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [result, setResult] = useState<{
    fileKey: string; byteLength: number; lastModifiedDateTime: string
    allSheetNames: string[]; usedSheet: UsedSheetCheck
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
      const [buffer, metadata] = await Promise.all([
        readNamedFile(accessToken, fileKey),
        getNamedFileMetadata(accessToken, fileKey),
      ])
      const wb = XLSX.read(buffer, { type: 'array' })
      const usedSheet = checkUsedSheet(fileKey, wb)

      setResult({
        fileKey,
        byteLength: buffer.byteLength,
        lastModifiedDateTime: metadata.lastModifiedDateTime,
        allSheetNames: wb.SheetNames,
        usedSheet,
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
        straight from SharePoint via Graph, and parses it exactly the way its real upload
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
              <p>Last modified: {new Date(result.lastModifiedDateTime).toLocaleString()}</p>
              <p>All sheets: {result.allSheetNames.join(', ')}</p>
              {result.usedSheet.usedSheetName && (
                <p>
                  Sheet Ciege uses: <span className="text-emerald-400">{result.usedSheet.usedSheetName}</span>
                  {result.usedSheet.rowCount != null && <> — {result.usedSheet.rowCount.toLocaleString()} rows</>}
                </p>
              )}
              {result.usedSheet.note && (
                <p className="text-amber-400">⚠️ {result.usedSheet.note}</p>
              )}
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
