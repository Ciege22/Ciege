// Shared SPO batch logic — the needs-SPO-requested bucket, its three-row
// payment split per HOP, and the email / Excel / tracker-stamp actions. Used
// by the SPO Requests page and by the dashboard's Needs SPO Requested window,
// so both always show and send exactly the same data.
import * as XLSX from 'xlsx'
import { supabase } from './supabase'
import { computeSpoStatus } from './spoStatus'
import { loadPendingUpdates, persistPendingUpdates, upsertPendingUpdate } from './pendingUpdates'

const DRAFT_ID = 'spo-request-draft'
const DEFAULT_SPLIT = [20, 60, 20]

export interface PhysicalRow {
  rowKey: string
  siteName: string
  pathId: string
  cpo: string
}

export interface HopSpo {
  hop: string
  appPathId: string
  pathId: string
  vendor: string
  gc: string
  siteNames: string[]
  cpos: string[]
  rows: PhysicalRow[]
}

export interface PaymentRow {
  pct: string
  value: string
}

export interface SpoDraft {
  to: string
  hops: Record<string, PaymentRow[]>
}

export const SPO_HEADERS = [
  'HOP (App Path ID)', 'SPO type', 'Customer Site ID', 'Site Name', 'General Contractor', '% payment', 'SPO value', 'CPO#', 'Vendor Name', 'SPO Request',
]

function validDate(v: unknown): boolean {
  if (!v) return false
  const d = new Date(String(v))
  return !isNaN(d.getTime()) && d.getFullYear() >= 1990
}

export function todayStr(): string {
  return new Date().toLocaleDateString('en-US')
}

export function money(v: string): string {
  const n = Number(String(v).replace(/[$,\s]/g, ''))
  if (!v || !Number.isFinite(n)) return v ? String(v) : '—'
  return n.toLocaleString('en-US', { style: 'currency', currency: 'USD' })
}

export function defaultPaymentRows(): PaymentRow[] {
  return DEFAULT_SPLIT.map(p => ({ pct: `${p}%`, value: '' }))
}

export function rowsFor(draft: SpoDraft, hop: string): PaymentRow[] {
  return draft.hops[hop] ?? defaultPaymentRows()
}

export function cpoText(h: HopSpo): string {
  return Array.from(new Set(h.cpos)).join(', ')
}

// Builds the needs_request bucket from a raw tracker snapshot. Same status
// rules as the dashboard (computeSpoStatus) — CPO received, no SPO requested.
export function buildSpoBucket(rows: unknown[][]): HopSpo[] {
  let headerIdx = -1
  for (let i = 0; i < 10; i++) {
    const r = rows[i] as unknown[]
    if (r && r.some(c => String(c ?? '').trim() === 'HOP')) { headerIdx = i; break }
  }
  if (headerIdx === -1) return []
  const headers = (rows[headerIdx] as unknown[]).map(h => String(h ?? '').trim())
  const col = (n: string) => headers.findIndex(h => h === n)
  const C = {
    app: col('App Path ID'), hop: col('HOP'), pid: col('Path ID'), site: col('Site Name'),
    siteNum: col('Site Number'), don: col('DON 444'), gc: col('General Contractor'),
    cpo: col('Service CPO Received'), spoIss: col('CX SPO issued'),
    spoReq: col('CX SPO Request'), vendor: col('CX SPO Vendor'),
  }

  const byHop = new Map<string, {
    appPathId: string; pathId: string; vendor: string; gc: string
    hasSpo: boolean; hasReq: boolean; hasCpo: boolean; rows: PhysicalRow[]
  }>()
  let counter = 0
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i] as unknown[]
    if (!row) continue
    if (String(row[C.don] ?? '').trim().toUpperCase() !== 'DON 444') continue
    const hop = String(row[C.hop] ?? '').trim()
    if (!hop || hop === 'undefined') continue
    const pathId = String(row[C.pid] ?? '').trim()
    const siteName = String(row[C.site] ?? '').trim()
    const siteNumber = String(row[C.siteNum] ?? '').trim()
    // Same rowKey construction as the Tracker grid, so the pending
    // "CX SPO Request" stamp lands on the same physical rows.
    const rowKey = (siteName || siteNumber)
      ? `${pathId}|${siteName}|${siteNumber}`
      : (pathId || `${hop}-row-${counter}`)
    counter++
    if (!byHop.has(hop)) {
      byHop.set(hop, {
        appPathId: String(row[C.app] ?? '').trim(),
        pathId, vendor: String(row[C.vendor] ?? '').trim(), gc: String(row[C.gc] ?? '').trim(),
        hasSpo: false, hasReq: false, hasCpo: false, rows: [],
      })
    }
    const h = byHop.get(hop)!
    if (validDate(row[C.spoIss])) h.hasSpo = true
    if (validDate(row[C.spoReq])) h.hasReq = true
    const cpo = String(row[C.cpo] ?? '').trim()
    if (cpo) h.hasCpo = true
    if (!h.vendor) h.vendor = String(row[C.vendor] ?? '').trim()
    if (!h.gc) h.gc = String(row[C.gc] ?? '').trim()
    h.rows.push({ rowKey, siteName, pathId, cpo })
  }

  const bucket: HopSpo[] = []
  byHop.forEach((h, hop) => {
    if (computeSpoStatus(h.hasSpo, h.hasCpo, h.hasReq).status !== 'needs_request') return
    bucket.push({
      hop, appPathId: h.appPathId, pathId: h.pathId, vendor: h.vendor, gc: h.gc,
      siteNames: Array.from(new Set(h.rows.map(r => r.siteName))).filter(Boolean),
      cpos: h.rows.map(r => r.cpo).filter(Boolean),
      rows: h.rows,
    })
  })
  bucket.sort((a, b) => (a.gc || '~').localeCompare(b.gc || '~') || a.hop.localeCompare(b.hop))
  return bucket
}

export async function loadSpoDraft(): Promise<SpoDraft> {
  const { data } = await supabase.from('pm_updates_cache').select('updates').eq('id', DRAFT_ID).single()
  if (data?.updates) {
    try { return JSON.parse(data.updates) as SpoDraft } catch { /* corrupt draft — start fresh */ }
  }
  return { to: '', hops: {} }
}

export function saveSpoDraft(draft: SpoDraft): void {
  supabase.from('pm_updates_cache').upsert({
    id: DRAFT_ID, updates: JSON.stringify(draft), updated_at: new Date().toISOString(),
  })
}

// One flat table: 3 payment rows per HOP, same columns as the SPO form.
export function batchTable(bucket: HopSpo[], draft: SpoDraft): string[][] {
  const out: string[][] = []
  bucket.forEach(h => {
    rowsFor(draft, h.hop).forEach(r => {
      out.push([h.appPathId, 'Base', h.pathId, h.hop, h.gc || '—', r.pct, money(r.value), cpoText(h), h.vendor, todayStr()])
    })
  })
  return out
}

export function emailBody(bucket: HopSpo[], draft: SpoDraft): string {
  const lines: string[] = [
    `Hi team,`,
    ``,
    `Please create the SPOs below for the GC. ${bucket.length} HOP${bucket.length === 1 ? '' : 's'}, 3 payment rows each.`,
    ``,
  ]
  bucket.forEach(h => {
    lines.push(`HOP ${h.appPathId}  |  ${h.hop}  |  GC: ${h.gc || '—'}  |  Customer Site ID ${h.pathId}  |  CPO# ${cpoText(h)}  |  Vendor ${h.vendor}`)
    rowsFor(draft, h.hop).forEach(r => {
      lines.push(`    Base  ${r.pct.padEnd(6)}  ${money(r.value).padEnd(14)}  Requested ${todayStr()}`)
    })
    lines.push('')
  })
  lines.push(`Thanks,`, `CJ`)
  return lines.join('\n')
}

export function openSpoEmail(bucket: HopSpo[], draft: SpoDraft): void {
  const subject = `SPO Requests — ${bucket.length} HOPs — ${todayStr()}`
  window.location.href = `mailto:${encodeURIComponent(draft.to)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(emailBody(bucket, draft))}`
}

export function downloadSpoExcel(bucket: HopSpo[], draft: SpoDraft): void {
  const ws = XLSX.utils.aoa_to_sheet([SPO_HEADERS, ...batchTable(bucket, draft)])
  ws['!cols'] = [12, 10, 16, 36, 24, 12, 14, 30, 26, 14].map(w => ({ wch: w }))
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, 'SPO Requests')
  XLSX.writeFile(wb, `SPO_Requests_${new Date().toISOString().slice(0, 10)}.xlsx`)
}

// Stages CX SPO Request = today on every physical row of every HOP in the
// batch, on the shared Pending Updates list (source: tracker).
export async function stageSpoRequested(bucket: HopSpo[]): Promise<number> {
  let list = await loadPendingUpdates()
  const stamp = todayStr()
  const ts = new Date().toISOString()
  bucket.forEach(h => {
    h.rows.forEach(r => {
      list = upsertPendingUpdate(list, {
        source: 'tracker', rowKey: r.rowKey, hop: h.hop, field: 'CX SPO Request',
        oldValue: '', newValue: stamp, timestamp: ts, user: 'CJ',
      })
    })
  })
  await persistPendingUpdates(list)
  return bucket.length
}
