'use client'

export const dynamic = 'force-dynamic'

import { useEffect, useState } from 'react'
import * as XLSX from 'xlsx'
import { supabase, loadTrackerSnapshot } from '../lib/supabase'
import BackToDashboard from '../components/BackToDashboard'
import { computeSpoStatus } from '../lib/spoStatus'
import { loadPendingUpdates, persistPendingUpdates, upsertPendingUpdate } from '../lib/pendingUpdates'

// SPO Request builder — lists every HOP in the "needs SPO requested" bucket
// (CPO received, no SPO requested yet), lets you enter each HOP's three
// payment-split rows, and emails the whole batch to the team as one
// table, plus an Excel download to attach.

const DRAFT_ID = 'spo-request-draft'
const DEFAULT_SPLIT = [20, 60, 20]

interface PhysicalRow {
  rowKey: string
  siteName: string
  pathId: string
  cpo: string
}

interface HopSpo {
  hop: string
  appPathId: string
  pathId: string
  vendor: string
  gc: string
  siteNames: string[]
  cpos: string[]
  rows: PhysicalRow[]
}

interface PaymentRow {
  pct: string
  value: string
}

interface Draft {
  to: string
  hops: Record<string, PaymentRow[]>
}

function validDate(v: unknown): boolean {
  if (!v) return false
  const d = new Date(String(v))
  return !isNaN(d.getTime()) && d.getFullYear() >= 1990
}

function todayStr(): string {
  return new Date().toLocaleDateString('en-US')
}

function money(v: string): string {
  const n = Number(String(v).replace(/[$,\s]/g, ''))
  if (!v || !Number.isFinite(n)) return v ? String(v) : '—'
  return n.toLocaleString('en-US', { style: 'currency', currency: 'USD' })
}

function defaultRows(): PaymentRow[] {
  return DEFAULT_SPLIT.map(p => ({ pct: `${p}%`, value: '' }))
}

export default function SpoRequestPage() {
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)
  const [hops, setHops] = useState<HopSpo[]>([])
  const [draft, setDraft] = useState<Draft>({ to: '', hops: {} })
  const [status, setStatus] = useState<string | null>(null)

  // Load tracker -> build the needs_request bucket.
  useEffect(() => {
    let cancelled = false
    loadTrackerSnapshot().then(snap => {
      if (cancelled) return
      if (!snap) { setError('No tracker found — upload one on the Dashboard first.'); setLoading(false); return }
      const rows = snap.data
      let headerIdx = -1
      for (let i = 0; i < 10; i++) {
        const r = rows[i] as unknown[]
        if (r && r.some(c => String(c ?? '').trim() === 'HOP')) { headerIdx = i; break }
      }
      if (headerIdx === -1) { setError('Could not find the HOP header row.'); setLoading(false); return }
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
        hasSpo: boolean; hasReq: boolean; hasCpo: boolean
        rows: PhysicalRow[]
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
        // "CX SPO Request" stamp below lands on the same physical rows.
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
        h.rows.push({ rowKey, siteName, pathId, cpo })
      }

      const bucket: HopSpo[] = []
      byHop.forEach((h, hop) => {
        const st = computeSpoStatus(h.hasSpo, h.hasCpo, h.hasReq)
        if (st.status !== 'needs_request') return
        bucket.push({
          hop, appPathId: h.appPathId, pathId: h.pathId, vendor: h.vendor, gc: h.gc,
          siteNames: Array.from(new Set(h.rows.map(r => r.siteName))).filter(Boolean),
          cpos: h.rows.map(r => r.cpo).filter(Boolean),
          rows: h.rows,
        })
      })
      bucket.sort((a, b) => a.hop.localeCompare(b.hop))
      setHops(bucket)
      setLoading(false)
    }).catch(err => {
      if (!cancelled) { setError(err instanceof Error ? err.message : 'Failed to load tracker.'); setLoading(false) }
    })
    return () => { cancelled = true }
  }, [])

  // Load saved drafts (recipient + entered values).
  useEffect(() => {
    supabase.from('pm_updates_cache').select('updates').eq('id', DRAFT_ID).single().then(({ data }) => {
      if (data?.updates) {
        try { setDraft(JSON.parse(data.updates)) } catch { /* ignore corrupt draft */ }
      }
    })
  }, [])

  const saveDraft = (next: Draft) => {
    setDraft(next)
    supabase.from('pm_updates_cache').upsert({
      id: DRAFT_ID, updates: JSON.stringify(next), updated_at: new Date().toISOString(),
    })
  }

  const rowsFor = (hop: string): PaymentRow[] => draft.hops[hop] ?? defaultRows()

  const updateRow = (hop: string, idx: number, field: keyof PaymentRow, val: string) => {
    const current = rowsFor(hop).map(r => ({ ...r }))
    current[idx][field] = val
    saveDraft({ ...draft, hops: { ...draft.hops, [hop]: current } })
  }

  const cpoText = (h: HopSpo) => Array.from(new Set(h.cpos)).join(', ')

  // One flat list: 3 payment rows per HOP, same columns as the SPO form.
  const table: string[][] = []
  hops.forEach(h => {
    rowsFor(h.hop).forEach(r => {
      table.push([h.appPathId, 'Base', h.pathId, h.hop, r.pct, money(r.value), cpoText(h), h.vendor, todayStr()])
    })
  })

  const headerRow = ['HOP (App Path ID)', 'SPO type', 'Customer Site ID', 'Site Name', '% payment', 'SPO value', 'CPO#', 'Vendor Name', 'SPO Request']

  const emailBody = () => {
    const lines: string[] = [
      `Hi team,`,
      ``,
      `Please create the SPOs below for the GC. ${hops.length} HOP${hops.length === 1 ? '' : 's'}, 3 payment rows each.`,
      ``,
    ]
    hops.forEach(h => {
      lines.push(`HOP ${h.appPathId}  |  ${h.hop}  |  Customer Site ID ${h.pathId}  |  CPO# ${cpoText(h)}  |  Vendor ${h.vendor}`)
      rowsFor(h.hop).forEach(r => {
        lines.push(`    Base  ${r.pct.padEnd(6)}  ${money(r.value).padEnd(14)}  Requested ${todayStr()}`)
      })
      lines.push('')
    })
    lines.push(`Thanks,`, `CJ`)
    return lines.join('\n')
  }

  const sendEmail = () => {
    const subject = `SPO Requests — ${hops.length} HOPs — ${todayStr()}`
    const href = `mailto:${encodeURIComponent(draft.to)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(emailBody())}`
    window.location.href = href
  }

  const downloadExcel = () => {
    const ws = XLSX.utils.aoa_to_sheet([headerRow, ...table])
    ws['!cols'] = [12, 10, 16, 36, 12, 14, 30, 26, 14].map(w => ({ wch: w }))
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'SPO Requests')
    XLSX.writeFile(wb, `SPO_Requests_${new Date().toISOString().slice(0, 10)}.xlsx`)
  }

  // Stamp CX SPO Request = today on every physical row of every HOP in the
  // batch, staged on the shared Pending Updates list (source: tracker), so
  // the bucket moves once those land in the master tracker.
  const stampRequested = async () => {
    setStatus('Staging CX SPO Request dates…')
    try {
      let list = await loadPendingUpdates()
      const stamp = todayStr()
      const ts = new Date().toISOString()
      hops.forEach(h => {
        h.rows.forEach(r => {
          list = upsertPendingUpdate(list, {
            source: 'tracker', rowKey: r.rowKey, hop: h.hop, field: 'CX SPO Request',
            oldValue: '', newValue: stamp, timestamp: ts, user: 'CJ',
          })
        })
      })
      await persistPendingUpdates(list)
      setStatus(`Staged CX SPO Request = ${stamp} on ${hops.length} HOPs. Review them on the Pending Updates list.`)
    } catch (err) {
      setStatus(err instanceof Error ? `Failed: ${err.message}` : 'Failed to stage updates.')
    }
  }

  if (loading) return <div className="p-6 text-zinc-400 text-sm">Loading SPO bucket…</div>
  if (error) return <div className="p-6 text-red-300 text-sm">{error}</div>

  return (
    <div className="min-h-screen bg-zinc-950 text-zinc-100 p-4">
      <BackToDashboard />
      <h1 className="text-2xl font-semibold text-white mb-1">SPO Requests</h1>
      <p className="text-sm text-zinc-400 mb-4">
        {hops.length} HOP{hops.length === 1 ? '' : 's'} with a CPO received and no SPO requested yet. Enter each payment split&apos;s SPO value, then email the batch.
      </p>

      <div className="flex flex-wrap items-center gap-3 mb-4 bg-zinc-900 border border-white/10 rounded-xl p-3">
        <input
          type="email"
          placeholder="Team email address"
          value={draft.to}
          onChange={e => saveDraft({ ...draft, to: e.target.value })}
          className="flex-1 min-w-[220px] bg-zinc-800 text-sm rounded px-3 py-2 border border-white/10"
        />
        <button onClick={sendEmail} disabled={hops.length === 0}
          className="bg-blue-600 hover:bg-blue-500 disabled:opacity-40 text-white text-sm font-semibold rounded px-4 py-2">
          ✉️ Email SPO requests
        </button>
        <button onClick={downloadExcel} disabled={hops.length === 0}
          className="bg-emerald-700 hover:bg-emerald-600 disabled:opacity-40 text-white text-sm font-semibold rounded px-4 py-2">
          ⬇️ Download Excel
        </button>
        <button onClick={stampRequested} disabled={hops.length === 0}
          className="bg-amber-700 hover:bg-amber-600 disabled:opacity-40 text-white text-sm font-semibold rounded px-4 py-2">
          📌 Stamp CX SPO Request = today
        </button>
      </div>
      {status && <div className="mb-4 text-sm text-zinc-300">{status}</div>}

      {hops.length === 0 ? (
        <div className="text-zinc-400 text-sm">Nothing waiting — every CPO-received HOP already has an SPO requested.</div>
      ) : (
        <div className="space-y-3">
          {hops.map(h => (
            <div key={h.hop} className="rounded-xl border border-white/10 bg-zinc-900 p-4">
              <div className="flex flex-wrap items-baseline gap-x-4 gap-y-1 mb-3">
                <span className="text-sm font-bold text-white">HOP {h.appPathId}</span>
                <span className="text-sm text-zinc-300">{h.hop}</span>
                <span className="text-xs text-zinc-500">Customer Site ID {h.pathId}</span>
                <span className="text-xs text-zinc-500">CPO# {cpoText(h) || '—'}</span>
                <span className="text-xs text-zinc-500">{h.vendor}</span>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-xs">
                  <thead>
                    <tr className="text-zinc-500 text-left">
                      <th className="py-1 pr-3">SPO type</th>
                      <th className="py-1 pr-3">% payment</th>
                      <th className="py-1 pr-3">SPO value</th>
                      <th className="py-1 pr-3">SPO Request</th>
                    </tr>
                  </thead>
                  <tbody>
                    {rowsFor(h.hop).map((r, i) => (
                      <tr key={i} className="border-t border-white/5">
                        <td className="py-1.5 pr-3 text-zinc-300">Base</td>
                        <td className="py-1.5 pr-3">
                          <input value={r.pct} onChange={e => updateRow(h.hop, i, 'pct', e.target.value)}
                            className="w-20 bg-zinc-800 rounded px-2 py-1 border border-white/10" />
                        </td>
                        <td className="py-1.5 pr-3">
                          <input value={r.value} placeholder="$0.00" onChange={e => updateRow(h.hop, i, 'value', e.target.value)}
                            className="w-32 bg-zinc-800 rounded px-2 py-1 border border-white/10" />
                        </td>
                        <td className="py-1.5 pr-3 text-zinc-300">{todayStr()}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  )
}
