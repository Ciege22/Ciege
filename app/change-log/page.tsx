'use client'

export const dynamic = 'force-dynamic'

import { useState, useEffect } from 'react'
import { getChangeLog, getSchemaChanges, getAllSnapshots } from '../lib/supabase'
import BackToDashboard from '../components/BackToDashboard'

interface Change {
  id: string
  upload_id: string
  uploaded_at: string
  hop_name: string
  gc: string
  cm: string
  nokia_pm: string
  change_type: string
  field_name: string
  old_value: string
  new_value: string
}

interface SchemaChange {
  id: string
  upload_id: string
  uploaded_at: string
  change_type: string
  column_name: string
}

interface Snapshot {
  id: string
  filename: string
  uploaded_at: string
  hop_count: number
}

const CHANGE_FILTERS = [
  'All',
  '🆕 New HOP',
  '❌ HOP Removed',
  '🚀 Site Started',
  '✅ Site Completed',
  '✅ NTP Confirmed',
  '📦 Material Received',
  '📋 SPO Issued',
  '📡 MSS Completed',
  '⚡ Power-Up Complete',
  '📅 FC Start',
  '📅 FC End',
  '🔄 GC Changed',
  '🔄 CM Changed',
  '📝 Updated',
]

export default function ChangeLogPage() {
  const [changes, setChanges] = useState<Change[]>([])
  const [schemaChanges, setSchemaChanges] = useState<SchemaChange[]>([])
  const [snapshots, setSnapshots] = useState<Snapshot[]>([])
  const [loaded, setLoaded] = useState(false)
  const [filter, setFilter] = useState('All')
  const [selectedUpload, setSelectedUpload] = useState('All')
  const [searchHop, setSearchHop] = useState('')

  useEffect(() => {
    const load = async () => {
      const [c, s, snaps] = await Promise.all([getChangeLog(), getSchemaChanges(), getAllSnapshots()])
      setChanges(c as Change[])
      setSchemaChanges(s as SchemaChange[])
      setSnapshots(snaps)
      setLoaded(true)
    }
    load()
  }, [])

  const filtered = changes.filter(c => {
    const typeMatch = filter === 'All' || c.change_type.startsWith(filter.split(' ')[0]) || c.change_type === filter
    const uploadMatch = selectedUpload === 'All' || c.upload_id === selectedUpload
    const hopMatch = !searchHop || c.hop_name.toLowerCase().includes(searchHop.toLowerCase())
    return typeMatch && uploadMatch && hopMatch
  })

  const groupedByUpload = snapshots.map(snap => ({
    snap,
    changes: filtered.filter(c => c.upload_id === snap.id),
    schemaChanges: schemaChanges.filter(s => s.upload_id === snap.id)
  })).filter(g => g.changes.length > 0 || g.schemaChanges.length > 0)

  const changeTypeBadgeColor = (type: string) => {
    if (type.includes('New HOP')) return 'bg-blue-800 text-blue-200'
    if (type.includes('Removed')) return 'bg-red-800 text-red-200'
    if (type.includes('Started') || type.includes('Completed') || type.includes('Confirmed') || type.includes('Received') || type.includes('Issued') || type.includes('MSS') || type.includes('Power-Up')) return 'bg-green-800 text-green-200'
    if (type.includes('Pushed')) return 'bg-orange-800 text-orange-200'
    if (type.includes('Pulled')) return 'bg-teal-800 text-teal-200'
    if (type.includes('Changed')) return 'bg-yellow-800 text-yellow-200'
    return 'bg-gray-700 text-gray-300'
  }

  return (
    <div className="min-h-screen bg-gray-950 text-white p-6">
      <div className="max-w-full mx-auto">

        <BackToDashboard />
        <div className="mb-6">
          <h1 className="text-3xl font-bold">Tracker Change Log</h1>
          <p className="text-gray-400 mt-1">What changed between each tracker upload — HOPs, dates, milestones, and schema</p>
        </div>

        {!loaded && <p className="text-gray-400">Loading change log...</p>}

        {loaded && changes.length === 0 && (
          <div className="bg-gray-900 rounded-xl border border-gray-700 p-12 text-center">
            <p className="text-gray-400 text-xl">No changes recorded yet</p>
            <p className="text-gray-600 mt-2">Upload a second tracker on the Dashboard to start tracking changes</p>
          </div>
        )}

        {loaded && changes.length > 0 && (
          <>
            {/* Summary Cards */}
            <div className="grid grid-cols-2 md:grid-cols-5 gap-3 mb-6">
              {[
                { label: 'Total Changes', value: changes.length, color: 'text-white' },
                { label: 'Sites Started', value: changes.filter(c => c.change_type === '🚀 Site Started').length, color: 'text-blue-400' },
                { label: 'Sites Completed', value: changes.filter(c => c.change_type === '✅ Site Completed').length, color: 'text-green-400' },
                { label: 'NTPs Confirmed', value: changes.filter(c => c.change_type === '✅ NTP Confirmed').length, color: 'text-teal-400' },
                { label: 'Dates Moved', value: changes.filter(c => c.change_type.includes('FC Start') || c.change_type.includes('FC End')).length, color: 'text-yellow-400' },
              ].map(({ label, value, color }) => (
                <div key={label} className="bg-gray-900 rounded-xl border border-gray-700 p-3 text-center">
                  <p className="text-gray-500 text-xs">{label}</p>
                  <p className={`text-2xl font-bold ${color}`}>{value}</p>
                </div>
              ))}
            </div>

            {/* Filters */}
            <div className="bg-gray-900 rounded-xl border border-gray-700 p-4 mb-6">
              <div className="flex gap-3 flex-wrap items-center mb-3">
                <span className="text-gray-500 text-xs font-semibold">Upload:</span>
                <select value={selectedUpload} onChange={e => setSelectedUpload(e.target.value)}
                  className="bg-gray-800 text-gray-300 text-xs rounded px-2 py-1 border border-gray-600">
                  <option value="All">All Uploads</option>
                  {snapshots.map(s => (
                    <option key={s.id} value={s.id}>
                      {new Date(s.uploaded_at).toLocaleDateString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric' })} at {new Date(s.uploaded_at).toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })} — {s.hop_count} HOPs
                    </option>
                  ))}
                </select>
                <span className="text-gray-500 text-xs font-semibold ml-2">Search HOP:</span>
                <input type="text" placeholder="HOP name..." value={searchHop}
                  onChange={e => setSearchHop(e.target.value)}
                  className="bg-gray-800 text-gray-300 text-xs rounded px-2 py-1 border border-gray-600 w-48 focus:outline-none focus:border-blue-500" />
              </div>
              <div className="flex gap-2 flex-wrap">
                {CHANGE_FILTERS.map(f => (
                  <button key={f} onClick={() => setFilter(f)}
                    className={`px-3 py-1 rounded-full text-xs font-semibold transition-all ${filter === f ? 'bg-blue-600 text-white' : 'bg-gray-800 text-gray-400 hover:bg-gray-700'}`}>
                    {f}
                  </button>
                ))}
              </div>
            </div>

            {/* Schema Changes */}
            {schemaChanges.length > 0 && (
              <div className="bg-gray-900 rounded-xl border border-yellow-700 p-4 mb-6">
                <h3 className="text-yellow-300 font-bold text-sm mb-3">📐 Schema Changes — Columns Added or Removed</h3>
                <div className="space-y-1">
                  {schemaChanges.map(s => (
                    <div key={s.id} className="flex items-center gap-3 text-xs">
                      <span className={`px-2 py-0.5 rounded-full font-semibold ${s.change_type.includes('Added') ? 'bg-green-800 text-green-200' : 'bg-red-800 text-red-200'}`}>
                        {s.change_type}
                      </span>
                      <span className="text-white font-mono">{s.column_name}</span>
                      <span className="text-gray-500">{new Date(s.uploaded_at).toLocaleDateString()}</span>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Change Log grouped by upload */}
            {selectedUpload === 'All' ? (
              groupedByUpload.map(({ snap, changes: snapChanges }) => (
                <div key={snap.id} className="mb-6 bg-gray-900 rounded-xl border border-gray-700 overflow-hidden">
                  <div className="bg-gray-800 px-4 py-3 flex items-center justify-between">
                    <div>
                      <span className="text-white font-bold text-sm">
                        {new Date(snap.uploaded_at).toLocaleDateString('en-US', { weekday: 'short', month: 'long', day: 'numeric', year: 'numeric' })} at {new Date(snap.uploaded_at).toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })}
                      </span>
                      <span className="text-gray-400 text-xs ml-3">{snap.filename}</span>
                    </div>
                    <span className="text-blue-400 text-xs font-bold">{snapChanges.length} changes</span>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full text-xs">
                      <thead>
                        <tr className="text-gray-400 border-b border-gray-700">
                          <th className="text-left p-2">Change</th>
                          <th className="text-left p-2">HOP</th>
                          <th className="text-left p-2">GC</th>
                          <th className="text-left p-2">CM</th>
                          <th className="text-left p-2">Nokia PM</th>
                          <th className="text-left p-2">Field</th>
                          <th className="text-left p-2">Was</th>
                          <th className="text-left p-2">Now</th>
                        </tr>
                      </thead>
                      <tbody>
                        {snapChanges.map(c => (
                          <tr key={c.id} className="border-t border-gray-800 hover:bg-gray-800">
                            <td className="p-2">
                              <span className={`px-2 py-0.5 rounded-full text-xs font-semibold whitespace-nowrap ${changeTypeBadgeColor(c.change_type)}`}>
                                {c.change_type}
                              </span>
                            </td>
                            <td className="p-2 font-semibold text-white whitespace-nowrap">{c.hop_name}</td>
                            <td className="p-2 text-gray-300 whitespace-nowrap">{c.gc || '—'}</td>
                            <td className="p-2 text-gray-300 whitespace-nowrap">{c.cm || '—'}</td>
                            <td className="p-2 text-gray-300 whitespace-nowrap">{c.nokia_pm || '—'}</td>
                            <td className="p-2 text-gray-400 whitespace-nowrap">{c.field_name}</td>
                            <td className="p-2 text-red-400 whitespace-nowrap">{c.old_value || '—'}</td>
                            <td className="p-2 text-green-400 whitespace-nowrap">{c.new_value || '—'}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              ))
            ) : (
              <div className="bg-gray-900 rounded-xl border border-gray-700 overflow-hidden">
                <div className="overflow-x-auto">
                  <table className="w-full text-xs">
                    <thead>
                      <tr className="text-gray-400 bg-gray-800 border-b border-gray-700">
                        <th className="text-left p-2">Change</th>
                        <th className="text-left p-2">HOP</th>
                        <th className="text-left p-2">GC</th>
                        <th className="text-left p-2">CM</th>
                        <th className="text-left p-2">Nokia PM</th>
                        <th className="text-left p-2">Field</th>
                        <th className="text-left p-2">Was</th>
                        <th className="text-left p-2">Now</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filtered.map(c => (
                        <tr key={c.id} className="border-t border-gray-800 hover:bg-gray-800">
                          <td className="p-2">
                            <span className={`px-2 py-0.5 rounded-full text-xs font-semibold whitespace-nowrap ${changeTypeBadgeColor(c.change_type)}`}>
                              {c.change_type}
                            </span>
                          </td>
                          <td className="p-2 font-semibold text-white whitespace-nowrap">{c.hop_name}</td>
                          <td className="p-2 text-gray-300 whitespace-nowrap">{c.gc || '—'}</td>
                          <td className="p-2 text-gray-300 whitespace-nowrap">{c.cm || '—'}</td>
                          <td className="p-2 text-gray-300 whitespace-nowrap">{c.nokia_pm || '—'}</td>
                          <td className="p-2 text-gray-400 whitespace-nowrap">{c.field_name}</td>
                          <td className="p-2 text-red-400 whitespace-nowrap">{c.old_value || '—'}</td>
                          <td className="p-2 text-green-400 whitespace-nowrap">{c.new_value || '—'}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </>
        )}
      </div>
    </div>
  )
}
