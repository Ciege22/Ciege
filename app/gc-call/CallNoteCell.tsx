'use client'

import React from 'react'

export interface CallNote {
  id: string
  hop_name: string
  note: string
  logged_at: string
}

// Call-note input + history cells, matching the Pipeline tab's "Call Notes
// (Today)" / "Notes History" pattern in this directory's page.tsx, keyed by
// an arbitrary string instead of a HOP so Decom/SCOP rows — which can repeat
// a HOP across sites — get their own independent note thread. Shared between
// page.tsx (DecomTab) and ScopGcTab.tsx.
export function CallNoteCell({ noteKey, sessionNotes, setSessionNotes, saveCallNote }: {
  noteKey: string
  sessionNotes: Record<string, string>
  setSessionNotes: React.Dispatch<React.SetStateAction<Record<string, string>>>
  saveCallNote: (key: string) => void
}) {
  return (
    <div className="flex gap-1">
      <input type="text" placeholder="Note..." value={sessionNotes[noteKey] || ''}
        onChange={(e) => setSessionNotes(s => ({ ...s, [noteKey]: e.target.value }))}
        onKeyDown={(e) => { if (e.key === 'Enter') saveCallNote(noteKey) }}
        className="w-36 bg-gray-800 text-white text-xs rounded px-2 py-1 border border-gray-600 focus:outline-none focus:border-blue-500" />
      <button onClick={() => saveCallNote(noteKey)} className="text-xs bg-blue-700 hover:bg-blue-600 text-white px-2 py-1 rounded">💾</button>
    </div>
  )
}

export function CallNoteHistoryCell({ noteKey, noteHistory }: {
  noteKey: string
  noteHistory: Record<string, CallNote[]>
}) {
  return (
    <div className="max-h-20 overflow-y-auto flex flex-col gap-1">
      {(noteHistory[noteKey] || []).slice(0, 5).map((n, i) => (
        <div key={i} className="text-xs text-gray-300 border-b border-gray-700 pb-1">
          <span className="text-gray-500 text-xs">{new Date(n.logged_at).toLocaleDateString()}</span>
          <span className="ml-1">{n.note}</span>
        </div>
      ))}
      {!noteHistory[noteKey]?.length && <span className="text-gray-600 text-xs">No history</span>}
    </div>
  )
}
