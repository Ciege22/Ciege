'use client'

import { useState, useRef, useEffect } from 'react'
import { usePathname } from 'next/navigation'
import { loadTrackerSnapshot } from '../lib/supabase'
import { parseDateAny, loadGrRows, groupGrRows } from '../lib/grTracker'
import { parseDecomRows } from '../lib/decom'
import { loadChunkedReport } from '../lib/reportChunks'

const NAVY = '#124191'
const TEAL = '#00A0B0'

// Same Railway Flask backend already used for /build, /ntp_emails, /gr_data —
// it holds ANTHROPIC_API_KEY server-side and has CORS open for this app's
// origin already, so the browser never touches Anthropic (or any key) directly.
const AI_ASSISTANT_ENDPOINT = 'https://ciege-production.up.railway.app/ai_assistant'

const GREETING = "Hey! Listen! 🧚 I'm Navi — your Ciege guide! I have access to your full program data. Ask me about any site, blocker, GC status, or anything else. I'll help you find what you need fast!"

const QUICK_ACTIONS = [
  '📊 Program summary',
  '🔴 What needs attention today?',
  '🚧 Sites in progress',
  '⚠️ Top blockers',
]

const PAGE_LABELS: Record<string, string> = {
  '/': 'User is on Dashboard',
  '/gc-call': 'User is on GC Call View',
  '/cm-view': 'User is on CM Call View',
  '/gr-tracker': 'User is on GR Tracker',
  '/reports': 'User is on Reports page',
  '/schedule': 'User is on Schedule Optimizer',
  '/tracker': 'User is on Tracker Grid',
  '/deck-builder': 'User is on Deck Builder',
}

const MAX_HISTORY_MESSAGES = 20 // 10 user/assistant pairs

interface ChatMessage {
  role: 'user' | 'assistant'
  content: string
}

interface HopSummary {
  hop: string
  gc: string
  cm: string
  nokiaPm: string
  ms15f: string
  ms15a: string
  ms16f: string
  ms16a: string
  hasNtp: boolean
  hasMat: boolean
  status: 'Complete' | 'Active' | 'Not Started'
  lastNote: string
}

function normHeader(h: unknown): string {
  return String(h ?? '').trim().replace(/^'+|'+$/g, '')
}

function fmtCell(val: unknown): string {
  if (val === null || val === undefined || val === '') return ''
  if (val instanceof Date) return val.toLocaleDateString('en-US')
  const d = parseDateAny(val)
  return d ? d.toLocaleDateString('en-US') : String(val).trim()
}

// Same DON 444 filter + dedup-by-HOP rule every page uses (prefer the row
// where GC and New CM are both populated). Deliberately keeps every Nokia PM
// so "Total HOPs" is program-wide while "CJ HOPs" narrows to CJ's own.
function parseHopSummaries(rows: unknown[][]): HopSummary[] {
  let headerRowIdx = -1
  for (let i = 0; i < 10; i++) {
    const row = rows[i] as unknown[]
    if (row && row.some(c => normHeader(c) === 'HOP')) { headerRowIdx = i; break }
  }
  if (headerRowIdx === -1) return []

  const headers = (rows[headerRowIdx] as unknown[]).map(normHeader)
  const col = (name: string) => headers.findIndex(h => h === name)
  const hopCol = col('HOP')
  const don444Col = col('DON 444')
  const gcCol = col('General Contractor')
  const cmCol = col('New CM')
  const pmCol = col('Nokia PM')
  const ms15fCol = col('MS15 Implementation Start F')
  const ms15aCol = col('MS15 Implementation Start A')
  const ms16fCol = col('MS16 Implementation Ends F')
  const ms16aCol = col('MS16 Implementation Ends A')
  const ntpCol = col('NTP A')
  const matCol = col('Material Received A')
  const cxNotesCol = headers.findIndex(h => h === 'CX Notes:')

  const hopMap = new Map<string, unknown[][]>()
  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i] as unknown[]
    if (!row) continue
    const don = normHeader(row[don444Col]).toUpperCase()
    if (don !== 'DON 444') continue
    const hop = normHeader(row[hopCol])
    if (!hop || hop === 'undefined') continue
    if (!hopMap.has(hop)) hopMap.set(hop, [])
    hopMap.get(hop)!.push(row)
  }

  const summaries: HopSummary[] = []
  hopMap.forEach((candidateRows, hop) => {
    const chosen = candidateRows.find(r => normHeader(r[gcCol]) && normHeader(r[cmCol])) || candidateRows[0]
    const ms15a = fmtCell(chosen[ms15aCol])
    const ms16a = fmtCell(chosen[ms16aCol])
    const status: HopSummary['status'] = ms16a ? 'Complete' : (ms15a ? 'Active' : 'Not Started')

    const notesRaw = String(chosen[cxNotesCol] ?? '').trim()
    const lines = notesRaw.split('\n').map(l => l.trim()).filter(Boolean)
    const last = lines[lines.length - 1] || ''
    const lastNote = last.length > 100 ? last.slice(-100) : last

    summaries.push({
      hop,
      gc: normHeader(chosen[gcCol]) || '—',
      cm: normHeader(chosen[cmCol]) || '—',
      nokiaPm: normHeader(chosen[pmCol]) || '—',
      ms15f: fmtCell(chosen[ms15fCol]) || '—',
      ms15a: ms15a || '—',
      ms16f: fmtCell(chosen[ms16fCol]) || '—',
      ms16a: ms16a || '—',
      hasNtp: !!fmtCell(chosen[ntpCol]),
      hasMat: !!fmtCell(chosen[matCol]),
      status,
      lastNote: lastNote || '—',
    })
  })

  return summaries
}

async function buildContextString(pathname: string): Promise<{ text: string; degraded: boolean }> {
  let hopSummaries: HopSummary[] = []
  let grReadyCount = 0
  let decomOutstanding = 0
  let decomPodGap = 0
  let degraded = false

  try {
    const snap = await loadTrackerSnapshot()
    if (snap) hopSummaries = parseHopSummaries(snap.data)
    else degraded = true
  } catch {
    degraded = true
  }

  try {
    const grRows = await loadGrRows()
    grReadyCount = groupGrRows(grRows).ready.length
  } catch {
    degraded = true
  }

  try {
    const decomReport = await loadChunkedReport('decom')
    if (decomReport) {
      const decomRows = parseDecomRows(decomReport.rows)
      decomOutstanding = decomRows.filter(r => r.status === 'outstanding' || r.status === 'pending').length
      decomPodGap = decomRows.filter(r => r.status === 'pod_gap').length
    }
  } catch {
    // Decom tracker may simply not be uploaded yet — not fatal to the rest of the context.
  }

  const cjHops = hopSummaries.filter(h => h.nokiaPm.toUpperCase() === 'CJ')
  const active = cjHops.filter(h => h.status === 'Active').length
  const complete = cjHops.filter(h => h.status === 'Complete').length
  const ntpComplete = cjHops.filter(h => h.hasNtp).length
  const materialReceived = cjHops.filter(h => h.hasMat).length

  const pageNote = PAGE_LABELS[pathname]

  const detailLines = [...hopSummaries]
    .sort((a, b) => a.hop.localeCompare(b.hop))
    .slice(0, 150)
    .map(h => `${h.hop} | ${h.gc} | ${h.cm} | ${h.nokiaPm} | ${h.ms15f} | ${h.ms15a} | ${h.ms16f} | ${h.ms16a} | ${h.hasNtp ? 'Yes' : 'No'} | ${h.hasMat ? 'Yes' : 'No'} | ${h.status} | ${h.lastNote}`)
    .join('\n')

  const text = `PROGRAM CONTEXT — Viaero MW Program, Nokia DON 444
Total HOPs: ${hopSummaries.length} | CJ HOPs: ${cjHops.length} | Active: ${active} | Complete: ${complete}
NTP Complete: ${ntpComplete} | Material Received: ${materialReceived}
GR Ready to Release: ${grReadyCount} SPOs pending
Decom Outstanding: ${decomOutstanding} sites | POD Gap: ${decomPodGap} sites
${pageNote ? `\n${pageNote}\n` : ''}
HOP DETAIL (first 150 HOPs, pipe-delimited):
HOP | GC | CM | Nokia PM | MS15F | MS15A | MS16F | MS16A | NTP | Material | Status | CX Notes (last note)
${detailLines}`

  return { text, degraded }
}

function trimHistory(history: ChatMessage[]): ChatMessage[] {
  return history.length > MAX_HISTORY_MESSAGES ? history.slice(history.length - MAX_HISTORY_MESSAGES) : history
}

// Navi — the animated fairy guide. Defined at module scope like every other
// piece of this component, for the same reason: a component defined inside
// another component's render body gets a fresh identity every render, which
// would restart her float/flutter/glow animations on every keystroke.
function NaviFairy({ size = 48, bright = false }: { size?: number; bright?: boolean }) {
  return (
    <div className="navi-float-layer" style={{ animationDuration: bright ? '1s' : '2s', width: size, height: size }}>
      <svg width={size} height={size} viewBox="0 0 48 48" style={{ overflow: 'visible' }}>
        <defs>
          <radialGradient id="navi-body-gradient" cx="50%" cy="50%" r="50%">
            <stop offset="0%" stopColor="#FFFFFF" />
            <stop offset="60%" stopColor="#E0FBFF" />
            <stop offset="100%" stopColor="#00A0B0" />
          </radialGradient>
          <filter id="navi-glow-filter" x="-150%" y="-150%" width="400%" height="400%">
            <feGaussianBlur in="SourceGraphic" stdDeviation="5" result="blur" />
            <feComposite in="blur" in2="SourceGraphic" operator="over" />
          </filter>
        </defs>

        {/* outer glow aura */}
        <circle
          cx="24" cy="24" r="10" fill="#00A0B0" filter="url(#navi-glow-filter)"
          className={bright ? undefined : 'navi-glow'}
          style={bright ? { opacity: 1 } : undefined}
        />

        {/* two pairs of translucent wings */}
        <g className="navi-wings">
          <ellipse cx="14" cy="18" rx="7" ry="4" fill="#FFFFFF" opacity="0.35" transform="rotate(-20 14 18)" />
          <ellipse cx="34" cy="18" rx="7" ry="4" fill="#FFFFFF" opacity="0.35" transform="rotate(20 34 18)" />
          <ellipse cx="14" cy="28" rx="6" ry="3.5" fill="#00A0B0" opacity="0.3" transform="rotate(15 14 28)" />
          <ellipse cx="34" cy="28" rx="6" ry="3.5" fill="#00A0B0" opacity="0.3" transform="rotate(-15 34 28)" />
        </g>

        {/* glowing body — white center fading to teal */}
        <circle cx="24" cy="24" r="6" fill="url(#navi-body-gradient)" />
        <circle cx="24" cy="24" r="2.5" fill="#FFFFFF" opacity="0.9" />
      </svg>
    </div>
  )
}

function LoadingDots() {
  return (
    <div className="flex gap-1 items-center px-3 py-2">
      <span className="w-1.5 h-1.5 rounded-full bg-gray-400 animate-bounce" style={{ animationDelay: '0ms' }} />
      <span className="w-1.5 h-1.5 rounded-full bg-gray-400 animate-bounce" style={{ animationDelay: '150ms' }} />
      <span className="w-1.5 h-1.5 rounded-full bg-gray-400 animate-bounce" style={{ animationDelay: '300ms' }} />
    </div>
  )
}

export default function AiAssistant() {
  const pathname = usePathname()
  const [isOpen, setIsOpen] = useState(false)
  const [messages, setMessages] = useState<ChatMessage[]>([])
  const [input, setInput] = useState('')
  const [loading, setLoading] = useState(false)
  const [contextLoading, setContextLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [degradedNote, setDegradedNote] = useState(false)

  const hasOpenedRef = useRef(false)
  // Cached context string, keyed by whatever localStorage's lastUploadTime
  // said when it was built — rebuilt only when that value changes (or on
  // first open). No page in the app writes lastUploadTime yet, so today this
  // just means "build once per session" — it's wired up for any upload flow
  // to opt into invalidating it later.
  const contextCacheRef = useRef<{ text: string; forUploadTime: string | null } | null>(null)
  const messagesEndRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    if (isOpen) messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' })
  }, [messages, isOpen, loading])

  const getLastUploadTime = (): string | null => {
    try { return localStorage.getItem('lastUploadTime') } catch { return null }
  }

  const ensureContext = async (): Promise<string> => {
    const currentUploadTime = getLastUploadTime()
    if (contextCacheRef.current && contextCacheRef.current.forUploadTime === currentUploadTime) {
      return contextCacheRef.current.text
    }
    setContextLoading(true)
    const { text, degraded } = await buildContextString(pathname)
    contextCacheRef.current = { text, forUploadTime: currentUploadTime }
    setDegradedNote(degraded)
    setContextLoading(false)
    return text
  }

  const openPanel = () => {
    setIsOpen(true)
    if (!hasOpenedRef.current) {
      hasOpenedRef.current = true
      setMessages([{ role: 'assistant', content: GREETING }])
      // Kick off context loading in the background — never blocks opening the panel.
      ensureContext().catch(() => setDegradedNote(true))
    }
  }

  const sendMessage = async (text: string) => {
    const trimmedText = text.trim()
    if (!trimmedText || loading) return
    setError(null)
    setInput('')
    setMessages(prev => trimHistory([...prev, { role: 'user', content: trimmedText }]))
    setLoading(true)
    try {
      const contextText = await ensureContext()
      const historyForApi = trimHistory([...messages, { role: 'user', content: trimmedText }])
        .map(m => ({ role: m.role, content: m.content }))

      const res = await fetch(AI_ASSISTANT_ENDPOINT, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ messages: historyForApi, context: contextText }),
      })
      if (!res.ok) throw new Error(`bad status ${res.status}`)
      const data = await res.json()
      if (!data.response) throw new Error('empty response')
      setMessages(prev => trimHistory([...prev, { role: 'assistant', content: data.response }]))
    } catch {
      setError('Something went wrong — try again')
    } finally {
      setLoading(false)
    }
  }

  const handleSend = () => sendMessage(input)

  return (
    <>
      {!isOpen && (
        <div
          className="group"
          style={{ position: 'fixed', bottom: 24, right: 24, zIndex: 9999 }}
        >
          <div
            className="absolute -top-9 right-0 whitespace-nowrap text-white text-xs font-semibold px-3 py-1.5 rounded-full opacity-0 group-hover:opacity-100 transition-opacity duration-300 pointer-events-none"
            style={{ backgroundColor: NAVY }}
          >
            Hey! Listen! Click me 🧚
          </div>
          <button
            onClick={openPanel}
            aria-label="Open Ciege AI Assistant — Navi"
            className="navi-hover-wrapper"
            style={{
              background: 'none',
              border: 'none',
              padding: 0,
              cursor: 'pointer',
              filter: 'drop-shadow(0 4px 8px rgba(0,0,0,0.35))',
            }}
          >
            <NaviFairy size={48} />
          </button>
        </div>
      )}

      {isOpen && (
        <div
          style={{
            position: 'fixed',
            bottom: 90,
            right: 24,
            width: 380,
            height: 520,
            zIndex: 9999,
            backgroundColor: '#FFFFFF',
            borderRadius: 12,
            boxShadow: '0 12px 40px rgba(0,0,0,0.35)',
            display: 'flex',
            flexDirection: 'column',
            overflow: 'hidden',
          }}
        >
          <div style={{ backgroundColor: NAVY, padding: '12px 16px', flexShrink: 0 }}>
            <div className="flex items-center justify-between">
              <span className="text-white font-bold text-sm">Ciege AI Assistant</span>
              <div className="flex items-center gap-2">
                <NaviFairy size={24} bright />
                <button
                  onClick={() => setIsOpen(false)}
                  aria-label="Close"
                  className="text-white hover:text-gray-300 text-lg leading-none"
                >
                  ×
                </button>
              </div>
            </div>
            <p style={{ color: TEAL }} className="text-xs mt-0.5">Ask anything about your program data</p>
          </div>

          <div className="ai-assistant-scroll flex-1 overflow-y-auto bg-white px-3 py-3 flex flex-col gap-2">
            {messages.map((m, i) => (
              <div
                key={i}
                className={`max-w-[85%] px-3 py-2 rounded-2xl text-sm whitespace-pre-wrap ${m.role === 'user' ? 'self-end text-white' : 'self-start text-gray-800'}`}
                style={{ backgroundColor: m.role === 'user' ? NAVY : '#F1F1F3' }}
              >
                {m.content}
              </div>
            ))}

            {messages.length <= 1 && !loading && (
              <div className="flex flex-wrap gap-1.5 mt-1">
                {QUICK_ACTIONS.map(action => (
                  <button
                    key={action}
                    onClick={() => sendMessage(action)}
                    className="text-xs bg-gray-100 hover:bg-gray-200 text-gray-700 px-2.5 py-1.5 rounded-full border border-gray-200"
                  >
                    {action}
                  </button>
                ))}
              </div>
            )}

            {contextLoading && (
              <div className="self-start text-xs text-gray-400 italic px-1">Loading program data...</div>
            )}
            {degradedNote && !contextLoading && (
              <div className="self-start text-xs text-amber-600 italic px-1">Some data unavailable — upload tracker to improve answers</div>
            )}

            {loading && (
              <div className="self-start rounded-2xl" style={{ backgroundColor: '#F1F1F3' }}>
                <LoadingDots />
              </div>
            )}

            {error && (
              <div className="self-start text-red-600 text-xs px-1">{error}</div>
            )}

            <div ref={messagesEndRef} />
          </div>

          <div className="flex-shrink-0 border-t border-gray-200 p-2.5 flex gap-2 bg-white">
            <input
              type="text"
              value={input}
              onChange={(e) => setInput(e.target.value)}
              onKeyDown={(e) => { if (e.key === 'Enter') handleSend() }}
              placeholder="Ask about any site, GC, blocker..."
              className="flex-1 text-sm border border-gray-300 rounded-full px-3 py-2 focus:outline-none focus:border-blue-500"
            />
            <button
              onClick={handleSend}
              disabled={!input.trim() || loading}
              style={{ backgroundColor: (!input.trim() || loading) ? '#9CA3AF' : NAVY }}
              className="text-white text-sm font-semibold px-4 py-2 rounded-full disabled:cursor-not-allowed flex-shrink-0"
            >
              Send
            </button>
          </div>
        </div>
      )}
    </>
  )
}
