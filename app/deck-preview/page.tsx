'use client'

export const dynamic = 'force-dynamic'

import BackToDashboard from '../components/BackToDashboard'
import DeckPreview from '../components/DeckPreview'

export default function DeckPreviewPage() {
  return (
    <div className="min-h-screen bg-zinc-950 text-zinc-100 p-4">
      <BackToDashboard />
      <h1 className="text-2xl font-semibold text-white mb-3">Deck Preview</h1>
      <DeckPreview />
    </div>
  )
}
