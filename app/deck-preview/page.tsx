'use client'

export const dynamic = 'force-dynamic'

import BackToDashboard from '../components/BackToDashboard'
import DeckPreview from '../components/DeckPreview'
import CustomerSlideDeckNav from '../components/CustomerSlideDeckNav'

export default function DeckPreviewPage() {
  return (
    <div className="min-h-screen bg-zinc-950 text-zinc-100 p-4">
      <div className="flex gap-6">
        <CustomerSlideDeckNav active="preview" />
        <div className="flex-1 min-w-0">
          <BackToDashboard />
          <h1 className="text-2xl font-semibold text-white mb-3">Deck Preview</h1>
          <DeckPreview />
        </div>
      </div>
    </div>
  )
}
