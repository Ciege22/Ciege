'use client'

export const dynamic = 'force-dynamic'

import dynamicImport from 'next/dynamic'
import BackToDashboard from '../components/BackToDashboard'

// Leaflet touches `window` at import time, so the actual map must never run
// during Next.js's build-time render pass — ssr:false guarantees that even
// though this page (like every page here) is otherwise a plain client
// component. The real implementation lives in MapView.tsx.
const MapView = dynamicImport(() => import('./MapView'), {
  ssr: false,
  loading: () => (
    <div className="flex items-center justify-center h-[70vh] text-zinc-400 text-sm">
      Loading map…
    </div>
  ),
})

export default function MapPage() {
  return (
    <div className="min-h-screen bg-zinc-950 text-zinc-100 p-4">
      <BackToDashboard />
      <h1 className="text-2xl font-semibold text-white mb-1">HOP Network Map</h1>
      <p className="text-sm text-zinc-400 mb-4">
        Every HOP drawn as a line between its two sites, from the latest tracker upload.
      </p>
      <MapView />
    </div>
  )
}
