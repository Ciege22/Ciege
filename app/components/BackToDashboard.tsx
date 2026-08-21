'use client'

import { useRouter } from 'next/navigation'

export default function BackToDashboard() {
  const router = useRouter()
  return (
    <button
      onClick={() => router.push('/')}
      type="button"
      className="mb-3 inline-flex items-center gap-1 text-xs text-gray-500 hover:text-gray-300 transition-colors"
    >
      ← Dashboard
    </button>
  )
}
