'use client'

// Shared left-side nav for the "Customer Slide Deck" group of pages — the
// dashboard now has one entry point (Deck Builder) instead of three; this is
// how you get to the other two (Deck Preview, the legacy builder) without
// going back to the dashboard first.
const LINKS = [
  { key: 'builder', label: 'Deck Builder', href: '/deck-builder-v2', emoji: '🃏' },
  { key: 'preview', label: 'Deck Preview', href: '/deck-preview', emoji: '🔍' },
  { key: 'legacy', label: 'Deck Builder (Legacy)', href: '/deck-builder', emoji: '🗄️' },
] as const

export default function CustomerSlideDeckNav({ active }: { active: 'builder' | 'preview' | 'legacy' }) {
  return (
    <nav className="w-56 shrink-0 space-y-2">
      <p className="px-2 mb-1 text-xs uppercase tracking-[0.3em] text-zinc-500">Customer Slide Deck</p>
      {LINKS.map(l => (
        <a
          key={l.key}
          href={l.href}
          className={`flex items-center gap-2 rounded-2xl px-4 py-3 text-sm font-medium transition ${
            active === l.key ? 'bg-emerald-500 text-zinc-950' : 'bg-white/5 text-zinc-300 hover:bg-white/10'
          }`}
        >
          <span>{l.emoji}</span>
          <span>{l.label}</span>
        </a>
      ))}
    </nav>
  )
}
