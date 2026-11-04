'use client'

// Deck Builder V2 preview — used to step through the last two V2 builds
// slide by slide (structured view of each slide's text, tables, and chart
// data), with a comment box under every slide.
//
// Stage 6, 11/2026: removed server-side. The backend used to parse every
// built deck's slide content (HOP names, dates, GC names — all Nokia-
// derived) into JSON and store it in Supabase (report_snapshots, ids
// "v2-build-latest"/"v2-build-prior") purely so this page could re-fetch
// and browse it later — a built deck is Nokia-derived content the same as
// the raw source files, and "must never be stored, no Supabase" applies to
// it the same way. Backend's /v2_slides route and save_v2_build() are
// gone, so there's no slide data left to show here.
//
// If browsing/annotating a build in the app is wanted again, the clean way
// to rebuild this is frontend-only: keep the deck Deck Builder V2 just
// built in memory (same as everything else in this migration) and parse it
// client-side for this view — nothing server-side needed.
export default function DeckPreview() {
  return (
    <div className="rounded-2xl border border-white/10 bg-white/5 p-6 text-sm text-zinc-300">
      <p className="font-semibold text-white mb-2">Slide preview is no longer available here.</p>
      <p className="text-zinc-400">
        Built decks are no longer stored on the server — open the .pptx you just downloaded from Deck Builder V2 directly to review it.
      </p>
    </div>
  )
}
