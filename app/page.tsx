const navItems = [
  { label: "Deck Builder", href: "/deck-builder", active: true },
  { label: "GC Call View", href: "/gc-call", active: false },
  { label: "CM Call View", href: "/cm-view", active: false },
];

const stats = [
  { label: "Active Decks", value: "12", delta: "+8%" },
  { label: "Open Actions", value: "34", delta: "-4%" },
  { label: "NTPs Pending", value: "7", delta: "+14%" },
  { label: "SCOP Invoices", value: "18", delta: "+22%" },
];

export default function Home() {
  return (
    <div className="min-h-screen bg-zinc-950 text-zinc-100">
      <div className="mx-auto max-w-[1600px] px-6 py-8">
        <div className="grid gap-6 lg:grid-cols-[280px_minmax(0,1fr)]">
          <aside className="rounded-[32px] border border-white/10 bg-white/5 p-6 shadow-[0_24px_120px_-80px_rgba(0,0,0,0.55)] backdrop-blur-xl">
            <div className="mb-8">
              <p className="text-xs uppercase tracking-[0.4em] text-emerald-300/80">Ciege</p>
              <h1 className="mt-3 text-3xl font-semibold text-white">PM Dashboard</h1>
              <p className="mt-2 text-sm leading-6 text-zinc-400">
                Plan Execute Conquer
              </p>
            </div>

            <nav className="space-y-2">
              {navItems.map((item) =>
                item.href ? (
                  <a
                    key={item.label}
                    href={item.href}
                    className={`block w-full rounded-2xl px-4 py-3 text-left text-sm font-medium transition hover:bg-white/10 ${
                      item.active ? "bg-emerald-400/10 text-emerald-300" : "text-zinc-300"
                    }`}
                  >
                    {item.label}
                  </a>
                ) : (
                  <button
                    key={item.label}
                    className={`w-full rounded-2xl px-4 py-3 text-left text-sm font-medium transition hover:bg-white/10 ${
                      item.active ? "bg-emerald-400/10 text-emerald-300" : "text-zinc-300"
                    }`}
                    type="button"
                  >
                    {item.label}
                  </button>
                )
              )}
            </nav>

            <div className="mt-10 rounded-3xl border border-white/10 bg-zinc-900/70 p-4 text-sm text-zinc-400">
              <p className="font-medium text-white">Platform insight</p>
              <p className="mt-3 leading-6">
                Keep your project teams aligned with a single source of truth for decks, actions, NTPs, and invoices.
              </p>
            </div>
          </aside>

          <main className="space-y-6">
            <section className="rounded-[32px] border border-white/10 bg-white/5 p-8 shadow-[0_24px_120px_-80px_rgba(0,0,0,0.55)] backdrop-blur-xl">
              <div className="flex flex-col gap-6 lg:flex-row lg:items-end lg:justify-between">
                <div className="max-w-2xl">
                  <p className="text-sm uppercase tracking-[0.4em] text-emerald-300/80">
                    Welcome back, team
                  </p>
                  <h2 className="mt-4 text-4xl font-semibold text-white">
                    Run your product operations with confidence.
                  </h2>
                  <p className="mt-4 text-base leading-7 text-zinc-400">
                    Ciege brings your Deck Builder, Action Board, NTP Tracker and SCOP Invoice flow together in one polished workspace.
                  </p>
                </div>

                <div className="rounded-3xl bg-zinc-900/90 px-5 py-4 text-sm text-zinc-300 ring-1 ring-white/10">
                  <p className="text-zinc-400">Next sync</p>
                  <p className="mt-1 text-xl font-semibold text-white">Thursday · 10:00 AM</p>
                </div>
              </div>
            </section>

            <section className="grid gap-4 sm:grid-cols-2 xl:grid-cols-4">
              {stats.map((stat) => (
                <article
                  key={stat.label}
                  className="rounded-[28px] border border-white/10 bg-zinc-900/80 p-6 shadow-xl shadow-black/20"
                >
                  <p className="text-sm font-medium uppercase tracking-[0.35em] text-zinc-500">
                    {stat.label}
                  </p>
                  <p className="mt-4 text-4xl font-semibold text-white">{stat.value}</p>
                  <p className="mt-3 text-sm text-emerald-300">{stat.delta} vs last week</p>
                </article>
              ))}
            </section>
          </main>
        </div>
      </div>
    </div>
  );
}
