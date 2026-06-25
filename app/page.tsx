'use client'

import { useEffect, useState } from 'react'

const navItems = [
  { label: "Deck Builder", href: "/deck-builder", active: true },
  { label: "GC Call View", href: "/gc-call", active: false },
  { label: "CM Call View", href: "/cm-view", active: false },
  { label: "Schedule Optimizer", href: "/schedule", active: false },
];

const stats = [
  { label: "Active Decks", value: "12", delta: "+8%" },
  { label: "Open Actions", value: "34", delta: "-4%" },
  { label: "NTPs Pending", value: "7", delta: "+14%" },
  { label: "SCOP Invoices", value: "18", delta: "+22%" },
];

function WeatherWidget() {
  const [weather, setWeather] = useState<{
    time: string
    temp: number
    condition: string
    wind: number
    windDir: number
    precip: number
    hourly: { time: string; temp: number; condition: string; precip: number; wind: number }[]
  } | null>(null)

  const getWindDir = (deg: number) => {
    const dirs = ['N','NE','E','SE','S','SW','W','NW']
    return dirs[Math.round(deg / 45) % 8]
  }

  useEffect(() => {
    const fetchWeather = async () => {
      try {
        const res = await fetch(
          'https://api.open-meteo.com/v1/forecast?latitude=40.2508&longitude=-103.7996&current=temperature_2m,weathercode,windspeed_10m,winddirection_10m,precipitation&hourly=temperature_2m,weathercode,precipitation_probability,windspeed_10m&temperature_unit=fahrenheit&timezone=America%2FDenver&forecast_days=1'
        )
        const data = await res.json()
        const now = new Date()
        const currentHour = now.getHours()

        const getCondition = (code: number) => {
          if (code === 0) return '☀️ Clear'
          if (code <= 3) return '⛅ Partly Cloudy'
          if (code <= 49) return '🌫️ Foggy'
          if (code <= 69) return '🌧️ Rain'
          if (code <= 79) return '🌨️ Snow'
          if (code <= 99) return '⛈️ Thunderstorm'
          return '🌤️ Mixed'
        }

        const hourly = data.hourly.time
          .slice(currentHour, currentHour + 6)
          .map((t: string, i: number) => ({
            time: new Date(t).toLocaleTimeString('en-US', { hour: 'numeric', hour12: true }),
            temp: Math.round(data.hourly.temperature_2m[currentHour + i]),
            condition: getCondition(data.hourly.weathercode[currentHour + i]),
            precip: data.hourly.precipitation_probability[currentHour + i] ?? 0,
            wind: Math.round(data.hourly.windspeed_10m[currentHour + i] ?? 0)
          }))

        setWeather({
          time: now.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit', hour12: true }),
          temp: Math.round(data.current.temperature_2m),
          condition: getCondition(data.current.weathercode),
          wind: Math.round(data.current.windspeed_10m ?? 0),
          windDir: data.current.winddirection_10m ?? 0,
          precip: data.current.precipitation ?? 0,
          hourly
        })
      } catch (e) {
        console.error('Weather fetch failed', e)
      }
    }
    fetchWeather()
    const interval = setInterval(fetchWeather, 300000)
    return () => clearInterval(interval)
  }, [])

  if (!weather) return <p className="text-gray-400 text-sm">Loading weather...</p>

  return (
    <div className="flex flex-col gap-2">
      <div className="flex items-center gap-4 flex-wrap">
        <span className="text-gray-300 text-sm font-semibold">{weather.time}</span>
        <span className="text-white text-lg font-bold">{weather.temp}°F</span>
        <span className="text-gray-300 text-sm">{weather.condition}</span>
        <span className="text-gray-400 text-sm">💨 {weather.wind} mph {getWindDir(weather.windDir)}</span>
        <span className="text-gray-400 text-sm">🌧️ {weather.precip} mm</span>
        <span className="text-gray-500 text-xs">Fort Morgan, CO</span>
      </div>
      <div className="flex gap-4 flex-wrap">
        {weather.hourly.map((h, i) => (
          <div key={i} className="text-center bg-gray-800 rounded-lg px-3 py-2">
            <p className="text-gray-500 text-xs">{h.time}</p>
            <p className="text-white text-sm font-semibold">{h.temp}°F</p>
            <p className="text-gray-400 text-xs">{h.condition.split(' ')[0]}</p>
            <p className="text-blue-400 text-xs">💧 {h.precip}%</p>
            <p className="text-gray-400 text-xs">💨 {h.wind} mph</p>
          </div>
        ))}
      </div>
    </div>
  )
}

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
          </aside>

          <main className="space-y-6">
            <section className="rounded-[32px] border border-white/10 bg-white/5 p-8 shadow-[0_24px_120px_-80px_rgba(0,0,0,0.55)] backdrop-blur-xl">
              <div className="flex flex-col gap-6 lg:flex-row lg:items-end lg:justify-between">
                <div className="max-w-2xl">
                  <p className="text-sm uppercase tracking-[0.4em] text-emerald-300/80">
                    Welcome back, team
                  </p>
                  <h2 className="mt-4 text-4xl font-semibold text-white">
                    Let's get to work!
                  </h2>
                  <div className="mt-4">
                    <WeatherWidget />
                  </div>
                </div>

                <img
                  src="/hylian-crest.png"
                  alt="Hylian Crest"
                  className="w-64 h-auto"
                  style={{
                    filter: 'invert(72%) sepia(98%) saturate(346%) hue-rotate(106deg) brightness(95%) contrast(88%)'
                  }}
                />
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
