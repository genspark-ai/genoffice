/**
 * "Recently used colors" for the ribbon fill/color pickers: a small
 * most-recent-first #RRGGBB list persisted in localStorage (survives reload,
 * per-app). Reads are on-demand (call on menu open) — no subscription needed.
 */

const KEY = 'slides:recent-colors'
const MAX = 10

const normalize = (hex: string): string | null => {
  const m = /^#?([0-9a-f]{6})/i.exec(hex.trim())
  return m ? `#${m[1]!.toUpperCase()}` : null
}

export function getRecentColors(): string[] {
  try {
    const raw = JSON.parse(localStorage.getItem(KEY) ?? '[]')
    if (!Array.isArray(raw)) return []
    return raw.map((h) => normalize(String(h))).filter((h): h is string => h != null)
  } catch {
    return []
  }
}

export function pushRecentColor(hex: string): void {
  const norm = normalize(hex)
  if (!norm) return
  const next = [norm, ...getRecentColors().filter((h) => h !== norm)].slice(0, MAX)
  try {
    localStorage.setItem(KEY, JSON.stringify(next))
  } catch {
    // storage full/unavailable: recents are a convenience, drop silently
  }
}
