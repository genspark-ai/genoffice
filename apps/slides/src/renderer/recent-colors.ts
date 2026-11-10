/**
 * "Recently used colors" for the ribbon fill/color pickers: a small
 * most-recent-first #RRGGBB list persisted in localStorage (survives reload,
 * per-app). Reads are on-demand (call on menu open) — no subscription needed.
 */

const KEY = 'slides:recent-colors'
const MAX = 10

const normalize = (hex: string): string | null => {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex.trim())
  return m ? `#${m[1]!.toUpperCase()}` : null
}

export function getRecentColors(): string[] {
  try {
    const raw = JSON.parse(localStorage.getItem(KEY) ?? '[]')
    if (!Array.isArray(raw)) return []
    // Enforce the push invariant on read too: legacy/corrupt stores may hold
    // duplicates or more than MAX entries, which would overflow the swatch row
    // until the next push self-heals it.
    const colors = raw.map((h) => normalize(String(h))).filter((h): h is string => h != null)
    return [...new Set(colors)].slice(0, MAX)
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

/**
 * History writer for a picker whose native color input previews live: the OS
 * panel fires a change event per drag movement, and pushing each one floods
 * all MAX slots with the drag trail of a single pick (gendesk #279). Previewed
 * values only stage here; exactly one — the last — enters the history, when
 * the menu closes. A swatch click is a single deliberate pick: it records
 * immediately and drops any staged preview it supersedes.
 */
export interface RecentColorSession {
  /** Live preview from the native input: stage, don't record. */
  preview(hex: string): void
  /** A single-click pick (swatch): record now, discard the staged preview. */
  commit(hex: string): void
  /** Menu closing by any path: record the staged preview, if one is left. */
  close(): void
}

export function createRecentColorSession(): RecentColorSession {
  let pending: string | null = null
  return {
    preview(hex) {
      pending = hex
    },
    commit(hex) {
      pending = null
      pushRecentColor(hex)
    },
    close() {
      if (pending !== null) pushRecentColor(pending)
      pending = null
    },
  }
}
