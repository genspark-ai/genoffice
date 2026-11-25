/**
 * Word's "Style area width" (Preferences, View): how wide the column naming
 * each paragraph's style is in Draft and Outline views. Zero, the default,
 * hides it. A setting of this machine, never of the document.
 */
const STYLE_AREA_KEY = 'aidocs.styleAreaWidth'
export const STYLE_AREA_MAX_TWIPS = 4320

const listeners = new Set<() => void>()

export function styleAreaTwips(): number {
  const stored = Number(localStorage.getItem(STYLE_AREA_KEY))
  return Number.isFinite(stored) ? Math.max(0, Math.min(STYLE_AREA_MAX_TWIPS, stored)) : 0
}

export function setStyleAreaTwips(twips: number): void {
  const next = Math.max(0, Math.min(STYLE_AREA_MAX_TWIPS, Math.round(twips)))
  if (next === 0) localStorage.removeItem(STYLE_AREA_KEY)
  else localStorage.setItem(STYLE_AREA_KEY, String(next))
  for (const listener of listeners) listener()
}

export function subscribeStyleArea(listener: () => void): () => void {
  listeners.add(listener)
  return () => listeners.delete(listener)
}
