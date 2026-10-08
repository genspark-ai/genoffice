/** Persisted "View ▸ Dark Mode" choice for the Word-style dark page
 * (editor/dark-page.ts). Only consulted in the dark UI theme: absent = on, and
 * an explicit "off" sticks across documents, windows and restarts. */
export const DARK_PAGE_KEY = 'aidocs.darkPage'

/** The user's explicit choice, or null when only the theme has spoken so far. */
export function readDarkPagePref(): boolean | null {
  const raw = globalThis.localStorage?.getItem(DARK_PAGE_KEY)
  if (raw === '1') return true
  if (raw === '0') return false
  return null
}

export function writeDarkPagePref(value: boolean): void {
  globalThis.localStorage?.setItem(DARK_PAGE_KEY, value ? '1' : '0')
}

/** the page state for a UI theme: a light theme never shows a dark page, whatever was stored */
export function resolveDarkPage(themeDark: boolean): boolean {
  return themeDark ? (readDarkPagePref() ?? true) : false
}
