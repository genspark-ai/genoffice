/** Persisted "View ▸ Dark Mode" choice for the Word-style dark page
 * (editor/dark-page.ts). Absent = follow the UI theme; once the user flips it
 * explicitly the choice sticks across documents, windows and restarts, and a
 * theme change no longer drops it — an explicit "off" cannot come back on its
 * own. */
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
