/** The layout-parking opt-in (route-B prototype, #2781): parking never
 * engages by default — a user turns it on once in Word ▸ Preferences ▸
 * General, and the choice sticks across documents and restarts. The sticky
 * per-block cost gate (editor/layout-parking.ts) still decides WHETHER an
 * opted-in document actually parks; the preference only lets the gate speak
 * at all.
 *
 * Priority, highest first:
 *   1. the development override `window.__layoutParkingOff` — forces parking
 *      off even for an opted-in user (A/B measurement)
 *   2. the development override `window.__layoutParkingOn` — forces parking
 *      on without flipping the stored preference (probe and dev use)
 *   3. this preference (default off)
 */
export const LAYOUT_PARKING_KEY = 'aidocs.layoutParking'
const CHANGE_EVENT = 'aidocs:layout-parking'

/** the slice of web storage this preference needs; injectable so tests run
    without a storage backend (headless jsdom ships none) */
export interface PrefStorage {
  getItem(key: string): string | null
  setItem(key: string, value: string): void
  removeItem(key: string): void
}

const browserStorage = (): PrefStorage | undefined =>
  (globalThis as { localStorage?: PrefStorage | null }).localStorage ?? undefined

/** true only on an explicit opt-in; absent or malformed storage means off */
export function layoutParkingOptedIn(storage: PrefStorage | undefined = browserStorage()): boolean {
  try {
    return storage?.getItem(LAYOUT_PARKING_KEY) === '1'
  } catch {
    return false
  }
}

export function setLayoutParkingOptedIn(
  on: boolean,
  storage: PrefStorage | undefined = browserStorage(),
): void {
  try {
    if (on) storage?.setItem(LAYOUT_PARKING_KEY, '1')
    else storage?.removeItem(LAYOUT_PARKING_KEY)
  } catch {
    /* storage unavailable: the choice just doesn't persist */
  }
  globalThis.dispatchEvent?.(new Event(CHANGE_EVENT))
}

export function subscribeLayoutParkingPref(listener: () => void): () => void {
  window.addEventListener(CHANGE_EVENT, listener)
  window.addEventListener('storage', listener)
  return () => {
    window.removeEventListener(CHANGE_EVENT, listener)
    window.removeEventListener('storage', listener)
  }
}
