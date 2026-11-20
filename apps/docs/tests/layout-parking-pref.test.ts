/**
 * The layout-parking opt-in (route-B prototype, #2781): parking must never
 * engage by default, an explicit opt-in sticks, and malformed storage reads
 * as off — the wake-up tax makes parking a net loss on documents that never
 * asked for it.
 */
import { describe, expect, it, vi } from 'vitest'
import {
  LAYOUT_PARKING_KEY,
  layoutParkingOptedIn,
  setLayoutParkingOptedIn,
  subscribeLayoutParkingPref,
  type PrefStorage,
} from '../src/renderer/layout-parking-pref'

const storageStub = (initial: Record<string, string> = {}): PrefStorage => {
  const map = new Map(Object.entries(initial))
  return {
    getItem: (k) => map.get(k) ?? null,
    setItem: (k, v) => void map.set(k, v),
    removeItem: (k) => void map.delete(k),
  }
}

describe('the layout-parking opt-in', () => {
  it('is off until the user opts in', () => {
    expect(layoutParkingOptedIn(storageStub())).toBe(false)
    expect(layoutParkingOptedIn(undefined)).toBe(false)
  })

  it('round-trips an explicit opt-in and opt-out', () => {
    const storage = storageStub()
    setLayoutParkingOptedIn(true, storage)
    expect(layoutParkingOptedIn(storage)).toBe(true)
    expect(storage.getItem(LAYOUT_PARKING_KEY)).toBe('1')
    setLayoutParkingOptedIn(false, storage)
    expect(layoutParkingOptedIn(storage)).toBe(false)
    expect(storage.getItem(LAYOUT_PARKING_KEY)).toBeNull()
  })

  it('reads values it did not write as off', () => {
    expect(layoutParkingOptedIn(storageStub({ [LAYOUT_PARKING_KEY]: 'yes' }))).toBe(false)
    expect(layoutParkingOptedIn(storageStub({ [LAYOUT_PARKING_KEY]: '0' }))).toBe(false)
  })

  it('a throwing storage reads as off and never propagates the write failure', () => {
    const hostile: PrefStorage = {
      getItem: () => {
        throw new DOMException('denied')
      },
      setItem: () => {
        throw new DOMException('denied')
      },
      removeItem: () => {
        throw new DOMException('denied')
      },
    }
    expect(layoutParkingOptedIn(hostile)).toBe(false)
    expect(() => setLayoutParkingOptedIn(true, hostile)).not.toThrow()
  })

  it('notifies subscribers on a preference flip and stops after unsubscribe', () => {
    const listener = vi.fn()
    const off = subscribeLayoutParkingPref(listener)
    setLayoutParkingOptedIn(true, storageStub())
    setLayoutParkingOptedIn(false, storageStub())
    off()
    setLayoutParkingOptedIn(true, storageStub())
    expect(listener).toHaveBeenCalledTimes(2)
  })

  it('re-reads storage changed from another tab through the storage event', () => {
    const listener = vi.fn()
    const off = subscribeLayoutParkingPref(listener)
    window.dispatchEvent(new StorageEvent('storage'))
    expect(listener).toHaveBeenCalledTimes(1)
    off()
  })
})
