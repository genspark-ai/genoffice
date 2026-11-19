import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, it } from 'vitest'

import { useStoredFlag } from '../src/renderer/stored-flag'

const KEY = 'aidocs.showRuler'

function mount() {
  const seen: { on: boolean; set: (on: boolean | ((prev: boolean) => boolean)) => void } = {
    on: false,
    set: () => {},
  }
  function Probe() {
    const [on, set] = useStoredFlag(KEY)
    seen.on = on
    seen.set = set
    return null
  }
  const container = document.createElement('div')
  const root = createRoot(container)
  act(() => root.render(createElement(Probe)))
  return { seen, unmount: () => act(() => root.unmount()) }
}

afterEach(() => localStorage.clear())

describe('View ▸ Ruler survives closing the window', () => {
  it('starts off, and a window opened after turning it on starts on', () => {
    const first = mount()
    expect(first.seen.on).toBe(false)
    act(() => first.seen.set(true))
    expect(first.seen.on).toBe(true)
    first.unmount()
    const reopened = mount()
    expect(reopened.seen.on).toBe(true)
    reopened.unmount()
  })

  it('toggles from the stored value, twice in one turn included', () => {
    const win = mount()
    act(() => win.seen.set((v) => !v))
    expect(win.seen.on).toBe(true)
    expect(localStorage.getItem(KEY)).toBe('1')
    act(() => {
      win.seen.set((v) => !v)
      win.seen.set((v) => !v)
    })
    expect(win.seen.on).toBe(true)
    win.unmount()
  })

  it('follows a change made in another Docs window', () => {
    const win = mount()
    act(() => {
      localStorage.setItem(KEY, '1')
      window.dispatchEvent(new StorageEvent('storage', { key: KEY }))
    })
    expect(win.seen.on).toBe(true)
    win.unmount()
  })
})
