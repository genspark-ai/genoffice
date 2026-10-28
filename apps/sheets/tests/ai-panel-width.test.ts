import { describe, expect, it } from 'vitest'
import { loadPanelWidth, panelDisplayWidth, PANEL_WIDTH_KEY } from '../src/renderer/ai/panel-width'

function storageWith(raw: string | null): Storage {
  const store = new Map<string, string>()
  if (raw !== null) store.set(PANEL_WIDTH_KEY, raw)
  return {
    get length(): number {
      return store.size
    },
    getItem: (key: string) => (store.has(key) ? store.get(key)! : null),
    setItem: (key: string, value: string) => void store.set(key, value),
    removeItem: (key: string) => void store.delete(key),
    key: (index: number) => Array.from(store.keys())[index] ?? null,
    clear: () => store.clear(),
  } as Storage
}

describe('loadPanelWidth', () => {
  it('reads the persisted dragged width', () => {
    expect(loadPanelWidth(storageWith('480'))).toBe(480)
  })

  it('returns null without a saved width', () => {
    expect(loadPanelWidth(storageWith(null))).toBeNull()
  })

  it('rejects missing, non-numeric, zero and negative values', () => {
    expect(loadPanelWidth(storageWith('wide'))).toBeNull()
    expect(loadPanelWidth(storageWith('0'))).toBeNull()
    expect(loadPanelWidth(storageWith('-40'))).toBeNull()
    expect(loadPanelWidth(storageWith('NaN'))).toBeNull()
  })

  it('keeps a legacy below-floor preference raw — migration happens at display time', () => {
    // the restored preference must not be baked smaller than the live window
    expect(loadPanelWidth(storageWith('300'))).toBe(300)
  })
})

describe('panelDisplayWidth', () => {
  it('tracks the window without a dragged preference: ~22%, floored at 360, capped at 720', () => {
    expect(panelDisplayWidth(null, 1920)).toBe(422)
    expect(panelDisplayWidth(null, 1000)).toBe(360)
    expect(panelDisplayWidth(null, 4000)).toBe(720)
  })

  it('clamps a dragged preference into [360, cap of the live window]', () => {
    expect(panelDisplayWidth(500, 1920)).toBe(500)
    expect(panelDisplayWidth(500, 800)).toBe(480)
    expect(panelDisplayWidth(900, 1920)).toBe(720)
  })

  it('migrates legacy below-floor widths up to the 360 floor', () => {
    expect(panelDisplayWidth(300, 1920)).toBe(360)
    expect(panelDisplayWidth(280, 700)).toBe(360)
  })

  it('never shrinks below 360 on a transiently tiny viewport', () => {
    // a WebContentsView is 0×0 until the shell lays it out
    expect(panelDisplayWidth(500, 100)).toBe(360)
    expect(panelDisplayWidth(null, 100)).toBe(360)
  })
})
