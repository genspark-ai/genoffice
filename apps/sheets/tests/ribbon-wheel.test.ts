import { describe, expect, it, vi } from 'vitest'
import { scrollRibbonOnWheel } from '../src/renderer/ribbon-wheel'

const ribbon = (scrollWidth = 2000) => ({ scrollWidth, clientWidth: 900, scrollLeft: 0 })

function wheel(
  overRibbon: ReturnType<typeof ribbon> | null,
  init: Partial<WheelEvent> = {},
  list?: { scrollHeight: number; clientHeight: number },
) {
  const target = {
    scrollHeight: list?.scrollHeight ?? 10,
    clientHeight: list?.clientHeight ?? 10,
    parentElement: overRibbon,
    closest: () => overRibbon,
  }
  return {
    deltaX: 0,
    deltaY: 100,
    ctrlKey: false,
    metaKey: false,
    target,
    preventDefault: vi.fn(),
    ...init,
  } as unknown as WheelEvent
}

describe('scrollRibbonOnWheel', () => {
  it('turns a vertical wheel over the ribbon into horizontal scrolling', () => {
    const el = ribbon()
    const event = wheel(el)
    scrollRibbonOnWheel(event)
    expect(el.scrollLeft).toBe(100)
    expect(event.preventDefault).toHaveBeenCalled()
  })

  it('leaves the wheel alone off the ribbon, or when nothing overflows', () => {
    const fits = ribbon(900)
    for (const event of [wheel(null), wheel(fits)]) {
      scrollRibbonOnWheel(event)
      expect(event.preventDefault).not.toHaveBeenCalled()
    }
    expect(fits.scrollLeft).toBe(0)
  })

  it('leaves zooming and sideways trackpad scrolling alone', () => {
    const el = ribbon()
    for (const init of [{ ctrlKey: true }, { metaKey: true }, { deltaX: 200 }]) {
      scrollRibbonOnWheel(wheel(el, init))
    }
    expect(el.scrollLeft).toBe(0)
  })

  it('leaves the wheel to a list that scrolls itself', () => {
    vi.stubGlobal('getComputedStyle', () => ({ overflowY: 'auto' }))
    const el = ribbon()
    const event = wheel(el, {}, { scrollHeight: 400, clientHeight: 200 })
    scrollRibbonOnWheel(event)
    expect(el.scrollLeft).toBe(0)
    expect(event.preventDefault).not.toHaveBeenCalled()
    vi.unstubAllGlobals()
  })
})
