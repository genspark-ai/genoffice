import { describe, expect, it } from 'vitest'

import { ICONS_ONLY_CLASS, TIGHT_CLASS, fitRibbon } from '../src/renderer/ribbon-fit'

type Widths = readonly [normal: number, tight: number, iconsOnly: number]

function classes() {
  const set = new Set<string>()
  return {
    set,
    classList: {
      add: (c: string) => void set.add(c),
      remove: (c: string) => void set.delete(c),
    },
  }
}

/// A band whose width depends on the classes fitRibbon sets, the way the real
/// layout does: each group has a normal width, a narrower one once the band is
/// tight, and a narrower one again once it is icon-only.
function band(clientWidth: number, groups: readonly Widths[]) {
  const own = classes()
  const state = groups.map((widths) => ({ ...classes(), widths }))
  const width = (g: (typeof state)[number]) => {
    if (g.set.has(ICONS_ONLY_CLASS)) return g.widths[2]
    return own.set.has(TIGHT_CLASS) ? g.widths[1] : g.widths[0]
  }
  return {
    own,
    state,
    clientWidth,
    classList: own.classList,
    get scrollWidth() {
      return state.reduce((sum, g) => sum + width(g), 0)
    },
    querySelectorAll: () => state,
  }
}

const three: readonly Widths[] = [
  [300, 280, 200],
  [300, 280, 200],
  [300, 280, 200],
]

const iconsOnly = (b: ReturnType<typeof band>) => b.state.map((g) => g.set.has(ICONS_ONLY_CLASS))
const tight = (b: ReturnType<typeof band>) => b.own.set.has(TIGHT_CLASS)

describe('fitRibbon', () => {
  it('leaves the band untouched when it fits', () => {
    const b = band(1000, three)
    expect(fitRibbon(b)).toEqual({ tight: false, iconsOnly: 0 })
    expect(tight(b)).toBe(false)
    expect(iconsOnly(b)).toEqual([false, false, false])
  })

  it('goes tight first and stops there when that is enough', () => {
    // 900 normal in 850: the tight band is 840
    const b = band(850, three)
    expect(fitRibbon(b)).toEqual({ tight: true, iconsOnly: 0 })
    expect(iconsOnly(b)).toEqual([false, false, false])
  })

  it('then squeezes from the rightmost group and stops once the band fits', () => {
    // tight 840 in 800: dropping the last group to 200 gives 760
    const b = band(800, three)
    expect(fitRibbon(b)).toEqual({ tight: true, iconsOnly: 1 })
    expect(iconsOnly(b)).toEqual([false, false, true])
  })

  it('keeps going left while it still overflows', () => {
    const b = band(700, three)
    expect(fitRibbon(b)).toEqual({ tight: true, iconsOnly: 2 })
    expect(iconsOnly(b)).toEqual([false, true, true])
  })

  it('gives the metrics and labels back when the window grows', () => {
    const b = band(700, three)
    fitRibbon(b)
    b.clientWidth = 1000
    expect(fitRibbon(b)).toEqual({ tight: false, iconsOnly: 0 })
    expect(tight(b)).toBe(false)
    expect(iconsOnly(b)).toEqual([false, false, false])
  })

  it('stops at the first group when even all icons do not fit, leaving the scrollbar', () => {
    const b = band(500, three)
    expect(fitRibbon(b)).toEqual({ tight: true, iconsOnly: 3 })
    expect(b.scrollWidth).toBe(600)
  })

  it('tolerates the 1px a fractional zoom rounds up', () => {
    const b = band(899, three)
    expect(fitRibbon(b)).toEqual({ tight: false, iconsOnly: 0 })
  })
})
