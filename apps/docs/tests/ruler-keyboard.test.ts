import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { Editor } from '@tiptap/core'
import type { SectionSettings, TabStop } from '@genoffice/docx-engine'
import {
  MAX_RULER_INCHES,
  Ruler,
  rulerDims,
  snapTabTwips,
  stableStopIds,
} from '../src/renderer/components/Ruler'
import { getLang, setModuleLang } from '../src/renderer/i18n/locale'
import { setMeasurementUnit } from '../src/renderer/units'

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

const originalLang = getLang()
afterEach(() => {
  setModuleLang(originalLang)
  setMeasurementUnit(null)
})

const section = (over: Partial<SectionSettings> = {}): SectionSettings => ({
  pageWidth: 12240,
  pageHeight: 15840,
  orientation: 'portrait',
  marginTop: 1440,
  marginRight: 1440,
  marginBottom: 1440,
  marginLeft: 1440,
  pageBorder: false,
  columns: 1,
  ...over,
})

function mount(
  stops: TabStop[],
  sectionOver: Partial<SectionSettings> = {},
): { container: HTMLElement; onTabStopsChange: ReturnType<typeof vi.fn>; cleanup: () => void } {
  const onTabStopsChange = vi.fn()
  const editor = {
    isActive: () => true,
    getAttributes: () => ({ tabStops: JSON.stringify(stops) }),
  } as unknown as Editor
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root: Root = createRoot(container)
  act(() => {
    root.render(createElement(Ruler, { section: section(sectionOver), editor, onTabStopsChange }))
  })
  return {
    container,
    onTabStopsChange,
    cleanup: () => {
      act(() => root.unmount())
      container.remove()
    },
  }
}

function keydown(el: Element, key: string, shift = false): void {
  act(() => {
    el.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, shiftKey: shift }))
  })
}

/**
 * A ruler whose paragraph attributes follow what it emits, the way the editor
 * does: every nudge writes back and the next render reads the new list.
 */
function mountLive(stops: TabStop[]): {
  container: HTMLElement
  render: () => void
  emitted: () => TabStop[] | null
  marker: (i: number) => HTMLElement | null
  cleanup: () => void
} {
  let current = stops
  let last: TabStop[] | null = null
  const onTabStopsChange = vi.fn((next: TabStop[] | null) => {
    last = next
    if (next) current = next
  })
  const editor = {
    isActive: () => true,
    getAttributes: () => ({ tabStops: JSON.stringify(current) }),
  } as unknown as Editor
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root: Root = createRoot(container)
  const render = () => {
    act(() => {
      root.render(createElement(Ruler, { section: section(), editor, onTabStopsChange }))
    })
  }
  render()
  return {
    container,
    render,
    emitted: () => last,
    marker: (i) => container.querySelector(`[data-ruler-stop="${i}"]`),
    cleanup: () => {
      act(() => root.unmount())
      container.remove()
    },
  }
}

describe('rulerDims', () => {
  it('passes sane geometry through', () => {
    expect(rulerDims(section())).toEqual({
      pageWidth: 12240,
      marginLeft: 1440,
      marginRight: 1440,
      inches: 8,
    })
  })

  it('bounds corrupt geometry instead of hanging', () => {
    const start = Date.now()
    const dims = rulerDims(section({ pageWidth: 1e12, marginLeft: NaN }))
    expect(Date.now() - start).toBeLessThan(5000)
    expect(dims.inches).toBeLessThanOrEqual(MAX_RULER_INCHES)
    expect(dims.pageWidth).toBeLessThanOrEqual(MAX_RULER_INCHES * 1440)
    const nan = rulerDims(section({ pageWidth: NaN }))
    expect(nan.pageWidth).toBe(12240)
  })

  it('snaps to the Word grid and maps non-finite to zero', () => {
    expect(snapTabTwips(100)).toBe(120)
    expect(snapTabTwips(30)).toBe(60)
    expect(snapTabTwips(Infinity)).toBe(0)
  })
})

describe('Ruler keyboard operation', () => {
  it('exposes stops as sliders with accessible names', () => {
    setModuleLang('en')
    setMeasurementUnit('cm')
    const { container, cleanup } = mount([{ pos: 1440, val: 'left' }])
    try {
      const stop = container.querySelector('[role="slider"]') as HTMLElement
      expect(stop).not.toBeNull()
      expect(stop.getAttribute('tabindex')).toBe('0')
      expect(stop.getAttribute('aria-label')).toContain('Left')
      expect(stop.getAttribute('aria-valuenow')).toBe('1440')
      expect(stop.getAttribute('aria-label')).toContain('@ 2.54')
      expect(container.querySelector('[role="group"]')).not.toBeNull()
    } finally {
      cleanup()
    }
  })

  it('arrow keys nudge and Delete removes the focused stop', () => {
    const { container, onTabStopsChange, cleanup } = mount([{ pos: 1440, val: 'left' }])
    try {
      const stop = container.querySelector('[role="slider"]') as HTMLElement
      keydown(stop, 'ArrowRight')
      expect(onTabStopsChange).toHaveBeenCalledTimes(1)
      expect(onTabStopsChange.mock.calls[0]![0]).toEqual([{ pos: 1500, val: 'left' }])
      keydown(stop, 'Delete')
      expect(onTabStopsChange).toHaveBeenCalledTimes(2)
      expect(onTabStopsChange.mock.calls[1]![0]).toBeNull()
    } finally {
      cleanup()
    }
  })

  it('renders bounded ticks for hostile section geometry', () => {
    const { container, cleanup } = mount([], {
      pageWidth: 1e12,
      marginLeft: NaN,
    } as unknown as Partial<SectionSettings>)
    try {
      const ticks = container.querySelectorAll('.ruler-num')
      expect(ticks.length).toBeLessThanOrEqual(MAX_RULER_INCHES)
      const ruler = container.querySelector('.ruler') as HTMLElement
      expect(ruler.style.width).toMatch(/^\d+(\.\d+)?px$/)
      expect(parseFloat(ruler.style.width)).toBeLessThanOrEqual(MAX_RULER_INCHES * 96)
      expect(ruler.getAttribute('aria-label')).toBeTruthy()
    } finally {
      cleanup()
    }
  })
})

describe('Ruler tab stop nudge across a neighbour', () => {
  it('keeps nudging the same marker after the stop list re-sorts', () => {
    // The first stop is nudged right past the second one, so the emitted list
    // re-sorts and the marker the user is holding the arrow on moves to another
    // index. The next arrow must still move that marker.
    const { render, emitted, marker, cleanup } = mountLive([
      { pos: 1440, val: 'left' },
      { pos: 1500, val: 'left' },
    ])
    try {
      const first = marker(0)!
      first.focus()
      expect(document.activeElement).toBe(first)
      // 1440 + 720 crosses the stop at 1500, so the list re-sorts
      keydown(first, 'ArrowRight', true)
      expect(emitted()).toEqual([
        { pos: 1500, val: 'left' },
        { pos: 2160, val: 'left' },
      ])
      render()
      // the focused marker followed its stop to the new index
      const moved = marker(1)!
      expect(moved.getAttribute('aria-valuenow')).toBe('2160')
      expect(document.activeElement).toBe(moved)

      // the queued arrow on the still-focused marker moves the same stop
      keydown(document.activeElement as Element, 'ArrowRight')
      expect(emitted()).toEqual([
        { pos: 1500, val: 'left' },
        { pos: 2220, val: 'left' },
      ])
    } finally {
      cleanup()
    }
  })

  it('gives a stop its own id when the list is replaced', () => {
    const prev = [
      { pos: 1440, val: 'left' },
      { pos: 1500, val: 'left' },
    ] as TabStop[]
    const ids = stableStopIds([], [], prev)
    expect(new Set(ids).size).toBe(2)
    // a stop that kept its position keeps its id
    expect(stableStopIds(prev, ids, [prev[1]!, prev[0]!])).toEqual([ids[1], ids[0]])
    // the nudged stop keeps its id even though it moved to another index
    const moved = stableStopIds(prev, ids, [prev[1]!, { pos: 2160, val: 'left' }])
    expect(moved[1]).toBe(ids[0])
    expect(new Set(moved).size).toBe(2)

    // a stop added after a delete must not land on a key that is still live:
    // A,B,C mint s0..s2, dropping B takes s1 out of circulation, and the new
    // stop D has to be given an id that no marker still holds
    const three = [
      { pos: 1440, val: 'left' },
      { pos: 1500, val: 'left' },
      { pos: 2160, val: 'left' },
    ] as TabStop[]
    const threeIds = stableStopIds([], [], three)
    expect(threeIds).toEqual(['s0', 's1', 's2'])
    const survivors = [three[0]!, three[2]!] as TabStop[]
    const withoutB = stableStopIds(three, threeIds, survivors)
    expect(withoutB).toEqual(['s0', 's2'])
    const appended = stableStopIds(survivors, withoutB, [
      survivors[0]!,
      survivors[1]!,
      { pos: 2880, val: 'left' },
    ])
    // the survivors keep their own ids, so D is the one that has to be distinct
    expect(appended[0]).toBe('s0')
    expect(appended[1]).toBe('s2')
    expect(new Set(appended).size).toBe(3)
    // a ref still holding a duplicate from the old scheme heals on the next
    // render instead of passing the colliding pair on
    const all = [survivors[0]!, survivors[1]!, { pos: 2880, val: 'left' }] as TabStop[]
    const healed = stableStopIds(all, ['s0', 's2', 's2'], all)
    expect(new Set(healed).size).toBe(3)
  })
})
