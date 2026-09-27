import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { Editor } from '@tiptap/core'
import type { SectionSettings, TabStop } from '@genoffice/docx-engine'
import { MAX_RULER_INCHES, Ruler, rulerDims, snapTabTwips } from '../src/renderer/components/Ruler'
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
