import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { Editor } from '@tiptap/core'
import type { SectionSettings } from '@genoffice/docx-engine'
import {
  Ruler,
  indentFromRuler,
  rulerDims,
  rulerIndentPositions,
  rulerIndents,
} from '../src/renderer/components/Ruler'
import { getLang, setModuleLang } from '../src/renderer/i18n/locale'

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

const originalLang = getLang()
afterEach(() => {
  setModuleLang(originalLang)
})

const section: SectionSettings = {
  pageWidth: 12240,
  pageHeight: 15840,
  orientation: 'portrait',
  marginTop: 1440,
  marginRight: 1440,
  marginBottom: 1440,
  marginLeft: 1440,
  pageBorder: false,
  columns: 1,
}

const dims = rulerDims(section)
/** pageWidth 12240 twips at 96dpi */
const WIDTH_PX = 816
const MARGIN_PX = 96

function mount(attrs: Record<string, unknown> = {}) {
  const onTabStopsChange = vi.fn()
  const onIndentsChange = vi.fn()
  const editor = {
    isActive: () => true,
    getAttributes: () => attrs,
  } as unknown as Editor
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root: Root = createRoot(container)
  act(() => {
    root.render(createElement(Ruler, { section, editor, onTabStopsChange, onIndentsChange }))
  })
  // jsdom lays nothing out; the drag reads the ruler's rect once at mousedown
  const ruler = container.querySelector('.ruler') as HTMLElement
  ruler.getBoundingClientRect = () =>
    ({
      left: 0,
      top: 0,
      width: WIDTH_PX,
      height: 20,
      right: WIDTH_PX,
      bottom: 20,
      x: 0,
      y: 0,
      toJSON: () => ({}),
    }) as DOMRect
  return {
    container,
    onTabStopsChange,
    onIndentsChange,
    cleanup: () => {
      act(() => root.unmount())
      container.remove()
    },
  }
}

const marker = (container: HTMLElement, kind: string): HTMLElement =>
  container.querySelector(`[data-ruler-indent="${kind}"]`) as HTMLElement

function mouse(el: Element, type: string, clientX: number): void {
  act(() => {
    el.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, clientX }))
  })
}

function drag(el: Element, toX: number): void {
  mouse(el, 'mousedown', 0)
  act(() => {
    document.dispatchEvent(new MouseEvent('mousemove', { bubbles: true, clientX: toX }))
    document.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, clientX: toX }))
  })
}

describe('rulerIndents', () => {
  it('reads an inherited (null/missing) indent as zero', () => {
    expect(rulerIndents({})).toEqual({ left: 0, right: 0, firstLine: 0 })
    expect(rulerIndents({ indentLeft: null, indentFirstLine: null })).toEqual({
      left: 0,
      right: 0,
      firstLine: 0,
    })
  })

  it('keeps a negative first-line indent: that is a hanging indent', () => {
    expect(rulerIndents({ indentLeft: 720, indentFirstLine: -360 })).toEqual({
      left: 720,
      right: 0,
      firstLine: -360,
    })
  })

  it('ignores non-finite values', () => {
    expect(rulerIndents({ indentLeft: NaN, indentRight: '720' })).toEqual({
      left: 0,
      right: 0,
      firstLine: 0,
    })
  })
})

describe('rulerIndentPositions', () => {
  it('places the marks inside the content area', () => {
    expect(rulerIndentPositions(dims, { left: 0, right: 0, firstLine: 0 })).toEqual({
      left: 1440,
      first: 1440,
      right: 10800,
    })
  })

  it('measures the right mark back from the right margin', () => {
    expect(rulerIndentPositions(dims, { left: 0, right: 720, firstLine: 0 }).right).toBe(10080)
  })

  it('puts a hanging first line to the left of the left-indent block', () => {
    const pos = rulerIndentPositions(dims, { left: 720, right: 0, firstLine: -360 })
    expect(pos.first).toBeLessThan(pos.left)
    expect(pos.first).toBe(1800)
  })
})

describe('indentFromRuler', () => {
  it('drops the left marker at the ruler position, minus the margin', () => {
    // x = 300px of 816px over 12240 twips
    const posTwips = (300 / WIDTH_PX) * 12240
    expect(indentFromRuler('left', dims, { left: 0, right: 0, firstLine: 0 }, posTwips)).toEqual({
      left: 3060,
      right: null,
      firstLine: null,
    })
  })

  it('carries the first-line offset when the left block moves', () => {
    const next = indentFromRuler('left', dims, { left: 720, right: 0, firstLine: 360 }, 4320)
    expect(next.left).toBe(2880)
    expect(next.firstLine).toBe(360)
  })

  it('writes the right indent as the distance back from the right margin', () => {
    const posTwips = (600 / WIDTH_PX) * 12240
    expect(
      indentFromRuler('right', dims, { left: 0, right: 0, firstLine: 0 }, posTwips).right,
    ).toBe(1800)
  })

  it('keeps a hanging indent negative and never past the left indent', () => {
    // dragged left of the left-indent block: clamped to -left
    const next = indentFromRuler('first', dims, { left: 720, right: 0, firstLine: 0 }, 1440)
    expect(next.firstLine).toBe(-720)
  })

  it('writes null rather than 0 so the style chain shows through', () => {
    const next = indentFromRuler('left', dims, { left: 720, right: 0, firstLine: 0 }, 1440)
    expect(next.left).toBeNull()
  })

  it('clamps inside the content area for hostile geometry', () => {
    // clamped to the content area; a zero indent writes null, not 0
    expect(
      indentFromRuler('left', dims, { left: 0, right: 0, firstLine: 0 }, -99999).left,
    ).toBeNull()
    expect(indentFromRuler('left', dims, { left: 0, right: 0, firstLine: 0 }, 99999).left).toBe(
      9360,
    )
  })
})

describe('Ruler indent markers', () => {
  it('renders all three, with accessible names', () => {
    setModuleLang('en')
    const { container, cleanup } = mount({ indentLeft: 720, indentFirstLine: 360 })
    try {
      for (const kind of ['first', 'left', 'right']) {
        const el = marker(container, kind)
        expect(el).not.toBeNull()
        expect(el.getAttribute('role')).toBe('slider')
        expect(el.getAttribute('tabindex')).toBe('0')
        expect(el.getAttribute('aria-label')).toBeTruthy()
      }
      expect(marker(container, 'first').getAttribute('aria-valuenow')).toBe('360')
      expect(marker(container, 'left').getAttribute('aria-valuenow')).toBe('720')
    } finally {
      cleanup()
    }
  })

  it('positions the markers from the paragraph indents', () => {
    const { container, cleanup } = mount({ indentLeft: 1440, indentFirstLine: 720 })
    try {
      // margin 1440 twips = 96px, indent 1440 twips = 96px, first line +48px
      expect(marker(container, 'left').style.left).toBe(`${MARGIN_PX + 96}px`)
      expect(marker(container, 'first').style.left).toBe(`${MARGIN_PX + 96 + 48}px`)
    } finally {
      cleanup()
    }
  })

  it('dragging the left block writes the new indent', () => {
    const { container, onIndentsChange, cleanup } = mount()
    try {
      drag(marker(container, 'left'), 300)
      expect(onIndentsChange).toHaveBeenCalledTimes(1)
      expect(onIndentsChange.mock.calls[0]![0]).toEqual({
        left: 3060,
        right: null,
        firstLine: null,
      })
    } finally {
      cleanup()
    }
  })

  it('dragging the first-line marker writes an offset from the left indent', () => {
    const { container, onIndentsChange, cleanup } = mount({ indentLeft: 720 })
    try {
      // 300px -> 4500 twips; minus margin 1440 and left indent 720
      drag(marker(container, 'first'), 300)
      expect(onIndentsChange.mock.calls[0]![0].firstLine).toBe(2340)
      expect(onIndentsChange.mock.calls[0]![0].left).toBe(720)
    } finally {
      cleanup()
    }
  })

  it('arrow keys nudge on the snap grid', () => {
    const { container, onIndentsChange, cleanup } = mount({ indentLeft: 720 })
    try {
      act(() => {
        marker(container, 'left').dispatchEvent(
          new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }),
        )
      })
      // the mark sits at margin + indent; one step right is +60 twips
      expect(onIndentsChange.mock.calls[0]![0].left).toBe(780)
    } finally {
      cleanup()
    }
  })

  it('releasing a marker does not also place a tab stop', () => {
    const { container, onTabStopsChange, cleanup } = mount()
    try {
      mouse(marker(container, 'left'), 'click', 300)
      expect(onTabStopsChange).not.toHaveBeenCalled()
    } finally {
      cleanup()
    }
  })
})
