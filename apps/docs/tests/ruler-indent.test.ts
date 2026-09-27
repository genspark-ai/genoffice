import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { Editor } from '@tiptap/core'
import type { SectionSettings } from '@genoffice/docx-engine'
import {
  Ruler,
  applyIndentDrag,
  clampMarginDrag,
  cssLengthToTwips,
  indentMarkerPositions,
  snapIndentTwips,
  type BodyBounds,
  type ParagraphIndents,
} from '../src/renderer/components/Ruler'
import { getLang, setModuleLang } from '../src/renderer/i18n/locale'

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

const originalLang = getLang()
afterEach(() => {
  setModuleLang(originalLang)
  vi.restoreAllMocks()
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
const body: BodyBounds = { left: 1440, right: 10800, inCell: false }
const hanging: ParagraphIndents = { left: 720, right: 0, firstLine: -360 }

describe('indent marker math', () => {
  it('snaps to 1/16 inch unless disabled', () => {
    expect(snapIndentTwips(100)).toBe(90)
    expect(snapIndentTwips(135)).toBe(180)
    expect(snapIndentTwips(100, false)).toBe(100)
    expect(snapIndentTwips(Number.NaN)).toBe(0)
  })

  it('places the markers from the body edge', () => {
    expect(indentMarkerPositions(hanging, body)).toEqual({
      firstLine: 1800,
      hanging: 2160,
      left: 2160,
      right: 10800,
    })
  })

  it('first-line marker moves only the first line', () => {
    const r = applyIndentDrag('firstLine', hanging, 2880, body, 12240)
    expect(r.attrs).toEqual({ indentFirstLine: 720 })
    expect(r.indents).toEqual({ left: 720, right: 0, firstLine: 720 })
  })

  it('hanging marker keeps the first line where it is', () => {
    const r = applyIndentDrag('hanging', hanging, 2880, body, 12240)
    expect(r.attrs).toEqual({ indentLeft: 1440, indentFirstLine: -1080 })
    expect(indentMarkerPositions(r.indents, body).firstLine).toBe(1800)
  })

  it('left box moves first line and hanging together', () => {
    const r = applyIndentDrag('left', hanging, 2880, body, 12240)
    expect(r.attrs).toEqual({ indentLeft: 1440 })
    expect(indentMarkerPositions(r.indents, body)).toMatchObject({ firstLine: 2520, left: 2880 })
  })

  it('right marker writes the right indent', () => {
    expect(applyIndentDrag('right', hanging, 9360, body, 12240).attrs).toEqual({
      indentRight: 1440,
    })
    expect(applyIndentDrag('right', hanging, 11520, body, 12240).attrs).toEqual({
      indentRight: -720,
    })
  })

  it('markers never cross each other or leave the page', () => {
    expect(applyIndentDrag('right', hanging, 1000, body, 12240).attrs).toEqual({
      indentRight: 10800 - 2250,
    })
    expect(applyIndentDrag('firstLine', hanging, -500, body, 12240).attrs).toEqual({
      indentFirstLine: -2160,
    })
    expect(applyIndentDrag('left', hanging, 20000, body, 12240).attrs).toEqual({
      indentLeft: 10710 - 1440,
    })
    // the box carries the first-line marker: it stops where that marker would leave the page
    expect(applyIndentDrag('left', hanging, -500, body, 12240).attrs).toEqual({
      indentLeft: 360 - 1440,
    })
    const positive: ParagraphIndents = { left: 0, right: 0, firstLine: 720 }
    expect(applyIndentDrag('left', positive, 20000, body, 12240).attrs).toEqual({
      indentLeft: 10710 - 720 - 1440,
    })
  })

  it('measures from the cell in a table', () => {
    const cell: BodyBounds = { left: 3000, right: 6000, inCell: true }
    expect(indentMarkerPositions({ left: 0, right: 0, firstLine: 0 }, cell)).toMatchObject({
      left: 3000,
      right: 6000,
    })
    expect(
      applyIndentDrag('left', { left: 0, right: 0, firstLine: 0 }, 3360, cell, 12240).attrs,
    ).toEqual({ indentLeft: 360 })
  })
})

describe('clampMarginDrag', () => {
  const dims = { pageWidth: 12240, marginLeft: 1440, marginRight: 1440 }
  it('keeps half an inch of body', () => {
    expect(clampMarginDrag('left', 2000, dims)).toBe(2000)
    expect(clampMarginDrag('left', 12000, dims)).toBe(10080)
    expect(clampMarginDrag('left', -50, dims)).toBe(0)
    expect(clampMarginDrag('right', 10000, dims)).toBe(2240)
    expect(clampMarginDrag('right', 500, dims)).toBe(10080)
    expect(clampMarginDrag('right', 13000, dims)).toBe(0)
  })
})

describe('cssLengthToTwips', () => {
  it('converts computed px and custom-property units', () => {
    expect(cssLengthToTwips('96px')).toBe(1440)
    expect(cssLengthToTwips('18pt')).toBe(360)
    expect(cssLengthToTwips(' 0.25in ')).toBe(360)
    expect(cssLengthToTwips('2.54cm')).toBe(1440)
    expect(cssLengthToTwips('-12px')).toBe(-180)
    expect(cssLengthToTwips('')).toBeNull()
    expect(cssLengthToTwips('auto')).toBeNull()
    expect(cssLengthToTwips(undefined)).toBeNull()
  })
})

function mount(
  attrs: Record<string, unknown>,
  editorOver: Record<string, unknown> = {},
  props: { indents?: boolean; margins?: boolean } = { indents: true, margins: true },
) {
  const onIndentsChange = vi.fn()
  const onMarginsChange = vi.fn()
  const editor = {
    isActive: (name: string) => name === 'docParagraph',
    getAttributes: () => attrs,
    ...editorOver,
  } as unknown as Editor
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root: Root = createRoot(container)
  act(() => {
    root.render(
      createElement(Ruler, {
        section,
        editor,
        onTabStopsChange: vi.fn(),
        onIndentsChange: props.indents ? onIndentsChange : undefined,
        onMarginsChange: props.margins ? onMarginsChange : undefined,
      }),
    )
  })
  return {
    container,
    onIndentsChange,
    onMarginsChange,
    cleanup: () => {
      act(() => root.unmount())
      container.remove()
    },
  }
}

const marker = (container: HTMLElement, label: string) =>
  container.querySelector(`[aria-label="${label}"]`) as HTMLElement

function keydown(el: Element, key: string, shift = false): void {
  act(() => {
    el.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, shiftKey: shift }))
  })
}

function mouse(type: string, target: EventTarget, clientX: number, altKey = false): void {
  act(() => {
    target.dispatchEvent(
      new MouseEvent(type, { bubbles: true, cancelable: true, clientX, clientY: 10, altKey }),
    )
  })
}

describe('Ruler indent and margin handles', () => {
  it('renders the four markers at the paragraph indents and the margin handles', () => {
    setModuleLang('en')
    const m = mount({ indentLeft: 720, indentFirstLine: -360, indentRight: 0 })
    expect(marker(m.container, 'First Line Indent').getAttribute('aria-valuenow')).toBe('1800')
    expect(marker(m.container, 'Hanging Indent').getAttribute('aria-valuenow')).toBe('2160')
    expect(marker(m.container, 'Left Indent').getAttribute('aria-valuenow')).toBe('2160')
    expect(marker(m.container, 'Right Indent').getAttribute('aria-valuenow')).toBe('10800')
    expect(marker(m.container, 'Left Margin').getAttribute('aria-valuenow')).toBe('1440')
    expect(m.container.querySelectorAll('.ruler-margin-handle')).toHaveLength(2)
    m.cleanup()
  })

  it('arrow keys nudge markers on the 1/16 inch grid, Shift by half an inch', () => {
    setModuleLang('en')
    const m = mount({ indentLeft: 720, indentFirstLine: -360 })
    keydown(marker(m.container, 'First Line Indent'), 'ArrowRight')
    expect(m.onIndentsChange).toHaveBeenLastCalledWith({ indentFirstLine: -270 })
    keydown(marker(m.container, 'Hanging Indent'), 'ArrowRight', true)
    expect(m.onIndentsChange).toHaveBeenLastCalledWith({
      indentLeft: 1440,
      indentFirstLine: -1080,
    })
    keydown(marker(m.container, 'Left Indent'), 'ArrowLeft')
    expect(m.onIndentsChange).toHaveBeenLastCalledWith({ indentLeft: 630 })
    keydown(marker(m.container, 'Right Indent'), 'ArrowLeft', true)
    expect(m.onIndentsChange).toHaveBeenLastCalledWith({ indentRight: 720 })
    keydown(marker(m.container, 'Left Margin'), 'ArrowLeft')
    expect(m.onMarginsChange).toHaveBeenLastCalledWith({ left: 1350, right: 1440 })
    keydown(marker(m.container, 'Right Margin'), 'ArrowLeft', true)
    expect(m.onMarginsChange).toHaveBeenLastCalledWith({ left: 1440, right: 2160 })
    m.cleanup()
  })

  it('mouse drag commits once on release, snapped unless Option is held', () => {
    setModuleLang('en')
    vi.spyOn(Element.prototype, 'getBoundingClientRect').mockReturnValue({
      left: 0,
      top: 0,
      right: 816,
      bottom: 20,
      width: 816,
      height: 20,
      x: 0,
      y: 0,
      toJSON: () => ({}),
    } as DOMRect)
    const m = mount({ indentLeft: 792, indentFirstLine: -360 })
    const box = marker(m.container, 'Left Indent')
    mouse('mousedown', box, 100)
    mouse('mouseup', document, 100)
    expect(m.onIndentsChange).not.toHaveBeenCalled()
    mouse('mousedown', box, 100)
    mouse('mousemove', document, 124)
    expect(document.querySelector('.ruler-drag-guide')).not.toBeNull()
    expect(m.onIndentsChange).not.toHaveBeenCalled()
    expect(box.style.left).toBe(`${(2610 / 1440) * 96}px`)
    mouse('mouseup', document, 148)
    expect(document.querySelector('.ruler-drag-guide')).toBeNull()
    expect(m.onIndentsChange).toHaveBeenCalledTimes(1)
    expect(m.onIndentsChange).toHaveBeenLastCalledWith({ indentLeft: 2970 - 1440 })

    const right = marker(m.container, 'Right Indent')
    mouse('mousedown', right, 600)
    mouse('mouseup', document, 593, true)
    expect(m.onIndentsChange).toHaveBeenLastCalledWith({ indentRight: 105 })

    const margin = marker(m.container, 'Left Margin')
    mouse('mousedown', margin, 96)
    mouse('mousemove', document, 144)
    expect((m.container.querySelector('.ruler-zone') as HTMLElement).style.width).toBe('144px')
    mouse('mouseup', document, 144)
    expect(m.onMarginsChange).toHaveBeenLastCalledWith({ left: 2160, right: 1440 })
    m.cleanup()
  })

  it('hides the margin handles and measures from the cell when the caret is in a table', () => {
    setModuleLang('en')
    const page = document.createElement('div')
    page.className = 'doc-page'
    const cell = document.createElement('td')
    cell.style.padding = '0'
    page.appendChild(cell)
    document.body.appendChild(page)
    const rect = (left: number, width: number) =>
      ({ left, right: left + width, width, top: 0, bottom: 10, height: 10 }) as DOMRect
    vi.spyOn(page, 'getBoundingClientRect').mockReturnValue(rect(0, 816))
    vi.spyOn(cell, 'getBoundingClientRect').mockReturnValue(rect(200, 100))
    const $from = {
      depth: 3,
      node: (d: number) => ({ type: { name: d === 2 ? 'docTableCell' : 'docParagraph' } }),
      before: () => 5,
    }
    const m = mount(
      { indentLeft: 0, indentFirstLine: 0, indentRight: 0 },
      { state: { selection: { $from } }, view: { nodeDOM: () => cell } },
    )
    expect(m.container.querySelectorAll('.ruler-margin-handle')).toHaveLength(0)
    expect(marker(m.container, 'Left Indent').getAttribute('aria-valuenow')).toBe('3000')
    expect(marker(m.container, 'Right Indent').getAttribute('aria-valuenow')).toBe('4500')
    m.cleanup()
    page.remove()
  })
})
