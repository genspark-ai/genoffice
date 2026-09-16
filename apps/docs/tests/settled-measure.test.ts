import { describe, expect, it, vi } from 'vitest'
import { SettledParagraphCache } from '../src/renderer/editor/settled-measure'
import type { EditorView } from '@tiptap/pm/view'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'

function fakeView(el: HTMLElement) {
  return {
    dom: document.createElement('div'),
    nodeDOM: () => el,
  } as unknown as EditorView
}

/** jsdom has no layout: give the paragraph the box the test wants */
function sized(el: HTMLElement, box: () => { width: number; height: number }) {
  Object.defineProperty(el, 'offsetWidth', { get: () => box().width })
  Object.defineProperty(el, 'offsetHeight', { get: () => box().height })
}

const shiftAll = (r: number[], d: number) => r.map((v) => v + d)

describe('SettledParagraphCache', () => {
  const node = {} as ProseMirrorNode

  it('re-measures until two consecutive passes agree, then reuses the result', () => {
    const el = document.createElement('p')
    const view = fakeView(el)
    const cache = new SettledParagraphCache<number[]>(shiftAll)
    const fn = vi.fn<() => number[]>()
    const pass = () => {
      cache.beginPass(view)
      return cache.measure(view, node, 1, fn)
    }
    fn.mockReturnValueOnce([1]).mockReturnValueOnce([2]).mockReturnValueOnce([2])
    expect(pass()).toEqual([1])
    expect(pass()).toEqual([2])
    expect(pass()).toEqual([2])
    expect(fn).toHaveBeenCalledTimes(3)
    expect(pass()).toEqual([2])
    expect(pass()).toEqual([2])
    expect(fn).toHaveBeenCalledTimes(3)
  })

  it('invalidates on clear(), a different node, a box change or an unmeasurable pass', () => {
    const el = document.createElement('p')
    const view = fakeView(el)
    const cache = new SettledParagraphCache<number[]>(shiftAll)
    let rect = { top: 0, bottom: 10, height: 10, width: 100 }
    sized(el, () => rect)
    const fn = vi.fn<() => number[] | null>(() => [7])
    const pass = (n: ProseMirrorNode = node) => {
      cache.beginPass(view)
      return cache.measure(view, n, 1, fn)
    }
    pass()
    pass()
    pass()
    expect(fn).toHaveBeenCalledTimes(2)
    rect = { ...rect, height: 20 }
    pass()
    expect(fn).toHaveBeenCalledTimes(3)
    pass()
    expect(fn).toHaveBeenCalledTimes(3)
    pass({} as ProseMirrorNode)
    expect(fn).toHaveBeenCalledTimes(4)
    cache.clear()
    pass()
    pass()
    pass()
    expect(fn).toHaveBeenCalledTimes(6)
    fn.mockReturnValueOnce(null)
    cache.clear()
    expect(pass()).toBeNull()
    pass()
    pass()
    pass()
    expect(fn).toHaveBeenCalledTimes(9)
  })

  it('keys on vertical position only beside a float', () => {
    const el = document.createElement('p')
    const view = fakeView(el)
    let top = 0
    el.getBoundingClientRect = () => ({ top, bottom: top + 10, height: 10, width: 100 }) as DOMRect
    sized(el, () => ({ width: 100, height: 10 }))
    const cache = new SettledParagraphCache<number[]>(shiftAll)
    const fn = vi.fn<() => number[]>(() => [1])
    const pass = () => {
      cache.beginPass(view)
      return cache.measure(view, node, 1, fn)
    }
    pass()
    pass()
    top = 50
    pass()
    expect(fn).toHaveBeenCalledTimes(2)
    const float = document.createElement('div')
    float.className = 'doc-table doc-table-float-left'
    float.getBoundingClientRect = () => ({ top: 40, bottom: 80 }) as DOMRect
    view.dom.appendChild(float)
    pass()
    expect(fn).toHaveBeenCalledTimes(3)
    top = 60
    pass()
    expect(fn).toHaveBeenCalledTimes(4)
    top = 200
    pass()
    pass()
    pass()
    expect(fn).toHaveBeenCalledTimes(5)
  })

  it('keeps a settled paragraph across an edit elsewhere and re-bases its positions', () => {
    const el = document.createElement('p')
    const view = fakeView(el)
    const cache = new SettledParagraphCache<number[]>(shiftAll)
    const fn = vi.fn<() => number[]>(() => [12, 15])
    const pass = (pos: number) => {
      cache.beginPass(view)
      return cache.measure(view, node, pos, fn)
    }
    expect(pass(10)).toEqual([12, 15])
    expect(pass(10)).toEqual([12, 15])
    expect(fn).toHaveBeenCalledTimes(2)
    // text inserted above the paragraph: same node, new position, no re-measure
    expect(pass(13)).toEqual([15, 18])
    expect(pass(13)).toEqual([15, 18])
    expect(fn).toHaveBeenCalledTimes(2)
    // a result re-measured at a new position settles against the position-independent key
    const other = {} as ProseMirrorNode
    fn.mockReturnValueOnce([3]).mockReturnValueOnce([5])
    cache.beginPass(view)
    expect(cache.measure(view, other, 1, fn)).toEqual([3])
    cache.beginPass(view)
    expect(cache.measure(view, other, 3, fn)).toEqual([5])
    cache.beginPass(view)
    expect(cache.measure(view, other, 3, fn)).toEqual([5])
    expect(fn).toHaveBeenCalledTimes(4)
  })

  it('prunes paragraphs a pass does not visit (deleted nodes)', () => {
    const el = document.createElement('p')
    const view = fakeView(el)
    const cache = new SettledParagraphCache<number[]>(shiftAll)
    const fn = vi.fn<() => number[]>(() => [1])
    const a = {} as ProseMirrorNode
    const b = {} as ProseMirrorNode
    for (let i = 0; i < 2; i++) {
      cache.beginPass(view)
      cache.measure(view, a, 1, fn)
      cache.measure(view, b, 5, fn)
    }
    expect(fn).toHaveBeenCalledTimes(4)
    // b disappears for two passes → its entry is dropped; a stays settled
    cache.beginPass(view)
    cache.measure(view, a, 1, fn)
    cache.beginPass(view)
    cache.measure(view, a, 1, fn)
    expect(fn).toHaveBeenCalledTimes(4)
    cache.beginPass(view)
    cache.measure(view, b, 5, fn)
    expect(fn).toHaveBeenCalledTimes(5)
  })

  it('reads top-level block DOM off the view descriptors', () => {
    const view = fakeView(document.createElement('p'))
    const a = {} as ProseMirrorNode
    const b = {} as ProseMirrorNode
    const pa = document.createElement('p')
    const pb = document.createElement('p')
    ;(pa as unknown as { pmViewDesc: unknown }).pmViewDesc = { node: a }
    ;(pb as unknown as { pmViewDesc: unknown }).pmViewDesc = { node: b }
    const widget = document.createElement('span') // no desc node
    view.dom.append(pa, widget, pb, document.createTextNode('x'))
    const map = SettledParagraphCache.topLevelDom(view)
    expect(map.get(a)).toBe(pa)
    expect(map.get(b)).toBe(pb)
    expect(map.size).toBe(2)
    // the hint replaces nodeDOM for the lookup
    const cache = new SettledParagraphCache<number[]>(shiftAll)
    const fn = vi.fn<(el: HTMLElement) => number[]>(() => [1])
    cache.beginPass(view)
    cache.measure(view, a, 1, fn, pa)
    expect(fn).toHaveBeenCalledWith(pa)
  })
})
