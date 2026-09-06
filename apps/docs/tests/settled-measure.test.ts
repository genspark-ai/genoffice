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

describe('SettledParagraphCache', () => {
  const node = {} as ProseMirrorNode

  it('re-measures until two consecutive passes agree, then reuses the result', () => {
    const el = document.createElement('p')
    const view = fakeView(el)
    const cache = new SettledParagraphCache<number[]>()
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
    const cache = new SettledParagraphCache<number[]>()
    let rect = { top: 0, bottom: 10, height: 10, width: 100 }
    el.getBoundingClientRect = () => rect as DOMRect
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
    const cache = new SettledParagraphCache<number[]>()
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
})
