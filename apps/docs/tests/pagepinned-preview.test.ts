import { describe, expect, it } from 'vitest'
import { pinnedFloatPage, type PageSlice } from '../src/renderer/pagination'
import { pinnedCloneCss } from '../src/renderer/components/PaginationPreview'

const slice = (start: number, end: number): PageSlice => ({ start, end, section: 0 })

describe('pinnedFloatPage', () => {
  const slices = [slice(0, 900), slice(900, 1800), slice(1800, 2700)]

  it('maps an anchor to its containing slice', () => {
    expect(pinnedFloatPage(slices, 0)).toBe(0)
    expect(pinnedFloatPage(slices, 899.9)).toBe(0)
    expect(pinnedFloatPage(slices, 900)).toBe(1)
    expect(pinnedFloatPage(slices, 2000)).toBe(2)
  })

  it('clamps out-of-range anchors', () => {
    expect(pinnedFloatPage(slices, -5)).toBe(0)
    expect(pinnedFloatPage(slices, 2700)).toBe(2)
    expect(pinnedFloatPage(slices, 99999)).toBe(2)
    expect(pinnedFloatPage([], 100)).toBe(0)
  })
})

describe('pinnedCloneCss', () => {
  const page = (index: number, html: string): HTMLElement => {
    const el = document.createElement('div')
    el.className = 'pv-page'
    el.dataset.pvPage = String(index)
    el.innerHTML = html
    return el
  }
  const pinned = (pin?: number) =>
    `<div class="doc-protected doc-protected-floating doc-protected-pagepinned"${
      pin === undefined ? '' : ` data-pin-page="${pin}"`
    }><div class="doc-textbox"></div></div>`
  const boxes = (el: HTMLElement): HTMLElement[] =>
    Array.from(el.querySelectorAll('.doc-protected-pagepinned'))
  const matches = (root: HTMLElement, css: string): boolean[] => {
    const selectors = css
      .split('\n')
      .filter(Boolean)
      .map((r) => r.slice(0, r.indexOf('{')))
    return boxes(root).map((b) => selectors.some((s) => b.matches(s)))
  }

  it('hides clones owned by other pages, keeps the owning page copy', () => {
    const css = pinnedCloneCss(2)
    // visibility, not display: stray-run wrappers keep their measured flow height
    expect(css).toContain('visibility:hidden')
    expect(css).not.toContain('display')
    // full-document clone: every page carries both covers (pin 0 and pin 1)
    const p0 = page(0, pinned(0) + pinned(1) + '<p>body</p>')
    const p1 = page(1, pinned(0) + pinned(1) + '<p>body</p>')
    expect(matches(p0, css)).toEqual([false, true])
    expect(matches(p1, css)).toEqual([true, false])
  })

  it('never targets unstamped wrappers', () => {
    const css = pinnedCloneCss(3)
    const p = page(2, pinned() + pinned(2) + pinned(0))
    expect(matches(p, css)).toEqual([false, false, true])
  })

  it('emits no rules for an empty preview', () => {
    expect(pinnedCloneCss(0)).toBe('')
  })
})
