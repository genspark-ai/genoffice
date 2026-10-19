import { afterEach, describe, expect, it } from 'vitest'
import { collectPdfInternalLinks } from '../src/renderer/pdf-internal-links'

type Box = { left: number; top: number; width: number; height: number }
const rect = (b: Box) =>
  ({ ...b, right: b.left + b.width, bottom: b.top + b.height, x: b.left, y: b.top }) as DOMRect

/**
 * Two 100x200 pages (top margin 20) stacked at y=0 and y=300. The flow is one
 * continuous column with the second page's slice starting at flow 150 (a
 * forced break: the heading at flow 120 is visually inside page 1's clone box
 * too but logically on page 0). Page 1's clip cuts its clone at flow 150 and
 * is preceded by a repeated-table-header clone that is not a body window.
 */
function mount(): HTMLElement[] {
  document.body.innerHTML = `
    <div class="pv-page" data-pv-page="0"><div class="pv-clip"><div class="pv-content" data-pv-body="">
      <div class="doc-toc-line" data-toc-anchor="_Toc1" data-toc-title="Intro"></div>
      <div class="doc-toc-line" data-toc-title="By Title"></div>
      <p><a href="#_Ref9">ref</a> <a href="https://x.test">ext</a></p>
      <h1 data-idx="5">Intro</h1>
      <h2 data-idx="7">By Title</h2>
      <p data-idx="9">target</p>
    </div></div></div>
    <div class="pv-page" data-pv-page="1"><div class="pv-clip"><div class="pv-content"><h1 data-idx="5">Intro</h1></div></div><div class="pv-clip"><div class="pv-content" data-pv-body="">
      <div class="doc-toc-line" data-toc-anchor="_Toc1" data-toc-title="Intro"></div>
      <div class="doc-toc-line" data-toc-title="By Title"></div>
      <p><a href="#_Ref9">ref</a> <a href="https://x.test">ext</a></p>
      <h1 data-idx="5">Intro</h1>
      <h2 data-idx="7">By Title</h2>
      <p data-idx="9">target</p>
    </div></div></div>`
  const pages = [...document.querySelectorAll<HTMLElement>('.pv-page')]
  const boxes = new Map<Element, Box[]>()
  const place = (page: HTMLElement, pageTop: number, flowShift: number) => {
    boxes.set(page, [{ left: 0, top: pageTop, width: 100, height: 200 }])
    const content = page.querySelector<HTMLElement>('.pv-content[data-pv-body]')!
    const origin = pageTop + 20 - flowShift
    boxes.set(content, [{ left: 0, top: origin, width: 100, height: 1000 }])
    boxes.set(content.parentElement!, [{ left: 0, top: pageTop + 20, width: 100, height: 160 }])
    const at = (sel: string, flowTop: number, h = 10, left = 0, width = 100) =>
      boxes.set(content.querySelector(sel)!, [{ left, top: origin + flowTop, width, height: h }])
    at('[data-toc-anchor]', 0, 10, 10, 80)
    at('.doc-toc-line:not([data-toc-anchor])', 10, 10, 10, 80)
    // the cross-reference wraps onto two lines
    boxes.set(content.querySelector('a[href="#_Ref9"]')!, [
      { left: 50, top: origin + 30, width: 50, height: 10 },
      { left: 0, top: origin + 40, width: 20, height: 10 },
    ])
    at('h1', 120)
    at('h2', 170)
    at('[data-idx="9"]', 230)
  }
  place(pages[0], 0, 0)
  place(pages[1], 300, 150)
  Element.prototype.getBoundingClientRect = function () {
    return rect(boxes.get(this)?.[0] ?? { left: 0, top: 0, width: 0, height: 0 })
  }
  Element.prototype.getClientRects = function () {
    const list = (boxes.get(this) ?? []).map(rect)
    return Object.assign(list, { item: (i: number) => list[i] ?? null }) as unknown as DOMRectList
  }
  return pages
}

const origBounding = Element.prototype.getBoundingClientRect
const origRects = Element.prototype.getClientRects
afterEach(() => {
  Element.prototype.getBoundingClientRect = origBounding
  Element.prototype.getClientRects = origRects
  document.body.innerHTML = ''
})

describe('collectPdfInternalLinks', () => {
  it('assigns sources and targets to pages by flow position, once each', () => {
    const pages = mount()
    const links = collectPdfInternalLinks(pages, [
      { docxIndex: 5, hiddenBookmarks: ['_Toc1'] },
      { docxIndex: 9, bookmarks: ['_Ref9'] },
      { docxIndex: 7 },
    ])
    // the heading at flow 120 precedes page 1's window start (150): page 0, y = 20 + 120;
    // the repeated-header clone on page 1 is not a body window
    expect(links).toEqual([
      { page: 0, x: 0.1, y: 0.1, w: 0.8, h: 0.05, toPage: 0, toY: 0.7 },
      { page: 0, x: 0.1, y: 0.15, w: 0.8, h: 0.05, toPage: 1, toY: 0.2 },
      { page: 0, x: 0.5, y: 0.25, w: 0.5, h: 0.05, toPage: 1, toY: 0.5 },
      { page: 0, x: 0, y: 0.3, w: 0.2, h: 0.05, toPage: 1, toY: 0.5 },
    ])
  })

  it('falls back to the heading text when the bookmark is unknown', () => {
    const pages = mount()
    expect(collectPdfInternalLinks(pages, [{ docxIndex: 9, bookmarks: ['_Ref9'] }])).toEqual([
      { page: 0, x: 0.1, y: 0.1, w: 0.8, h: 0.05, toPage: 0, toY: 0.7 },
      { page: 0, x: 0.1, y: 0.15, w: 0.8, h: 0.05, toPage: 1, toY: 0.2 },
      { page: 0, x: 0.5, y: 0.25, w: 0.5, h: 0.05, toPage: 1, toY: 0.5 },
      { page: 0, x: 0, y: 0.3, w: 0.2, h: 0.05, toPage: 1, toY: 0.5 },
    ])
  })
})
