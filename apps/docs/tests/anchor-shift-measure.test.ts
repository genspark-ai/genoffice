import { describe, expect, it } from 'vitest'
import { anchorBoxLift, computeSectionedSlicesF2, measureBlocks } from '../src/renderer/pagination'
import { hoistWindow, pinnedCloneCss } from '../src/renderer/components/PaginationPreview'

const rectOf = (top: number, height: number) =>
  ({
    top,
    height,
    bottom: top + height,
    left: 0,
    right: 100,
    width: 100,
    x: 0,
    y: top,
    toJSON: () => ({}),
  }) as DOMRect

// paragraph, paragraph, then a wrapSquare textbox paragraph whose anchor offset
// (-150px) is painted as a wrapper translate: its rect sits 150px above its flow slot
function flow(zoom: number, shifted = true) {
  const pm = document.createElement('div')
  const add = (tag: string, top: number, height: number, html: string) => {
    const el = document.createElement(tag)
    el.innerHTML = html
    el.getBoundingClientRect = () => rectOf(top * zoom, height * zoom)
    pm.appendChild(el)
    return el
  }
  add('p', 0, 100, 'one')
  add('p', 100, 100, 'two')
  const box = add(
    'div',
    shifted ? 50 : 200,
    120,
    '<div class="doc-textbox-stray">anchor text</div><div class="doc-textbox">box</div>',
  )
  box.className = 'doc-protected doc-protected-passthrough doc-protected-textboxes'
  if (shifted) {
    box.style.transform = 'translate(40px,-150px)'
    box.dataset.anchorDy = '-150'
  }
  add('p', 320, 50, 'after')
  return pm
}

describe('measureBlocks — column breaks', () => {
  function flowWithColBreak(html: string) {
    const pm = document.createElement('div')
    const add = (top: number, height: number, inner: string) => {
      const el = document.createElement('p')
      el.innerHTML = inner
      el.getBoundingClientRect = () => rectOf(top, height)
      pm.appendChild(el)
    }
    add(0, 20, 'one')
    add(20, 20, html)
    add(40, 20, 'after')
    return pm
  }

  it('a leading (break-only) column break opens the next column with its own block', () => {
    const { blocks } = measureBlocks(flowWithColBreak('<br class="doc-col-br">'), 0, 1)
    expect(blocks[1].colBreakBefore).toBe(true)
    expect(blocks[1].colBreakAfter).toBeUndefined()
  })

  it('a trailing column break pushes the next block', () => {
    const { blocks } = measureBlocks(flowWithColBreak('text<br class="doc-col-br">'), 0, 1)
    expect(blocks[1].colBreakBefore).toBeUndefined()
    expect(blocks[1].colBreakAfter).toBe(true)
  })
})

describe('hoistWindow', () => {
  it('plain page: the page window, dy from the flow start', () => {
    expect(hoistWindow({ start: 100, end: 500 } as never, 160)).toEqual({
      start: 100,
      end: 500,
      dy: 60,
      multiCol: false,
    })
  })

  it('regioned page: the anchor column window, dy from the region top', () => {
    const slice = {
      start: 0,
      end: 300,
      regions: [
        { top: 0, height: 100, section: 0, columns: [{ start: 0, end: 100 }] },
        {
          top: 100,
          height: 200,
          section: 1,
          columns: [
            { start: 100, end: 200 },
            { start: 200, end: 300 },
          ],
        },
      ],
    } as never
    expect(hoistWindow(slice, 230)).toEqual({ start: 200, end: 300, dy: 130, multiCol: true })
    expect(hoistWindow(slice, 50)).toEqual({ start: 0, end: 100, dy: 50, multiCol: false })
    expect(hoistWindow(slice, 400)).toBeNull()
  })
})

describe('measureBlocks — anchor-shifted textbox wrapper', () => {
  it('measures the flow slot, not the translated paint rect', () => {
    const { blocks } = measureBlocks(flow(1), 0, 1)
    expect(blocks.map((b) => b.top)).toEqual([0, 100, 200, 320])
    expect(blocks[2].height).toBe(120)
  })

  it('scales the shift with the zoom factor', () => {
    const { blocks } = measureBlocks(flow(1.5), 0, 1.5)
    expect(blocks.map((b) => Math.round(b.top))).toEqual([0, 100, 200, 320])
  })

  it('leaves unshifted wrappers alone', () => {
    const { blocks } = measureBlocks(flow(1, false), 0, 1)
    expect(blocks[2].top).toBe(200)
  })

  it('slices the block on its own page instead of over the previous text', () => {
    const { blocks, totalHeight } = measureBlocks(flow(1), 0, 1)
    const slices = computeSectionedSlicesF2(
      blocks,
      [{ contentHeight: 250, forceBreak: false }],
      totalHeight,
    )
    expect(slices.map((s) => s.start)).toEqual([0, 200])
    for (let i = 1; i < slices.length; i++) {
      expect(slices[i].start).toBeGreaterThanOrEqual(slices[i - 1].end)
    }
  })
})

describe('pinnedCloneCss — anchor-shifted wrappers', () => {
  it('hides the ride-along copies on non-owning pages', () => {
    const css = pinnedCloneCss(2)
    expect(css).toContain(
      '.pv-page[data-pv-page="1"] [data-anchor-dy][data-pin-page]:not([data-pin-page="1"]){visibility:hidden;}',
    )
  })
})

describe('anchorBoxLift', () => {
  it('leaves a box inside the window alone', () => {
    expect(anchorBoxLift(1100, 1500, 1000, 1900)).toBe(0)
  })

  it('brings a box reaching above the window top down onto the page', () => {
    expect(anchorBoxLift(850, 1300, 1000, 1900)).toBe(150)
  })

  it('moves a box overrunning the window bottom back up', () => {
    expect(anchorBoxLift(1700, 2100, 1000, 1900)).toBe(-200)
  })

  it('never lifts a window-tall box above the top', () => {
    expect(anchorBoxLift(1200, 2300, 1000, 1900)).toBe(-200)
  })
})

describe('measureBlocks — first-block lead fold', () => {
  it('records the folded lead so the slot top can be recovered', () => {
    const pm = document.createElement('div')
    const el = document.createElement('div')
    el.className = 'doc-protected doc-protected-passthrough doc-protected-textboxes'
    el.innerHTML = '<div class="doc-textbox">box</div>'
    el.dataset.anchorDy = '-100'
    el.getBoundingClientRect = () => rectOf(-80, 120)
    pm.appendChild(el)
    const { blocks } = measureBlocks(pm, 0, 1)
    expect(blocks[0].top).toBe(0)
    expect(blocks[0].leadFoldPx).toBe(20)
    expect(blocks[0].height).toBe(140)
  })
})
