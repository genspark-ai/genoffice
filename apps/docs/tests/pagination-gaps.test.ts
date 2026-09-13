import { describe, expect, it } from 'vitest'
import {
  GAP_BAND,
  alignGapHfStrips,
  alignTableGapFills,
  makeGapEl,
  pageFramesFromGaps,
  syncCutOverlays,
  syncPageSheets,
  syncPhantomRowspans,
  clampCellBoxTops,
  pageBorderArtStripStyle,
  pageBorderStyleOf,
  rowFillAttrs,
  syncPageBorders,
} from '../src/renderer/editor/pagination-gaps'
import { createLineRectsCache, singleCutCell } from '../src/renderer/pagination'

const rowOf = (cells: number): HTMLTableRowElement => {
  const tr = document.createElement('tr')
  for (let i = 0; i < cells; i++) tr.appendChild(document.createElement('td'))
  return tr
}

const m = { marginTop: 96, marginBottom: 96, marginLeft: 90, marginRight: 90 }

describe('page gap header push', () => {
  it('writes the push even when zero so a later page clears an earlier one', () => {
    expect(makeGapEl({ ...m, sectionMarginTop: 70 }, 'cell').dataset.topPush).toBe('26.0')
    expect(makeGapEl({ ...m, sectionMarginTop: 96 }, 'cell').dataset.topPush).toBe('0.0')
    expect(makeGapEl(m, 'cell').dataset.topPush).toBeUndefined()
  })
})

describe('in-row table cut decorations', () => {
  it('single-cell row: real inline gap band, not a zero-height cut marker', () => {
    const tr = rowOf(1)
    expect(singleCutCell(tr)).toBe(tr.firstElementChild)
    const el = makeGapEl(m, 'cell')
    // page-gap-inline is what the measurement gap-subtraction keys on
    expect(el.classList.contains('page-gap')).toBe(true)
    expect(el.classList.contains('page-gap-inline')).toBe(true)
    expect(el.classList.contains('page-gap-cut')).toBe(false)
    expect(el.style.height).toBe(`${96 + GAP_BAND + 96}px`)
    expect(el.style.getPropertyValue('--gap-mb')).toBe('96px')
    expect(el.style.width).toBe('calc(100% + 180px)')
  })

  it('multi-cell rows without an anchor / missing rows keep the zero-height cut marker', () => {
    expect(singleCutCell(rowOf(2))).toBeNull()
    expect(singleCutCell(null)).toBeNull()
    expect(makeGapEl(m, 'cut').className).toBe('page-gap-cut')
  })

  it('table gap row spans exactly the host grid: phantom columns collapse colgroup-less fixed-layout tables', () => {
    const gap = makeGapEl(m, 'table', 3)
    expect(gap.tagName).toBe('TR')
    const cell = gap.firstElementChild as HTMLTableCellElement
    expect(cell.tagName).toBe('TD')
    expect(cell.colSpan).toBe(3)
    // no count supplied: never widen the grid
    const bare = makeGapEl(m, 'table').firstElementChild as HTMLTableCellElement
    expect(bare.colSpan).toBe(1)
  })

  const rectOf = (top: number, height: number) =>
    ({ top, bottom: top + height, height, width: 100 }) as DOMRect
  const cellAnchorRow = (siblingBottom: number) => {
    // host cell: article content with the cut anchor at y=500; sibling: spacer content
    const tr = rowOf(2)
    const host = tr.children[0] as HTMLElement
    const p = document.createElement('p')
    const text = document.createTextNode('article body')
    p.appendChild(text)
    p.getBoundingClientRect = () => rectOf(500, 20)
    host.appendChild(p)
    const sibling = tr.children[1] as HTMLElement
    const sp = document.createElement('p')
    sp.getBoundingClientRect = () => rectOf(0, siblingBottom)
    sibling.appendChild(sp)
    return { tr, anchor: { node: text, charOffset: 0 } }
  }

  it('multi-cell row hosts a real band when sibling content ends above the cut', () => {
    const { tr, anchor } = cellAnchorRow(4) // 1px-spacer-gif sliver cell
    expect(singleCutCell(tr, anchor)).toBe(tr.children[0])
  })

  it('multi-cell row keeps the cut marker when a sibling has content below the cut', () => {
    const { tr, anchor } = cellAnchorRow(900)
    expect(singleCutCell(tr, anchor)).toBeNull()
  })

  it('a sibling block holding only a spacer strut below the cut still hosts a band', () => {
    const { tr, anchor } = cellAnchorRow(900)
    const sp = tr.children[1].firstElementChild as HTMLElement
    const strut = document.createElement('img')
    // 1px-wide vertical strut gif reaching below the cut (the sample-72 spacer)
    strut.getBoundingClientRect = () => ({ top: 0, bottom: 900, height: 900, width: 1 }) as DOMRect
    sp.appendChild(strut)
    expect(singleCutCell(tr, anchor)).toBe(tr.children[0])
  })
})

describe('phantom-row rowspan bridging', () => {
  const rootOf = (rows: string): HTMLElement => {
    const root = document.createElement('div')
    root.innerHTML = `<table><tbody>${rows}</tbody></table>`
    return root
  }
  const cell = (root: HTMLElement, id: string) =>
    root.querySelector(`#${id}`) as HTMLTableCellElement

  it('grows a rowspan crossing the insertion point and records the base', () => {
    const root = rootOf(`
      <tr><td id="a" rowspan="3"></td><td></td></tr>
      <tr><td></td></tr>
      <tr class="page-gap"><td colspan="1000"></td></tr>
      <tr><td></td></tr>
      <tr><td id="b" rowspan="2"></td><td></td></tr>
      <tr><td></td></tr>`)
    syncPhantomRowspans(root)
    expect(cell(root, 'a').rowSpan).toBe(4)
    expect(cell(root, 'a').getAttribute('data-base-rowspan')).toBe('3')
    // span entirely below the gap is untouched
    expect(cell(root, 'b').rowSpan).toBe(2)
    expect(cell(root, 'b').hasAttribute('data-base-rowspan')).toBe(false)
  })

  it('accumulates per insertion point, ignores spans ending at the boundary, restores on removal', () => {
    const root = rootOf(`
      <tr><td id="a" rowspan="5"></td><td id="c" rowspan="2"></td><td></td></tr>
      <tr><td></td></tr>
      <tr class="page-gap"><td></td></tr>
      <tr class="page-repeat-header"><td></td><td></td><td></td></tr>
      <tr><td></td><td></td></tr>
      <tr><td></td><td></td></tr>
      <tr class="page-gap"><td></td></tr>
      <tr><td></td><td></td></tr>`)
    syncPhantomRowspans(root)
    // gap + header clone before real row 2, gap before real row 4: +3
    expect(cell(root, 'a').rowSpan).toBe(8)
    expect(cell(root, 'a').getAttribute('data-base-rowspan')).toBe('5')
    // c spans real rows 0-1: ends exactly at the first insertion point
    expect(cell(root, 'c').rowSpan).toBe(2)
    expect(cell(root, 'c').hasAttribute('data-base-rowspan')).toBe(false)
    // idempotent
    syncPhantomRowspans(root)
    expect(cell(root, 'a').rowSpan).toBe(8)
    for (const tr of root.querySelectorAll('tr.page-gap, tr.page-repeat-header')) tr.remove()
    syncPhantomRowspans(root)
    expect(cell(root, 'a').rowSpan).toBe(5)
    expect(cell(root, 'a').hasAttribute('data-base-rowspan')).toBe(false)
  })

  it('leaves tables without phantom rows alone', () => {
    const root = rootOf(`
      <tr><td id="a" rowspan="2"></td><td></td></tr>
      <tr><td></td></tr>`)
    syncPhantomRowspans(root)
    expect(cell(root, 'a').rowSpan).toBe(2)
    expect(cell(root, 'a').hasAttribute('data-base-rowspan')).toBe(false)
  })
})

describe('overlay cut markers (read-only nested-table anchors)', () => {
  const rectOf = (top: number) => ({ top, height: 10 }) as DOMRect
  const wrapAt = (top: number): HTMLElement => {
    const wrap = document.createElement('div')
    wrap.getBoundingClientRect = () => rectOf(top)
    document.body.appendChild(wrap)
    return wrap
  }
  const anchorAt = (top: number): { node: Text; charOffset: number } => {
    const p = document.createElement('p')
    // empty text node: anchorTop falls back to the parent element's rect (jsdom has no Range rects)
    const node = document.createTextNode('')
    p.appendChild(node)
    p.getBoundingClientRect = () => rectOf(top)
    document.body.appendChild(p)
    return { node, charOffset: 0 }
  }

  it('creates one marker per anchor at the zoom-normalized offset', () => {
    const wrap = wrapAt(100)
    syncCutOverlays(wrap, [anchorAt(300), anchorAt(500)], 2)
    const layer = wrap.querySelector(':scope > .page-cut-overlays') as HTMLElement
    expect(layer).not.toBeNull()
    const marks = layer.querySelectorAll('.page-gap-cut.page-cut-overlay')
    expect(marks.length).toBe(2)
    expect((marks[0] as HTMLElement).style.top).toBe('100px')
    expect((marks[1] as HTMLElement).style.top).toBe('200px')
  })

  it('rebuild replaces markers; empty anchors remove the layer', () => {
    const wrap = wrapAt(0)
    syncCutOverlays(wrap, [anchorAt(50), anchorAt(60), anchorAt(70)], 1)
    syncCutOverlays(wrap, [anchorAt(80)], 1)
    const layer = wrap.querySelector('.page-cut-overlays') as HTMLElement
    expect(layer.children.length).toBe(1)
    expect((layer.children[0] as HTMLElement).style.top).toBe('80px')
    syncCutOverlays(wrap, [], 1)
    expect(wrap.querySelector('.page-cut-overlays')).toBeNull()
  })
})

describe('createLineRectsCache', () => {
  it('memoizes per element within one pass', () => {
    const rectsOf = createLineRectsCache()
    const el = document.createElement('p')
    const a = rectsOf(el, 1)
    expect(rectsOf(el, 1)).toBe(a)
    const other = document.createElement('p')
    expect(rectsOf(other, 1)).not.toBe(a)
  })
})

describe('clampCellBoxTops', () => {
  const boxAt = (pm: HTMLElement, top: number, height = 45): HTMLElement => {
    const cell = document.createElement('div')
    cell.className = 'doc-cell-boxes'
    const box = document.createElement('div')
    box.className = 'doc-textbox'
    box.getBoundingClientRect = () => ({ top, bottom: top + height, height, width: 100 }) as DOMRect
    cell.appendChild(box)
    pm.appendChild(cell)
    return box
  }

  it('pushes a box lifted above the paper top back to the edge; leaves on-page boxes alone', () => {
    const pm = document.createElement('div')
    const above = boxAt(pm, -62)
    const inside = boxAt(pm, 30)
    clampCellBoxTops(pm, 0, 1)
    expect(above.style.getPropertyValue('--page-float-dy')).toBe('62.0px')
    expect(inside.style.getPropertyValue('--page-float-dy')).toBe('')
  })

  it('is idempotent: a re-run against the already-shifted rect keeps the same dy', () => {
    const pm = document.createElement('div')
    const box = boxAt(pm, -62)
    clampCellBoxTops(pm, 0, 1)
    // after the translate the live rect reads at the paper top
    box.getBoundingClientRect = () => ({ top: 0, bottom: 45, height: 45, width: 100 }) as DOMRect
    clampCellBoxTops(pm, 0, 1)
    expect(box.style.getPropertyValue('--page-float-dy')).toBe('62.0px')
  })
})

describe('alignTableGapFills', () => {
  /** paper at x=0, 816 wide; the gap cell follows its table's box */
  const pmWithFill = (cellLeft: number, cellWidth: number) => {
    const pm = document.createElement('div')
    pm.getBoundingClientRect = () => ({ left: 0, width: 816 }) as DOMRect
    const gap = makeGapEl(m, 'table', 2)
    const cell = gap.firstElementChild as HTMLElement
    cell.getBoundingClientRect = () => ({ left: cellLeft, width: cellWidth }) as DOMRect
    const fill = cell.firstElementChild as HTMLElement
    pm.appendChild(gap)
    return { pm, fill }
  }

  it('re-anchors the fill to the paper box (issue #174: indented full-width table)', () => {
    // table starts 32px right of the paper edge and spills past the right margin
    const { pm, fill } = pmWithFill(32, 794)
    alignTableGapFills(pm, 1)
    expect(fill.style.left).toBe('-32px')
    expect(fill.style.width).toBe('816px')
    expect(fill.style.right).toBe('auto')
  })

  it('widens a narrower-than-paper table band to the full paper width', () => {
    const { pm, fill } = pmWithFill(196, 424) // centered half-width table
    alignTableGapFills(pm, 1)
    expect(fill.style.left).toBe('-196px')
    expect(fill.style.width).toBe('816px')
  })

  it('divides zoomed rect deltas by the zoom factor', () => {
    const { pm, fill } = pmWithFill(64, 1588)
    pm.getBoundingClientRect = () => ({ left: 0, width: 1632 }) as DOMRect
    alignTableGapFills(pm, 2)
    expect(fill.style.left).toBe('-32px')
    expect(fill.style.width).toBe('816px')
  })

  it('is idempotent: the cell rect does not move with the fill styles', () => {
    const { pm, fill } = pmWithFill(32, 794)
    alignTableGapFills(pm, 1)
    alignTableGapFills(pm, 1)
    expect(fill.style.left).toBe('-32px')
    expect(fill.style.width).toBe('816px')
  })

  it('differing-width documents: the fill covers the table’s own centered page, not the paper', () => {
    // portrait page (816) centered on a landscape paper (1056): gap carries the page box
    const pm = document.createElement('div')
    pm.getBoundingClientRect = () => ({ left: 0, width: 1056 }) as DOMRect
    const gap = makeGapEl({ ...m, pageLeft: 120, pageWidth: 816 }, 'table', 2)
    const cell = gap.firstElementChild as HTMLElement
    cell.getBoundingClientRect = () => ({ left: 120 + 96, width: 624 }) as DOMRect
    pm.appendChild(gap)
    const fill = cell.firstElementChild as HTMLElement
    alignTableGapFills(pm, 1)
    expect(fill.style.left).toBe('-96px')
    expect(fill.style.width).toBe('816px')
  })
})

describe('page frames and sheets (differing-width documents, genoffice#246)', () => {
  const rect = (top: number, height: number, left = 0, width = 1056) =>
    ({ top, height, left, width, bottom: top + height, right: left + width }) as DOMRect
  /** wrap holding a paper and two gaps: page 0 (portrait) → page 1 (landscape) → page 2 (portrait) */
  const wrapEl = () => {
    const wrap = document.createElement('div')
    wrap.getBoundingClientRect = () => rect(0, 3000)
    const paper = document.createElement('div')
    paper.className = 'doc-page'
    paper.getBoundingClientRect = () => rect(0, 3000, 0, 1056)
    wrap.appendChild(paper)
    const g1 = makeGapEl({ ...m, pageLeft: 0, pageWidth: 1056 }, 'block')
    g1.getBoundingClientRect = () => rect(1000, 96 + GAP_BAND + 96)
    const g2 = makeGapEl({ ...m, pageLeft: 120, pageWidth: 816 }, 'block')
    g2.getBoundingClientRect = () => rect(2000, 96 + GAP_BAND + 96)
    paper.append(g1, g2)
    return wrap
  }

  it('derives each page’s box from the first page and the gaps’ next-page geometry', () => {
    const frames = pageFramesFromGaps(wrapEl(), 1, { left: 120, width: 816 })
    expect(frames).toHaveLength(3)
    expect(frames[0]).toEqual({ top: 0, bottom: 1096, left: 120, width: 816 })
    expect(frames[1]).toEqual({
      top: 1000 + 96 + GAP_BAND + 96 - 96,
      bottom: 2096,
      left: 0,
      width: 1056,
    })
    expect(frames[2]).toEqual({ top: 2000 + 96 + GAP_BAND, bottom: 3000, left: 120, width: 816 })
  })

  it('a boundary without page geometry (cut marker) keeps the page above it', () => {
    const wrap = wrapEl()
    const cut = makeGapEl(m, 'cut')
    cut.getBoundingClientRect = () => rect(1500, 0)
    wrap.firstElementChild!.appendChild(cut)
    const frames = pageFramesFromGaps(wrap, 1, { left: 120, width: 816 })
    expect(frames.map((f) => f.left)).toEqual([120, 0, 0, 120])
  })

  it('without a first frame the whole paper is the page', () => {
    const wrap = document.createElement('div')
    wrap.getBoundingClientRect = () => rect(0, 1056)
    const paper = document.createElement('div')
    paper.className = 'doc-page'
    paper.getBoundingClientRect = () => rect(0, 1056, 0, 816)
    wrap.appendChild(paper)
    expect(pageFramesFromGaps(wrap, 1)).toEqual([{ top: 0, bottom: 1056, left: 0, width: 816 }])
  })

  it('paints one sheet per page and clears the layer when the paper is uniform again', () => {
    const wrap = wrapEl()
    syncPageSheets(wrap, 1, { left: 120, width: 816 })
    const sheets = Array.from(wrap.querySelectorAll<HTMLElement>('.page-sheets > .page-sheet'))
    expect(sheets.map((s) => [s.style.left, s.style.width])).toEqual([
      ['120px', '816px'],
      ['0px', '1056px'],
      ['120px', '816px'],
    ])
    expect(sheets[0].style.top).toBe('0px')
    expect(sheets[0].style.height).toBe('1096px')
    syncPageSheets(wrap, 1, null)
    expect(wrap.querySelector('.page-sheets')).toBeNull()
  })
})

describe('alignGapHfStrips', () => {
  const STRIP_W = 600
  /** emulates layout: stylesheet centering (left:50% of a 786px gap + translateX(-50%))
   *  until inline left/transform pin the strip */
  const stripEl = (pm: HTMLElement): HTMLElement => {
    const el = document.createElement('div')
    el.className = 'page-gap-hf'
    el.getBoundingClientRect = () => {
      const left = el.style.left ? parseFloat(el.style.left) : 393
      const tx = el.style.transform === 'none' ? 0 : -STRIP_W / 2
      return { left: left + tx, width: STRIP_W } as DOMRect
    }
    pm.appendChild(el)
    return el
  }
  const pmEl = (): HTMLElement => {
    const pm = document.createElement('div')
    pm.getBoundingClientRect = () => ({ left: 0, width: 816 }) as DOMRect
    return pm
  }

  it('pins reused stylesheet-centered strips (left:50% + translateX(-50%)) before aligning', () => {
    const pm = pmEl()
    const strip = stripEl(pm)
    alignGapHfStrips(pm, 96, 1)
    expect(strip.style.transform).toBe('none')
    expect(strip.getBoundingClientRect().left).toBeCloseTo(96, 1)
  })

  it('is idempotent across widget reuse: a second pass leaves the pinned strip alone', () => {
    const pm = pmEl()
    const strip = stripEl(pm)
    alignGapHfStrips(pm, 96, 1)
    const left = strip.style.left
    alignGapHfStrips(pm, 96, 1)
    expect(strip.style.left).toBe(left)
    expect(strip.getBoundingClientRect().left).toBeCloseTo(96, 1)
  })

  it('re-anchors footnote areas and floating images from their gap’s origin (data-paper-x)', () => {
    // block gap starting 20px right of the paper edge (next section's narrower margin)
    const pm = pmEl()
    const gap = makeGapEl(
      { marginTop: 96, marginBottom: 96, marginLeft: 76, marginRight: 96 },
      'block',
    )
    gap.getBoundingClientRect = () => ({ left: 20, width: 796 }) as DOMRect
    pm.appendChild(gap)
    const notes = document.createElement('div')
    notes.className = 'page-gap-notes'
    notes.dataset.paperX = '216.0'
    notes.style.left = '216px'
    Object.defineProperty(notes, 'offsetParent', { get: () => gap })
    gap.appendChild(notes)
    alignGapHfStrips(pm, 96, 1)
    expect(notes.style.left).toBe('196px')
    alignGapHfStrips(pm, 96, 1)
    expect(notes.style.left).toBe('196px')
  })

  // A full-bleed cover section (w:pgMar left="0") ahead of body sections with real
  // margins: the canvas pads by the cover, so the body pages' strips must follow
  // their own section inset (--gap-ml), not the canvas target
  const COVER_ML = 0
  const BODY_ML = 113.4 // 1701 twips
  const INDENT = 48 // 720 twips
  const gapWithStrip = (
    pm: HTMLElement,
    metrics: Parameters<typeof makeGapEl>[0],
    kind: 'block' | 'inline',
  ) => {
    const gap = makeGapEl(metrics, kind)
    pm.appendChild(gap)
    const strip = stripEl(gap)
    return { gap, strip }
  }

  it('block gaps place the strips on the next section’s own left margin', () => {
    const pm = pmEl()
    const { strip } = gapWithStrip(
      pm,
      { marginTop: 96, marginBottom: 96, marginLeft: BODY_ML, marginRight: 94.5 },
      'block',
    )
    alignGapHfStrips(pm, COVER_ML, 1)
    expect(strip.getBoundingClientRect().left).toBeCloseTo(BODY_ML, 1)
  })

  it('mid-paragraph gaps in an indented paragraph keep the strips on the section margin, not the indent', () => {
    // inline gaps fold the host paragraph's paper offset (margin + indent) into
    // marginLeft so the gray band spans the paper; the strip must not inherit that
    const pm = pmEl()
    const { gap, strip } = gapWithStrip(
      pm,
      {
        marginTop: 96,
        marginBottom: 96,
        marginLeft: BODY_ML + INDENT,
        marginRight: 94.5,
        sectionMarginLeft: BODY_ML,
        sectionMarginRight: 94.5,
      },
      'inline',
    )
    expect(gap.style.marginLeft).toBe(`-${BODY_ML + INDENT}px`) // bleed unchanged
    expect(gap.style.getPropertyValue('--gap-ml')).toBe(`${BODY_ML}px`)
    alignGapHfStrips(pm, COVER_ML, 1)
    expect(strip.getBoundingClientRect().left).toBeCloseTo(BODY_ML, 1)
  })

  it('equal-margin landscape sections: the indent never leaks into the strip position', () => {
    // the pre-existing differing-width case (portrait + landscape, 1in margins):
    // strips at block and mid-paragraph gaps must agree
    const pm = pmEl()
    const ML = 96
    const block = gapWithStrip(
      pm,
      { marginTop: 96, marginBottom: 96, marginLeft: ML, marginRight: ML },
      'block',
    )
    const inline = gapWithStrip(
      pm,
      {
        marginTop: 96,
        marginBottom: 96,
        marginLeft: ML + INDENT,
        marginRight: ML,
        sectionMarginLeft: ML,
        sectionMarginRight: ML,
      },
      'inline',
    )
    alignGapHfStrips(pm, ML, 1)
    expect(block.strip.getBoundingClientRect().left).toBeCloseTo(ML, 1)
    expect(inline.strip.getBoundingClientRect().left).toBeCloseTo(ML, 1)
  })

  it('scales the section inset by the zoom factor', () => {
    // zoomed canvas: inline left is CSS px, the measured rect is screen px (×2)
    const pm = pmEl()
    const gap = makeGapEl(
      { marginTop: 96, marginBottom: 96, marginLeft: BODY_ML, marginRight: 94.5 },
      'block',
    )
    pm.appendChild(gap)
    const strip = document.createElement('div')
    strip.className = 'page-gap-hf'
    strip.getBoundingClientRect = () =>
      ({ left: (parseFloat(strip.style.left) || 0) * 2, width: STRIP_W * 2 }) as DOMRect
    strip.style.left = '0px'
    strip.style.transform = 'none'
    gap.appendChild(strip)
    alignGapHfStrips(pm, COVER_ML, 2)
    expect(parseFloat(strip.style.left)).toBeCloseTo(BODY_ML, 1)
    expect(strip.getBoundingClientRect().left).toBeCloseTo(BODY_ML * 2, 1)
  })

  it('section-break gaps: the outgoing footer keeps the previous section’s margin, the incoming header takes the next one’s', () => {
    // the gap between a full-bleed cover and the body hosts the cover's footer and
    // the body's header; only the header belongs on the body column
    const pm = pmEl()
    const gap = makeGapEl(
      { marginTop: 96, marginBottom: 96, marginLeft: BODY_ML, marginRight: 94.5 },
      'block',
    )
    pm.appendChild(gap)
    const footer = stripEl(gap)
    footer.style.setProperty('--hf-ml', `${COVER_ML}px`)
    const header = stripEl(gap)
    header.style.setProperty('--hf-ml', `${BODY_ML}px`)
    const legacy = stripEl(gap) // no --hf-ml: falls back to the gap's inset
    alignGapHfStrips(pm, COVER_ML, 1)
    expect(footer.getBoundingClientRect().left).toBeCloseTo(COVER_ML, 1)
    expect(header.getBoundingClientRect().left).toBeCloseTo(BODY_ML, 1)
    expect(legacy.getBoundingClientRect().left).toBeCloseTo(BODY_ML, 1)
  })

  it('table gap rows carry the section inset as well', () => {
    const tr = makeGapEl(
      {
        marginTop: 96,
        marginBottom: 96,
        marginLeft: BODY_ML,
        marginRight: 94.5,
        sectionMarginLeft: BODY_ML,
        sectionMarginRight: 94.5,
      },
      'table',
    )
    expect(tr.classList.contains('page-gap')).toBe(true)
    expect(tr.style.getPropertyValue('--gap-ml')).toBe(`${BODY_ML}px`)
    expect(tr.style.getPropertyValue('--gap-mr')).toBe('94.5px')
  })
})

describe('pageBorderStyleOf', () => {
  const margins = { marginTop: 1440, marginRight: 1080, marginBottom: 1440, marginLeft: 1080 }

  it('maps single lines to solid with the document color and page-edge inset', () => {
    const style = pageBorderStyleOf({
      ...margins,
      pageBorder: true,
      pageBorderProps: {
        offsetFrom: 'page',
        spacePt: 24,
        widthPt: 0.5,
        color: '80340D',
        sides: {
          top: { val: 'single', widthPt: 0.5, spacePt: 24, color: '80340D' },
          left: { val: 'single', widthPt: 0.5, spacePt: 24, color: '80340D' },
        },
      },
    })
    expect(style).not.toBeNull()
    expect(style!.sides.top).toEqual({ css: '1px solid #80340D', insetPx: 32 })
    expect(style!.sides.left).toEqual({ css: '1px solid #80340D', insetPx: 32 })
    expect(style!.sides.right).toBeUndefined()
    expect(style!.sides.bottom).toBeUndefined()
  })

  it('maps compound thin/thick lines to CSS double; an unflagged art val stays a solid line', () => {
    const style = pageBorderStyleOf({
      ...margins,
      pageBorder: true,
      pageBorderProps: {
        offsetFrom: 'page',
        spacePt: 24,
        widthPt: 3,
        sides: {
          top: { val: 'thinThickSmallGap', widthPt: 3, spacePt: 24 },
          bottom: { val: 'twistedLines1', widthPt: 2.25, spacePt: 24 },
        },
      },
    })
    expect(style!.sides.top!.css).toBe('8px double #000000')
    expect(style!.sides.bottom!.css).toBe('3px solid #000000')
  })

  describe('art borders', () => {
    const artSide = (val: string, color?: string) => ({
      val,
      widthPt: 16,
      spacePt: 24,
      art: true as const,
      ...(color ? { color } : {}),
    })
    const styleOf = (
      sides: NonNullable<Parameters<typeof pageBorderStyleOf>[0]['pageBorderProps']>['sides'],
    ) =>
      pageBorderStyleOf({
        ...margins,
        pageBorder: true,
        pageBorderProps: { offsetFrom: 'page', spacePt: 24, widthPt: 16, sides },
      })!

    it('stitch family: a 16pt-tall zigzag strip in the bitmap grey, tiled along the side', () => {
      const side = styleOf({ bottom: artSide('zigZagStitch') }).sides.bottom!
      expect(side.css).toBeUndefined()
      expect(side.insetPx).toBe(32)
      expect(side.art!.sizePx).toBe(21)
      expect(side.art!.backgroundRepeat).toBe('repeat-x')
      expect(side.art!.backgroundSize).toBe('21px 21px')
      const svg = decodeURIComponent(side.art!.backgroundImage)
      expect(svg).toContain('<polyline')
      expect(svg).toContain('stroke="#E2E2E2"')
      const strip = pageBorderArtStripStyle('bottom', side.art!)
      expect(strip).toMatchObject({ bottom: '0', left: '0', right: '0', height: '21px' })
      expect(strip.top).toBeUndefined()
    })

    it('pictogram families: dot rows in the main colour with the outline ink, vertical sides repeat-y', () => {
      const gems = styleOf({ left: artSide('gems') }).sides.left!
      expect(gems.art!.backgroundRepeat).toBe('repeat-y')
      const svg = decodeURIComponent(gems.art!.backgroundImage)
      expect(svg).toContain('<circle')
      expect(svg).toContain('fill="#BFBFBF"')
      expect(svg).toContain('stroke="#000000"')
      expect(pageBorderArtStripStyle('left', gems.art!)).toMatchObject({
        left: '0',
        top: '0',
        bottom: '0',
        width: '21px',
      })
      const cake = styleOf({ right: artSide('cakeSlice') }).sides.right!
      expect(decodeURIComponent(cake.art!.backgroundImage)).toContain('fill="#FEEDC9"')
      // an unknown pictogram name defaults to black dots
      const apples = styleOf({ top: artSide('apples') }).sides.top!
      expect(decodeURIComponent(apples.art!.backgroundImage)).toContain('fill="#000000"')
    })

    it('line-work families reuse CSS borders sized from the pattern height', () => {
      const sides = styleOf({
        top: artSide('handmade2'),
        right: artSide('whiteFlowers'),
        bottom: artSide('basicBlackDashes'),
        left: artSide('weavingAngles'),
      }).sides
      expect(sides.top).toEqual({ css: '11px double #000000', insetPx: 32 })
      expect(sides.right).toEqual({ css: '3px solid #000000', insetPx: 32 })
      expect(sides.bottom).toEqual({ css: '21px dashed #000000', insetPx: 32 })
      expect(sides.left).toEqual({ css: '11px double #000000', insetPx: 32 })
    })

    it('an explicit w:color recolours the pattern and drops the bitmap ink', () => {
      const side = styleOf({ top: artSide('gems', 'FF0000') }).sides.top!
      const svg = decodeURIComponent(side.art!.backgroundImage)
      expect(svg).toContain('fill="#FF0000"')
      expect(svg).not.toContain('stroke=')
      expect(styleOf({ top: artSide('handmade1', '1F497D') }).sides.top!.css).toBe(
        '11px double #1F497D',
      )
    })

    it('syncPageBorders draws the strip inside the box and keeps back borders out of the editor root', () => {
      const wrap = document.createElement('div')
      wrap.className = 'page-wrap'
      const paper = document.createElement('div')
      paper.className = 'doc-page'
      wrap.appendChild(paper)
      const rect = { top: 0, left: 0, width: 816, height: 1056, right: 816, bottom: 1056 }
      wrap.getBoundingClientRect = () => ({ ...rect, x: 0, y: 0, toJSON: () => rect })
      paper.getBoundingClientRect = wrap.getBoundingClientRect
      const style = styleOf({
        top: artSide('gems'),
        left: { val: 'single', widthPt: 1, spacePt: 24 },
      })
      syncPageBorders(wrap, style, 1)
      const box = wrap.querySelector('.page-border-overlays > .page-border-overlay') as HTMLElement
      expect(box.style.borderLeft).toBe('1px solid rgb(0, 0, 0)')
      expect(box.style.borderTop).toBe('')
      const strip = box.querySelector('.page-border-art') as HTMLElement
      expect(strip.style.height).toBe('21px')
      expect(strip.style.backgroundRepeat).toBe('repeat-x')
      syncPageBorders(wrap, { ...style, zOrder: 'back' }, 1)
      const layer = wrap.querySelector('.page-border-overlays') as HTMLElement
      expect(layer.parentElement).toBe(wrap)
      expect(paper.querySelector('.page-border-overlays')).toBeNull()
      expect(layer.classList.contains('page-border-overlays-back')).toBe(true)
      syncPageBorders(wrap, style, 1)
      expect(layer.classList.contains('page-border-overlays-back')).toBe(false)
      expect(wrap.querySelectorAll('.page-border-overlays').length).toBe(1)
      syncPageBorders(wrap, null, 1)
      expect(wrap.querySelector('.page-border-overlays')).toBeNull()
    })

    it('zOrder=back carries through to the style', () => {
      const style = pageBorderStyleOf({
        ...margins,
        pageBorder: true,
        pageBorderProps: {
          zOrder: 'back',
          spacePt: 24,
          widthPt: 16,
          sides: { top: artSide('stars') },
        },
      })
      expect(style!.zOrder).toBe('back')
      expect(styleOf({ top: artSide('stars') }).zOrder).toBeUndefined()
    })
  })

  it('offsetFrom=text measures the inset back from the margin edge', () => {
    const style = pageBorderStyleOf({
      ...margins,
      pageBorder: true,
      pageBorderProps: {
        offsetFrom: 'text',
        spacePt: 10,
        widthPt: 0.75,
        sides: { left: { val: 'single', widthPt: 0.75, spacePt: 10 } },
      },
    })
    // 1080 twips = 72px margin, minus 10pt (13.33px) space
    expect(style!.sides.left!.insetPx).toBeCloseTo(72 - (10 * 96) / 72, 5)
  })

  it('pageBorder without parse-side details falls back to a plain box', () => {
    const style = pageBorderStyleOf({ ...margins, pageBorder: true })
    for (const side of ['top', 'right', 'bottom', 'left'] as const) {
      expect(style!.sides[side]).toEqual({ css: '1px solid #000000', insetPx: 32 })
    }
  })

  it('display carries through and no border yields null', () => {
    const style = pageBorderStyleOf({
      ...margins,
      pageBorder: true,
      pageBorderProps: {
        display: 'firstPage',
        spacePt: 24,
        widthPt: 0.75,
        sides: { top: { val: 'single', widthPt: 0.75, spacePt: 24 } },
      },
    })
    expect(style!.display).toBe('firstPage')
    expect(pageBorderStyleOf({ ...margins, pageBorder: false })).toBeNull()
  })
  it('an auto-colored side stays black, never a sibling side color', () => {
    const style = pageBorderStyleOf({
      ...margins,
      pageBorder: true,
      pageBorderProps: {
        offsetFrom: 'page',
        spacePt: 24,
        widthPt: 0.75,
        color: 'FF0000',
        sides: {
          top: { val: 'single', widthPt: 0.75, spacePt: 24, color: 'FF0000' },
          bottom: { val: 'single', widthPt: 0.75, spacePt: 24 },
        },
      },
    })
    expect(style!.sides.top!.css).toBe('1px solid #FF0000')
    expect(style!.sides.bottom!.css).toBe('1px solid #000000')
  })
})

describe('rowFillAttrs', () => {
  it('measures the stored split extra against the rounded CSS height', () => {
    expect(rowFillAttrs(310.4, 10.4)).toEqual({ style: 'height:310px', 'data-split-extra': '10.0' })
    expect(rowFillAttrs(309.6, 9.6)).toEqual({ style: 'height:310px', 'data-split-extra': '10.0' })
    expect(rowFillAttrs(310.4)).toEqual({ style: 'height:310px' })
  })
})
