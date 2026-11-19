import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { renderTextboxSpec } from '../src/renderer/editor/protected-render'
import { syncFloatShifts } from '../src/renderer/editor/pagination-gaps'
import { applyGapLedger, buildPageCuts, learnGapFlowTops } from '../src/renderer/pagination-cuts'

const box = {
  paras: [{ runs: [{ text: 'band' }] }],
  floating: true,
  pageRelV: true,
  offsetXEmu: 28575,
  offsetYEmu: 100000,
} as unknown as Parameters<typeof renderTextboxSpec>[0]

describe('page-edge V anchors ignore the header push of the body top (Word)', () => {
  it('stamps data-page-rel-from only for page anchors', () => {
    const page = renderTextboxSpec({ ...box, pageRelVFrom: 'page' })
    expect(page[1]['data-page-rel-v']).toBe('1')
    expect(page[1]['data-page-rel-from']).toBe('page')
    const margin = renderTextboxSpec({ ...box, pageRelVFrom: 'margin' })
    expect(margin[1]['data-page-rel-v']).toBe('1')
    expect(margin[1]['data-page-rel-from']).toBeUndefined()
  })

  it('preview CSS re-pins page anchors with the pgMar top, margin anchors with the effective top', () => {
    const css = readFileSync(
      join(dirname(fileURLToPath(import.meta.url)), '../src/renderer/styles.css'),
      'utf8',
    )
    expect(css).toMatch(
      /\[data-page-rel-from='page'\][^}]*translate: var\(--pv-ml, 0px\) var\(--pv-mt-page, var\(--pv-mt, 0px\)\)/,
    )
  })

  it('syncFloatShifts subtracts the first page push for page anchors only', () => {
    const pm = document.createElement('div')
    document.body.appendChild(pm)
    const mk = () => {
      const el = document.createElement('div')
      el.getBoundingClientRect = () => ({ top: 0 }) as DOMRect
      pm.appendChild(el)
      return el
    }
    const pageEl = mk()
    const marginEl = mk()
    // anchor at 0, box offset 100 (content coords), body pushed 26px by the header
    const base = { top: 100, height: 20, anchorTop: 0, pinned: false, pageRelV: true }
    syncFloatShifts(
      pm,
      [
        { ...base, el: pageEl, pageRelFromPage: true },
        { ...base, el: marginEl },
      ],
      0,
      1,
      26,
    )
    expect(parseFloat(pageEl.style.getPropertyValue('--page-float-dy'))).toBeCloseTo(74, 1)
    expect(parseFloat(marginEl.style.getPropertyValue('--page-float-dy'))).toBeCloseTo(100, 1)
    pm.remove()
  })

  it('a float host or repeated header at the page start does not clear the page push', () => {
    const pm = document.createElement('div')
    document.body.appendChild(pm)
    for (const cls of ['page-float-host', 'page-repeat-header']) {
      const w = document.createElement('div')
      w.className = cls
      w.getBoundingClientRect = () => ({ top: 0, height: 0 }) as DOMRect
      pm.appendChild(w)
    }
    const pageEl = document.createElement('div')
    pageEl.getBoundingClientRect = () => ({ top: 0 }) as DOMRect
    pm.appendChild(pageEl)
    syncFloatShifts(
      pm,
      [
        {
          top: 100,
          anchorTop: 0,
          pinned: false,
          pageRelV: true,
          pageRelFromPage: true,
          el: pageEl,
        },
      ],
      0,
      1,
      26,
    )
    expect(parseFloat(pageEl.style.getPropertyValue('--page-float-dy'))).toBeCloseTo(74, 1)
    pm.remove()
  })

  it('a later page gap with a zero push clears the earlier page push', () => {
    const pm = document.createElement('div')
    document.body.appendChild(pm)
    const pageEl = document.createElement('div')
    pageEl.getBoundingClientRect = () => ({ top: 0 }) as DOMRect
    pm.appendChild(pageEl)
    const cuts = buildPageCuts(
      [
        { top: 0, height: 500, el: document.createElement('p') },
        { top: 500, height: 400, el: document.createElement('p') },
      ],
      [
        { start: 0, end: 500, section: 0 },
        { start: 500, end: 900, section: 0 },
      ],
      [{ contentHeight: 500, forceBreak: false }],
      [{ headerPx: 0, footerPx: 0 }],
    )
    applyGapLedger(cuts, [
      {
        page: 1,
        kind: 'block',
        heightPx: 40,
        marginBottomPx: 6,
        marginTopPx: 6,
        marginLeftPx: 0,
        topPushPx: 0,
        pullUpPx: 0,
        repeatHeaderPx: 0,
      },
    ])
    // anchor on page 2 (virtual 600), box 100 below it; first page was pushed 26
    syncFloatShifts(
      pm,
      [
        {
          top: 700,
          anchorTop: 600,
          pinned: false,
          pageRelV: true,
          pageRelFromPage: true,
          el: pageEl,
        },
      ],
      0,
      1,
      26,
      { cuts, carries: [] },
    )
    // page start 500 + rel 100 + gap height 40, no push subtracted
    expect(parseFloat(pageEl.style.getPropertyValue('--page-float-dy'))).toBeCloseTo(640, 1)
    pm.remove()
  })

  it('an anchor leading the page the widget opens still counts that gap once the widget top is learned', () => {
    const pm = document.createElement('div')
    document.body.appendChild(pm)
    const pageEl = document.createElement('div')
    pageEl.getBoundingClientRect = () => ({ top: 0 }) as DOMRect
    pm.appendChild(pageEl)
    const cuts = buildPageCuts(
      [
        { top: 0, height: 500, el: document.createElement('p') },
        { top: 510, height: 400, el: document.createElement('p') },
      ],
      [
        { start: 0, end: 500, section: 0 },
        { start: 500, end: 910, section: 0 },
      ],
      [{ contentHeight: 500, forceBreak: false }],
      [{ headerPx: 0, footerPx: 0 }],
    )
    applyGapLedger(cuts, [
      {
        page: 1,
        kind: 'block',
        heightPx: 40,
        marginBottomPx: 6,
        marginTopPx: 6,
        marginLeftPx: 0,
        topPushPx: 0,
        pullUpPx: 0,
        repeatHeaderPx: 0,
      },
    ])
    // the widget landed 10 px under the boundary (the previous block's space
    // after): the lead block's virtual top is the widget's bottom edge
    learnGapFlowTops(cuts, [510])
    syncFloatShifts(
      pm,
      [{ top: 610, anchorTop: 510, pinned: false, pageRelV: true, el: pageEl }],
      0,
      1,
      0,
      { cuts, carries: [] },
    )
    expect(parseFloat(pageEl.style.getPropertyValue('--page-float-dy'))).toBeCloseTo(640, 1)
    pm.remove()
  })
})
