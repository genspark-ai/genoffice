// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'
import { anchorPageHoldsFloat, syncFloatShifts } from '../src/renderer/editor/pagination-gaps'
import type { PageCut } from '../src/renderer/pagination-cuts'

/**
 * Issue #2482: a Word cover anchors its title/subtitle/signature boxes to the
 * first paragraph, which ends in a section break. The page's flow ends after
 * that paragraph's 45px anchor line; a 1077px page-turn gap then opens before
 * page 2's body. The boxes hang below their anchor (offsets 458..811), still
 * on page 1 in Word.
 */
function coverCuts(): PageCut[] {
  const base = {
    section: 0,
    lead: { block: -1, kind: 'tail' },
    usedPx: 0,
    physUsedPx: 0,
    padPx: 0,
    hf: { variant: 'default', headerPx: 0, footerPx: 0 },
    leadDropsSpaceBefore: false,
    pageNo: 1,
    display: '1',
    blank: false,
    floats: [] as number[],
    gapOffsetPx: 0,
    flowOffsetPx: 0,
  }
  return [
    {
      ...base,
      start: 0,
      end: 45,
      capacityPx: 930,
      gap: {
        kind: 'block',
        heightPx: 1077,
        marginBottomPx: 96,
        marginTopPx: 96,
        marginLeftPx: 90,
        pullUpPx: 0,
        repeatHeaderPx: 0,
        flowTopPx: 45,
      },
    },
    {
      ...base,
      start: 1122,
      end: 1200,
      capacityPx: 930,
      pageNo: 2,
      display: '2',
      gapOffsetPx: 1077,
    },
  ] as unknown as PageCut[]
}

function rect(top: number, height: number): DOMRect {
  return {
    top,
    bottom: top + height,
    left: 0,
    right: 0,
    width: 0,
    height,
    x: 0,
    y: top,
    toJSON: () => ({}),
  } as DOMRect
}

function floatHost(): { pm: HTMLElement; box: HTMLElement } {
  document.body.innerHTML = ''
  const pm = document.createElement('div')
  const wrapper = document.createElement('div')
  wrapper.className = 'doc-protected-floating'
  const box = document.createElement('div')
  box.className = 'doc-textbox'
  wrapper.append(box)
  pm.append(wrapper)
  document.body.append(pm)
  return { pm, box }
}

const shiftOf = (box: HTMLElement): number => parseFloat(box.dataset.pageFloatDy ?? '0') || 0

describe('paragraph-anchored floats on their anchor page (#2482)', () => {
  it('a cover box hanging below a section-break paragraph stays on its anchor page', () => {
    const { pm, box } = floatHost()
    // the box paints 458.5px below the anchor line, still inside page 1's paper
    box.getBoundingClientRect = () => rect(458.5, 106)
    syncFloatShifts(
      pm,
      [
        {
          el: box,
          top: 458.5,
          height: 106,
          anchorTop: 0,
          pinned: false,
          pageRelV: false,
          noSpill: true,
        },
      ],
      0,
      1,
      0,
      { cuts: coverCuts(), carries: [] },
    )
    // the page-turn gap after the anchor must not drag the box onto page 2
    expect(shiftOf(box)).toBe(0)
  })

  it('a box hanging past the anchor page bottom keeps the flow placement', () => {
    const { pm, box } = floatHost()
    // 900 + 200 = 1100 past page 1's 930px content window: not held
    box.getBoundingClientRect = () => rect(900, 200)
    syncFloatShifts(
      pm,
      [
        {
          el: box,
          top: 900,
          height: 200,
          anchorTop: 0,
          pinned: false,
          pageRelV: false,
          noSpill: true,
        },
      ],
      0,
      1,
      0,
      { cuts: coverCuts(), carries: [] },
    )
    expect(shiftOf(box)).toBe(1077)
  })

  it('a box above the anchor page start is held too (top-margin cover art)', () => {
    const { pm, box } = floatHost()
    // the sample's behind image: 1in above the paragraph top, still on page 1
    box.getBoundingClientRect = () => rect(-95.7, 532)
    syncFloatShifts(
      pm,
      [
        {
          el: box,
          top: -95.7,
          height: 532,
          anchorTop: 0,
          pinned: false,
          pageRelV: false,
          noSpill: true,
        },
      ],
      0,
      1,
      0,
      { cuts: coverCuts(), carries: [] },
    )
    expect(shiftOf(box)).toBe(0)
  })

  it('the helper reports which boxes the anchor page can hold', () => {
    const cuts = coverCuts()
    expect(anchorPageHoldsFloat({ top: 458.5, height: 106, anchorTop: 0 }, cuts)).toBe(true)
    expect(anchorPageHoldsFloat({ top: 900, height: 200, anchorTop: 0 }, cuts)).toBe(false)
    // the second page's window: flow start 1122 + 930px capacity
    expect(anchorPageHoldsFloat({ top: 1122 + 458, height: 106, anchorTop: 1122 }, cuts)).toBe(true)
    expect(anchorPageHoldsFloat({ top: 1122 + 900, height: 200, anchorTop: 1122 }, cuts)).toBe(
      false,
    )
  })

  it('a flow-positioned float still rides the page gap', () => {
    const { pm, box } = floatHost()
    box.getBoundingClientRect = () => rect(458.5, 106)
    syncFloatShifts(
      pm,
      [{ el: box, top: 458.5, height: 106, pinned: false, pageRelV: false }],
      0,
      1,
      0,
      {
        cuts: coverCuts(),
        carries: [],
      },
    )
    expect(shiftOf(box)).toBe(1077)
  })
})
