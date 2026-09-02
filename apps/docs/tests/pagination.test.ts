import { describe, expect, it } from 'vitest'
import type { SectionInfo } from '@genoffice/docx-engine'
import {
  type SliceOutputs,
  appendEndnotesBlock,
  assignSections,
  cellCutYs,
  columnLayoutSpecs,
  vAlignShiftSpecs,
  sectionWidthSpecs,
  sectionGridPitchSpecs,
  docGridPitchPt,
  sectionCharSpaceSpecs,
  docCharSpacePt,
  sectionColGeom,
  computePageSlices,
  computeSectionedSlices,
  computeSectionedSlicesF2,
  effectiveHfRefs,
  hasPrintableHeaderFooter,
  insertParityBlanks,
  lineBreakBoundaries,
  lineStartAnchor,
  nextLineAnchor,
  anchorElement,
  pageAt,
  visiblePageCount,
  liveSections,
  pageNumbers,
  pageStartBlocks,
  applyBlockMeta,
  fillLineBoxes,
  measureBlocks,
  sectionFirstPages,
  sectionGeoms,
  sectionPageBox,
  tableHeaderFlags,
  tableRowFlags,
  type BlockBox,
  type PageSlice,
  type SectionGeom,
  type TableRowBox,
} from '../src/renderer/pagination'

const block = (top: number, height: number, extra?: Partial<BlockBox>): BlockBox => ({
  top,
  height,
  ...extra,
})

describe('computePageSlices', () => {
  it('empty document yields one page', () => {
    expect(computePageSlices([], 800, 0)).toEqual([{ start: 0, end: 0, section: 0 }])
  })

  it('content shorter than a page does not break', () => {
    const slices = computePageSlices([block(0, 100), block(100, 200)], 800, 300)
    expect(slices).toEqual([{ start: 0, end: 300, section: 0 }])
  })

  it('block spanning pages is pushed whole to the next page', () => {
    // 2nd block 700→900 crosses the 800 boundary and fits on one page → new page starts at 700
    const slices = computePageSlices([block(0, 700), block(700, 200)], 800, 900)
    expect(slices).toEqual([
      { start: 0, end: 700, section: 0 },
      { start: 700, end: 900, section: 0 },
    ])
  })

  it('block exactly at the page boundary does not break', () => {
    const slices = computePageSlices([block(0, 300), block(300, 500)], 800, 800)
    expect(slices).toEqual([{ start: 0, end: 800, section: 0 }])
  })

  it('block taller than a page is hard-split by pixel', () => {
    const slices = computePageSlices([block(0, 100), block(100, 1900)], 800, 2000)
    expect(slices).toEqual([
      { start: 0, end: 800, section: 0 },
      { start: 800, end: 1600, section: 0 },
      { start: 1600, end: 2000, section: 0 },
    ])
  })

  it('oversized block with lineOffsets splits at line boundaries', () => {
    // Big block 100→1990, line height 90 (boundaries 90,180,...): line boundaries before page limits 800/1600 are 730/1450
    const offsets = Array.from({ length: 20 }, (_, i) => (i + 1) * 90)
    const slices = computePageSlices(
      [block(0, 100), block(100, 1890, { lineOffsets: offsets })],
      800,
      1990,
    )
    expect(slices).toEqual([
      { start: 0, end: 730, section: 0 },
      { start: 730, end: 1450, section: 0 },
      { start: 1450, end: 1990, section: 0 },
    ])
  })

  it('line boundary exactly at the page boundary does not split early', () => {
    const offsets = Array.from({ length: 18 }, (_, i) => (i + 1) * 100)
    const slices = computePageSlices(
      [block(0, 100), block(100, 1900, { lineOffsets: offsets })],
      800,
      2000,
    )
    expect(slices).toEqual([
      { start: 0, end: 800, section: 0 },
      { start: 800, end: 1600, section: 0 },
      { start: 1600, end: 2000, section: 0 },
    ])
  })

  it('no line boundary within the page (single line taller than the page) falls back to pixel split to guarantee progress', () => {
    const slices = computePageSlices([block(0, 1000, { lineOffsets: [900] })], 800, 1000)
    expect(slices).toEqual([
      { start: 0, end: 800, section: 0 },
      { start: 800, end: 1000, section: 0 },
    ])
  })

  it('spanning block with line boundaries splits in place (not pushed whole to the next page)', () => {
    // Block 600→1000 (height 400, fits on one page): has line boundaries → cut at 790, the last boundary before page limit 800
    const slices = computePageSlices(
      [block(0, 600), block(600, 400, { lineOffsets: [40, 90, 140, 190, 240, 290, 340] })],
      800,
      1000,
    )
    expect(slices).toEqual([
      { start: 0, end: 790, section: 0 },
      { start: 790, end: 1000, section: 0 },
    ])
  })

  it('orphan/widow lines: push the whole block when the head cannot fit splitMinLines lines', () => {
    // Block 750→1050, line height 30: only 1 line fits before page limit 800, minLines=2 → push whole block
    const slices = computePageSlices(
      [
        block(0, 750),
        block(750, 300, {
          lineOffsets: [30, 60, 90, 120, 150, 180, 210, 240, 270],
          splitMinLines: 2,
        }),
      ],
      800,
      1050,
    )
    expect(slices).toEqual([
      { start: 0, end: 750, section: 0 },
      { start: 750, end: 1050, section: 0 },
    ])
  })

  it('orphan/widow lines: split one line earlier when the tail has fewer than splitMinLines lines', () => {
    // Block 700→1000, line height 30 (10 lines): page limit 800 is exactly a line boundary, but the tail after the cut is too short…
    // After boundary 90 (y=790) the tail has 7 lines ✓; boundary 240 (y=940) exceeds the page; optimum y=790
    // Construct a case where the tail constraint kicks in: block 700→820, 4 lines (30): boundary before limit 800 is y=790 (k=2, tail 1 line < 2)
    // → back off to k=1 (y=760, tail 2 lines)
    const slices = computePageSlices(
      [block(0, 700), block(700, 120, { lineOffsets: [30, 60, 90], splitMinLines: 2 })],
      800,
      820,
    )
    expect(slices).toEqual([
      { start: 0, end: 760, section: 0 },
      { start: 760, end: 820, section: 0 },
    ])
  })

  it('table rows (splitMinLines defaults to 1) may split right before the last row', () => {
    const slices = computePageSlices(
      [block(0, 700), block(700, 200, { lineOffsets: [50, 100, 150] })],
      800,
      900,
    )
    expect(slices).toEqual([
      { start: 0, end: 800, section: 0 },
      { start: 800, end: 900, section: 0 },
    ])
  })

  it('breakBefore forces a break before the block', () => {
    const slices = computePageSlices(
      [block(0, 100), block(100, 100, { breakBefore: true })],
      800,
      200,
    )
    expect(slices).toEqual([
      { start: 0, end: 100, section: 0 },
      { start: 100, end: 200, section: 0 },
    ])
  })

  it('breakBefore at the top of a page does not create an empty page', () => {
    const slices = computePageSlices([block(0, 100, { breakBefore: true })], 800, 100)
    expect(slices).toEqual([{ start: 0, end: 100, section: 0 }])
  })

  it('breakAfter (page-break field) breaks after the block', () => {
    const slices = computePageSlices(
      [block(0, 100, { breakAfter: true }), block(100, 100), block(200, 100)],
      800,
      300,
    )
    expect(slices).toEqual([
      { start: 0, end: 100, section: 0 },
      { start: 100, end: 300, section: 0 },
    ])
  })

  it('breakAfter at the end of the document keeps its deliberate blank page (tdf#99090)', () => {
    const slices = computePageSlices([block(0, 100, { breakAfter: true })], 800, 100)
    expect(slices).toEqual([
      { start: 0, end: 100, section: 0 },
      { start: 100, end: 100, section: 0 },
    ])
  })

  it('push and hard split combined: normal spanning block before a huge block', () => {
    // Block A 0→600; block B 600→800 fits; block C 800→2500 starts on a new page then gets hard-cut
    const slices = computePageSlices(
      [block(0, 600), block(600, 200), block(800, 1700, { breakBefore: true })],
      800,
      2500,
    )
    expect(slices).toEqual([
      { start: 0, end: 800, section: 0 },
      { start: 800, end: 1600, section: 0 },
      { start: 1600, end: 2400, section: 0 },
      { start: 2400, end: 2500, section: 0 },
    ])
  })

  it('invalid contentHeight falls back to a single page', () => {
    expect(computePageSlices([block(0, 100)], 0, 100)).toEqual([{ start: 0, end: 100, section: 0 }])
  })
})

describe('lineBreakBoundaries', () => {
  const ln = (offset: number, bottom: number) => ({ offset, bottom })

  it('ignores the first glyph line and cuts midway through each ink gap', () => {
    // ink gap of 6px on each break: the boundary sits 3px above the next ink top
    expect(lineBreakBoundaries([ln(3.25, 16.5), ln(22.5, 35.75), ln(41.75, 55)])).toEqual([
      19.5, 38.75,
    ])
  })

  it('returns no boundary for a single visual line', () => {
    expect(lineBreakBoundaries([ln(2.75, 16)])).toEqual([])
  })

  it('overlapping ink keeps the ink-top cut (never below the next line start)', () => {
    expect(lineBreakBoundaries([ln(0, 24), ln(22, 46)])).toEqual([22])
  })
})

describe('pageAt', () => {
  const slices = [
    { start: 0, end: 800, section: 0 },
    { start: 800, end: 1600, section: 0 },
    { start: 1600, end: 2000, section: 0 },
  ]

  it('locates the page number by content Y', () => {
    expect(pageAt(slices, 0)).toBe(1)
    expect(pageAt(slices, 799)).toBe(1)
    expect(pageAt(slices, 800)).toBe(2)
    expect(pageAt(slices, 1999)).toBe(3)
  })

  it('out-of-range values clamp to the first/last page', () => {
    expect(pageAt(slices, -50)).toBe(1)
    expect(pageAt(slices, 99999)).toBe(3)
    expect(pageAt([], 100)).toBe(1)
  })
})

describe('visiblePageCount', () => {
  // insertParityBlanks puts the zero-height blank before the real page, sharing its start
  const withBlank = [
    { start: 0, end: 800, section: 0 },
    { start: 800, end: 800, section: 0 },
    { start: 800, end: 1600, section: 1 },
    { start: 1600, end: 2000, section: 1 },
  ]

  it('counts a zero-height blank as its own page (parity/deliberate blanks are drawn sheets)', () => {
    expect(visiblePageCount(withBlank)).toBe(4)
    expect(visiblePageCount([{ start: 0, end: 800, section: 0 }])).toBe(1)
    expect(visiblePageCount([])).toBe(0)
  })

  it('maps a physical pageAt index to its visible page number', () => {
    // y=900 lands on physical slice 3 (the real page after the blank) = visible page 3
    expect(visiblePageCount(withBlank, pageAt(withBlank, 900))).toBe(3)
    expect(visiblePageCount(withBlank, pageAt(withBlank, 0))).toBe(1)
    expect(visiblePageCount(withBlank, pageAt(withBlank, 1700))).toBe(4)
  })
})

describe('pageStartBlocks', () => {
  it('returns the index of the first block on each non-first page', () => {
    const blocks = [block(0, 700), block(700, 200), block(900, 100)]
    const slices = computePageSlices(blocks, 800, 1000)
    expect(slices.length).toBe(2)
    expect(pageStartBlocks(blocks, slices)).toEqual([1])
  })

  it('pixel hard-split boundary (inside a huge block) has no matching block and is skipped', () => {
    const blocks = [block(0, 100), block(100, 1900)]
    const slices = computePageSlices(blocks, 800, 2000)
    expect(slices.length).toBe(3)
    // 800/1600 both fall inside the big block; no block starts there
    expect(pageStartBlocks(blocks, slices)).toEqual([])
  })

  it('single page has no boundaries', () => {
    const blocks = [block(0, 100)]
    expect(pageStartBlocks(blocks, computePageSlices(blocks, 800, 100))).toEqual([])
  })
})

const sec = (
  over: Partial<SectionInfo['settings']>,
  extra?: Partial<SectionInfo>,
): SectionInfo => ({
  settings: {
    pageWidth: 11906,
    pageHeight: 16838,
    orientation: 'portrait',
    marginTop: 1440,
    marginRight: 1440,
    marginBottom: 1440,
    marginLeft: 1440,
    pageBorder: false,
    columns: 1,
    ...over,
  },
  startType: 'nextPage',
  firstBlockIndex: 0,
  lastBlockIndex: 0,
  sectPrXml: '',
  titlePg: false,
  headerRefs: {},
  footerRefs: {},
  ...extra,
})

describe('multi-section slicing', () => {
  it('section switch forces a page break; each section uses its own content height', () => {
    // Section 0 content height 800, section 1 content height 400
    const geoms = [
      { contentHeight: 800, forceBreak: false },
      { contentHeight: 400, forceBreak: true },
    ]
    const blocks = [
      block(0, 300, { section: 0 }),
      block(300, 100, { section: 0 }),
      block(400, 500, { section: 1 }), // Section 1 start: forced new page at 400, and hard-cut once past its 400 height
    ]
    expect(computeSectionedSlices(blocks, geoms, 900)).toEqual([
      { start: 0, end: 400, section: 0 },
      { start: 400, end: 800, section: 1 },
      { start: 800, end: 900, section: 1 },
    ])
  })

  it('continuous section does not force a page break', () => {
    const geoms = [
      { contentHeight: 800, forceBreak: false },
      { contentHeight: 800, forceBreak: false },
    ]
    const blocks = [block(0, 300, { section: 0 }), block(300, 100, { section: 1 })]
    expect(computeSectionedSlices(blocks, geoms, 400)).toEqual([{ start: 0, end: 400, section: 0 }])
  })

  it('a leading block-less section keeps its own blank first page (lone sectPr paragraph)', () => {
    // section 0 = a lone sectPr paragraph rendered as a zero-height chip: no
    // measured blocks, but Word still shows its blank page (tdf128156)
    const geoms = [
      { contentHeight: 800, forceBreak: true },
      { contentHeight: 800, forceBreak: true },
    ]
    const blocks = [block(0, 30, { section: 1 }), block(30, 30, { section: 1 })]
    expect(computeSectionedSlices(blocks, geoms, 60)).toEqual([
      { start: 0, end: 0, section: 0 },
      { start: 0, end: 60, section: 1 },
    ])
  })

  it('sectionGeoms: continuous with identical geometry keeps the flow; differing geometry promotes to a page break', () => {
    const a = sec({})
    const cont = sec({}, { startType: 'continuous' })
    const contWide = sec({ pageWidth: 16838, pageHeight: 11906 }, { startType: 'continuous' })
    const next = sec({})
    const geoms = sectionGeoms([a, cont, contWide, next])
    expect(geoms.map((g) => g.forceBreak)).toEqual([false, false, true, true])
    expect(Math.round(geoms[0].contentHeight)).toBe(Math.round(((16838 - 2880) / 1440) * 96))
  })

  it('sectionGeoms: contentWidth follows each section page width and side margins', () => {
    const portrait = sec({})
    const landscape = sec({ pageWidth: 16838, pageHeight: 11906, orientation: 'landscape' })
    const narrowMargins = sec({ marginLeft: 720, marginRight: 720 })
    const geoms = sectionGeoms([portrait, landscape, narrowMargins])
    expect(geoms[0].contentWidth).toBeCloseTo(((11906 - 2880) / 1440) * 96, 1)
    expect(geoms[1].contentWidth).toBeCloseTo(((16838 - 2880) / 1440) * 96, 1)
    expect(geoms[2].contentWidth).toBeCloseTo(((11906 - 1440) / 1440) * 96, 1)
  })

  it('sectionGeoms: nextColumn starts a new page in a single-column layout, flows in a multi-column one (n750255)', () => {
    const single = sec({})
    const afterSingle = sec({}, { startType: 'nextColumn' })
    const twoCol = sec({ columns: 2 }, { startType: 'nextColumn' })
    const afterTwoCol = sec({ columns: 2 }, { startType: 'nextColumn' })
    const geoms = sectionGeoms([single, afterSingle, twoCol, afterTwoCol])
    expect(geoms.map((g) => g.forceBreak)).toEqual([false, true, true, false])
  })

  it('effectiveHfRefs: undefined variants inherit forward section by section', () => {
    const sections = [
      sec({}, { headerRefs: { default: 'rH1', first: 'rHF1' }, footerRefs: { default: 'rF1' } }),
      sec({}, { headerRefs: { default: 'rH2' }, footerRefs: {} }),
      sec({}, { headerRefs: {}, footerRefs: { default: 'rF3' } }),
    ]
    const eff = effectiveHfRefs(sections)
    expect(eff[0].header).toEqual({ default: 'rH1', first: 'rHF1' })
    expect(eff[1].header).toEqual({ default: 'rH2', first: 'rHF1' })
    expect(eff[1].footer).toEqual({ default: 'rF1' })
    expect(eff[2].header).toEqual({ default: 'rH2', first: 'rHF1' })
    expect(eff[2].footer).toEqual({ default: 'rF3' })
  })

  it('hasPrintableHeaderFooter: empty header/footer does not force; non-empty (including inherited refs/variants/images) forces', () => {
    const empty = { text: ' ', hasPageNumber: false, paras: [] }
    const filled = { text: 'Confidential', hasPageNumber: false, paras: [] }
    const pageNum = { text: '', hasPageNumber: true, paras: [] }
    const imageOnly = { text: '', hasPageNumber: false, paras: [], images: [{ dataUrl: 'd' }] }

    expect(hasPrintableHeaderFooter({ edited: [null], sections: [sec({})] })).toBe(false)
    expect(hasPrintableHeaderFooter({ edited: [{ text: '  ' }], sections: [] })).toBe(false)
    expect(
      hasPrintableHeaderFooter({ edited: [{ text: '', pageNumber: true }], sections: [] }),
    ).toBe(true)
    expect(hasPrintableHeaderFooter({ edited: [{ text: 'Chapter 1' }], sections: [] })).toBe(true)
    expect(
      hasPrintableHeaderFooter({
        edited: [{ text: '', paras: [{ runs: [{ text: 'footer' }] }] }],
        sections: [],
      }),
    ).toBe(true)

    const refs = [sec({}, { headerRefs: { default: 'rH' }, footerRefs: {} })]
    expect(hasPrintableHeaderFooter({ edited: [], sections: refs, hfParts: { rH: empty } })).toBe(
      false,
    )
    expect(hasPrintableHeaderFooter({ edited: [], sections: refs, hfParts: { rH: filled } })).toBe(
      true,
    )
    expect(
      hasPrintableHeaderFooter({ edited: [], sections: refs, hfParts: { rH: imageOnly } }),
    ).toBe(true)

    // a later section inherits the previous section's refs
    const inherit = [sec({}, { footerRefs: { default: 'rF' } }), sec({})]
    expect(
      hasPrintableHeaderFooter({ edited: [], sections: inherit, hfParts: { rF: pageNum } }),
    ).toBe(true)

    // first/even variants only count when titlePg / evenOddHf is enabled
    const first = (titlePg: boolean) => [sec({}, { titlePg, headerRefs: { first: 'rHF' } })]
    expect(
      hasPrintableHeaderFooter({ edited: [], sections: first(false), hfParts: { rHF: filled } }),
    ).toBe(false)
    expect(
      hasPrintableHeaderFooter({ edited: [], sections: first(true), hfParts: { rHF: filled } }),
    ).toBe(true)
    const even = [sec({}, { headerRefs: { even: 'rHE' } })]
    expect(hasPrintableHeaderFooter({ edited: [], sections: even, hfParts: { rHE: filled } })).toBe(
      false,
    )
    expect(
      hasPrintableHeaderFooter({
        edited: [],
        sections: even,
        hfParts: { rHE: filled },
        evenOddHf: true,
      }),
    ).toBe(true)
  })

  it('pageNumbers: pgNumType w:start restarts numbering, otherwise continuous', () => {
    const sections = [sec({}), sec({}, { pageNumberStart: 1 }), sec({})]
    const slices = [
      { start: 0, end: 1, section: 0 },
      { start: 1, end: 2, section: 0 },
      { start: 2, end: 3, section: 1 }, // renumbered from 1
      { start: 3, end: 4, section: 1 },
      { start: 4, end: 5, section: 2 }, // no start: continues from previous page
    ]
    expect(pageNumbers(slices, sections)).toEqual([1, 2, 1, 2, 3])
    expect(sectionFirstPages(slices)).toEqual([true, false, true, false, true])
  })

  it('pageNumbers: evenPage/oddPage section breaks skip numbers to restore parity', () => {
    const sections = [
      sec({}),
      sec({}, { startType: 'evenPage' }),
      sec({}, { startType: 'oddPage' }),
    ]
    const slices = [
      { start: 0, end: 1, section: 0 },
      { start: 1, end: 2, section: 1 }, // continued number is 2, already even, no skip
      { start: 2, end: 3, section: 1 },
      { start: 3, end: 4, section: 2 }, // continued number is 4, even → skip to 5 (odd page)
    ]
    expect(pageNumbers(slices, sections)).toEqual([1, 2, 3, 5])
    // Explicit renumbering takes precedence over parity padding
    const withStart = [sec({}), sec({}, { startType: 'evenPage', pageNumberStart: 7 })]
    const slices2 = [
      { start: 0, end: 1, section: 0 },
      { start: 1, end: 2, section: 1 },
    ]
    expect(pageNumbers(slices2, withStart)).toEqual([1, 7])
  })

  it('liveSections: deleting a section-break block merges that section into the next one live', () => {
    const sections = [
      sec({}, { firstBlockIndex: 0, lastBlockIndex: 2 }),
      sec({}, { firstBlockIndex: 3, lastBlockIndex: 9 }),
    ]
    const allPresent = [
      block(0, 10, { docxIndex: 0 }),
      block(10, 10, { docxIndex: 2 }),
      block(20, 10, { docxIndex: 5 }),
    ]
    expect(liveSections(sections, allPresent)).toBe(sections)
    // Section-break block (docxIndex 2) gone from canvas → section 1 merges into section 2 (using section 2's settings)
    const deleted = [block(0, 10, { docxIndex: 0 }), block(20, 10, { docxIndex: 5 })]
    const merged = liveSections(sections, deleted)
    expect(merged.length).toBe(1)
    expect(merged[0].firstBlockIndex).toBe(0)
    expect(merged[0].lastBlockIndex).toBe(9)
    // The last section's "section break" is a hidden block, excluded from the check: deleting only body blocks does not merge
    const bodyDeleted = [block(10, 10, { docxIndex: 2 })]
    expect(liveSections(sections, bodyDeleted).length).toBe(2)
  })

  it('assignSections: assigns sections by docxIndex; new blocks inherit from the previous block', () => {
    const sections = [
      sec({}, { lastBlockIndex: 1 }),
      sec({}, { firstBlockIndex: 2, lastBlockIndex: 9 }),
    ]
    const blocks = [
      block(0, 10, { docxIndex: 0 }),
      block(10, 10, { docxIndex: 1 }),
      block(20, 10, {}), // new block (no docxIndex): inherit section 0? previous block is section 0's break → inherits 0
      block(30, 10, { docxIndex: 2 }),
      block(40, 10, { docxIndex: 99 }), // out of range, clamped to the last section
    ]
    assignSections(blocks, sections)
    expect(blocks.map((b) => b.section)).toEqual([0, 0, 0, 1, 1])
  })
})

// ─── F2: line-level page splitting + pagination constraints ────────────────

/** Helper: build a block with line boxes */
const lineBlock = (top: number, heights: number[], extra?: Partial<BlockBox>): BlockBox => {
  const spaceBeforePx = extra?.spaceBeforePx ?? 0
  const spaceAfterPx = extra?.spaceAfterPx ?? 0
  let offset = spaceBeforePx
  const lineBoxes = heights.map((h) => {
    const lb = { offsetInBlock: offset, height: h }
    offset += h
    return lb
  })
  const totalHeight = heights.reduce((s, h) => s + h, 0) + spaceBeforePx + spaceAfterPx
  return {
    top,
    height: totalHeight,
    lineBoxes,
    spaceBeforePx,
    spaceAfterPx,
    ...extra,
  }
}

const geoms1 = [{ contentHeight: 200, forceBreak: false }]

describe('sectionWidthSpecs — differing-width sections wrap at their own content width', () => {
  const el = (tag = 'p', marginLeft?: string) => {
    const e = document.createElement(tag)
    if (marginLeft) e.style.marginLeft = marginLeft
    return e
  }

  it('equal-width sections produce no specs', () => {
    const secsList = [sec({}), sec({ pageHeight: 20160 })]
    const blocks = [
      block(0, 100, { section: 0, el: el() }),
      block(100, 100, { section: 1, el: el() }),
    ]
    expect(sectionWidthSpecs(blocks, secsList, sectionGeoms(secsList))).toEqual([])
  })

  it('every block wraps at its own section width; tables get vars only; floats are skipped', () => {
    const landscape = sec({
      pageWidth: 16838,
      pageHeight: 11906,
      orientation: 'landscape',
      marginLeft: 720,
    })
    const secsList = [sec({}), landscape, sec({})]
    const portraitW = ((11906 - 2880) / 1440) * 96
    const landscapeW = ((16838 - 720 - 1440) / 1440) * 96
    const para = el('p', '30px')
    const table = el('table')
    const blocks = [
      block(0, 100, { section: 0, el: el() }),
      block(100, 100, { section: 1, el: para }),
      block(200, 100, { section: 1, el: table }),
      block(300, 100, { section: 1, el: el(), floated: true }),
      block(400, 100, { section: 2, el: el() }),
    ]
    const specs = sectionWidthSpecs(blocks, secsList, sectionGeoms(secsList))
    expect(specs).toHaveLength(4)
    // canvas-width blocks get explicit widths too: preview clones render into
    // per-section wrap widths, so container-relative blocks would reflow there
    expect(specs[0].widthPx).toBeCloseTo(portraitW, 1)
    expect(specs[0].contentWPx).toBeCloseTo(portraitW, 1)
    expect(specs[1].el).toBe(para)
    expect(specs[1].widthPx).toBeCloseTo(landscapeW - 30, 1)
    expect(specs[1].contentWPx).toBeCloseTo(landscapeW, 1)
    expect(specs[1].marginLeftPx).toBeCloseTo((720 / 1440) * 96, 1)
    expect(specs[1].marginRightPx).toBeCloseTo(96, 1)
    expect(specs[2].el).toBe(table)
    expect(specs[2].widthPx).toBeUndefined()
    expect(specs[2].contentWPx).toBeCloseTo(landscapeW, 1)
    expect(specs[3].widthPx).toBeCloseTo(portraitW, 1)
  })

  it('blocks are placed at their own section’s left margin (full-bleed cover section)', () => {
    // A cover section with w:pgMar w:left="0" must not strip the body sections'
    // margins: the canvas pads by the first section, so every other section also
    // needs a horizontal placement offset (dx), not just its own wrap width.
    const cover = sec({ marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0 })
    const body = sec({ marginLeft: 1701, marginRight: 1417 })
    const secsList = [cover, body]
    const blocks = [
      block(0, 100, { section: 0, el: el() }),
      block(100, 100, { section: 1, el: el() }),
    ]
    const specs = sectionWidthSpecs(blocks, secsList, sectionGeoms(secsList))
    expect(specs).toHaveLength(2)
    // the canvas section sits on the page padding: no offset
    expect(specs[0].dx).toBe(0)
    expect(specs[1].dx).toBeCloseTo((1701 / 1440) * 96, 1)
    // width is the section's own content width, so the shifted block's right edge
    // lands on its right margin (1701 + 8788 + 1417 = 11906)
    expect(specs[1].widthPx).toBeCloseTo(((11906 - 1701 - 1417) / 1440) * 96, 1)
  })

  it('side-margin-only differences still get placement specs', () => {
    // Mirrored margins keep the content width identical: nothing would be emitted
    // on the width comparison alone, but the text column must move right.
    const a = sec({ marginLeft: 720, marginRight: 2160 })
    const b = sec({ marginLeft: 2160, marginRight: 720 })
    const secsList = [a, b]
    const blocks = [
      block(0, 100, { section: 0, el: el() }),
      block(100, 100, { section: 1, el: el() }),
    ]
    const specs = sectionWidthSpecs(blocks, secsList, sectionGeoms(secsList))
    expect(specs).toHaveLength(2)
    expect(specs[0].dx).toBe(0)
    expect(specs[1].dx).toBeCloseTo(((2160 - 720) / 1440) * 96, 1)
    expect(specs[1].widthPx).toBeCloseTo(((11906 - 2160 - 720) / 1440) * 96, 1)
  })
})

describe('sectionGridPitchSpecs — per-section typed docGrid pitch', () => {
  const gridSec = (linePitch?: number, type: 'lines' | 'linesAndChars' = 'lines') =>
    sec(linePitch ? { docGrid: { type, linePitch } } : {})
  const el = () => document.createElement('p')

  it('uniform typed pitch produces no specs (single .doc-page injection)', () => {
    const secsList = [gridSec(307), gridSec(307)]
    const blocks = [block(0, 100, { section: 0, el: el() })]
    expect(sectionGridPitchSpecs(blocks, secsList)).toEqual([])
    expect(docGridPitchPt(secsList)).toBeCloseTo(307 / 20, 5)
  })

  it('untyped documents produce no specs', () => {
    const secsList = [sec({}), sec({})]
    expect(sectionGridPitchSpecs([block(0, 100, { section: 0, el: el() })], secsList)).toEqual([])
    expect(docGridPitchPt(secsList)).toBeNull()
  })

  it('mixed pitches: each block carries its own section pitch (prod-sas 043)', () => {
    const secsList = [gridSec(307), gridSec(329, 'linesAndChars')]
    const blocks = [
      block(0, 100, { section: 0, el: el() }),
      block(100, 100, { section: 1, el: el() }),
    ]
    const specs = sectionGridPitchSpecs(blocks, secsList)
    expect(specs).toHaveLength(2)
    expect(specs[0].gridPitchPt).toBeCloseTo(307 / 20, 5)
    expect(specs[1].gridPitchPt).toBeCloseTo(329 / 20, 5)
    expect(docGridPitchPt(secsList)).toBeNull()
  })

  it('typed + untyped mix: untyped-section blocks opt out (pitch 0); own doc-nosnap wins', () => {
    const secsList = [gridSec(307), sec({})]
    const nosnap = el()
    nosnap.classList.add('doc-nosnap')
    const blocks = [
      block(0, 100, { section: 0, el: el() }),
      block(100, 100, { section: 0, el: nosnap }),
      block(200, 100, { section: 1, el: el() }),
    ]
    const specs = sectionGridPitchSpecs(blocks, secsList)
    expect(specs).toHaveLength(2)
    expect(specs[0].gridPitchPt).toBeCloseTo(307 / 20, 5)
    expect(specs[1].gridPitchPt).toBe(0)
  })

  it('the channel-applied doc-grid-nosnap class does not drop the spec on the next pass', () => {
    const secsList = [gridSec(307), sec({})]
    const marked = el()
    marked.classList.add('doc-grid-block', 'doc-grid-nosnap') // applied by setColumnLayout
    const blocks = [block(0, 100, { section: 1, el: marked })]
    const specs = sectionGridPitchSpecs(blocks, secsList)
    expect(specs).toHaveLength(1)
    expect(specs[0].gridPitchPt).toBe(0)
  })
})

// Word probes 2026-09-02 (MS Mincho/Arial, 10.5/12pt): under w:docGrid
// type="linesAndChars" every character advances natural width + charSpace/4096
// pt (10.5pt EA with charSpace=-820 → 10.2998pt), any script and size; types
// lines/default ignore charSpace.
describe('sectionCharSpaceSpecs / docCharSpacePt — docGrid character grid', () => {
  const csSec = (charSpace?: number, type: 'lines' | 'linesAndChars' = 'linesAndChars') =>
    sec({ docGrid: { type, linePitch: 329, ...(charSpace !== undefined ? { charSpace } : {}) } })
  const el = () => document.createElement('p')

  it('uniform nonzero charSpace produces no specs (single .doc-page injection)', () => {
    const secsList = [csSec(-820), csSec(-820)]
    const blocks = [block(0, 100, { section: 0, el: el() })]
    expect(sectionCharSpaceSpecs(blocks, secsList)).toEqual([])
    expect(docCharSpacePt(secsList)).toBeCloseTo(-820 / 4096, 6)
  })

  it('positive charSpace widens (probe: +820 → 10.7002pt at 10.5pt)', () => {
    expect(docCharSpacePt([csSec(820)])).toBeCloseTo(820 / 4096, 6)
  })

  it('lines/default grids and charSpace 0 have no effect', () => {
    expect(docCharSpacePt([csSec(-820, 'lines')])).toBeNull()
    expect(docCharSpacePt([csSec(0)])).toBeNull()
    expect(docCharSpacePt([sec({})])).toBeNull()
    expect(sectionCharSpaceSpecs([block(0, 100, { section: 0, el: el() })], [csSec(0)])).toEqual([])
  })

  it('mixed sections (prod-sas 043): only charSpace-section blocks carry the delta', () => {
    const secsList = [csSec(undefined, 'lines'), csSec(-820)]
    const blocks = [
      block(0, 100, { section: 0, el: el() }),
      block(100, 100, { section: 1, el: el() }),
    ]
    const specs = sectionCharSpaceSpecs(blocks, secsList)
    expect(specs).toHaveLength(1)
    expect(specs[0].el).toBe(blocks[1].el)
    expect(specs[0].charSpacePt).toBeCloseTo(-820 / 4096, 6)
    expect(docCharSpacePt(secsList)).toBeNull()
  })
})

describe('computeSectionedSlicesF2 — line-level pagination', () => {
  it('content shorter than a page does not break', () => {
    const b = lineBlock(0, [50, 50, 50])
    const slices = computeSectionedSlicesF2([b], geoms1, 150)
    expect(slices.length).toBe(1)
  })

  it('a leading block-less section keeps its own blank first page (lone sectPr paragraph)', () => {
    const geoms = [
      { contentHeight: 200, forceBreak: true },
      { contentHeight: 200, forceBreak: true },
    ]
    const blocks = [block(0, 50, { section: 1 }), block(50, 50, { section: 1 })]
    const slices = computeSectionedSlicesF2(blocks, geoms, 100)
    expect(slices.map((s) => [s.start, s.section])).toEqual([
      [0, 0],
      [0, 1],
    ])
  })

  it('a promoted nextColumn start absorbs the leading block-less section (no blank page)', () => {
    // single-column nextColumn acts as a page break (n#750255) but Word skips
    // the blank first page a lone leading sectPr paragraph would otherwise get
    const geoms: SectionGeom[] = [
      { contentHeight: 200, forceBreak: true, startType: 'nextPage' },
      { contentHeight: 200, forceBreak: true, startType: 'nextColumn' },
    ]
    const blocks = [block(0, 50, { section: 1 }), block(50, 50, { section: 1 })]
    const slices = computeSectionedSlicesF2(blocks, geoms, 100)
    expect(slices.map((s) => [s.start, s.section])).toEqual([[0, 1]])
  })

  it('a continuous section after an empty next-page section flows onto the blank page', () => {
    // Word: the lone sectPr paragraph opens the page, the continuous section
    // continues right below it on the same page — the page keeps the empty
    // section's attribution (headers follow the section at the page top)
    const geoms: SectionGeom[] = [
      { contentHeight: 200, forceBreak: false, startType: 'nextPage' },
      { contentHeight: 200, forceBreak: true, startType: 'nextPage' },
      { contentHeight: 200, forceBreak: false, startType: 'continuous' },
    ]
    const blocks = [block(0, 50, { section: 0 }), block(50, 50, { section: 2 })]
    const slices = computeSectionedSlicesF2(blocks, geoms, 100)
    expect(slices.map((s) => [s.start, s.end, s.section])).toEqual([
      [0, 50, 0],
      [50, 100, 1],
    ])
  })

  it('a mid-document block-less next-page section claims a blank page between its neighbours', () => {
    const geoms = [
      { contentHeight: 200, forceBreak: false },
      { contentHeight: 200, forceBreak: true },
      { contentHeight: 200, forceBreak: true },
    ]
    const blocks = [block(0, 50, { section: 0 }), block(50, 50, { section: 2 })]
    const slices = computeSectionedSlicesF2(blocks, geoms, 100)
    expect(slices.map((s) => [s.start, s.section])).toEqual([
      [0, 0],
      [50, 1],
      [50, 2],
    ])
  })

  it('a floated block consumes no column height (wrapped text carries the extent)', () => {
    // float 180px tall, wrapped paragraphs stack to 150 ≤ 200: everything is one
    // page — counting the float would double-book the overlap and break early
    const blocks = [
      { ...block(0, 180), floated: true },
      block(0, 60),
      block(60, 60),
      block(120, 30),
    ]
    const slices = computeSectionedSlicesF2(blocks, geoms1, 180)
    expect(slices.length).toBe(1)
  })

  it('a floated block taller than the remaining column moves whole to the next page', () => {
    const blocks = [block(0, 150), { ...block(150, 180), floated: true }, block(150, 40)]
    const slices = computeSectionedSlicesF2(blocks, geoms1, 330)
    expect(slices.length).toBe(2)
    expect(slices[1].start).toBe(150)
  })

  it('a section break right after a page-filling floated table starts below its band', () => {
    // a landscape form built as one positioned table: the float consumes no
    // column height, but the next section must not cut into its band — the
    // form page keeps the whole table, the next page starts below it
    const geoms: SectionGeom[] = [
      { contentHeight: 900, forceBreak: false, topPx: 96 },
      { contentHeight: 550, forceBreak: true, topPx: 113 },
      { contentHeight: 900, forceBreak: true, topPx: 96 },
    ]
    const blocks = [
      block(0, 20, { section: 0 }),
      {
        ...block(20, 480, { section: 1 }),
        floated: true,
        pageRelVyPx: 33,
        pageRelVAnchor: 'margin' as const,
      },
      block(25, 20, { section: 2 }),
    ]
    // dy = 33 (margin target) → float band ends at 20 + 33 + 480 = 533
    const slices = computeSectionedSlicesF2(blocks, geoms, 553)
    expect(slices.map((s) => [s.start, s.end, s.section])).toEqual([
      [0, 20, 0],
      [20, 533, 1],
      [533, 553, 2],
    ])
  })

  it('an empty section crossed after a float band claims its blank page below it, not above', () => {
    // startPage resets the float bottom: the crossed empty section's second
    // forced start must stay clamped at the previous start (no inverted slice)
    const geoms: SectionGeom[] = [
      { contentHeight: 900, forceBreak: false },
      { contentHeight: 550, forceBreak: true },
      { contentHeight: 900, forceBreak: true },
      { contentHeight: 900, forceBreak: true },
    ]
    const blocks = [
      block(0, 20, { section: 0 }),
      { ...block(20, 480, { section: 1 }), floated: true },
      block(25, 20, { section: 3 }),
    ]
    const slices = computeSectionedSlicesF2(blocks, geoms, 520)
    expect(slices.map((s) => [s.start, s.end, s.section])).toEqual([
      [0, 20, 0],
      [20, 500, 1],
      [500, 500, 2],
      [500, 520, 3],
    ])
  })

  it('a margin-anchored floated table shifts down to its tblpY target on the landing page', () => {
    const out: SliceOutputs = { floatVShifts: [] }
    const blocks = [
      block(0, 100),
      { ...block(100, 60), floated: true, pageRelVyPx: 150, pageRelVAnchor: 'margin' as const },
      block(100, 40),
    ]
    computeSectionedSlicesF2(blocks, geoms1, 140, out)
    expect(out.floatVShifts).toEqual([{ blockTop: 100, dyPx: 50 }])
  })

  it('a page-anchored tblpY target converts through the section top margin', () => {
    const out: SliceOutputs = { floatVShifts: [] }
    const geoms: SectionGeom[] = [{ contentHeight: 200, forceBreak: false, topPx: 40 }]
    const blocks = [
      block(0, 100),
      { ...block(100, 60), floated: true, pageRelVyPx: 150, pageRelVAnchor: 'page' as const },
    ]
    computeSectionedSlicesF2(blocks, geoms, 100, out)
    expect(out.floatVShifts).toEqual([{ blockTop: 100, dyPx: 10 }])
  })

  it('an anchor target above the flow position clamps to zero (floats never move up)', () => {
    const out: SliceOutputs = { floatVShifts: [] }
    const blocks = [
      block(0, 100),
      { ...block(100, 60), floated: true, pageRelVyPx: 30, pageRelVAnchor: 'margin' as const },
    ]
    computeSectionedSlicesF2(blocks, geoms1, 100, out)
    expect(out.floatVShifts).toEqual([{ blockTop: 100, dyPx: 0 }])
  })

  it('an anchored float pushed to the next page targets that page instead', () => {
    const out: SliceOutputs = { floatVShifts: [] }
    const blocks = [
      block(0, 150),
      { ...block(150, 180), floated: true, pageRelVyPx: 20, pageRelVAnchor: 'margin' as const },
    ]
    computeSectionedSlicesF2(blocks, geoms1, 330, out)
    expect(out.floatVShifts).toEqual([{ blockTop: 150, dyPx: 20 }])
  })

  it('a leading w:br on the first content keeps the blank first page (fdo#78907)', () => {
    const b = { ...block(0, 100), breakBefore: true, breakBeforeBr: true }
    const slices = computeSectionedSlicesF2([b], geoms1, 100)
    expect(slices).toEqual([
      { start: 0, end: 0, section: 0 },
      { start: 0, end: 100, section: 0 },
    ])
  })

  it('a pageBreakBefore property on the first content stays suppressed', () => {
    const b = { ...block(0, 100), breakBefore: true }
    const slices = computeSectionedSlicesF2([b], geoms1, 100)
    expect(slices.length).toBe(1)
  })

  it('a pending w:br plus a leading w:br are two page turns with a blank sheet between (tdf#154478)', () => {
    const blocks = [
      { ...block(0, 100), breakAfter: true },
      { ...block(100, 100), breakBefore: true, breakBeforeBr: true },
    ]
    const slices = computeSectionedSlicesF2(blocks, geoms1, 200)
    expect(slices).toEqual([
      { start: 0, end: 100, section: 0 },
      { start: 100, end: 100, section: 0 },
      { start: 100, end: 200, section: 0 },
    ])
  })

  it('a trailing w:br keeps its deliberate blank last page (tdf#99090)', () => {
    const blocks = [{ ...block(0, 100), breakAfter: true }]
    const slices = computeSectionedSlicesF2(blocks, geoms1, 100)
    expect(slices).toEqual([
      { start: 0, end: 100, section: 0 },
      { start: 100, end: 100, section: 0 },
    ])
  })

  it('paragraph that fits entirely is not split', () => {
    // 3 lines × 50px = 150 < 200, no page break
    const b = lineBlock(0, [50, 50, 50])
    const slices = computeSectionedSlicesF2([b], geoms1, 150)
    expect(slices).toHaveLength(1)
  })

  it('spanning paragraph splits at line level', () => {
    // Page height 200; block has 5 lines of 50px = 250px, should break before line 5
    // 4 lines = 200px fill exactly; line 5 breaks to the next page
    const b = lineBlock(0, [50, 50, 50, 50, 50])
    const slices = computeSectionedSlicesF2([b], geoms1, 250)
    expect(slices.length).toBe(2)
    // Page 1 holds 4 lines (4×50=200); line 5 starts on the next page at offset 200
    expect(slices[1].start).toBe(200) // offset of line 4 = 4×50
  })

  it('widowControl: a single line at the page bottom pushes the whole paragraph to the next page', () => {
    // Page height 160; paragraph has 4 lines of 50px = 200px
    // 160px left at page end fits only 3 lines → 3 lines at page end / 1 at page top → widow
    // widowControl should push the whole paragraph down (after the 1-line-orphan adjustment, 0 lines remain at page end → push all)
    const beforeBlock = lineBlock(0, [50, 50, 50]) // 150px used
    const mainBlock = lineBlock(150, [50, 50, 50, 50], { widowControl: true }) // 200px
    const geoms = [{ contentHeight: 200, forceBreak: false }]
    const slices = computeSectionedSlicesF2([beforeBlock, mainBlock], geoms, 350)
    // mainBlock 150+200=350; page height 200, previous block uses 150, only 50px left for mainBlock → just 1 line
    // 1 line = orphan; whole paragraph pushed to page 2
    expect(slices.length).toBe(2)
    expect(slices[1].start).toBe(150) // mainBlock pushed down whole; new page starts at top=150
  })

  it('widowControl: gives up one line at the page bottom so the next page gets 2 lines', () => {
    // Page height 300; 150px used → exactly 3 of 4 lines fit → a naive split leaves a 1-line widow
    // Word takes one line back: 2 lines stay, 2 go to the next page (no whole-paragraph push)
    const beforeBlock = lineBlock(0, [50, 50, 50])
    const mainBlock = lineBlock(150, [50, 50, 50, 50], { widowControl: true })
    const geoms = [{ contentHeight: 300, forceBreak: false }]
    const slices = computeSectionedSlicesF2([beforeBlock, mainBlock], geoms, 350)
    expect(slices.length).toBe(2)
    expect(slices[1].start).toBe(250) // 2+2 split; a whole push would start at 150
  })

  it('widowControl=false: a single line may remain at the page bottom', () => {
    const beforeBlock = lineBlock(0, [50, 50, 50]) // 150px
    const mainBlock = lineBlock(150, [50, 50, 50, 50], { widowControl: false }) // 200px
    const geoms = [{ contentHeight: 200, forceBreak: false }]
    const slices = computeSectionedSlicesF2([beforeBlock, mainBlock], geoms, 350)
    // No widow/orphan control: line 1 fits, line 2 breaks to the next page
    expect(slices.length).toBe(2)
  })

  it('keepLines: whole paragraph stays on one page', () => {
    const b1 = lineBlock(0, [50, 50, 50]) // 150px
    const b2 = lineBlock(150, [50, 50, 50], { keepLines: true }) // 150px, total 300 > 200
    const geoms = [{ contentHeight: 200, forceBreak: false }]
    const slices = computeSectionedSlicesF2([b1, b2], geoms, 300)
    // b2 pushed whole to page 2 (150px < 200px, it fits)
    expect(slices.length).toBe(2)
    expect(slices[1].start).toBe(150)
  })

  it('keepLines: paragraph taller than a page is handled best-effort (no infinite loop)', () => {
    // Paragraph 300px > page height 200px
    const b = lineBlock(0, [60, 60, 60, 60, 60], { keepLines: true })
    expect(() => computeSectionedSlicesF2([b], geoms1, 300)).not.toThrow()
  })

  it('keepLines without line data taller than a page terminates with hard cuts', () => {
    // First slicing pass has no lineBoxes yet; a keepLines paragraph taller
    // than the page used to re-test its full height after every column turn
    // and loop forever (form-gov renderer hang).
    const b = block(0, 1000, { keepLines: true })
    const slices = computeSectionedSlicesF2([b], geoms1, 1000)
    expect(slices.length).toBe(5) // 1000px / 200px pages
    expect(slices[0].start).toBe(0)
    expect(slices[slices.length - 1].end).toBe(1000)
    // pages advance monotonically
    for (let i = 1; i < slices.length; i++) {
      expect(slices[i].start).toBeGreaterThan(slices[i - 1].start)
    }
  })

  it('keepLines without line data on a partly used page fills the remainder then hard-cuts', () => {
    const before = block(0, 150)
    const b = block(150, 500, { keepLines: true })
    const slices = computeSectionedSlicesF2([before, b], geoms1, 650)
    // 50px fills page 1 (matches the _hardCutLines policy), then 200px cuts
    expect(slices.map((s) => s.start)).toEqual([0, 200, 400, 600])
    expect(slices[slices.length - 1].end).toBe(650)
  })

  it('keepNext: chain head stays on the same page as the first line of the next paragraph', () => {
    // Block A (100px keepNext) + block B (100px) — page height 200
    // 120px of the page is used; A alone fits (100px → 220px > 200), but A + B's first line must share a page
    // A + B's first line = 100+50 = 150 > remaining 80px → A pushed to page 2
    const beforeBlock = block(0, 120)
    const a = lineBlock(120, [50, 50], { keepNext: true }) // 100px
    const b = lineBlock(220, [50, 50]) // 100px
    const geoms = [{ contentHeight: 200, forceBreak: false }]
    const slices = computeSectionedSlicesF2([beforeBlock, a, b], geoms, 320)
    expect(slices.length).toBeGreaterThanOrEqual(2)
    // A should be on page 2 (120+100+50 > 200, page break needed)
    const aInPage2 = slices.some((s) => Math.abs(s.start - 120) < 2)
    expect(aInPage2).toBe(true)
  })

  it('pageBreakBefore has the highest priority (overrides keepNext)', () => {
    const a = block(0, 100, { keepNext: true })
    const b = block(100, 100, { breakBefore: true })
    const slices = computeSectionedSlicesF2([a, b], geoms1, 200)
    // b has breakBefore; even with a's keepNext, a page break is forced before b
    expect(slices.length).toBe(2)
    expect(slices[1].start).toBe(100)
  })

  it("keepNext on the document's last block places normally (POI headerPic: lone keepNext paragraph crashed)", () => {
    const only = block(0, 100, { keepNext: true })
    const slices = computeSectionedSlicesF2([only], geoms1, 100)
    expect(slices.length).toBe(1)
    expect(slices[0].end).toBe(100)

    const a = block(0, 100)
    const last = block(100, 100, { keepNext: true })
    const two = computeSectionedSlicesF2([a, last], geoms1, 200)
    expect(two[two.length - 1].end).toBe(200)
  })

  it('keepNext+keepLines heading at the page bottom moves with the next paragraph', () => {
    // Word Heading styles carry both flags; the chain decides the push even when
    // the heading alone would fit at the page bottom.
    const before = block(0, 120)
    const heading = lineBlock(120, [60], { keepNext: true, keepLines: true })
    const body = lineBlock(180, [50, 50])
    const slices = computeSectionedSlicesF2([before, heading, body], geoms1, 280)
    expect(slices.length).toBe(2)
    expect(slices[1].start).toBe(120) // heading pushed together with the body
  })

  it('keepLines without keepNext stays at the page bottom (no chain push)', () => {
    const before = block(0, 120)
    const kl = lineBlock(120, [60], { keepLines: true })
    const body = lineBlock(180, [50, 50])
    const slices = computeSectionedSlicesF2([before, kl, body], geoms1, 280)
    expect(slices.length).toBe(2)
    expect(slices[1].start).toBe(180) // only the body moves
  })

  it('keepNext chain taller than a page degrades to per-block placement', () => {
    const a = lineBlock(0, [60, 60], { keepNext: true, widowControl: false })
    const b = lineBlock(120, [60, 60], { keepNext: true, widowControl: false })
    const c = lineBlock(240, [60, 60], { keepNext: true, widowControl: false })
    const anchor = lineBlock(360, [50], { widowControl: false })
    const slices = computeSectionedSlicesF2([a, b, c, anchor], geoms1, 410)
    expect(slices.map((s) => s.start)).toEqual([0, 180, 360])
    expect(slices[slices.length - 1].end).toBe(410)
  })

  it('overlong keepNext chain with a keepLines head keeps the head unsplit', () => {
    const before = block(0, 120)
    const head = lineBlock(120, [50, 50, 50], { keepNext: true, keepLines: true })
    const mid = lineBlock(270, [60, 60], { keepNext: true })
    const anchor = lineBlock(390, [50, 50], { widowControl: false })
    const slices = computeSectionedSlicesF2([before, head, mid, anchor], geoms1, 490)
    // head (120..270) pushed whole to page 2, never split
    expect(slices.map((s) => s.start)).toEqual([0, 120, 270, 440])
    expect(slices[slices.length - 1].end).toBe(490)
  })

  it('keepNext heading stays at the page bottom when the anchor brings its first 2 lines', () => {
    // Word: the anchor paragraph is not atomic — the heading only needs the anchor's
    // first 2 lines (widow minimum) on the same page; the paragraph splits normally.
    const before = block(0, 140)
    const heading = lineBlock(140, [20], { keepNext: true, keepLines: true })
    const body = lineBlock(160, [20, 20, 20, 20, 20]) // widow control on by default
    const slices = computeSectionedSlicesF2([before, heading, body], geoms1, 260)
    // 140 + heading 20 + 2 body lines 40 = 200 fits → heading stays, body splits 2/3
    expect(slices.map((s) => s.start)).toEqual([0, 200])
    expect(slices[slices.length - 1].end).toBe(260)
  })

  it("keepNext heading pushes when the anchor's first 2 lines do not fit; the anchor flows on", () => {
    const before = block(0, 170)
    const heading = lineBlock(170, [20], { keepNext: true })
    const body = lineBlock(190, [20, 20, 20])
    const slices = computeSectionedSlicesF2([before, heading, body], geoms1, 250)
    // 170 + 20 + 40 > 200 → heading pushed; body follows and fits whole on page 2
    expect(slices.map((s) => s.start)).toEqual([0, 170])
    expect(slices[slices.length - 1].end).toBe(250)
  })

  it('keepLines anchor follows the chain whole (atomic exception)', () => {
    const before = block(0, 140)
    const heading = lineBlock(140, [20], { keepNext: true })
    const body = lineBlock(160, [20, 20, 20, 20], { keepLines: true })
    const slices = computeSectionedSlicesF2([before, heading, body], geoms1, 240)
    // anchor demand = whole 80px block → 140+20+80 > 200 → heading + body pushed together
    expect(slices.map((s) => s.start)).toEqual([0, 140])
    expect(slices[slices.length - 1].end).toBe(240)
  })

  it('anchor without line data (first pass) is kept whole conservatively', () => {
    const before = block(0, 140)
    const heading = block(140, 20, { keepNext: true })
    const body = block(160, 80)
    const slices = computeSectionedSlicesF2([before, heading, body], geoms1, 240)
    expect(slices.map((s) => s.start)).toEqual([0, 140])
  })
})

describe('computeSectionedSlicesF2 — table row-level page breaks', () => {
  const makeTableBlock = (top: number, rows: TableRowBox[]): BlockBox => {
    const h = rows.reduce((s, r) => s + r.height, 0)
    return { top, height: h, tableRows: rows }
  }

  it('table splits at row level: breaks at row boundaries', () => {
    // Table: 5 rows of 50px = 250px, page height 200px
    const rows: TableRowBox[] = Array.from({ length: 5 }, () => ({ height: 50 }))
    const b = makeTableBlock(0, rows)
    const slices = computeSectionedSlicesF2([b], geoms1, 250)
    expect(slices.length).toBe(2)
  })

  it('cantSplit row: whole row is pushed to the next page (no break inside the row)', () => {
    // 3 rows: 50+50+50, page height 120, the third row does not fit
    const rows: TableRowBox[] = [
      { height: 50 },
      { height: 50 },
      { height: 50, cantSplit: true }, // does not fit; whole row pushed to the next page
    ]
    const b = makeTableBlock(0, rows)
    const geoms = [{ contentHeight: 120, forceBreak: false }]
    const slices = computeSectionedSlicesF2([b], geoms, 150)
    // Row 3 is on page 2
    expect(slices.length).toBe(2)
  })

  it('empty table does not crash', () => {
    const b = makeTableBlock(0, [])
    expect(() => computeSectionedSlicesF2([b], geoms1, 0)).not.toThrow()
  })

  it('split atLeast row: the continuation fragment re-honors the declared height (Word probe 2026-08-27)', () => {
    // page 200px; a 40px row leaves 160 >= the declared 150, so no fresh-page
    // turn: the 300px-content row starts mid-page and splits (cut at 200).
    // The final fragment holds 100px of content; Word stretches it to the full
    // declared 150 -> row target = 200 + 150 = 350
    const rows: TableRowBox[] = [
      { height: 40 },
      { height: 300, minHPx: 150, cutYs: [200], contentBottom: 300 },
    ]
    const b = makeTableBlock(0, rows)
    const out: SliceOutputs = { rowFills: [] }
    const slices = computeSectionedSlicesF2(
      [b],
      [{ contentHeight: 200, forceBreak: false }],
      340,
      out,
    )
    expect(out.rowFills).toEqual([{ blockTop: 0, row: 1, targetPx: 350 }])
    expect(slices.length).toBe(3)
  })

  it('split row without a declared height reports no fill', () => {
    const rows: TableRowBox[] = [{ height: 80 }, { height: 160, cutYs: [80], contentBottom: 160 }]
    const b = makeTableBlock(0, rows)
    const out: SliceOutputs = { rowFills: [] }
    computeSectionedSlicesF2([b], [{ contentHeight: 200, forceBreak: false }], 240, out)
    expect(out.rowFills).toEqual([])
  })

  it('a previously patched split row re-emits the same target (no oscillation)', () => {
    // same card as above but the tr already carries the 350px patch: the target
    // must be re-emitted level-triggered, or clearing the decoration would
    // shrink the row and pagination would oscillate (Bugbot 2026-08-27)
    const rows: TableRowBox[] = [
      { height: 40 },
      { height: 350, minHPx: 150, cutYs: [200], contentBottom: 300 },
    ]
    const b = makeTableBlock(0, rows)
    const out: SliceOutputs = { rowFills: [] }
    computeSectionedSlicesF2([b], [{ contentHeight: 200, forceBreak: false }], 390, out)
    expect(out.rowFills).toEqual([{ blockTop: 0, row: 1, targetPx: 350 }])
  })

  it('unsplit declared row reports no fill', () => {
    const rows: TableRowBox[] = [{ height: 120, minHPx: 150, cutYs: [60], contentBottom: 120 }]
    const b = makeTableBlock(0, rows)
    const out: SliceOutputs = { rowFills: [] }
    computeSectionedSlicesF2([b], [{ contentHeight: 200, forceBreak: false }], 130, out)
    expect(out.rowFills).toEqual([])
  })

  it('tblHeader: continuation pages record repeatHeader and reserve header space', () => {
    // Header 40 + 6 rows × 50 = 340px, page height 200
    const rows: TableRowBox[] = [
      { height: 40, isHeader: true },
      ...Array.from({ length: 6 }, () => ({ height: 50 })),
    ]
    const b = makeTableBlock(0, rows)
    const slices = computeSectionedSlicesF2([b], [{ contentHeight: 200, forceBreak: false }], 340)
    expect(slices.length).toBeGreaterThan(1)
    for (const s of slices.slice(1)) {
      expect(s.repeatHeader).toEqual({ top: 0, height: 40 })
    }
    // From page 2 the header takes 40px: rows per page = (200-40)/50 = 3
    expect(slices[1].end - slices[1].start).toBeLessThanOrEqual(160)
  })

  it('tblHeader: header at 75% of the page still repeats (Word probe 2026-08-16)', () => {
    const rows: TableRowBox[] = [
      { height: 150, isHeader: true },
      ...Array.from({ length: 4 }, () => ({ height: 50 })),
    ]
    const b = makeTableBlock(0, rows)
    const slices = computeSectionedSlicesF2([b], [{ contentHeight: 200, forceBreak: false }], 350)
    for (const s of slices.slice(1)) expect(s.repeatHeader).toEqual({ top: 0, height: 150 })
  })

  it('tblHeader: header block taller than a full page is not repeated', () => {
    const rows: TableRowBox[] = [
      { height: 120, isHeader: true },
      { height: 120, isHeader: true },
      ...Array.from({ length: 4 }, () => ({ height: 50 })),
    ]
    const b = makeTableBlock(0, rows)
    const slices = computeSectionedSlicesF2([b], [{ contentHeight: 200, forceBreak: false }], 440)
    for (const s of slices) expect(s.repeatHeader).toBeUndefined()
  })

  it('modernTableHeaders: a header block that misses the remaining space pushes the table to a fresh page', () => {
    const para: BlockBox = { top: 0, height: 60 }
    const rows: TableRowBox[] = [
      { height: 90, isHeader: true },
      { height: 90, isHeader: true },
      ...Array.from({ length: 3 }, () => ({ height: 40 })),
    ]
    const b = makeTableBlock(60, rows)
    b.modernTableHeaders = true
    const slices = computeSectionedSlicesF2(
      [para, b],
      [{ contentHeight: 200, forceBreak: false }],
      360,
    )
    // page 1 keeps only the paragraph; the table starts page 2 with its header block
    expect(slices[1].start).toBe(60)
    // legacy mode places header row 1 on page 1 and splits in place
    const b2 = makeTableBlock(
      60,
      rows.map((r) => ({ ...r })),
    )
    const legacy = computeSectionedSlicesF2(
      [{ top: 0, height: 60 }, b2],
      [{ contentHeight: 200, forceBreak: false }],
      360,
    )
    expect(legacy[1].start).toBe(150)
  })

  it('split table without tblHeader carries no repeatHeader', () => {
    const rows: TableRowBox[] = Array.from({ length: 5 }, () => ({ height: 50 }))
    const b = makeTableBlock(0, rows)
    const slices = computeSectionedSlicesF2([b], geoms1, 250)
    for (const s of slices) expect(s.repeatHeader).toBeUndefined()
  })

  it('in-row cut points: oversized row splits across pages by cutYs (Word allows in-row breaks by default)', () => {
    // Row 1 50px + row 2 300px (cut points 100/200), page height 120
    const rows: TableRowBox[] = [{ height: 50 }, { height: 300, cutYs: [100, 200] }]
    const b = makeTableBlock(0, rows)
    const slices = computeSectionedSlicesF2([b], [{ contentHeight: 120, forceBreak: false }], 350)
    // First segment (100px) does not fit at page end (50+100>120) → row 2 starts on a new page and continues segment by segment
    expect(slices.map((s) => s.start)).toEqual([0, 50, 150, 250])
  })

  it('in-row cut points: breaks at the page bottom when the first segment fits (row not pushed whole)', () => {
    const rows: TableRowBox[] = [{ height: 50 }, { height: 200, cutYs: [60, 120] }]
    const b = makeTableBlock(0, rows)
    const slices = computeSectionedSlicesF2([b], [{ contentHeight: 120, forceBreak: false }], 250)
    // Page 1: row 1 (50) + row 2's first segment (60) = 110 ≤ 120; page 2 starts at in-row cut point 50+60=110
    expect(slices[1].start).toBe(110)
  })

  it('plain first row splits at cutYs like any row (Word probe: no first-row rule)', () => {
    // 80px block, then a table whose first row (100px, cuts 40/70) exceeds the 40px remainder
    const rows: TableRowBox[] = [{ height: 100, cutYs: [40, 70] }, { height: 20 }]
    const table = makeTableBlock(80, rows)
    const slices = computeSectionedSlicesF2(
      [block(0, 80), table],
      [{ contentHeight: 120, forceBreak: false }],
      200,
    )
    // first segment (40) fills page 1; page 2 starts at the in-row cut 80+40=120
    expect(slices.map((s) => s.start)).toEqual([0, 120])
  })

  it('1x1 table: page remainder holds 2 of 3 segments, row splits leaving them behind', () => {
    // sample 14_10f3d2ed cover banner: Word leaves 2 paragraphs on page 1
    const rows: TableRowBox[] = [{ height: 90, cutYs: [20, 40] }]
    const table = makeTableBlock(80, rows)
    const slices = computeSectionedSlicesF2(
      [block(0, 80), table],
      [{ contentHeight: 120, forceBreak: false }],
      170,
    )
    expect(slices.map((s) => s.start)).toEqual([0, 120])
  })

  it('tblHeader first row is pushed whole despite cutYs', () => {
    const rows: TableRowBox[] = [{ height: 50, cutYs: [20, 35], isHeader: true }, { height: 20 }]
    const table = makeTableBlock(100, rows)
    const slices = computeSectionedSlicesF2(
      [block(0, 100), table],
      [{ contentHeight: 120, forceBreak: false }],
      170,
    )
    expect(slices.map((s) => s.start)).toEqual([0, 100])
  })

  it('tblHeader row taller than half a page still pushes whole (only repetition is dropped)', () => {
    // header block > contentH/2 disables per-page repetition, not the no-split rule:
    // remaining space (50) fits the first cut segment (30), so a split-allowed row
    // would leave it behind — the header row must still push whole
    const rows: TableRowBox[] = [{ height: 80, cutYs: [30, 55], isHeader: true }, { height: 20 }]
    const table = makeTableBlock(70, rows)
    const slices = computeSectionedSlicesF2(
      [block(0, 70), table],
      [{ contentHeight: 120, forceBreak: false }],
      170,
    )
    expect(slices.map((s) => s.start)).toEqual([0, 70])
  })

  it('first row taller than an empty page still splits at cutYs', () => {
    const rows: TableRowBox[] = [{ height: 300, cutYs: [100, 200] }]
    const b = makeTableBlock(0, rows)
    const slices = computeSectionedSlicesF2([b], [{ contentHeight: 120, forceBreak: false }], 300)
    expect(slices.map((s) => s.start)).toEqual([0, 100, 200])
  })

  it('cantSplit row ignores cutYs and stays atomic', () => {
    const rows: TableRowBox[] = [{ height: 50 }, { height: 200, cutYs: [60, 120], cantSplit: true }]
    const b = makeTableBlock(0, rows)
    const slices = computeSectionedSlicesF2([b], [{ contentHeight: 220, forceBreak: false }], 250)
    expect(slices[1].start).toBe(50)
  })

  it('atLeast trHeight overflowing the remainder pushes the whole row (Word probe 2026-08-23)', () => {
    // reserved 150px min with only 30px of content: the declared height does
    // not fit the 100px remainder, so the row pushes whole instead of leaving
    // a mostly-empty fragment on page 1
    const rows: TableRowBox[] = [
      { height: 100 },
      { height: 150, minHPx: 150, contentBottom: 30, cutYs: [30] },
    ]
    const b = makeTableBlock(0, rows)
    const slices = computeSectionedSlicesF2([b], [{ contentHeight: 200, forceBreak: false }], 250)
    expect(slices.map((s) => s.start)).toEqual([0, 100])
  })

  it('atLeast minH turn with a repeated header does not double the page break', () => {
    // header repeats on the fresh page (usedInCol = 40), so the atomic path
    // must not see the non-empty page and turn again
    const rows: TableRowBox[] = [
      { height: 40, isHeader: true },
      { height: 150 },
      { height: 100, minHPx: 100, cantSplit: true },
    ]
    const b = makeTableBlock(0, rows)
    const slices = computeSectionedSlicesF2([b], [{ contentHeight: 200, forceBreak: false }], 290)
    expect(slices.map((s) => s.start)).toEqual([0, 190])
  })

  it('over-page atLeast trHeight row starts on a fresh page, then splits at cutYs', () => {
    const rows: TableRowBox[] = [
      { height: 100 },
      { height: 500, minHPx: 450, cutYs: [150, 300, 450] },
    ]
    const b = makeTableBlock(0, rows)
    const slices = computeSectionedSlicesF2([b], [{ contentHeight: 200, forceBreak: false }], 600)
    expect(slices.map((s) => s.start)).toEqual([0, 100, 250, 400])
  })

  it('atLeast trHeight row whose declared minimum fits the remainder still splits at cutYs', () => {
    // content grew past the 120px minimum; the minimum fits the 140px
    // remainder, so the declared height plays no role and the row splits
    const rows: TableRowBox[] = [
      { height: 60 },
      { height: 150, minHPx: 120, contentBottom: 150, cutYs: [50, 100] },
    ]
    const b = makeTableBlock(0, rows)
    const slices = computeSectionedSlicesF2([b], [{ contentHeight: 200, forceBreak: false }], 210)
    expect(slices.map((s) => s.start)).toEqual([0, 160])
  })

  it('over-page fixed row without natural cut points advances by content bands', () => {
    const rows: TableRowBox[] = [{ height: 550 }]
    const b = makeTableBlock(0, rows)
    const slices = computeSectionedSlicesF2([b], [{ contentHeight: 200, forceBreak: false }], 550)
    expect(slices.map((slice) => [slice.start, slice.end])).toEqual([
      [0, 200],
      [200, 400],
      [400, 550],
    ])
  })

  it('over-page fixed row prefers natural cut points and fills overly long content bands', () => {
    const rows: TableRowBox[] = [{ height: 600, cutYs: [150, 500] }]
    const b = makeTableBlock(0, rows)
    const slices = computeSectionedSlicesF2([b], [{ contentHeight: 200, forceBreak: false }], 600)
    expect(slices.map((slice) => slice.start)).toEqual([0, 150, 350, 500])
  })

  it('page-sized declared-fill row is pushed whole, not clipped (fill clipping is over-tall only)', () => {
    const rows: TableRowBox[] = [{ height: 50 }, { height: 180, contentBottom: 20 }]
    const b = makeTableBlock(0, rows)
    const slices = computeSectionedSlicesF2([b], geoms1, 230)
    expect(slices.map((slice) => slice.start)).toEqual([0, 50])
  })

  it('page-sized rows after a pushed fill row keep whole-row placement', () => {
    const rows: TableRowBox[] = [
      { height: 50 },
      { height: 180, contentBottom: 20 },
      { height: 40, contentBottom: 30 },
    ]
    const b = makeTableBlock(0, rows)
    const slices = computeSectionedSlicesF2([b], geoms1, 270)
    expect(slices.map((slice) => slice.start)).toEqual([0, 50, 230])
  })

  it('empty rows that do not fit are pushed whole, not absorbed by clipped-fill bookkeeping', () => {
    // empty rows report contentBottom 0 and no cuts; each must turn the page like an atomic row
    const rows: TableRowBox[] = [
      { height: 190, contentBottom: 190 },
      ...Array.from({ length: 4 }, () => ({ height: 30, contentBottom: 0 })),
    ]
    const b = makeTableBlock(0, rows)
    const slices = computeSectionedSlicesF2([b], geoms1, 310)
    expect(slices.map((slice) => slice.start)).toEqual([0, 190])
  })

  it('trailing padding of a page-sized row stays glued to its last text band (no fill strip past the page)', () => {
    // two text bands (cut at 45, content ends at 80) + shaded bottom padding to 120
    const rows: TableRowBox[] = [{ height: 110 }, { height: 120, cutYs: [45], contentBottom: 80 }]
    const b = makeTableBlock(0, rows)
    const slices = computeSectionedSlicesF2([b], geoms1, 230)
    // last segment spans 45..120 and does not fit the 45px remainder → breaks at the cut
    expect(slices.map((slice) => slice.start)).toEqual([0, 155])
  })

  it('declared row taller than a page with top-only content: one clipped page, no empty segment pages', () => {
    const rows: TableRowBox[] = [{ height: 600, contentBottom: 40 }]
    const b = makeTableBlock(0, rows)
    const slices = computeSectionedSlicesF2([b], geoms1, 600)
    expect(slices).toEqual([{ start: 0, end: 600, section: 0 }])
  })

  it('content below the fold still pushes the row to the next page (not clipped)', () => {
    const rows: TableRowBox[] = [{ height: 50 }, { height: 100, contentBottom: 80 }]
    const b = makeTableBlock(0, rows)
    const slices = computeSectionedSlicesF2([b], [{ contentHeight: 120, forceBreak: false }], 150)
    expect(slices.map((slice) => slice.start)).toEqual([0, 50])
  })

  it('rows with natural cut points keep the segment path even with contentBottom set', () => {
    const rows: TableRowBox[] = [
      { height: 50 },
      { height: 200, cutYs: [60, 120], contentBottom: 190 },
    ]
    const b = makeTableBlock(0, rows)
    const slices = computeSectionedSlicesF2([b], [{ contentHeight: 120, forceBreak: false }], 250)
    expect(slices[1].start).toBe(110)
  })

  it('fill below the last content band collapses into a clipped remainder, not empty pages', () => {
    const rows: TableRowBox[] = [{ height: 600, cutYs: [50], contentBottom: 80 }]
    const b = makeTableBlock(0, rows)
    const slices = computeSectionedSlicesF2([b], geoms1, 600)
    expect(slices).toEqual([{ start: 0, end: 600, section: 0 }])
  })

  it('multi-page content keeps page-sized segments; only the fill tail is clipped', () => {
    const rows: TableRowBox[] = [{ height: 900, contentBottom: 450 }]
    const b = makeTableBlock(0, rows)
    const slices = computeSectionedSlicesF2([b], geoms1, 900)
    expect(slices.map((slice) => slice.start)).toEqual([0, 200, 400])
  })
})

describe('computeSectionedSlicesF2 — keepNext chain anchored by a table', () => {
  it('the heading only keeps with the first table row; the table breaks by rows', () => {
    const filler = block(0, 80)
    const heading = block(80, 20, { keepNext: true })
    const rows: TableRowBox[] = Array.from({ length: 6 }, () => ({ height: 30 }))
    const table = block(100, 180, { tableRows: rows })
    const slices = computeSectionedSlicesF2([filler, heading, table], geoms1, 280)
    // whole-table anchor height would push heading + table to page 2 (start 80)
    expect(slices.map((slice) => slice.start)).toEqual([0, 190])
  })
})

describe('fillLineBoxes — keepNext chain anchors', () => {
  const geoms = [{ contentHeight: 200, forceBreak: false }]
  const tableEl = () => {
    const el = document.createElement('div')
    el.innerHTML = '<table><tbody><tr><td></td></tr></tbody></table>'
    return el
  }

  it('samples table rows even when the table fits on one page', () => {
    const heading: BlockBox = { top: 0, height: 20, keepNext: true }
    const table: BlockBox = { top: 20, height: 60, el: tableEl() }
    expect(fillLineBoxes([heading, table], geoms, 1)).toBe(true)
    expect(table.tableRows).toEqual([{ height: 60, contentBottom: 0 }])
  })

  it('leaves non-anchored fitting tables unsampled', () => {
    const para: BlockBox = { top: 0, height: 20 }
    const table: BlockBox = { top: 20, height: 60, el: tableEl() }
    expect(fillLineBoxes([para, table], geoms, 1)).toBe(false)
    expect(table.tableRows).toBeUndefined()
  })

  it('samples a table that fits the page but not its mixed-column region', () => {
    // real_run2/61: a table after a balanced 3-col region has only the second
    // region's height; gating on the full page height left it row-less and the
    // first pass placed it whole into the short region (collapse fixed point)
    const table: BlockBox = { top: 1000, height: 150, section: 1, el: tableEl() }
    const slices: PageSlice[] = [
      {
        start: 0,
        end: 1150,
        section: 0,
        regions: [
          {
            top: 0,
            height: 120,
            section: 0,
            columns: [
              { start: 0, end: 500 },
              { start: 500, end: 1000 },
            ],
          },
          { top: 120, height: 80, section: 1, columns: [{ start: 1000, end: 1150 }] },
        ],
      },
    ]
    expect(fillLineBoxes([table], geoms, 1, slices)).toBe(true)
    expect(table.tableRows).toBeDefined()
  })

  it('keeps a fitting table unsampled when its region holds it', () => {
    // top offset from the column start: a block exactly at a column top is
    // always sampled by the existing atPageTop rule
    const table: BlockBox = { top: 1010, height: 60, section: 1, el: tableEl() }
    const slices: PageSlice[] = [
      {
        start: 0,
        end: 1150,
        section: 0,
        regions: [
          { top: 0, height: 120, section: 0, columns: [{ start: 0, end: 1000 }] },
          { top: 120, height: 80, section: 1, columns: [{ start: 1000, end: 1150 }] },
        ],
      },
    ]
    expect(fillLineBoxes([table], geoms, 1, slices)).toBe(false)
    expect(table.tableRows).toBeUndefined()
  })

  it('samples line boxes for paragraph anchors even when they fit on one page', () => {
    const heading: BlockBox = { top: 0, height: 20, keepNext: true }
    const el = document.createElement('p')
    el.textContent = 'two lines'
    const para: BlockBox = { top: 20, height: 40, el }
    const rects = [
      { top: 0, bottom: 20, height: 20, width: 50, left: 0 },
      { top: 20, bottom: 40, height: 20, width: 50, left: 0 },
    ]
    const orig = Range.prototype.getClientRects
    Range.prototype.getClientRects = () => rects as unknown as DOMRectList
    try {
      expect(fillLineBoxes([heading, para], geoms, 1)).toBe(true)
    } finally {
      Range.prototype.getClientRects = orig
    }
    expect(para.lineBoxes).toEqual([
      { offsetInBlock: 0, height: 20 },
      { offsetInBlock: 20, height: 20 },
    ])
  })
})

describe('measureBlocks — page-anchored floated tables', () => {
  it('reads the tblp target and strips the applied shift back to the natural position', () => {
    const rect = {
      top: 140,
      height: 60,
      bottom: 200,
      left: 0,
      right: 100,
      width: 100,
      x: 0,
      y: 140,
      toJSON: () => ({}),
    } as DOMRect
    const pm = document.createElement('div')
    const tbl = document.createElement('table')
    tbl.className = 'doc-table-float-left'
    tbl.dataset.tblpVy = '150'
    tbl.dataset.tblpVanchor = 'page'
    tbl.dataset.tblpDy = '40'
    tbl.innerHTML = '<tbody><tr><td>x</td></tr></tbody>'
    tbl.getBoundingClientRect = () => rect
    pm.appendChild(tbl)
    const { blocks } = measureBlocks(pm, 0, 1)
    expect(blocks[0]).toMatchObject({
      top: 100,
      floated: true,
      pageRelVyPx: 150,
      pageRelVAnchor: 'page',
    })
  })
})

describe('measureBlocks — break-only paragraphs', () => {
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
  // page height 200; the break paragraph's DOM height is two line boxes (br + trailingBreak) = 44px
  const breakDoc = (fillerH: number, gap = 0) => {
    const pm = document.createElement('div')
    const addPara = (top: number, height: number, html: string) => {
      const el = document.createElement('p')
      el.innerHTML = html
      el.getBoundingClientRect = () => rectOf(top, height)
      pm.appendChild(el)
    }
    addPara(0, fillerH, 'filler text')
    addPara(fillerH + gap, 44, '<br class="doc-page-br"><br class="ProseMirror-trailingBreak">')
    addPara(fillerH + gap + 44, 100, 'after the break')
    return measureBlocks(pm, 0, 1)
  }
  const geoms = [{ contentHeight: 200, forceBreak: false }]

  // Word probe 20260901 (break-only fit matrix): the break line must FULLY fit
  // below the preceding content — an exact-12pt break line fits at exactly 12pt
  // remaining and blanks at 11pt (no partial absorb into the bottom margin),
  // trailing space-after is charged (13pt remaining + 8pt space-after blanks),
  // and auto line-spacing multiples above 1 are not charged (a double-spaced
  // Calibri 11pt break line absorbs at ~14pt remaining).
  it('absorbs at the page bottom when the break line fits (no blank page)', () => {
    const { blocks, totalHeight } = breakDoc(173) // 27px left >= the 22px line
    expect(blocks[1].breakAfter).toBe(true)
    // one line's share of the two DOM line boxes (br + trailingBreak)
    expect(blocks[1].breakOnlyLineH).toBe(22)
    const slices = computeSectionedSlicesF2(blocks, geoms, totalHeight)
    expect(slices.map((s) => s.start)).toEqual([0, 217])
  })

  it('opens a Word-style blank page when the previous paragraph exactly fills the page', () => {
    const { blocks, totalHeight } = breakDoc(200)
    const slices = computeSectionedSlicesF2(blocks, geoms, totalHeight)
    // page 2 holds only the break paragraph (blank), which then pushes the rest to page 3
    expect(slices.map((s) => s.start)).toEqual([0, 200, 244])
    expect(slices[1]).toMatchObject({ start: 200, end: 244 })
  })

  it('opens a blank page when less than the full break line remains (probe: 11pt of 12pt blanks)', () => {
    const { blocks, totalHeight } = breakDoc(185) // 15px left < the 22px line
    const slices = computeSectionedSlicesF2(blocks, geoms, totalHeight)
    expect(slices.map((s) => s.start)).toEqual([0, 185, 229])
  })

  it('judges the fit by a single line share, not the phantom-inflated DOM height', () => {
    const { blocks, totalHeight } = breakDoc(170) // 30px left: one 22px line fits, the 44px box does not
    const slices = computeSectionedSlicesF2(blocks, geoms, totalHeight)
    expect(slices.map((s) => s.start)).toEqual([0, 214])
  })

  it('charges the previous block trailing space-after in the page-bottom fit', () => {
    // filler text ends at 180; its 15px space-after folds into usedInCol and Word
    // charges it (probe: 13pt remaining with an 8pt space-after still blanks)
    const { blocks, totalHeight } = breakDoc(180, 15)
    expect(blocks[0].spaceAfterPx).toBe(15)
    const slices = computeSectionedSlicesF2(blocks, geoms, totalHeight)
    expect(slices.map((s) => s.start)).toEqual([0, 195, 239])
  })

  it('charges footnote reservations in the page-bottom fit', () => {
    // reservations ride the height like the trailing space: the page stays full
    // and keeps its deliberate blank page
    const { blocks, totalHeight } = breakDoc(180, 15)
    blocks[0].footnoteExtraPx = 15
    blocks[0].height += 15 // reservation rides the height (applyBlockMeta)
    const slices = computeSectionedSlicesF2(blocks, geoms, totalHeight)
    expect(slices.map((s) => s.start)).toEqual([0, 195, 239])
  })

  it('divides an auto line-spacing multiple out of the fit height (probe: multiples are not charged)', () => {
    const pm = document.createElement('div')
    const addPara = (top: number, height: number, html: string, style = '') => {
      const el = document.createElement('p')
      el.innerHTML = html
      if (style) el.setAttribute('style', style)
      el.getBoundingClientRect = () => rectOf(top, height)
      pm.appendChild(el)
    }
    addPara(0, 170, 'filler text')
    // double-spaced break paragraph: 88px DOM box (2 line boxes x 44), natural line 22
    addPara(
      170,
      88,
      '<br class="doc-page-br"><br class="ProseMirror-trailingBreak">',
      '--doc-line-mult:2',
    )
    addPara(258, 100, 'after the break')
    const { blocks, totalHeight } = measureBlocks(pm, 0, 1)
    expect(blocks[1].breakOnlyLineH).toBe(22)
    const slices = computeSectionedSlicesF2(blocks, geoms, totalHeight)
    // 30px left fits the 22px natural line: absorbed, no blank page
    expect(slices.map((s) => s.start)).toEqual([0, 258])
  })

  it('keeps the full exact-rule line height in the fit (probe: exact lines demand their box)', () => {
    const pm = document.createElement('div')
    const addPara = (top: number, height: number, html: string, cls = '') => {
      const el = document.createElement('p')
      el.innerHTML = html
      if (cls) el.className = cls
      el.getBoundingClientRect = () => rectOf(top, height)
      pm.appendChild(el)
    }
    addPara(0, 170, 'filler text')
    addPara(
      170,
      88,
      '<br class="doc-page-br"><br class="ProseMirror-trailingBreak">',
      'doc-lh-fixed',
    )
    addPara(258, 100, 'after the break')
    const { blocks, totalHeight } = measureBlocks(pm, 0, 1)
    expect(blocks[1].breakOnlyLineH).toBe(44)
    const slices = computeSectionedSlicesF2(blocks, geoms, totalHeight)
    // 30px left < the 44px exact line: deliberate blank page
    expect(slices.map((s) => s.start)).toEqual([0, 170, 258])
  })

  it('keeps the box of a style-level fixed line (--doc-line-fixed marker, no class)', () => {
    // style-level exact/atLeast lines carry no doc-lh-fixed class; doc-style-css
    // marks them with --doc-line-fixed so an inherited document auto multiple is
    // not divided out (Word probe 20260901: a style-level exact break line
    // demands its full box like a direct one)
    const pm = document.createElement('div')
    const el = document.createElement('p')
    el.innerHTML = '<br class="doc-page-br"><br class="ProseMirror-trailingBreak">'
    el.setAttribute('style', '--doc-line-fixed:1')
    el.getBoundingClientRect = () => rectOf(0, 88)
    pm.appendChild(el)
    const { blocks } = measureBlocks(pm, 0, 1)
    expect(blocks[0].breakOnlyLineH).toBe(44)
  })

  it('lets a direct auto multiple override a fixed-line style in the fit height', () => {
    const pm = document.createElement('div')
    const el = document.createElement('p')
    el.innerHTML = '<br class="doc-page-br"><br class="ProseMirror-trailingBreak">'
    el.setAttribute('style', '--doc-line-fixed:1;--doc-line-mult:2')
    el.getBoundingClientRect = () => rectOf(0, 88)
    pm.appendChild(el)
    const { blocks } = measureBlocks(pm, 0, 1)
    expect(blocks[0].breakOnlyLineH).toBe(22)
  })

  it('applyBlockMeta charges footnotes through the height, never the space-after', () => {
    const blocks = [{ top: 0, height: 100, docxIndex: 0, spaceAfterPx: 5 }]
    applyBlockMeta(blocks, () => ({ footnoteExtraPx: 12 }))
    expect(blocks[0]).toMatchObject({ height: 112, spaceAfterPx: 5, footnoteExtraPx: 12 })
  })

  it('applyBlockMeta resolves footnote bands at their marker offsets', () => {
    const el = document.createElement('p')
    el.getBoundingClientRect = () => ({ top: 100, height: 80 }) as DOMRect
    for (const top of [110, 152]) {
      const sup = document.createElement('sup')
      sup.className = 'doc-note-ref'
      sup.setAttribute('data-note-ref', String(top))
      sup.setAttribute('data-note-kind', 'footnote')
      sup.getBoundingClientRect = () => ({ top, height: 8 }) as DOMRect
      el.appendChild(sup)
    }
    const blocks: BlockBox[] = [{ top: 0, height: 80, docxIndex: 0, el }]
    applyBlockMeta(blocks, () => ({
      footnoteExtraPx: 30,
      footnoteBands: [{ heightPx: 10 }, { heightPx: 20 }],
    }))
    expect(blocks[0].noteBands).toEqual([
      { offset: 10, height: 10 },
      { offset: 52, height: 20 },
    ])
    expect(blocks[0].height).toBe(110)
  })

  it('a split paragraph charges each note on the page holding its reference line', () => {
    // 4 lines x 40 after a 100px filler; the line at offset 120 carries a 60px
    // note → its page must reserve line + note, splitting the paragraph earlier
    const filler = block(0, 100)
    const lines = [0, 40, 80, 120].map((off) => ({ offsetInBlock: off, height: 40 }))
    const b = block(100, 160 + 60, {
      lineBoxes: lines,
      footnoteExtraPx: 60,
      noteBands: [{ offset: 130, height: 60 }],
    })
    const geoms = [{ contentHeight: 300, forceBreak: false }]
    const slices = computeSectionedSlicesF2([filler, b], geoms, 260)
    // lines 0-1 on page 1 (widow rule keeps 2 on the next page), lines 2-3 + note on page 2
    expect(slices.map((s) => s.start)).toEqual([0, 180])
  })

  it('a reference line whose note cannot fit moves to the next page with its note', () => {
    const filler = block(0, 150)
    const lines = [0, 30].map((off) => ({ offsetInBlock: off, height: 30 }))
    const para = block(150, 60 + 50, {
      lineBoxes: lines,
      widowControl: false,
      footnoteExtraPx: 50,
      noteBands: [{ offset: 0, height: 50 }],
    })
    const geoms = [{ contentHeight: 200, forceBreak: false }]
    // text alone would fit (150+60 <= 200 - separator? no notes on page 1), but
    // line 0 demands its 50px note area → the whole paragraph turns the page
    const slices = computeSectionedSlicesF2([filler, para], geoms, 210)
    expect(slices.map((s) => s.start)).toEqual([0, 150])
  })
})

describe('cellCutYs — line-level in-row cut candidates', () => {
  it('emits a boundary between adjacent lines even with zero gap', () => {
    const lines: Array<[number, number]> = [
      [0, 15],
      [15, 30],
      [30, 45],
    ]
    expect(cellCutYs([lines], 45)).toEqual([15, 30])
  })

  it('clusters same-line rects (overlapping spans) into one line box', () => {
    const rects: Array<[number, number]> = [
      [0, 15],
      [1, 14],
      [15, 30],
    ]
    expect(cellCutYs([rects], 30)).toEqual([15])
  })

  it('drops a candidate that crosses another cell line box', () => {
    const cellA: Array<[number, number]> = [
      [0, 15],
      [15, 30],
    ]
    const cellB: Array<[number, number]> = [[5, 25]]
    expect(cellCutYs([cellA, cellB], 30)).toEqual([])
  })

  it('keeps only boundaries safe across all cells and dedupes near-identical ones', () => {
    const cellA: Array<[number, number]> = [
      [0, 20],
      [20, 60],
    ]
    const cellB: Array<[number, number]> = [
      [0, 20.3],
      [20.3, 40],
      [40, 60],
    ]
    // 20/20.3 coincide within jitter → one cut; B's 40 falls inside A's [20,60] line
    expect(cellCutYs([cellA, cellB], 60)).toEqual([20])
  })

  it('rejects candidates hugging the row edges', () => {
    const lines: Array<[number, number]> = [
      [0, 1.5],
      [2, 28],
      [28.5, 30],
    ]
    expect(cellCutYs([lines], 30)).toEqual([])
  })

  it('returns nothing for empty or single-line rows', () => {
    expect(cellCutYs([], 100)).toEqual([])
    expect(cellCutYs([[[0, 20]]], 100)).toEqual([])
  })

  it('widow/orphan: never splits a two-line paragraph (whole row pushes instead)', () => {
    const lines: Array<[number, number]> = [
      [0, 15],
      [15, 30],
    ]
    expect(cellCutYs([lines], 30, [lines])).toEqual([])
  })

  it('widow/orphan: a four-line paragraph only cuts at its midpoint', () => {
    const lines: Array<[number, number]> = [
      [0, 15],
      [15, 30],
      [30, 45],
      [45, 60],
    ]
    expect(cellCutYs([lines], 60, [lines])).toEqual([30])
  })

  it('widow/orphan: a paragraph boundary between two short paragraphs stays cuttable', () => {
    const paraA: Array<[number, number]> = [
      [0, 15],
      [15, 30],
    ]
    const paraB: Array<[number, number]> = [
      [30, 45],
      [45, 60],
    ]
    expect(cellCutYs([[...paraA, ...paraB]], 60, [paraA, paraB])).toEqual([30])
  })
})

describe('insertParityBlanks — even/odd section blank pages', () => {
  const geoms = (types: Array<'nextPage' | 'evenPage' | 'oddPage' | 'continuous'>) =>
    types.map((t) => ({ contentHeight: 800, forceBreak: t !== 'continuous', startType: t }))

  it('inserts a blank page when an evenPage section starts on an odd physical page', () => {
    const slices = [
      { start: 0, end: 500, section: 0 },
      { start: 500, end: 900, section: 1 },
    ]
    const out = insertParityBlanks(slices, geoms(['nextPage', 'evenPage']))
    // Section 2 should start on page 2 (even) → already even, no insert; a 3-page scenario verifies insertion
    expect(out).toHaveLength(2)
    const slices3 = [
      { start: 0, end: 500, section: 0 },
      { start: 500, end: 900, section: 0 },
      { start: 900, end: 1200, section: 1 },
    ]
    const out3 = insertParityBlanks(slices3, geoms(['nextPage', 'evenPage']))
    expect(out3).toHaveLength(4)
    expect(out3[2]).toEqual({ start: 900, end: 900, section: 0 })
    expect(out3[3].section).toBe(1)
  })

  it('inserts a blank page when an oddPage section starts on an even physical page', () => {
    const slices = [
      { start: 0, end: 500, section: 0 },
      { start: 500, end: 900, section: 1 },
    ]
    const out = insertParityBlanks(slices, geoms(['nextPage', 'oddPage']))
    expect(out).toHaveLength(3)
    expect(out[1]).toEqual({ start: 500, end: 500, section: 0 })
  })

  it('returns input unchanged when there are no even/odd sections', () => {
    const slices = [{ start: 0, end: 500, section: 0 }]
    expect(insertParityBlanks(slices, geoms(['nextPage']))).toBe(slices)
  })
})

describe('sectionGeoms — header/footer pushing the body', () => {
  const sec = (): SectionInfo =>
    ({
      settings: {
        pageWidth: 12240,
        pageHeight: 15840,
        orientation: 'portrait',
        marginTop: 1440,
        marginRight: 1440,
        marginBottom: 1440,
        marginLeft: 1440,
        pageBorder: false,
        columns: 1,
        headerDist: 720,
        footerDist: 720,
      },
      startType: 'nextPage',
      firstBlockIndex: 0,
      lastBlockIndex: 0,
      sectPrXml: '',
      titlePg: false,
      headerRefs: {},
      footerRefs: {},
    }) as SectionInfo

  it('header within the top margin: capacity unchanged', () => {
    // marginTop 96px; headerDist 48px + header 30px = 78 < 96
    const [g] = sectionGeoms([sec()], [{ headerPx: 30, footerPx: 0 }])
    expect(g.contentHeight).toBeCloseTo(864, 0)
  })

  it('oversized header pushes the body down: capacity shrinks', () => {
    // headerDist 48 + header 100 = 148 > marginTop 96 → capacity 1056-148-96 = 812
    const [g] = sectionGeoms([sec()], [{ headerPx: 100, footerPx: 0 }])
    expect(g.contentHeight).toBeCloseTo(812, 0)
  })

  it('oversized footer pushes the body up', () => {
    const [g] = sectionGeoms([sec()], [{ headerPx: 0, footerPx: 100 }])
    expect(g.contentHeight).toBeCloseTo(812, 0)
  })
})

describe('sectionPageBox — editor/preview physical page parity', () => {
  it('uses each section page height, orientation, and footer distance', () => {
    const a4 = sectionPageBox(sec({ pageWidth: 11906, pageHeight: 16838 }).settings)
    const letterLandscape = sectionPageBox(
      sec({ pageWidth: 15840, pageHeight: 12240, orientation: 'landscape', footerDist: 360 })
        .settings,
    )
    expect(a4.height).toBeCloseTo((16838 / 1440) * 96)
    expect(letterLandscape.width).toBeCloseTo(1056)
    expect(letterLandscape.height).toBeCloseTo(816)
    expect(letterLandscape.footerDist).toBeCloseTo(24)
  })
})

describe('tableHeaderFlags', () => {
  it('parses the tblHeader flag of each tr', () => {
    const xml =
      '<w:tbl><w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc/></w:tr>' +
      '<w:tr><w:tc/></w:tr>' +
      '<w:tr><w:trPr><w:tblHeader w:val="false"/></w:trPr><w:tc/></w:tr></w:tbl>'
    expect(tableHeaderFlags(xml)).toEqual([true, false, false])
  })

  it('tableRowFlags also parses cantSplit', () => {
    const xml =
      '<w:tbl><w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr><w:tc/></w:tr>' +
      '<w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc/></w:tr></w:tbl>'
    expect(tableRowFlags(xml)).toEqual([
      { isHeader: true, cantSplit: true },
      { isHeader: false, cantSplit: true },
    ])
  })

  it('tableRowFlags parses atLeast trHeight into minHPx (exact/auto-only rows excluded)', () => {
    const xml =
      '<w:tbl><w:tr><w:trPr><w:trHeight w:val="1500"/></w:trPr><w:tc/></w:tr>' +
      '<w:tr><w:trPr><w:trHeight w:val="1500" w:hRule="atLeast"/></w:trPr><w:tc/></w:tr>' +
      '<w:tr><w:trPr><w:trHeight w:val="1500" w:hRule="exact"/></w:trPr><w:tc/></w:tr>' +
      '<w:tr><w:trPr><w:trHeight w:val="99999"/></w:trPr><w:tc/></w:tr>' +
      '<w:tr><w:tc/></w:tr></w:tbl>'
    expect(tableRowFlags(xml)).toEqual([
      { isHeader: false, cantSplit: false, minHPx: 100 },
      { isHeader: false, cantSplit: false, minHPx: 100 },
      { isHeader: false, cantSplit: false },
      // Word clamps trHeight to 31680 twips / 22in (MS-OI29500 2.1.51)
      { isHeader: false, cantSplit: false, minHPx: 2112 },
      { isHeader: false, cantSplit: false },
    ])
  })
})

describe('applyBlockMeta / trailing spacing overflow', () => {
  it('applyBlockMeta injects keepNext/keepLines/widowControl by docxIndex', () => {
    const blocks: BlockBox[] = [
      { top: 0, height: 50, docxIndex: 0 },
      { top: 50, height: 50, docxIndex: 1 },
      { top: 100, height: 50 }, // no docxIndex, untouched
    ]
    applyBlockMeta(blocks, (idx) =>
      idx === 0
        ? { keepNext: true, widowControl: false }
        : idx === 1
          ? { keepLines: true, breakBefore: true }
          : undefined,
    )
    expect(blocks[0].keepNext).toBe(true)
    expect(blocks[0].widowControl).toBe(false)
    expect(blocks[1].keepLines).toBe(true)
    expect(blocks[1].breakBefore).toBe(true)
    expect(blocks[2].keepNext).toBeUndefined()
  })

  it('breakBefore meta (style-level pageBreakBefore) forces a page break, first block excepted', () => {
    const blocks: BlockBox[] = [
      { top: 0, height: 100, docxIndex: 0 },
      { top: 100, height: 100, docxIndex: 1 },
    ]
    // both paragraphs use a pageBreakBefore style; the document's first block must not open an empty page
    applyBlockMeta(blocks, () => ({ breakBefore: true }))
    const slices = computePageSlices(blocks, 800, 200)
    expect(slices).toEqual([
      { start: 0, end: 100, section: 0 },
      { start: 100, end: 200, section: 0 },
    ])
  })

  it('trailing spaceAfter overflow does not push to the next page (Word breaks by text only)', () => {
    // Page height 100: block text 60+38=98 fits; b's 10px space-after overflows — b should stay on page 1
    const a: BlockBox = { top: 0, height: 60 }
    const b: BlockBox = { top: 60, height: 48, spaceAfterPx: 10 }
    const c: BlockBox = { top: 108, height: 50 }
    const slices = computeSectionedSlicesF2(
      [a, b, c],
      [{ contentHeight: 100, forceBreak: false }],
      158,
    )
    expect(slices.length).toBe(2)
    expect(slices[1].start).toBe(108) // break before c, not before b
  })
})

describe('formatPageNumber', () => {
  it('supports each numeric format', async () => {
    const { formatPageNumber } = await import('../src/renderer/pagination')
    expect(formatPageNumber(3)).toBe('3')
    expect(formatPageNumber(3, 'numberInDash')).toBe('- 3 -')
    expect(formatPageNumber(3, 'lowerRoman')).toBe('iii')
    expect(formatPageNumber(49, 'upperRoman')).toBe('XLIX')
    expect(formatPageNumber(1, 'lowerLetter')).toBe('a')
    expect(formatPageNumber(27, 'upperLetter')).toBe('AA')
    expect(formatPageNumber(3, 'chineseCounting')).toBe('三')
    expect(formatPageNumber(21, 'chineseCounting')).toBe('二十一')
    expect(formatPageNumber(10, 'chineseCounting')).toBe('十')
  })
})

describe('appendEndnotesBlock — endnote layout', () => {
  const item = (id: string, height: number) => ({ no: 1, id, text: 'x', height })

  it('returns null without endnotes; block list unchanged', () => {
    const blocks: BlockBox[] = [{ top: 0, height: 100 }]
    expect(appendEndnotesBlock(blocks, 100, [], 16)).toBeNull()
    expect(blocks).toHaveLength(1)
  })

  it('endnotes fit on the last page: they follow the body on the same page', () => {
    const blocks: BlockBox[] = [{ top: 0, height: 500 }]
    const r = appendEndnotesBlock(blocks, 500, [item('1', 40), item('2', 40)], 16)!
    expect(r.top).toBe(500)
    expect(r.totalHeight).toBe(500 + 16 + 80)
    const slices = computeSectionedSlicesF2(
      blocks,
      [{ contentHeight: 800, forceBreak: false }],
      r.totalHeight,
    )
    expect(slices).toHaveLength(1)
  })

  it('endnotes do not fit on the last page: break between entries and continue on the next page', () => {
    const blocks: BlockBox[] = [{ top: 0, height: 760 }]
    // Page height 800: first item (16+40=56) fits? (760+56=816>800) → no, pushed to the next page
    const r = appendEndnotesBlock(blocks, 760, [item('1', 40), item('2', 40)], 16)!
    const slices = computeSectionedSlicesF2(
      blocks,
      [{ contentHeight: 800, forceBreak: false }],
      r.totalHeight,
    )
    expect(slices).toHaveLength(2)
    expect(slices[1].start).toBeGreaterThanOrEqual(760)
  })

  it('endnote area taller than a page: splits at entry boundaries (widow off, any gap between entries may break)', () => {
    const blocks: BlockBox[] = [{ top: 0, height: 100 }]
    const items = Array.from({ length: 10 }, (_, i) => item(String(i), 100))
    const r = appendEndnotesBlock(blocks, 100, items, 16)!
    const slices = computeSectionedSlicesF2(
      blocks,
      [{ contentHeight: 400, forceBreak: false }],
      r.totalHeight,
    )
    expect(slices.length).toBeGreaterThan(2)
    // Every cut point lands on an item boundary
    const boundaries = new Set<number>()
    let off = r.top
    items.forEach((it, i) => {
      boundaries.add(off)
      off += (i === 0 ? 16 : 0) + it.height
    })
    for (const s of slices.slice(1)) {
      expect([...boundaries].some((b) => Math.abs(b - s.start) < 0.01)).toBe(true)
    }
  })

  it('inherits the section of the last block', () => {
    const blocks: BlockBox[] = [{ top: 0, height: 100, section: 2 }]
    appendEndnotesBlock(blocks, 100, [item('1', 40)], 16)
    expect(blocks[1].section).toBe(2)
    expect(blocks[1].isEndnotes).toBe(true)
  })
})

describe('computeSectionedSlicesF2 — multi-column flow', () => {
  const twoCol = [{ contentHeight: 200, forceBreak: false, cols: 2 }]

  it('single-column document outputs no regions (compatible with existing consumers)', () => {
    const slices = computeSectionedSlicesF2([lineBlock(0, [50, 50])], geoms1, 100)
    expect(slices[0].regions).toBeUndefined()
  })

  it('content shorter than a column: one region per page, first column filled, second column empty', () => {
    const slices = computeSectionedSlicesF2([lineBlock(0, [50, 50])], twoCol, 100)
    expect(slices).toHaveLength(1)
    expect(slices[0].regions).toHaveLength(1)
    const cols = slices[0].regions![0].columns
    expect(cols).toHaveLength(1) // no overflow, no second column opened
    expect(cols[0]).toMatchObject({ start: 0, end: 100 })
  })

  it('overflow moves to the next column: page capacity = column count × column height', () => {
    // 6 lines × 50 = 300 > column height 200; from line 5 into the second column; total < 400, no page turn
    const slices = computeSectionedSlicesF2([lineBlock(0, [50, 50, 50, 50, 50, 50])], twoCol, 300)
    expect(slices).toHaveLength(1)
    const cols = slices[0].regions![0].columns
    expect(cols).toHaveLength(2)
    expect(cols[0]).toMatchObject({ start: 0, end: 200 })
    expect(cols[1]).toMatchObject({ start: 200, end: 300 })
  })

  it('overflow in the last column turns the page', () => {
    // 10 lines × 50 = 500 > 2 × 200; page 2 starts at 400
    const slices = computeSectionedSlicesF2(
      [lineBlock(0, [50, 50, 50, 50, 50, 50, 50, 50, 50, 50])],
      twoCol,
      500,
    )
    expect(slices).toHaveLength(2)
    expect(slices[0]).toMatchObject({ start: 0, end: 400 })
    expect(slices[1].start).toBe(400)
    expect(slices[1].regions![0].columns[0]).toMatchObject({ start: 400, end: 500 })
  })

  it('column boundaries also honor orphan/widow constraints', () => {
    // Previous block uses 150; the 4-line paragraph has room for only 1 line → pushed whole into the second column (no orphan)
    const blocks = [lineBlock(0, [50, 50, 50]), lineBlock(150, [50, 50, 50, 50])]
    const slices = computeSectionedSlicesF2(blocks, twoCol, 350)
    expect(slices).toHaveLength(1)
    const cols = slices[0].regions![0].columns
    expect(cols[1].start).toBe(150)
  })

  it('pageBreakBefore inside a column turns the page (not just the column)', () => {
    const blocks = [lineBlock(0, [50, 50]), lineBlock(100, [50], { breakBefore: true })]
    const slices = computeSectionedSlicesF2(blocks, twoCol, 150)
    expect(slices).toHaveLength(2)
    expect(slices[1].start).toBe(100)
  })

  it('colBreakAfter moves to the next column; turns the page at the last column', () => {
    const blocks = [
      lineBlock(0, [50], { colBreakAfter: true }),
      lineBlock(50, [50], { colBreakAfter: true }),
      lineBlock(100, [50]),
    ]
    const slices = computeSectionedSlicesF2(blocks, twoCol, 150)
    // Block 1 → column 1, block 2 → column 2, block 3 → page 2 column 1
    expect(slices).toHaveLength(2)
    expect(slices[0].regions![0].columns.map((c) => c.start)).toEqual([0, 50])
    expect(slices[1].start).toBe(100)
  })

  it('continuous with a changed column count: opens a new region in the remaining page height', () => {
    // Section 0 single column (60px title), section 1 two columns: region top 60, column height 140
    const blocks: BlockBox[] = [
      { ...lineBlock(0, [60]), section: 0 },
      { ...lineBlock(60, [50, 50, 50, 50]), section: 1 },
    ]
    const geoms = [
      { contentHeight: 200, forceBreak: false },
      { contentHeight: 200, forceBreak: false, cols: 2 },
    ]
    const slices = computeSectionedSlicesF2(blocks, geoms, 260)
    expect(slices).toHaveLength(1)
    const regions = slices[0].regions!
    expect(regions).toHaveLength(2)
    expect(regions[0]).toMatchObject({ top: 0, height: 200, section: 0 })
    expect(regions[1].top).toBe(60)
    expect(regions[1].height).toBe(140)
    // Two-column region: 4 lines 200px > column height 140 → from line 3 (offset 60+100=160) into the second column
    expect(regions[1].columns).toHaveLength(2)
    expect(regions[1].columns[1].start).toBe(160)
  })

  it('a region opened on a shorter host page is bounded by the host capacity (prod-sas 043)', () => {
    // Section 0's first (titlePg) page holds only 150px; the continuous
    // two-column section declares 200 — its region on the host page must end
    // at the host footer (height 90 = 150 - 60), not run 140px past it
    const blocks: BlockBox[] = [
      { ...lineBlock(0, [60]), section: 0 },
      { ...lineBlock(60, [40, 40, 40, 40, 40, 40]), section: 1, widowControl: false },
    ]
    const geoms = [
      { contentHeight: 200, firstContentHeight: 150, forceBreak: false },
      { contentHeight: 200, forceBreak: false, cols: 2 },
    ]
    const slices = computeSectionedSlicesF2(blocks, geoms, 300)
    const regions = slices[0].regions!
    expect(regions[1].top).toBe(60)
    expect(regions[1].height).toBe(90)
    // 2 lines per 90px column (80 ≤ 90): col2 starts after line 2, page 2 holds the rest
    expect(regions[1].columns[1].start).toBe(140)
    expect(slices).toHaveLength(2)
    expect(slices[1].start).toBe(220)
    expect(slices[1].regions![0]).toMatchObject({ top: 0, height: 200 })
  })

  it('a host-page region already at the host footer turns the page instead', () => {
    // single-column content fills the 150px host page exactly: the two-column
    // region has no room left on it (with the new section's 200 it would open
    // a 50px region running past the host footer)
    const blocks: BlockBox[] = [
      { ...lineBlock(0, [150]), section: 0 },
      { ...lineBlock(150, [40, 40]), section: 1 },
    ]
    const geoms = [
      { contentHeight: 200, firstContentHeight: 150, forceBreak: false },
      { contentHeight: 200, forceBreak: false, cols: 2 },
    ]
    const slices = computeSectionedSlicesF2(blocks, geoms, 230)
    expect(slices).toHaveLength(2)
    expect(slices[1].start).toBe(150)
  })

  it('over-column blocks without line data advance instead of stacking in one column', () => {
    // 3 line-less blocks of 150 in 100px columns: one per column, page turn
    // after the second (they must not all pile into the first column)
    const blocks = [block(0, 150), block(150, 150), block(300, 150)]
    const slices = computeSectionedSlicesF2(
      blocks,
      [{ contentHeight: 100, forceBreak: false, cols: 2 }],
      450,
    )
    expect(slices).toHaveLength(2)
    expect(slices[0].regions![0].columns.map((c) => c.start)).toEqual([0, 150])
    expect(slices[1].start).toBe(300)
  })

  it('nextPage break into a multi-column section: full-page column flow', () => {
    const blocks: BlockBox[] = [
      { ...lineBlock(0, [50]), section: 0 },
      { ...lineBlock(50, [50, 50, 50, 50, 50]), section: 1 },
    ]
    const geoms = [
      { contentHeight: 200, forceBreak: false },
      { contentHeight: 200, forceBreak: true, cols: 2 },
    ]
    const slices = computeSectionedSlicesF2(blocks, geoms, 300)
    expect(slices).toHaveLength(2)
    expect(slices[0].regions).toBeUndefined()
    const cols = slices[1].regions![0].columns
    expect(cols[0]).toMatchObject({ start: 50, end: 250 })
    expect(cols[1]).toMatchObject({ start: 250, end: 300 })
  })

  it('tblHeader table split across columns: header repeats at the top of the column', () => {
    const rows: TableRowBox[] = [
      { height: 30, isHeader: true },
      ...Array.from({ length: 10 }, () => ({ height: 30 })),
    ]
    const b: BlockBox = { top: 0, height: 330, tableRows: rows }
    const slices = computeSectionedSlicesF2([b], twoCol, 330)
    expect(slices).toHaveLength(1)
    const cols = slices[0].regions![0].columns
    expect(cols.length).toBe(2)
    expect(cols[1].repeatHeader).toMatchObject({ top: 0, height: 30 })
  })

  it('sectionGeoms: nextColumn with the same multi-column count advances a column; a changed count breaks the page (tdf135343)', () => {
    const fourCol = sec({ columns: 4 })
    const twoColRtl = sec({ columns: 2 }, { startType: 'nextColumn' })
    const geoms = sectionGeoms([fourCol, twoColRtl])
    expect(geoms[1].forceBreak).toBe(true) // 4 → 2: acts like a page break (c12v3)
    expect(geoms[1].colBreakStart).toBeUndefined()

    const threeCol = sec({ columns: 3 }, { startType: 'continuous' })
    const threeColNext = sec({ columns: 3 }, { startType: 'nextColumn' })
    const geoms2 = sectionGeoms([threeCol, threeColNext])
    expect(geoms2[1].forceBreak).toBe(false) // 3 → 3: column advance (c14/c15)
    expect(geoms2[1].colBreakStart).toBe(true)
  })

  it('colBreakStart advances one column at the section boundary (0876 shape)', () => {
    const blocks: BlockBox[] = [
      { ...lineBlock(0, [40, 40]), section: 0 },
      { ...lineBlock(80, [40]), section: 1 },
    ]
    const geoms = [
      { contentHeight: 200, forceBreak: false, cols: 3 },
      { contentHeight: 200, forceBreak: false, cols: 3, colBreakStart: true },
    ]
    const slices = computeSectionedSlicesF2(blocks, geoms, 120)
    expect(slices).toHaveLength(1)
    const cols = slices[0].regions![0].columns
    expect(cols).toHaveLength(2)
    expect(cols[1]).toMatchObject({ start: 80, end: 120 })
  })

  it('continuous column-count change balances the closed region at line granularity (0089 shape)', () => {
    // 1-col title (60), 4-col index of 6 short lines (20 each), 1-col body: the
    // index balances 2/2/2/0, so the body region starts 40px below the index top
    const blocks: BlockBox[] = [
      { ...lineBlock(0, [60]), section: 0 },
      ...Array.from({ length: 6 }, (_, i) => ({
        ...lineBlock(60 + i * 20, [20]),
        section: 1,
      })),
      { ...lineBlock(180, [50]), section: 2 },
    ]
    const geoms = [
      { contentHeight: 400, forceBreak: false },
      { contentHeight: 400, forceBreak: false, cols: 4 },
      { contentHeight: 400, forceBreak: false },
    ]
    const slices = computeSectionedSlicesF2(blocks, geoms, 230)
    expect(slices).toHaveLength(1)
    const regions = slices[0].regions!
    expect(regions).toHaveLength(3)
    const idx = regions[1]
    expect(idx.height).toBe(40)
    expect(idx.columns.map((c) => [c.start, c.end])).toEqual([
      [60, 100],
      [100, 140],
      [140, 180],
      [180, 180],
    ])
    // the body region opens right under the balanced index
    expect(regions[2].top).toBe(100)
    expect(slices[0].physHeight).toBe(150) // 100 + 50 body line
  })

  it('natural column overflow still balances on a continuous close (0876 shape)', () => {
    // five 2-line paragraphs (10 lines) in a 2-col region of height 160: col1
    // naturally overflows at 8 lines, but the continuous single-column close
    // re-balances to the line quota (5) — widow/orphan atomicity keeps the
    // 2-line paragraph whole, so the cut lands at the next paragraph top (6/4)
    const blocks: BlockBox[] = [
      ...Array.from({ length: 5 }, (_, i) => ({
        ...lineBlock(i * 40, [20, 20]),
        section: 0,
      })),
      { ...lineBlock(200, [30]), section: 1 },
    ]
    const geoms = [
      { contentHeight: 160, forceBreak: false, cols: 2 },
      { contentHeight: 160, forceBreak: false },
    ]
    const slices = computeSectionedSlicesF2(blocks, geoms, 230)
    expect(slices).toHaveLength(1)
    const regions = slices[0].regions!
    expect(regions[0].columns.map((c) => [c.start, c.end])).toEqual([
      [0, 120],
      [120, 200],
    ])
    expect(regions[0].height).toBe(120)
    expect(regions[1].top).toBe(120)
  })

  it('balance splits a long paragraph at a widow/orphan-safe line boundary', () => {
    // one 10-line paragraph, 2 columns: quota 5, k=5 leaves 5 lines each side
    const blocks: BlockBox[] = [
      {
        ...lineBlock(
          0,
          Array.from({ length: 10 }, () => 20),
        ),
        section: 0,
      },
      { ...lineBlock(200, [30]), section: 1 },
    ]
    const geoms = [
      { contentHeight: 400, forceBreak: false, cols: 2 },
      { contentHeight: 400, forceBreak: false },
    ]
    const regions = computeSectionedSlicesF2(blocks, geoms, 230)[0].regions!
    expect(regions[0].columns.map((c) => [c.start, c.end])).toEqual([
      [0, 100],
      [100, 200],
    ])
  })

  it('empty paragraphs flow into balanced columns but do not count toward the quota', () => {
    // 4 text lines + 2 trailing empties, 2 cols: quota 2 → cut after 2 text lines
    const blocks: BlockBox[] = [
      { ...lineBlock(0, [20]), section: 0 },
      { ...lineBlock(20, [20]), section: 0 },
      { ...lineBlock(40, [20]), section: 0 },
      { ...lineBlock(60, [20]), section: 0 },
      { ...lineBlock(80, [20]), section: 0, emptyPara: true },
      { ...lineBlock(100, [20]), section: 0, emptyPara: true },
      { ...lineBlock(120, [30]), section: 1 },
    ]
    const geoms = [
      { contentHeight: 400, forceBreak: false, cols: 2 },
      { contentHeight: 400, forceBreak: false },
    ]
    const regions = computeSectionedSlicesF2(blocks, geoms, 150)[0].regions!
    expect(regions[0].columns.map((c) => c.start)).toEqual([0, 40])
    expect(regions[0].height).toBe(40) // trailing empties are absorbed: extent = 2 text lines
  })

  it('a manual column break disables balancing for its region', () => {
    const blocks: BlockBox[] = [
      { ...lineBlock(0, [20], { colBreakAfter: true }), section: 0 },
      { ...lineBlock(20, [20]), section: 0 },
      { ...lineBlock(40, [50]), section: 1 },
    ]
    const geoms = [
      { contentHeight: 400, forceBreak: false, cols: 2 },
      { contentHeight: 400, forceBreak: false },
    ]
    const slices = computeSectionedSlicesF2(blocks, geoms, 90)
    // unbalanced (a turn happened) but the region still ends at its tallest
    // column's content: the next continuous section stacks below on the same
    // page (Word packs a short col-broken letterhead row, prod100r3/45)
    expect(slices).toHaveLength(1)
    expect(slices[0].regions![0].columns.map((c) => c.start)).toEqual([0, 20])
    expect(slices[0].regions).toHaveLength(2)
    expect(slices[0].regions![1].top).toBe(20)
  })

  it('columnLayoutSpecs: width + per-column constant translate, later regions pull up', () => {
    const els = Array.from({ length: 4 }, () => ({}) as HTMLElement)
    const blocks: BlockBox[] = [
      { ...lineBlock(0, [20]), section: 1, el: els[0] },
      { ...lineBlock(20, [20]), section: 1, el: els[1] },
      { ...lineBlock(40, [20]), section: 1, el: els[2] },
      { ...lineBlock(60, [50]), section: 2, el: els[3] },
    ]
    const slices: PageSlice[] = [
      {
        start: 0,
        end: 110,
        section: 1,
        physHeight: 70,
        regions: [
          {
            top: 0,
            height: 20,
            section: 1,
            columns: [
              { start: 0, end: 20 },
              { start: 20, end: 40 },
              { start: 40, end: 60 },
            ],
          },
          { top: 20, height: 100, section: 2, columns: [{ start: 60, end: 110 }] },
        ],
      },
    ]
    const secs = [
      sec({}),
      sec({ columns: 3, colSpace: 720 }, { startType: 'continuous' }),
      sec({}, { startType: 'continuous' }),
    ]
    const specs = columnLayoutSpecs(blocks, slices, secs)
    const g = sectionColGeom(secs[1])
    expect(specs).toHaveLength(4)
    expect(specs[0]).toMatchObject({ el: els[0], widthPx: g.colWidthPx, dx: 0, dy: 0 })
    expect(specs[1]).toMatchObject({ el: els[1], dx: g.colWidthPx + g.gapPx, dy: -20 })
    expect(specs[2]).toMatchObject({ el: els[2], dx: 2 * (g.colWidthPx + g.gapPx), dy: -40 })
    // the single-column body pulls up over the vacated stacked space (region top 20, col start offset 60)
    expect(specs[3]).toMatchObject({ el: els[3], dx: 0, dy: -40 })
    expect(specs[3].widthPx).toBeUndefined()
  })

  it('vAlignShiftSpecs: center/bottom pages translate whole blocks into the free space', () => {
    const els = Array.from({ length: 4 }, () => ({}) as HTMLElement)
    const blocks: BlockBox[] = [
      { ...lineBlock(0, [20]), section: 0, el: els[0] },
      { ...lineBlock(20, [20]), section: 0, el: els[1] },
      { ...lineBlock(40, [30]), section: 1, el: els[2] },
      { ...lineBlock(70, [30]), section: 2, el: els[3] },
    ]
    const slices: PageSlice[] = [
      { start: 0, end: 40, section: 0 },
      { start: 40, end: 70, section: 1 },
      { start: 70, end: 100, section: 2 },
    ]
    const secs = [
      sec({ vAlign: 'center' }),
      sec({ vAlign: 'bottom' }, { startType: 'nextPage' }),
      sec({}, { startType: 'nextPage' }),
    ]
    const geoms = [
      { contentHeight: 100, forceBreak: false },
      { contentHeight: 100, forceBreak: true },
      { contentHeight: 100, forceBreak: true },
    ]
    const specs = vAlignShiftSpecs(blocks, slices, secs, geoms)
    // center page: free = 100-40 = 60 → dy 30 for both blocks; bottom page: dy 70
    expect(specs).toHaveLength(3)
    expect(specs[0]).toMatchObject({ el: els[0], dx: 0, dy: 30 })
    expect(specs[1]).toMatchObject({ el: els[1], dx: 0, dy: 30 })
    expect(specs[2]).toMatchObject({ el: els[2], dx: 0, dy: 70 })
  })

  it('vAlignShiftSpecs: a block crossing the page boundary keeps the page top-aligned', () => {
    const els = Array.from({ length: 2 }, () => ({}) as HTMLElement)
    const blocks: BlockBox[] = [
      { ...lineBlock(0, [30, 30]), section: 0, el: els[0] },
      { ...lineBlock(60, [20]), section: 0, el: els[1] },
    ]
    // the first block spans the page-1/page-2 boundary at 40
    const slices: PageSlice[] = [
      { start: 0, end: 40, section: 0 },
      { start: 40, end: 80, section: 0 },
    ]
    const secs = [sec({ vAlign: 'center' })]
    const geoms = [{ contentHeight: 100, forceBreak: false }]
    expect(vAlignShiftSpecs(blocks, slices, secs, geoms)).toHaveLength(0)
  })

  it('sectionColGeom: w:equalWidth="0" reads the explicit w:col width/space list (1290 shape)', () => {
    const s = sec(
      { columns: 2, colSpace: 720 },
      {
        sectPrXml:
          '<w:sectPr><w:cols w:num="2" w:space="720" w:equalWidth="0"><w:col w:w="2640" w:space="720"/><w:col w:w="6000"/></w:cols></w:sectPr>',
      },
    )
    const g = sectionColGeom(s)
    expect(g.equalWidth).toBe(false)
    expect(g.widths.map(Math.round)).toEqual([176, 400])
    expect(g.gaps.map(Math.round)).toEqual([48])
    expect(Math.round(g.colWidthPx)).toBe(176)
    // equal-width fallback when the list is absent
    const eq = sectionColGeom(sec({ columns: 2, colSpace: 720 }))
    expect(eq.equalWidth).toBe(true)
    expect(eq.widths).toHaveLength(2)
  })

  it('columnLayoutSpecs: unequal columns place blocks at cumulative offsets with per-column widths', () => {
    const els = Array.from({ length: 2 }, () => ({}) as HTMLElement)
    const blocks: BlockBox[] = [
      { ...lineBlock(0, [20]), section: 0, el: els[0] },
      { ...lineBlock(20, [20]), section: 0, el: els[1] },
    ]
    const slices: PageSlice[] = [
      {
        start: 0,
        end: 40,
        section: 0,
        physHeight: 20,
        regions: [
          {
            top: 0,
            height: 100,
            section: 0,
            columns: [
              { start: 0, end: 20 },
              { start: 20, end: 40 },
            ],
          },
        ],
      },
    ]
    const secs = [
      sec(
        { columns: 2, colSpace: 720 },
        {
          sectPrXml:
            '<w:sectPr><w:cols w:num="2" w:space="720" w:equalWidth="0"><w:col w:w="2640" w:space="720"/><w:col w:w="6000"/></w:cols></w:sectPr>',
        },
      ),
    ]
    const specs = columnLayoutSpecs(blocks, slices, secs)
    const g = sectionColGeom(secs[0])
    expect(specs[0]).toMatchObject({ el: els[0], widthPx: g.widths[0], dx: 0, dy: 0 })
    expect(specs[1].widthPx).toBeCloseTo(g.widths[1], 3)
    expect(specs[1].dx).toBeCloseTo(g.widths[0] + g.gaps[0], 3)
    expect(specs[1].dy).toBe(-20)
  })

  it('a fixed-width block never advances into a narrower column: the page turns instead (1270 shape)', () => {
    // region cols [400, 200]; a 300px-wide table overflowing col1 must not land in the 200px col2
    const blocks: BlockBox[] = [
      { ...lineBlock(0, [50]), section: 0 },
      { ...lineBlock(50, [80]), section: 0, fixedWidthPx: 300 },
    ]
    const geoms = [{ contentHeight: 100, forceBreak: false, cols: 2, colWidths: [400, 200] }]
    const slices = computeSectionedSlicesF2(blocks, geoms, 130)
    expect(slices).toHaveLength(2)
    expect(slices[1].start).toBe(50)
    // reflowable text still advances into the narrow column
    const flowBlocks: BlockBox[] = [
      { ...lineBlock(0, [50]), section: 0 },
      { ...lineBlock(50, [80]), section: 0 },
    ]
    const flow = computeSectionedSlicesF2(flowBlocks, geoms, 130)
    expect(flow).toHaveLength(1)
    expect(flow[0].regions![0].columns).toHaveLength(2)
  })

  it('pageAt locates by page start (column spans do not affect it)', () => {
    const slices = computeSectionedSlicesF2(
      [lineBlock(0, [50, 50, 50, 50, 50, 50, 50, 50, 50, 50])],
      twoCol,
      500,
    )
    expect(pageAt(slices, 300)).toBe(1) // content in column 2 is still page 1
    expect(pageAt(slices, 450)).toBe(2)
  })
})

describe('fillLineBoxes — picture-only paragraphs (inline image lines)', () => {
  const geoms = [{ contentHeight: 200, forceBreak: false }]
  const rectOf = (top: number, height: number, left = 0, width = 100) =>
    ({
      top,
      height,
      bottom: top + height,
      left,
      right: left + width,
      width,
      x: left,
      y: top,
      toJSON: () => ({}),
    }) as DOMRect
  const imageStack = (tops: number[], lineH: number, style?: string) => {
    const el = document.createElement('p')
    for (const t of tops) {
      const img = document.createElement('img')
      img.className = 'doc-inline-img'
      if (style) img.setAttribute('style', style)
      img.getBoundingClientRect = () => rectOf(t, lineH - 5)
      el.appendChild(img)
    }
    el.getBoundingClientRect = () => rectOf(0, tops.length * lineH)
    return el
  }

  it('breaks an over-page image stack between image lines, not synthesized pixel cuts', () => {
    const el = imageStack([0, 120, 240, 360], 120)
    const b: BlockBox = { top: 0, height: 480, el }
    expect(fillLineBoxes([b], geoms, 1)).toBe(true)
    // each image's ink is 115px tall: the cut sits midway through the 5px gap
    expect(b.lineBoxes).toEqual([
      { offsetInBlock: 0, height: 117.5 },
      { offsetInBlock: 117.5, height: 120 },
      { offsetInBlock: 237.5, height: 120 },
      { offsetInBlock: 357.5, height: 122.5 },
    ])
    // one 120px image line per 200px page (two lines exceed a page)
    const slices = computeSectionedSlicesF2([b], geoms, 480)
    expect(slices.map((s) => s.start)).toEqual([0, 117.5, 237.5, 357.5])
  })

  it('floated and absolutely positioned images do not form lines', () => {
    for (const style of ['float:left', 'position:absolute']) {
      const el = imageStack([0, 120, 240, 360], 120, style)
      const b: BlockBox = { top: 0, height: 480, el }
      expect(fillLineBoxes([b], geoms, 1)).toBe(true)
      // no line data: synthesized page-height cuts remain
      expect(b.lineBoxes).toEqual([
        { offsetInBlock: 0, height: 200 },
        { offsetInBlock: 200, height: 200 },
        { offsetInBlock: 400, height: 80 },
      ])
    }
  })

  it('anchors a picture-only line at its image element', () => {
    const el = imageStack([0, 120, 240, 360], 120)
    const anchor = nextLineAnchor(el, 120, 1)
    expect(anchor).toEqual({ node: el.children[1], charOffset: 0 })
    expect(anchorElement(anchor!)).toBe(el.children[1])
  })

  it('a line holding both text and a taller image keeps the text anchor', () => {
    const el = imageStack([0], 120)
    el.appendChild(document.createTextNode('caption'))
    const glyph = rectOf(10, 20, 60, 50)
    const orig = Range.prototype.getClientRects
    Range.prototype.getClientRects = () => [glyph] as unknown as DOMRectList
    try {
      const anchor = lineStartAnchor(el, 0, 1)
      expect(anchor?.node).toBe(el.lastChild)
    } finally {
      Range.prototype.getClientRects = orig
    }
  })
})

describe('fillLineBoxes — protected image blocks are atomic', () => {
  const geoms = [{ contentHeight: 200, forceBreak: false }]
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
  const imageBlockEl = (tops: number[], lineH: number, totalH: number) => {
    const el = document.createElement('div')
    el.className = 'doc-protected doc-protected-image'
    for (const t of tops) {
      const img = document.createElement('img')
      img.className = 'doc-protected-img'
      img.getBoundingClientRect = () => rectOf(t, lineH - 5)
      el.appendChild(img)
    }
    el.getBoundingClientRect = () => rectOf(0, totalH)
    return el
  }

  it('a page-crossing image that fits a page gets no line data (pushes whole)', () => {
    // a tiny lead line (anchor marker / leading spaces) + the photo would
    // otherwise form a boundary that cuts a sliver strip off the photo
    const el = imageBlockEl([0, 20], 90, 180)
    const b: BlockBox = { top: 100, height: 180, el }
    const slices: PageSlice[] = [
      { start: 0, end: 200, section: 0 },
      { start: 200, end: 400, section: 0 },
    ]
    expect(fillLineBoxes([b], geoms, 1, slices)).toBe(false)
    expect(b.lineBoxes).toBeUndefined()
  })

  it('an over-page sole-line image gets the oversize-clip flag, not pixel cuts', () => {
    const el = imageBlockEl([0], 480, 480)
    const b: BlockBox = { top: 0, height: 480, el }
    expect(fillLineBoxes([b], geoms, 1)).toBe(true)
    expect(b.lineBoxes).toBeUndefined()
    // ink height of the sole image line (480 - the 5px stub gap)
    expect(b.oversizeLineH).toBe(475)
  })

  it('a clip-marked block re-qualifies from its sole line ink (no oscillation)', () => {
    // renderer clip applied: the block measures at the fitting 200px, but the
    // image ink is still 475px — the flag (and thus the clip patch) must survive
    const el = imageBlockEl([0], 480, 200)
    el.dataset.oversizeClip = '200.0'
    const b: BlockBox = { top: 0, height: 200, el }
    expect(fillLineBoxes([b], geoms, 1)).toBe(true)
    expect(b.oversizeLineH).toBe(475)
    // re-run: the flag is already in place, nothing changes
    expect(fillLineBoxes([b], geoms, 1)).toBe(false)
    expect(b.oversizeLineH).toBe(475)
  })

  it('a clip-marked block whose line fits again drops the flag', () => {
    const el = imageBlockEl([0], 150, 150)
    el.dataset.oversizeClip = '200.0'
    const b: BlockBox = { top: 0, height: 150, el }
    expect(fillLineBoxes([b], geoms, 1)).toBe(false)
    expect(b.oversizeLineH).toBeUndefined()
  })
})

describe('computeSectionedSlicesF2 — oversized sole-line blocks clip at the page bottom', () => {
  it('starts on a fresh page, emits a one-page clip patch, no pixel cuts', () => {
    const geoms: SectionGeom[] = [{ contentHeight: 200, forceBreak: false }]
    const blocks: BlockBox[] = [
      { top: 0, height: 100 },
      { top: 100, height: 600, oversizeLineH: 600 },
      { top: 700, height: 50 },
    ]
    const out: SliceOutputs = { oversizeClips: [] }
    const slices = computeSectionedSlicesF2(blocks, geoms, 750, out)
    expect(out.oversizeClips).toEqual([{ blockTop: 100, clipPx: 200 }])
    // the oversized block owns one page; the next block starts the following page
    expect(slices.map((s) => s.start)).toEqual([0, 100, 700])
  })

  it('a renderer-clipped block fills exactly one page and re-emits the patch', () => {
    const geoms: SectionGeom[] = [{ contentHeight: 200, forceBreak: false }]
    const blocks: BlockBox[] = [
      { top: 0, height: 100 },
      { top: 100, height: 200, oversizeLineH: 600 },
      { top: 300, height: 50 },
    ]
    const out: SliceOutputs = { oversizeClips: [] }
    const slices = computeSectionedSlicesF2(blocks, geoms, 350, out)
    expect(out.oversizeClips).toEqual([{ blockTop: 100, clipPx: 200 }])
    expect(slices.map((s) => s.start)).toEqual([0, 100, 300])
  })

  it('keepNext heading pushes with the picture and clips it below (no orphan)', () => {
    const geoms: SectionGeom[] = [{ contentHeight: 200, forceBreak: false }]
    const blocks: BlockBox[] = [
      { top: 0, height: 150 },
      { top: 150, height: 30, keepNext: true },
      { top: 180, height: 600, oversizeLineH: 600 },
      { top: 780, height: 50 },
    ]
    const out: SliceOutputs = { oversizeClips: [] }
    const slices = computeSectionedSlicesF2(blocks, geoms, 830, out)
    // heading and picture share the fresh page; the clip is the remainder below the heading
    expect(out.oversizeClips).toEqual([{ blockTop: 180, clipPx: 170 }])
    expect(slices.map((s) => s.start)).toEqual([0, 150, 780])
  })

  it('keepNext + renderer-clipped picture is a fixed point', () => {
    const geoms: SectionGeom[] = [{ contentHeight: 200, forceBreak: false }]
    const blocks: BlockBox[] = [
      { top: 0, height: 150 },
      { top: 150, height: 30, keepNext: true },
      { top: 180, height: 170, oversizeLineH: 600 },
      { top: 350, height: 50 },
    ]
    const out: SliceOutputs = { oversizeClips: [] }
    const slices = computeSectionedSlicesF2(blocks, geoms, 400, out)
    expect(out.oversizeClips).toEqual([{ blockTop: 180, clipPx: 170 }])
    expect(slices.map((s) => s.start)).toEqual([0, 150, 350])
  })

  it('multi-column: advances one column and clips to the column height', () => {
    const geoms: SectionGeom[] = [{ contentHeight: 200, forceBreak: false, cols: 2 }]
    const blocks: BlockBox[] = [
      { top: 0, height: 50 },
      { top: 50, height: 600, oversizeLineH: 600 },
    ]
    const out: SliceOutputs = { oversizeClips: [] }
    const slices = computeSectionedSlicesF2(blocks, geoms, 650, out)
    expect(out.oversizeClips).toEqual([{ blockTop: 50, clipPx: 200 }])
    expect(slices).toHaveLength(1)
    expect(slices[0].regions?.[0].columns.map((c) => c.start)).toEqual([0, 50])
  })
})
