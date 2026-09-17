import { describe, expect, it } from 'vitest'
import type { SectionInfo } from '@genoffice/docx-engine'
import {
  type SliceOutputs,
  appendEndnotesBlock,
  assignSections,
  cellCutYs,
  columnLayoutSpecs,
  columnLineSplits,
  fillColWraps,
  splitFloatStyle,
  widthPassGate,
  newWidthPassState,
  resetWidthPassHistory,
  vAlignShiftSpecs,
  verticalTextSpecs,
  sectionWidthSpecs,
  paperWidthPx,
  pageLeftPx,
  sectionGridPitchSpecs,
  sectionTopMarginSpecs,
  docGridPitchPt,
  sectionCharSpaceSpecs,
  docCharSpacePt,
  sectionColGeom,
  computePageSlices,
  computeSectionedSlices,
  computeSectionedSlicesF2,
  planRowSplit,
  cellLinesOf,
  rowSplitCss,
  effectiveHfRefs,
  hasPrintableHeaderFooter,
  hfVariantOf,
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
  applyRowNotes,
  fillLineBoxes,
  measureBlocks,
  sectionFirstPages,
  sectionGeoms,
  sectionPageBox,
  tableHeaderFlags,
  tableRowFlags,
  type BlockBox,
  type ColWrapTable,
  type PageSlice,
  type SectionGeom,
  type TableRowBox,
} from '../src/renderer/pagination'
import { rowFillAttrs } from '../src/renderer/editor/pagination-gaps'
import { verticalPageCss } from '../src/renderer/components/PaginationPreview'

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

  it('pages narrower than the widest section are centered on the shared paper (genoffice#246)', () => {
    // portrait pages in a document with a landscape section: the canvas paper is
    // the landscape width, so every portrait block carries the centering offset
    const landscape = sec({ pageWidth: 16838, pageHeight: 11906, orientation: 'landscape' })
    const secsList = [sec({}), landscape, sec({})]
    const paperW = (16838 / 1440) * 96
    expect(paperWidthPx(secsList)).toBeCloseTo(paperW, 5)
    const offset = ((16838 - 11906) / 1440 / 2) * 96
    expect(pageLeftPx(secsList[0].settings, paperW)).toBeCloseTo(offset, 5)
    expect(pageLeftPx(landscape.settings, paperW)).toBe(0)
    const blocks = [
      block(0, 100, { section: 0, el: el() }),
      block(100, 100, { section: 1, el: el() }),
      block(200, 100, { section: 2, el: el() }),
    ]
    const specs = sectionWidthSpecs(blocks, secsList, sectionGeoms(secsList))
    expect(specs).toHaveLength(3)
    expect(specs[0].pageDx).toBeCloseTo(offset, 1)
    expect(specs[1].pageDx).toBe(0)
    expect(specs[2].pageDx).toBeCloseTo(offset, 1)
    // the column shift (dx) stays the margin difference: page-relative anchors undo it alone
    expect(specs[0].dx).toBe(0)
    expect(specs[1].dx).toBe(0)
  })

  it('a page-width-only difference (same margins and content width) still emits specs', () => {
    // A4 vs. a 1in wider page whose extra width sits in the right margin only
    const wide = sec({ pageWidth: 11906 + 1440, marginRight: 2880 })
    const secsList = [sec({}), wide]
    const blocks = [
      block(0, 100, { section: 0, el: el() }),
      block(100, 100, { section: 1, el: el() }),
    ]
    const specs = sectionWidthSpecs(blocks, secsList, sectionGeoms(secsList))
    expect(specs).toHaveLength(2)
    expect(specs[0].pageDx).toBeCloseTo(48, 1)
    expect(specs[1].pageDx).toBe(0)
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

describe('sectionTopMarginSpecs — per-section top margin for page-relative anchors', () => {
  const el = () => document.createElement('p')

  it('uniform top margins produce no specs (single .doc-page injection)', () => {
    const secsList = [sec({ marginTop: 860 }), sec({ marginTop: 860 })]
    expect(sectionTopMarginSpecs([block(0, 100, { section: 0, el: el() })], secsList)).toEqual([])
  })

  it('a final section with a different top margin: each block carries its own section top', () => {
    // a letterhead section (pgMar top 860) closed by a continuous 1in section:
    // the header logo pinned to the page top must subtract 860, not 1440
    const secsList = [sec({ marginTop: 860 }), sec({ marginTop: 1440 })]
    const blocks = [
      block(0, 100, { section: 0, el: el() }),
      block(100, 100, { section: 1, el: el() }),
    ]
    const specs = sectionTopMarginSpecs(blocks, secsList)
    expect(specs).toHaveLength(2)
    expect(specs[0].marginTopPx).toBeCloseTo((860 / 1440) * 96, 3)
    expect(specs[1].marginTopPx).toBe(96)
    expect(specs.every((sp) => sp.dx === 0 && sp.dy === 0)).toBe(true)
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

  it('a continuous section starting mid-page keeps the host page capacity (Word: the page is laid out with the section it begins in)', () => {
    // prod-sas 087: nextPage sections with 720/800 margins host continuous
    // 120/280 column blocks; Word keeps the host's shorter body on that page
    const geoms: SectionGeom[] = [
      { contentHeight: 400, forceBreak: false, startType: 'nextPage' },
      { contentHeight: 500, forceBreak: false, startType: 'continuous' },
    ]
    const blocks = [
      block(0, 300, { section: 0 }),
      block(300, 60, { section: 1 }),
      block(360, 60, { section: 1 }), // 360..420 crosses the host's 400 capacity
    ]
    expect(
      computeSectionedSlicesF2(blocks, geoms, 420).map((s) => [s.start, s.end, s.section]),
    ).toEqual([
      [0, 360, 0],
      [360, 420, 1],
    ])
    expect(
      computeSectionedSlices(blocks, geoms, 420).map((s) => [s.start, s.end, s.section]),
    ).toEqual([
      [0, 360, 0],
      [360, 420, 1],
    ])
  })

  it('a continuous section that begins on a page overflow opened takes that page over', () => {
    const geoms: SectionGeom[] = [
      { contentHeight: 400, forceBreak: false, startType: 'nextPage' },
      { contentHeight: 500, forceBreak: false, startType: 'continuous' },
    ]
    // section 0 fills its page exactly; section 1 starts at the top of page 2
    // and paginates with its own (taller) capacity
    const blocks = [
      block(0, 400, { section: 0 }),
      block(400, 450, { section: 1 }),
      block(850, 100, { section: 1 }),
    ]
    expect(
      computeSectionedSlicesF2(blocks, geoms, 950).map((s) => [s.start, s.end, s.section]),
    ).toEqual([
      [0, 400, 0],
      [400, 850, 1],
      [850, 950, 1],
    ])
    expect(
      computeSectionedSlices(blocks, geoms, 950).map((s) => [s.start, s.end, s.section]),
    ).toEqual([
      [0, 400, 0],
      [400, 850, 1],
      [850, 950, 1],
    ])
  })

  it("a foreign-section block measures oversize against that section's first-page capacity", () => {
    const geoms: SectionGeom[] = [
      { contentHeight: 400, forceBreak: false, startType: 'nextPage' },
      { contentHeight: 600, firstContentHeight: 450, forceBreak: false, startType: 'continuous' },
    ]
    // 430 fits section 1's shortened first page: pushed there whole
    const pushed = computeSectionedSlicesF2(
      [block(0, 300, { section: 0 }), block(300, 430, { section: 1 })],
      geoms,
      730,
    )
    expect(pushed.map((s) => [s.start, s.end, s.section])).toEqual([
      [0, 300, 0],
      [300, 730, 1],
    ])
    // 480 fits neither the host page nor that first page (only the 600 default
    // would take it): it stays oversized in place instead of being pushed onto
    // a page it then overflows
    const kept = computeSectionedSlicesF2(
      [block(0, 300, { section: 0 }), block(300, 480, { section: 1 })],
      geoms,
      780,
    )
    expect(kept.map((s) => [s.start, s.end, s.section])).toEqual([[0, 780, 0]])
    // a later block of a section that already began on the host page opens a
    // non-first page: the 600 default capacity takes the 480 block
    const later = computeSectionedSlicesF2(
      [
        block(0, 300, { section: 0 }),
        block(300, 60, { section: 1 }),
        block(360, 480, { section: 1 }),
      ],
      geoms,
      840,
    )
    expect(later.map((s) => [s.start, s.end, s.section, s.continuedSection ?? false])).toEqual([
      [0, 360, 0, false],
      [360, 840, 1, true],
    ])
  })

  describe('titlePg first page: the page a section starts on, only when it owns that page', () => {
    const geoms: SectionGeom[] = [
      { contentHeight: 400, forceBreak: false, startType: 'nextPage' },
      { contentHeight: 400, firstContentHeight: 300, forceBreak: false, startType: 'continuous' },
    ]
    const shape = (slices: PageSlice[]) =>
      slices.map((s) => [s.start, s.end, s.section, s.continuedSection ?? false])

    it('nextPage section: its first page takes the first-page variant and capacity', () => {
      const next: SectionGeom[] = [
        geoms[0],
        { ...geoms[1], forceBreak: true, startType: 'nextPage' },
      ]
      const blocks = [
        block(0, 300, { section: 0 }),
        block(300, 200, { section: 1 }),
        block(500, 150, { section: 1 }), // 500..650 overflows the 300 first-page capacity
      ]
      const slices = computeSectionedSlicesF2(blocks, next, 650)
      expect(shape(slices)).toEqual([
        [0, 300, 0, false],
        [300, 500, 1, false],
        [500, 650, 1, false],
      ])
      expect(sectionFirstPages(slices)).toEqual([true, true, false])
      expect(hfVariantOf(true, true, false, 2)).toBe('first')
    })

    it('continuous section starting mid-page: the host page keeps its section, later pages are not first pages', () => {
      const blocks = [
        block(0, 300, { section: 0 }),
        block(300, 60, { section: 1 }),
        block(360, 60, { section: 1 }), // crosses the host's 400 capacity
        block(420, 280, { section: 1 }), // 360..700 fits the 400 default capacity, not the 300 first-page one
      ]
      const slices = computeSectionedSlicesF2(blocks, geoms, 700)
      expect(shape(slices)).toEqual([
        [0, 360, 0, false],
        [360, 700, 1, true],
      ])
      const firsts = sectionFirstPages(slices)
      expect(firsts).toEqual([true, false])
      expect(hfVariantOf(true, firsts[1], false, 2)).toBe('default')
      expect(hfVariantOf(true, firsts[1], true, 2)).toBe('even')
    })

    it('continuous section starting exactly at a page top: that page is its first page', () => {
      const blocks = [
        block(0, 400, { section: 0 }),
        block(400, 200, { section: 1 }),
        block(600, 150, { section: 1 }), // 600..750 overflows the 300 first-page capacity
      ]
      const slices = computeSectionedSlicesF2(blocks, geoms, 750)
      expect(shape(slices)).toEqual([
        [0, 400, 0, false],
        [400, 600, 1, false],
        [600, 750, 1, false],
      ])
      expect(sectionFirstPages(slices)).toEqual([true, true, false])
      expect(hfVariantOf(true, true, true, 2)).toBe('first')
    })

    it('a page break right after a mid-page continuous break opens the section at a page top', () => {
      const blocks = [
        block(0, 300, { section: 0 }),
        block(300, 200, { section: 1, breakBefore: true }),
      ]
      const slices = computeSectionedSlicesF2(blocks, geoms, 500)
      expect(shape(slices)).toEqual([
        [0, 300, 0, false],
        [300, 500, 1, false],
      ])
      expect(sectionFirstPages(slices)).toEqual([true, true])
    })
  })

  it("a page that begins with the previous section's float keeps that section", () => {
    const geoms: SectionGeom[] = [
      { contentHeight: 400, forceBreak: false, startType: 'nextPage' },
      { contentHeight: 500, forceBreak: false, startType: 'continuous' },
    ]
    // the float lands on page 2 without consuming column height; the continuous
    // section starting right after it must not take the page over
    const blocks = [
      block(0, 400, { section: 0 }),
      block(400, 80, { section: 0, floated: true }),
      block(400, 300, { section: 1 }),
    ]
    expect(
      computeSectionedSlicesF2(blocks, geoms, 700).map((s) => [s.start, s.end, s.section]),
    ).toEqual([
      [0, 400, 0],
      [400, 700, 0],
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

  it('two leading w:br on one paragraph leave a blank sheet before it', () => {
    const blocks = [
      block(0, 100),
      { ...block(100, 100), breakBefore: true, breakBeforeBr: true, extraBreaksBefore: 1 },
    ]
    const slices = computeSectionedSlicesF2(blocks, geoms1, 200)
    expect(slices).toEqual([
      { start: 0, end: 100, section: 0 },
      { start: 100, end: 100, section: 0 },
      { start: 100, end: 200, section: 0 },
    ])
  })

  it('a trailing w:br on a keepNext block ends the chain: the next block starts the new page', () => {
    const marker = block(0, 20, { keepNext: true, breakAfter: true })
    const heading = block(20, 30, { keepNext: true })
    const body = block(50, 100)
    const slices = computeSectionedSlicesF2([marker, heading, body], geoms1, 150)
    expect(slices).toEqual([
      { start: 0, end: 20, section: 0 },
      { start: 20, end: 150, section: 0 },
    ])
  })

  it('a leading-break cut that snaps into the first line box turns the page before the block', () => {
    const blocks = [
      block(0, 100),
      block(100, 60, {
        innerBreaks: [10],
        lineBoxes: [
          { offsetInBlock: 0, height: 40 },
          { offsetInBlock: 40, height: 20 },
        ],
      }),
    ]
    const slices = computeSectionedSlicesF2(blocks, geoms1, 160)
    expect(slices).toEqual([
      { start: 0, end: 100, section: 0 },
      { start: 100, end: 160, section: 0 },
    ])
  })

  it('two trailing w:br leave a blank sheet before the next block and two blank last pages', () => {
    const blocks = [
      { ...block(0, 100), breakAfter: true, extraBreaksAfter: 1 },
      { ...block(100, 100), breakAfter: true, extraBreaksAfter: 1 },
    ]
    const slices = computeSectionedSlicesF2(blocks, geoms1, 200)
    expect(slices).toEqual([
      { start: 0, end: 100, section: 0 },
      { start: 100, end: 100, section: 0 },
      { start: 100, end: 200, section: 0 },
      { start: 200, end: 200, section: 0 },
      { start: 200, end: 200, section: 0 },
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

  it('widowControl: an over-page paragraph whose first cut would orphan one line starts on the next page', () => {
    // 150px used of 300; the paragraph (100 + 300 = 400px) exceeds a page, and
    // only its first line fits here -> Word moves the start down, then cuts
    const beforeBlock = lineBlock(0, [50, 50, 50])
    const mainBlock = lineBlock(150, [100, 300], { widowControl: true })
    const geoms = [{ contentHeight: 300, forceBreak: false }]
    const slices = computeSectionedSlicesF2([beforeBlock, mainBlock], geoms, 550)
    expect(slices.map((s) => s.start)).toEqual([0, 150, 250])
  })

  it('widowControl: an over-page paragraph keeps its start when two lines fit before the cut', () => {
    const beforeBlock = lineBlock(0, [50, 50, 50])
    const mainBlock = lineBlock(150, [50, 50, 300], { widowControl: true })
    const geoms = [{ contentHeight: 300, forceBreak: false }]
    const slices = computeSectionedSlicesF2([beforeBlock, mainBlock], geoms, 550)
    expect(slices.map((s) => s.start)).toEqual([0, 250])
  })

  it('widowControl=false: an over-page paragraph cuts after its first line', () => {
    const beforeBlock = lineBlock(0, [50, 50, 50])
    const mainBlock = lineBlock(150, [100, 300], { widowControl: false })
    const geoms = [{ contentHeight: 300, forceBreak: false }]
    const slices = computeSectionedSlicesF2([beforeBlock, mainBlock], geoms, 550)
    expect(slices.map((s) => s.start)).toEqual([0, 250])
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

  it('keepNext heading whose anchor row cannot share any page still moves whole', () => {
    // the chain + anchor demand exceeds a fresh page (a table head row taller
    // than the page), so the keepNext constraint is dropped; the heading must
    // then place like any block instead of being cut by the page bottom
    const before = block(0, 180)
    const heading = lineBlock(180, [30], { keepNext: true, keepLines: true })
    const table = block(210, 400, {
      tableRows: [
        { height: 250, contentBottom: 0 },
        { height: 150, contentBottom: 0 },
      ],
    })
    const slices = computeSectionedSlicesF2([before, heading, table], geoms1, 610)
    expect(slices[1].start).toBe(180)
  })

  it('keepNext heading demands the footnote separator its table anchor will open', () => {
    // heading (30) + head row (20) + its notes (30) fit the 200px page by 10px,
    // but the row's notes also open the 16px separator strip; the heading must
    // move with the table instead of staying orphaned at the page bottom
    const before = block(0, 110)
    const heading = lineBlock(110, [30], { keepNext: true, keepLines: true })
    const table = block(140, 130, {
      footnoteExtraPx: 30,
      tableRows: [
        { height: 20, contentBottom: 0, notesPx: 30 },
        { height: 80, contentBottom: 0 },
      ],
    })
    const slices = computeSectionedSlicesF2([before, heading, table], geoms1, 270)
    expect(slices.length).toBe(2)
    expect(slices[1].start).toBe(110)
  })

  it('keepLines without keepNext stays at the page bottom (no chain push)', () => {
    const before = block(0, 120)
    const kl = lineBlock(120, [60], { keepLines: true })
    const body = lineBlock(180, [50, 50])
    const slices = computeSectionedSlicesF2([before, kl, body], geoms1, 280)
    expect(slices.length).toBe(2)
    expect(slices[1].start).toBe(180) // only the body moves
  })

  it('keepLines: only the text must fit, the space-after may overflow the bottom margin', () => {
    const before = block(0, 120)
    // text 75px ends at 195 (fits), text + 15px space-after ends at 210 (does not)
    const kl = lineBlock(120, [40, 35], { keepLines: true, spaceAfterPx: 15 })
    const body = lineBlock(210, [50, 50])
    const slices = computeSectionedSlicesF2([before, kl, body], geoms1, 310)
    expect(slices.map((s) => s.start)).toEqual([0, 210])
  })

  it('keepLines anchor of a keepNext chain: its space-after may overflow too', () => {
    const before = block(0, 140)
    const heading = lineBlock(140, [20], { keepNext: true })
    // anchor text ends exactly at 200; only its space-after runs past the page
    const body = lineBlock(160, [10, 10, 10, 10], { keepLines: true, spaceAfterPx: 15 })
    const next = lineBlock(215, [50])
    const slices = computeSectionedSlicesF2([before, heading, body, next], geoms1, 265)
    expect(slices.map((s) => s.start)).toEqual([0, 215])
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

describe('computeSectionedSlicesF2 — mid-paragraph page breaks (innerBreaks)', () => {
  it('turns the page at the break line and keeps the text before on the current page', () => {
    // 5 lines of 20px; a page-type w:br sits at the end of line 2, so line 3 starts a page
    const para = lineBlock(0, [20, 20, 20, 20, 20], { innerBreaks: [40] })
    const next = block(100, 30)
    const slices = computeSectionedSlicesF2([para, next], geoms1, 130)
    expect(slices.map((s) => [s.start, s.end])).toEqual([
      [0, 40],
      [40, 130],
    ])
  })

  it('snaps an ink-top break Y to the line box holding it', () => {
    const para = lineBlock(0, [20, 20, 20, 20, 20], { innerBreaks: [43] })
    const slices = computeSectionedSlicesF2([para], geoms1, 100)
    expect(slices.map((s) => s.start)).toEqual([0, 40])
  })

  it('splits at the raw Y when no line data exists yet (first pass)', () => {
    const para = block(0, 100, { innerBreaks: [40] })
    const slices = computeSectionedSlicesF2([para, block(100, 30)], geoms1, 130)
    expect(slices.map((s) => [s.start, s.end])).toEqual([
      [0, 40],
      [40, 130],
    ])
  })

  it('the text before the break still splits by widow/orphan rules when it overflows', () => {
    // 12 lines of 20px before the break (240px > 200px page): lines wrap to page 2,
    // then the break forces page 3 for the trailing lines
    const heights = Array.from({ length: 14 }, () => 20)
    const para = lineBlock(0, heights, { innerBreaks: [240] })
    const slices = computeSectionedSlicesF2([para], geoms1, 280)
    expect(slices.map((s) => s.start)).toEqual([0, 200, 240])
  })

  it('overrides keepLines: the page still turns at the break line', () => {
    const para = lineBlock(0, [20, 20, 20, 20], { innerBreaks: [40], keepLines: true })
    const slices = computeSectionedSlicesF2([para, block(80, 30)], geoms1, 110)
    expect(slices.map((s) => [s.start, s.end])).toEqual([
      [0, 40],
      [40, 110],
    ])
  })

  it('truncates a keepNext chain: a chain member with an inner break turns the page', () => {
    const head = lineBlock(0, [20, 20], { keepNext: true })
    const member = lineBlock(40, [20, 20, 20, 20], { innerBreaks: [40], keepNext: true })
    const anchor = lineBlock(120, [20, 20])
    const slices = computeSectionedSlicesF2([head, member, anchor], geoms1, 160)
    expect(slices.map((s) => [s.start, s.end])).toEqual([
      [0, 80],
      [80, 160],
    ])
  })

  it('truncates a keepNext chain at a forced section start: the next section is not pulled back', () => {
    // a keep-on title closing section 0, then two keep-on headings opening a
    // nextPage section 1 (the chain must not hold the section break's page open)
    const geoms = [
      { contentHeight: 800, forceBreak: false, startType: 'nextPage' as const },
      { contentHeight: 800, forceBreak: true, startType: 'nextPage' as const },
    ]
    const title = lineBlock(0, [60], { keepNext: true, section: 0 })
    const h1 = lineBlock(60, [50], { keepNext: true, section: 1 })
    const h2 = lineBlock(110, [40], { keepNext: true, section: 1 })
    const body = lineBlock(150, [20, 20, 20, 20, 20], { section: 1 })
    const slices = computeSectionedSlicesF2([title, h1, h2, body], geoms, 250)
    expect(slices.map((s) => [s.start, s.end, s.section])).toEqual([
      [0, 60, 0],
      [60, 250, 1],
    ])
  })

  it('a break with no text after it still pushes the next block (breakAfter unchanged)', () => {
    const para = lineBlock(0, [20, 20], { breakAfter: true })
    const slices = computeSectionedSlicesF2([para, block(40, 30)], geoms1, 70)
    expect(slices.map((s) => s.start)).toEqual([0, 40])
  })
})

describe('computeSectionedSlicesF2 — table row-level page breaks', () => {
  const makeTableBlock = (top: number, rows: TableRowBox[]): BlockBox => {
    const h = rows.reduce((s, r) => s + r.height, 0)
    return { top, height: h, tableRows: rows }
  }

  it('a row fits a page only together with the footnotes it references', () => {
    // page 100 (84 below the note separator); rows 30 x 3, the first row carries
    // a 50px footnote: rows 0..1 would fit by height alone, but the note area
    // leaves room for row 0 only
    const rows: TableRowBox[] = Array.from({ length: 3 }, () => ({ height: 30 }))
    rows[0].notesPx = 50
    const b = makeTableBlock(0, rows)
    b.height += 50
    b.footnoteExtraPx = 50
    const slices = computeSectionedSlicesF2([b], [{ contentHeight: 100, forceBreak: false }], 90)
    expect(slices.map((s) => s.start)).toEqual([0, 30])
  })

  it('a keepNext heading demands the table head together with its footnotes', () => {
    // filler 30 + heading 20 + head row 30 fit the 84px capacity, but the row's
    // 30px footnote does not: the chain pushes instead of orphaning the heading
    const rows: TableRowBox[] = [{ height: 30, notesPx: 30 }]
    const table = makeTableBlock(50, rows)
    table.height += 30
    table.footnoteExtraPx = 30
    const blocks = [block(0, 30), block(30, 20, { keepNext: true }), table]
    const slices = computeSectionedSlicesF2(blocks, [{ contentHeight: 100, forceBreak: false }], 80)
    expect(slices.map((s) => s.start)).toEqual([0, 30])
  })

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

  it('over-page atLeast row: a trailing end-of-cell mark past the page bottom is overflow, not a fragment (Word probe 2026-09-05)', () => {
    // page 200px; the row starts after a 40px row and spans two pages. Its text
    // bands end exactly at the second page's bottom (contentBottom 360) and
    // only the band-less end-of-cell mark (20px) hangs over: Word keeps that
    // mark on the page after a nested table, so no third page opens and the
    // fragment target stays at the content bottom
    const rows: TableRowBox[] = [
      { height: 40 },
      { height: 380, minHPx: 150, cutYs: [160, 280], contentBottom: 360 },
    ]
    const out: SliceOutputs = { rowFills: [] }
    const slices = computeSectionedSlicesF2(
      [makeTableBlock(0, rows)],
      [{ contentHeight: 200, forceBreak: false }],
      420,
      out,
    )
    expect(slices.length).toBe(2)
    expect(out.rowFills).toEqual([{ blockTop: 0, row: 1, targetPx: 360 }])
    // the same 20px as a text line (a paragraph after the nested table) does
    // spill and re-honors the declared height on the new page
    const spill: TableRowBox[] = [
      { height: 40 },
      { height: 380, minHPx: 150, cutYs: [160, 280], contentBottom: 380 },
    ]
    const out2: SliceOutputs = { rowFills: [] }
    const slices2 = computeSectionedSlicesF2(
      [makeTableBlock(0, spill)],
      [{ contentHeight: 200, forceBreak: false }],
      420,
      out2,
    )
    expect(slices2.length).toBe(3)
    expect(out2.rowFills).toEqual([{ blockTop: 0, row: 1, targetPx: 430 }])
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

  it('keepNext rows chain to the next row; a lone last row may open a page', () => {
    // 200px pages, paragraph 150 leaves 50: rows 0 (keepNext) and 1 (30 each)
    // only fit one at a time, so row 0 moves down with row 1 (probe 2026-09-17,
    // cases 10-12); without keepNext row 0 stays. A chain taller than a page is
    // placed normally. A table's last row alone on the next page is Word's
    // behaviour too (case 14).
    const page = [{ contentHeight: 200, forceBreak: false }]
    const slicesOf = (rows: TableRowBox[]) =>
      computeSectionedSlicesF2(
        [{ top: 0, height: 150 }, makeTableBlock(150, rows)],
        page,
        150 + rows.reduce((s, r) => s + r.height, 0),
      ).map((s) => [s.start, s.end])
    expect(slicesOf([{ height: 30, keepNext: true }, { height: 30 }, { height: 30 }])).toEqual([
      [0, 150],
      [150, 240],
    ])
    expect(slicesOf([{ height: 30 }, { height: 30 }, { height: 30 }])).toEqual([
      [0, 180],
      [180, 240],
    ])
    // chain of two keepNext rows with the anchor's first cut segment as its demand
    expect(
      slicesOf([
        { height: 20, keepNext: true },
        { height: 20, keepNext: true },
        { height: 60, cutYs: [20, 40] },
      ]),
    ).toEqual([
      [0, 150],
      [150, 250],
    ])
    expect(
      slicesOf([{ height: 120, keepNext: true }, { height: 120, keepNext: true }, { height: 30 }]),
    ).toEqual([
      [0, 150],
      [150, 270],
      [270, 420],
    ])
    expect(slicesOf([{ height: 25 }, { height: 25 }, { height: 25 }])).toEqual([
      [0, 200],
      [200, 225],
    ])
    // the anchor's demand is widow-aware: its 3-line cell paragraph cannot be
    // cut 2+1, so the chain needs the whole 60px paragraph and moves although a
    // raw first cut (20px) would have fit below the keepNext row
    const lines = (n: number, pitch: number) => ({
      lines: Array.from({ length: n }, (_, i): [number, number] => [i * pitch, (i + 1) * pitch]),
      childOf: Array.from({ length: n }, () => 0),
      paraOf: Array.from({ length: n }, () => 0),
      alignDy: 0,
      alignFrac: 0,
    })
    const anchor: TableRowBox = {
      height: 60,
      contentBottom: 60,
      cutYs: [20, 40],
      cells: [lines(3, 20)],
    }
    const chain = (rows: TableRowBox[], out?: SliceOutputs) =>
      computeSectionedSlicesF2(
        [
          { top: 0, height: 150 },
          { ...makeTableBlock(150, rows), modernTableHeaders: true },
        ],
        page,
        230,
        out,
      ).map((s) => [s.start, s.end])
    const moved = [
      [0, 150],
      [150, 230],
    ]
    expect(chain([{ height: 20, keepNext: true }, anchor], { rowSplits: [] })).toEqual(moved)
    // without a cell split path the anchor is atomic and demands its full height;
    // a declared-height anchor demands at least its (page-capped) minimum
    expect(
      chain([
        { height: 20, keepNext: true },
        { ...anchor, cutYs: [] },
      ]),
    ).toEqual(moved)
    expect(
      chain(
        [
          { height: 20, keepNext: true },
          { ...anchor, cells: [lines(1, 20)], minHPx: 60 },
        ],
        {
          rowSplits: [],
        },
      ),
    ).toEqual(moved)
  })

  it('tblHeader on every row: no repetition and every row reaches a page (LO tdf104492)', () => {
    // 12 header rows x 150 = 1800px on 680px pages: Word lays the rows out
    // normally (a header set that is the whole table cannot repeat)
    const rows: TableRowBox[] = Array.from({ length: 12 }, () => ({
      height: 150,
      isHeader: true,
      cutYs: [50, 100],
    }))
    const b = makeTableBlock(0, rows)
    b.modernTableHeaders = true
    const slices = computeSectionedSlicesF2([b], [{ contentHeight: 680, forceBreak: false }], 1800)
    expect(slices.map((s) => [s.start, s.end])).toEqual([
      [0, 600],
      [600, 1200],
      [1200, 1800],
    ])
    for (const s of slices) expect(s.repeatHeader).toBeUndefined()
  })

  it('tblHeader: header block taller than a page drops repetition but keeps every body row', () => {
    const rows: TableRowBox[] = [
      { height: 150, isHeader: true },
      { height: 150, isHeader: true },
      ...Array.from({ length: 4 }, () => ({ height: 50 })),
    ]
    const b = makeTableBlock(0, rows)
    const slices = computeSectionedSlicesF2([b], [{ contentHeight: 200, forceBreak: false }], 500)
    for (const s of slices) expect(s.repeatHeader).toBeUndefined()
    expect(slices[0].start).toBe(0)
    expect(slices[slices.length - 1].end).toBe(500)
    for (let i = 1; i < slices.length; i++) expect(slices[i].start).toBe(slices[i - 1].end)
  })

  it('floating table flows only when one of its rows is taller than the page', () => {
    const oneRow: BlockBox = {
      top: 0,
      height: 300,
      floated: true,
      floatTable: true,
      tableRows: [{ height: 300 }],
      pageRelVyPx: 40,
      pageRelVAnchor: 'margin',
    }
    const out: SliceOutputs = { floatFlows: [], floatVShifts: [], floatSplits: [] }
    computeSectionedSlicesF2([oneRow, { top: 0, height: 320 }], geoms1, 320, out)
    expect(out.floatFlows).toEqual([{ blockTop: 0 }])
    // a table about to flow takes no --tblp-dy shift and is not split
    expect(out.floatVShifts).toEqual([])
    expect(out.floatSplits).toEqual([])
    // rows not sampled yet: the float is placed whole, never flowed
    const unsampled: BlockBox = { top: 0, height: 300, floated: true, floatTable: true }
    const out2: SliceOutputs = { floatFlows: [], floatSplits: [] }
    computeSectionedSlicesF2([unsampled, { top: 0, height: 320 }], geoms1, 320, out2)
    expect(out2.floatFlows).toEqual([])
    expect(out2.floatSplits).toEqual([])
    // a flowed table whose rows all fit a page goes back to floating
    const flowed = makeTableBlock(
      0,
      Array.from({ length: 6 }, () => ({ height: 50 })),
    )
    flowed.floatTable = true
    flowed.floatFlowed = true
    const out3: SliceOutputs = { floatFlows: [] }
    computeSectionedSlicesF2([flowed], geoms1, 300, out3)
    expect(out3.floatFlows).toEqual([])
  })

  it('floating table taller than the page splits at row boundaries, the last portion floating', () => {
    const rows: TableRowBox[] = Array.from({ length: 6 }, () => ({ height: 50 }))
    const b = makeTableBlock(0, rows)
    b.floated = true
    b.floatTable = true
    // converged state: the anchor paragraph already sits beside the last portion
    const anchor: BlockBox = { top: 200, height: 60 }
    const out: SliceOutputs = { floatFlows: [], floatSplits: [] }
    const slices = computeSectionedSlicesF2([b, anchor], geoms1, 300, out)
    expect(out.floatFlows).toEqual([])
    expect(out.floatSplits).toEqual([{ blockTop: 0, dyPx: 0, cutYs: [200], carryPx: 200 }])
    // rows 0-3 fill page 1 alone; rows 4-5 float on page 2 with the anchor beside
    expect(slices.map((s) => [s.start, s.end])).toEqual([
      [0, 200],
      [200, 300],
    ])
  })

  it('floating table not fitting the remainder splits; the text after it starts beside the continuation', () => {
    const lead: BlockBox = { top: 0, height: 120 }
    const b = makeTableBlock(
      120,
      Array.from({ length: 3 }, () => ({ height: 50 })),
    )
    b.floated = true
    b.floatTable = true
    const anchor: BlockBox = { top: 170, height: 50 }
    const out: SliceOutputs = { floatSplits: [] }
    const slices = computeSectionedSlicesF2([lead, b, anchor], geoms1, 270, out)
    expect(out.floatSplits).toEqual([{ blockTop: 120, dyPx: 0, cutYs: [170], carryPx: 50 }])
    expect(slices.map((s) => [s.start, s.end])).toEqual([
      [0, 170],
      [170, 270],
    ])
    // the whole anchor paragraph fits page 2 next to the float: no further turn
    expect(slices).toHaveLength(2)
  })

  it('floating table whose first row does not fit the remainder moves whole to the next page', () => {
    const lead: BlockBox = { top: 0, height: 180 }
    const b = makeTableBlock(
      180,
      Array.from({ length: 3 }, () => ({ height: 50 })),
    )
    b.floated = true
    b.floatTable = true
    const out: SliceOutputs = { floatSplits: [] }
    const slices = computeSectionedSlicesF2([lead, b, { top: 180, height: 100 }], geoms1, 330, out)
    expect(out.floatSplits).toEqual([])
    expect(slices.map((s) => [s.start, s.end])).toEqual([
      [0, 180],
      [180, 330],
    ])
  })

  it('cantSplit rows of a split floating table stay whole', () => {
    const rows: TableRowBox[] = [
      { height: 50 },
      { height: 50 },
      { height: 50 },
      { height: 80, cantSplit: true },
      { height: 50 },
    ]
    const b = makeTableBlock(0, rows)
    b.floated = true
    b.floatTable = true
    const out: SliceOutputs = { floatSplits: [] }
    computeSectionedSlicesF2([b, { top: 150, height: 130 }], geoms1, 280, out)
    // 150 + 80 > 200: the cantSplit row opens page 2
    expect(out.floatSplits).toEqual([{ blockTop: 0, dyPx: 0, cutYs: [150], carryPx: 150 }])
  })

  it('page-anchored split floating table: cuts are measured from the shifted top', () => {
    const b = makeTableBlock(
      0,
      Array.from({ length: 6 }, () => ({ height: 50 })),
    )
    b.floated = true
    b.floatTable = true
    b.pageRelVyPx = 40
    b.pageRelVAnchor = 'margin'
    const out: SliceOutputs = { floatSplits: [], floatVShifts: [] }
    const slices = computeSectionedSlicesF2([b, { top: 190, height: 100 }], geoms1, 340, out)
    expect(out.floatVShifts).toEqual([{ blockTop: 0, dyPx: 40 }])
    // 40 + 3 rows = 190 fit; the fourth row would end at 240
    expect(out.floatSplits).toEqual([{ blockTop: 0, dyPx: 40, cutYs: [190], carryPx: 190 }])
    expect(slices[0]).toMatchObject({ start: 0, end: 190 })
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

  const twoPages = (cut: number): PageSlice[] => [
    { start: 0, end: cut, section: 0 },
    { start: cut, end: 200, section: 0 },
  ]

  it('keeps lines apart when tall glyph boxes overlap the next line (KR 1.3029 pitch)', () => {
    // GenOffice Sans KR content area 1.448em under a 1.3029 line: each rect
    // runs 2.35px into the next line, so a fixed 1px tolerance merged whole
    // paragraphs into one line and made them atomic at page bottoms
    const el = document.createElement('p')
    el.textContent = 'four lines'
    const para: BlockBox = { top: 0, height: 83.4, el }
    const rects = [0, 20.85, 41.7, 62.55].map((top) => ({
      top,
      bottom: top + 23.2,
      height: 23.2,
      width: 50,
      left: 0,
    }))
    const orig = Range.prototype.getClientRects
    Range.prototype.getClientRects = () => rects as unknown as DOMRectList
    try {
      expect(fillLineBoxes([para], geoms, 1, twoPages(50))).toBe(true)
    } finally {
      Range.prototype.getClientRects = orig
    }
    expect(para.lineBoxes?.length).toBe(4)
    expect(para.lineBoxes?.map((b) => Math.round(b.offsetInBlock * 10) / 10)).toEqual([
      0, 20.9, 41.7, 62.6,
    ])
  })

  it('still joins same-line rects of different font sizes', () => {
    // a 12pt run baseline-aligned on a 24pt line sits 18px below the line top
    const el = document.createElement('p')
    el.textContent = 'one line'
    const para: BlockBox = { top: 0, height: 46, el }
    const rects = [
      { top: 0, bottom: 46, height: 46, width: 80, left: 0 },
      { top: 18.4, bottom: 41.6, height: 23.2, width: 40, left: 80 },
    ]
    const orig = Range.prototype.getClientRects
    Range.prototype.getClientRects = () => rects as unknown as DOMRectList
    try {
      fillLineBoxes([para], geoms, 1, twoPages(30))
    } finally {
      Range.prototype.getClientRects = orig
    }
    expect(para.lineBoxes).toBeUndefined()
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

describe('measureBlocks — adjacent page breaks in one paragraph', () => {
  const para = (html: string) => {
    const pm = document.createElement('div')
    const el = document.createElement('p')
    el.innerHTML = html
    el.getBoundingClientRect = () =>
      ({ top: 0, height: 60, bottom: 60, left: 0, right: 100, width: 100 }) as DOMRect
    pm.appendChild(el)
    return measureBlocks(pm, 0, 1).blocks[0]
  }

  it('two leading breaks before the text: breakBefore plus one extra turn', () => {
    const b = para('<br class="doc-page-br"><br class="doc-page-br">Another Look')
    expect(b).toMatchObject({ breakBefore: true, breakBeforeBr: true, extraBreaksBefore: 1 })
    expect(b.breakAfter).toBeUndefined()
    expect(b.innerBreaks).toBeUndefined()
  })

  it('a leading break cuts at the measured text line: the break line stays on the page before', () => {
    const proto = Range.prototype as unknown as { getClientRects?: () => DOMRect[] }
    const orig = proto.getClientRects
    proto.getClientRects = () =>
      [{ top: 20, height: 20, width: 80, left: 0, right: 80, bottom: 40 }] as DOMRect[]
    try {
      const b = para('<br class="doc-page-br">Business Plan')
      expect(b.innerBreaks).toEqual([20])
      expect(b.breakBefore).toBeUndefined()
      expect(b.breakBeforeBr).toBeUndefined()
      const two = para('<br class="doc-page-br"><br class="doc-page-br">Another Look')
      expect(two).toMatchObject({ innerBreaks: [20], extraBreaksBefore: 1 })
      expect(two.breakBefore).toBeUndefined()
    } finally {
      proto.getClientRects = orig
    }
  })

  it('two trailing breaks after the text: breakAfter plus one extra turn', () => {
    const b = para('Closing line<br class="doc-page-br"><br class="doc-page-br">')
    expect(b).toMatchObject({ breakAfter: true, extraBreaksAfter: 1 })
    expect(b.breakBefore).toBeUndefined()
  })

  it('a break-only paragraph with two breaks queues one extra turn', () => {
    const b = para(
      '<br class="doc-page-br"><br class="doc-page-br"><br class="ProseMirror-trailingBreak">',
    )
    expect(b).toMatchObject({ breakAfter: true, extraBreaksAfter: 1 })
    expect(b.breakBefore).toBeUndefined()
  })

  it('a field page break with only its own label is a break-only paragraph (one turn)', () => {
    const b = para('<div class="doc-field-pagebreak"><span>Page Break</span></div>')
    expect(b).toMatchObject({ breakAfter: true })
    expect(b.breakBefore).toBeUndefined()
    expect(b.extraBreaksAfter).toBeUndefined()
  })

  it('a single leading and a single trailing break keep both turns', () => {
    const b = para('<br class="doc-page-br">Middle<br class="doc-page-br">')
    expect(b).toMatchObject({ breakBefore: true, breakBeforeBr: true, breakAfter: true })
    expect(b.extraBreaksBefore).toBeUndefined()
    expect(b.extraBreaksAfter).toBeUndefined()
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

  it("applyBlockMeta resolves a table block's note markers into per-row bands", () => {
    const el = document.createElement('div')
    el.innerHTML =
      '<table><tbody>' +
      '<tr><td>head<sup class="doc-note-ref" data-note-kind="footnote">1</sup></td></tr>' +
      '<tr><td>plain</td></tr>' +
      '<tr class="page-repeat-header"><td>head<sup class="doc-note-ref" data-note-kind="footnote">1</sup></td></tr>' +
      '<tr><td>tail<sup class="doc-note-ref" data-note-kind="footnote">2</sup></td></tr>' +
      '</tbody></table>'
    el.getBoundingClientRect = () => rectOf(0, 90)
    const trs = Array.from(el.querySelectorAll('tr')).filter(
      (tr) => !tr.classList.contains('page-repeat-header'),
    )
    trs.forEach((tr, i) => (tr.getBoundingClientRect = () => rectOf(i * 30, 30)))
    el.querySelectorAll('sup').forEach((sup) => {
      const top = sup.textContent === '1' ? 5 : 65
      sup.getBoundingClientRect = () => rectOf(top, 10)
    })
    const blocks: BlockBox[] = [{ top: 0, height: 90, docxIndex: 0, el }]
    const bands = [{ heightPx: 20 }, { heightPx: 40 }]
    applyBlockMeta(blocks, () => ({ footnoteExtraPx: 60, footnoteBands: bands }))
    expect(blocks[0].noteBands).toEqual([
      { offset: 5, height: 20 },
      { offset: 65, height: 40 },
    ])
    const rows: TableRowBox[] = [{ height: 30 }, { height: 30 }, { height: 30 }]
    applyRowNotes(el, rows, bands)
    expect(rows.map((r) => r.notesPx)).toEqual([20, undefined, 40])
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

  it("keeps each section's own box when pages differ in width (issue #246)", () => {
    // A4 portrait followed by a narrower page: mixed-width documents must
    // keep per-section geometry so every page measures against its own width.
    const portrait = sectionPageBox(sec({ pageWidth: 11906 }).settings)
    const narrow = sectionPageBox(sec({ pageWidth: 8391 }).settings)
    expect(portrait.width).toBeCloseTo((11906 / 1440) * 96)
    expect(narrow.width).toBeCloseTo((8391 / 1440) * 96)
    expect(narrow.contentWidth).toBeCloseTo(((8391 - 1440 - 1440) / 1440) * 96)
    expect(narrow.width).toBeLessThan(portrait.width)
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

  it('tableRowFlags: a keepNext paragraph in any cell (direct or via pStyle) marks the row keepNext', () => {
    const xml =
      '<w:tbl><w:tr><w:tc><w:p><w:pPr><w:keepNext/></w:pPr></w:p></w:tc><w:tc><w:p/></w:tc></w:tr>' +
      '<w:tr><w:tc><w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr></w:p></w:tc></w:tr>' +
      '<w:tr><w:tc><w:p><w:pPr><w:pStyle w:val="Heading2"/><w:keepNext w:val="0"/></w:pPr></w:p></w:tc></w:tr>' +
      '<w:tr><w:tc><w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr></w:p></w:tc></w:tr></w:tbl>'
    const keep = (id: string) => id === 'Heading2'
    expect(tableRowFlags(xml, keep).map((f) => f.keepNext ?? false)).toEqual([
      true,
      true,
      false,
      false,
    ])
    expect(tableRowFlags(xml).map((f) => f.keepNext ?? false)).toEqual([true, false, false, false])
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

  it('colBreakBefore opens the next column with the block itself (leading w:br column)', () => {
    const blocks = [
      lineBlock(0, [50]),
      lineBlock(50, [20], { colBreakBefore: true }),
      lineBlock(70, [50]),
    ]
    const slices = computeSectionedSlicesF2(blocks, twoCol, 150)
    // the break-only paragraph is the first line of column 2, not the last of column 1
    expect(slices).toHaveLength(1)
    expect(slices[0].regions![0].columns.map((c) => c.start)).toEqual([0, 50])
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

  it('a continuous section keeping the column count still balances the closed region (Word balance-columns idiom)', () => {
    // five 2-line paragraphs in a 2-col section closed by a continuous 2-col section
    const blocks: BlockBox[] = [
      ...Array.from({ length: 5 }, (_, i) => ({
        ...lineBlock(i * 40, [20, 20]),
        section: 0,
      })),
      { ...lineBlock(200, [20]), section: 1, emptyPara: true },
    ]
    const geoms = [
      { contentHeight: 400, forceBreak: false, cols: 2 },
      { contentHeight: 400, forceBreak: false, cols: 2 },
    ]
    const slices = computeSectionedSlicesF2(blocks, geoms, 220)
    expect(slices).toHaveLength(1)
    const regions = slices[0].regions!
    expect(regions).toHaveLength(2)
    expect(regions[0].columns.map((c) => [c.start, c.end])).toEqual([
      [0, 120],
      [120, 200],
    ])
    expect(regions[0].height).toBe(120)
    expect(regions[1].top).toBe(120)
    expect(regions[1].section).toBe(1)
    // a nextColumn boundary keeps flowing in the region (no balance)
    const flow = computeSectionedSlicesF2(
      blocks,
      [geoms[0], { ...geoms[1], colBreakStart: true }],
      220,
    )
    expect(flow[0].regions).toHaveLength(1)
  })

  it('unequal column widths balance at block tops with extents scaled by the column width', () => {
    // cols [400, 200]: measured in the wide column, the overflowing paragraph
    // (140px) would wrap to ~280px in the narrow one. The balance weighs each
    // block by measured width / column width and cuts on block tops only (a
    // block keeps one wrap width), so the last paragraph moves whole to column 2
    // instead of the line-quota cut inside the third paragraph
    const wide = { widthPx: 400 }
    const blocks: BlockBox[] = [
      { ...lineBlock(0, [20], wide), section: 0 },
      { ...lineBlock(20, [20, 20, 20, 20, 20, 20], wide), section: 0 },
      { ...lineBlock(140, [20, 20, 20, 20, 20, 20], wide), section: 0 },
      { ...lineBlock(260, [20, 20, 20, 20, 20, 20, 20], wide), section: 0 },
      { ...lineBlock(400, [20]), section: 1, emptyPara: true },
    ]
    const geoms = [
      { contentHeight: 800, forceBreak: false, cols: 2, colWidths: [400, 200] },
      { contentHeight: 800, forceBreak: false, cols: 2, colWidths: [400, 200] },
    ]
    const slices = computeSectionedSlicesF2(blocks, geoms, 420)
    const regions = slices[0].regions!
    expect(regions[0].columns.map((c) => [c.start, c.end])).toEqual([
      [0, 260],
      [260, 400],
    ])
    // after the remeasure with the paragraph wrapped at the narrow width, the
    // same cut holds and the region grows to the narrow column's extent
    const narrow = blocks.map((b, i) =>
      i === 3
        ? {
            ...lineBlock(
              260,
              Array.from({ length: 14 }, () => 20),
              { widthPx: 200 },
            ),
            section: 0,
          }
        : i === 4
          ? { ...b, top: 540 }
          : b,
    )
    const again = computeSectionedSlicesF2(narrow, geoms, 560)
    const r2 = again[0].regions!
    expect(r2[0].columns.map((c) => [c.start, c.end])).toEqual([
      [0, 260],
      [260, 540],
    ])
    expect(r2[0].height).toBe(280)
    expect(r2[1].top).toBe(280)
  })

  // 171 shape: cols [400, 200] (wide / narrow), heading + two 6-line paragraphs +
  // a 6-line paragraph (wide). Rewrapped narrow the last paragraph runs 14 lines;
  // after k wide head lines its tail runs [14, 13, 10, 8, 6, 3, 0] narrow lines.
  const wrap171 = (): ColWrapTable => {
    const tailLines = [14, 13, 10, 8, 6, 3, 0]
    const headH = Array.from({ length: 7 }, (_, k) => 20 * k)
    const tailH = tailLines.map((n) => 20 * n)
    return {
      headWidthPx: 400,
      tailWidthPx: 200,
      n: 6,
      headH,
      tailH,
      tailBottom: headH.map((h, k) => h + tailH[k]),
      // ink edges 2px inside the 20px line boxes
      shapeY: headH.map((h, k) => (k === 0 || k === 6 ? h : h + 2)),
    }
  }
  const blocks171 = (extra?: Partial<BlockBox>): BlockBox[] => {
    const wide = { widthPx: 400 }
    return [
      { ...lineBlock(0, [20], wide), section: 0 },
      { ...lineBlock(20, [20, 20, 20, 20, 20, 20], wide), section: 0 },
      { ...lineBlock(140, [20, 20, 20, 20, 20, 20], wide), section: 0 },
      {
        ...lineBlock(260, [20, 20, 20, 20, 20, 20], wide),
        section: 0,
        el: document.createElement('p'),
        ...extra,
      },
      { ...lineBlock(380, [20]), section: 1, emptyPara: true },
    ]
  }
  const geoms171 = [
    { contentHeight: 800, forceBreak: false, cols: 2, colWidths: [400, 200] },
    { contentHeight: 800, forceBreak: false, cols: 2, colWidths: [400, 200] },
  ]

  it('unequal columns: with a wrap table the balance cuts inside the paragraph at the line minimizing the taller column, left column first on ties', () => {
    // block-top cut: 13 wide lines / 14 narrow lines; one head line: 14 / 13
    // (same maximum, fuller left column, Word); two head lines: 15 / 12
    const blocks = blocks171({ widowControl: false, colWraps: [wrap171()] })
    const out: SliceOutputs = { colWrapRequests: [] }
    const slices = computeSectionedSlicesF2(blocks, geoms171, 400, out)
    const region = slices[0].regions![0]
    expect(region.columns.map((c) => [c.start, c.end])).toEqual([
      [0, 280],
      [280, 380],
    ])
    expect(out.colWrapRequests).toEqual([])
    // widow control on: a single head line is not allowed, the block-top cut
    // (13 / 14) beats two head lines (15 / 12)
    const widow = blocks171({ colWraps: [wrap171()] })
    const again = computeSectionedSlicesF2(widow, geoms171, 400)
    expect(again[0].regions![0].columns.map((c) => [c.start, c.end])).toEqual([
      [0, 260],
      [260, 380],
    ])
    // keepLines: never cut inside
    const keep = blocks171({ widowControl: false, keepLines: true, colWraps: [wrap171()] })
    expect(computeSectionedSlicesF2(keep, geoms171, 400)[0].regions![0].columns[1].start).toBe(260)
  })

  it('unequal columns: a paragraph near the quota without a wrap table is requested and the pass cuts on block tops', () => {
    const blocks = blocks171()
    const out: SliceOutputs = { colWrapRequests: [] }
    const slices = computeSectionedSlicesF2(blocks, geoms171, 400, out)
    expect(slices[0].regions![0].columns[1].start).toBe(260)
    expect(out.colWrapRequests).toEqual([{ blockTop: 260, headWidthPx: 400, tailWidthPx: 200 }])
  })

  it('columnLineSplits: a cut on a line boundary between windows of different widths (in a region, across a page turn) yields one split per block', () => {
    const els = Array.from({ length: 3 }, () => document.createElement('p'))
    const blocks: BlockBox[] = [
      { ...lineBlock(0, [20]), section: 0, el: els[0] },
      { ...lineBlock(20, [20, 20, 20, 20]), section: 0, el: els[1], colWraps: [wrap171()] },
      { ...lineBlock(100, [20, 20]), section: 0, el: els[2] },
    ]
    const ws = () => [400, 200]
    const region = (start: number, end: number, cut: number) => ({
      top: 0,
      height: end - start,
      section: 0,
      columns: [
        { start, end: cut },
        { start: cut, end },
      ],
    })
    // balance cut after the second line of block 1 (table match), narrow tail
    const inRegion = columnLineSplits(
      blocks,
      [{ start: 0, end: 140, section: 0, regions: [region(0, 140, 60)] }],
      ws,
    )
    expect(inRegion).toEqual([
      {
        bi: 1,
        line: 2,
        headWidthPx: 400,
        tailWidthPx: 200,
        cutY: 40,
        tailBottom: 240,
        shapeY: 42,
      },
    ])
    // page turn from a narrow last column into a wide first column: the block
    // has no table for that pair, the line boxes locate the cut
    const turn = columnLineSplits(
      blocks,
      [
        { start: 0, end: 80, section: 0, regions: [region(0, 80, 10)] },
        { start: 80, end: 140, section: 0, regions: [region(80, 140, 120)] },
      ],
      ws,
    )
    expect(turn).toEqual([
      { bi: 1, line: 3, headWidthPx: 200, tailWidthPx: 400, cutY: 60 },
      { bi: 2, line: 1, headWidthPx: 400, tailWidthPx: 200, cutY: 20 },
    ])
    // an overflow cut that is not one of the table's boundaries (the DOM wraps
    // at another width right now) falls back to the measured line boxes
    const shifted: BlockBox[] = [
      blocks[0],
      {
        ...blocks[1],
        lineBoxes: [16, 32, 48, 64].map((h, m) => ({ offsetInBlock: m * 16, height: 16 })),
      },
      blocks[2],
    ]
    expect(
      columnLineSplits(
        shifted,
        [{ start: 0, end: 140, section: 0, regions: [region(0, 140, 52)] }],
        ws,
      ),
    ).toEqual([{ bi: 1, line: 2, headWidthPx: 400, tailWidthPx: 200, cutY: 32 }])
    // block-top cuts and equal widths never split
    expect(
      columnLineSplits(
        blocks,
        [{ start: 0, end: 140, section: 0, regions: [region(0, 140, 20)] }],
        ws,
      ),
    ).toEqual([])
    expect(
      columnLineSplits(
        blocks,
        [{ start: 0, end: 140, section: 0, regions: [region(0, 140, 60)] }],
        () => [300, 300],
      ),
    ).toEqual([])
  })

  it('fillColWraps: probes once per element and text, reuses the table for fresh block records', () => {
    const el = document.createElement('p')
    el.textContent = 'straddling paragraph'
    let probes = 0
    const probe = () => {
      probes++
      return wrap171()
    }
    const req = [{ blockTop: 100, headWidthPx: 400, tailWidthPx: 200 }]
    const a: BlockBox[] = [{ top: 100, height: 60, el }]
    expect(fillColWraps(a, req, 1, probe)).toBe(true)
    expect(a[0].colWraps).toHaveLength(1)
    expect(fillColWraps(a, req, 1, probe)).toBe(false)
    const b: BlockBox[] = [{ top: 100, height: 60, el }]
    expect(fillColWraps(b, req, 1, probe)).toBe(true)
    expect(probes).toBe(1)
    el.textContent = 'edited paragraph'
    const c: BlockBox[] = [{ top: 100, height: 60, el }]
    expect(fillColWraps(c, req, 1, probe)).toBe(true)
    expect(probes).toBe(2)
    // a failed probe attaches nothing
    expect(
      fillColWraps([{ top: 100, height: 60, el: document.createElement('p') }], req, 1, () => null),
    ).toBe(false)
  })

  it('splitFloatStyle: right float (left in RTL) with the shape inset', () => {
    expect(splitFloatStyle(224, 80, 20.4, undefined)).toBe(
      'float:right;width:224px;height:80px;shape-outside:inset(20.4px 0 0 0)',
    )
    expect(splitFloatStyle(224, 80, 0, true)).toContain('float:left')
  })

  it('widthPassGate: a changed split shape asks for another pass at unchanged widths; split blocks are never pinned', () => {
    const el = document.createElement('p')
    const state = newWidthPassState()
    const at = (h: number) => [
      { el, widthPx: 400, dx: 0, dy: 0, split: { floatPx: 224, heightPx: h, insetPx: 20 } },
    ]
    expect(widthPassGate(state, at(80))).toBe(true)
    expect(widthPassGate(state, at(80))).toBe(false)
    expect(widthPassGate(state, at(120))).toBe(true)
    state.forced.set(el, 200)
    const specs = at(120)
    widthPassGate(state, specs)
    expect(specs[0].widthPx).toBe(400)
  })

  it('unequal columns: an empty paragraph mid-region does not force the cut', () => {
    // cols [400, 200], every block measured wide: the balance target is 2/3 of
    // the content (a block doubles in the narrow column). The blank line after
    // the first paragraph has zero extent and must not be a cut point; the cut
    // lands where the accumulated extent is nearest the quota (top of p4)
    const wide = { widthPx: 400 }
    const blocks: BlockBox[] = [
      { ...lineBlock(0, [20], wide), section: 0 },
      { ...lineBlock(20, [20], wide), section: 0, emptyPara: true },
      { ...lineBlock(40, [20, 20, 20], wide), section: 0 },
      { ...lineBlock(100, [20, 20, 20], wide), section: 0 },
      { ...lineBlock(160, [20, 20, 20], wide), section: 0 },
      { ...lineBlock(220, [20]), section: 1, emptyPara: true },
    ]
    const geoms = [
      { contentHeight: 800, forceBreak: false, cols: 2, colWidths: [400, 200] },
      { contentHeight: 800, forceBreak: false, cols: 2, colWidths: [400, 200] },
    ]
    const slices = computeSectionedSlicesF2(blocks, geoms, 240)
    expect(slices[0].regions![0].columns.map((c) => [c.start, c.end])).toEqual([
      [0, 160],
      [160, 220],
    ])
  })

  it('columnLayoutSpecs: width + per-column constant translate, later regions pull up', () => {
    const els = Array.from({ length: 4 }, () => document.createElement('p'))
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
    const els = Array.from({ length: 4 }, () => document.createElement('p'))
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
    const els = Array.from({ length: 2 }, () => document.createElement('p'))
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
    const els = Array.from({ length: 2 }, () => document.createElement('p'))
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

  it('columnLayoutSpecs: a block cut at a line boundary from a wide column into a narrower one is split: wide width plus a float shape that narrows the tail', () => {
    const els = Array.from({ length: 3 }, () => document.createElement('p'))
    const blocks: BlockBox[] = [
      { ...lineBlock(0, [20]), section: 0, el: els[0] },
      // starts in the wide first column of page 1, continues on page 2's narrow first column
      { ...lineBlock(20, [20, 20, 20, 20]), section: 0, el: els[1] },
      { ...lineBlock(100, [20]), section: 0, el: els[2] },
    ]
    const sectPrXml =
      '<w:sectPr><w:cols w:num="2" w:space="720" w:equalWidth="0"><w:col w:w="6000" w:space="720"/><w:col w:w="2640"/></w:cols></w:sectPr>'
    const secs = [sec({ columns: 2, colSpace: 720 }, { sectPrXml })]
    const g = sectionColGeom(secs[0])
    const page = (start: number, end: number, cut: number): PageSlice => ({
      start,
      end,
      section: 0,
      physHeight: end - start,
      regions: [
        {
          top: 0,
          height: end - start,
          section: 0,
          columns: [
            { start, end: cut },
            { start: cut, end },
          ],
        },
      ],
    })
    // page 1: wide column [0,40) holds block 0 and the first line of block 1; the
    // narrow column [40,60) shows its next line; page 2's wide column the rest
    const specs = columnLayoutSpecs(blocks, [page(0, 60, 40), page(60, 120, 100)], secs)
    expect(specs.map((s) => s.el)).toEqual(els)
    expect(specs[0].widthPx).toBeCloseTo(g.widths[0], 3)
    // one element at the wide width; the float covers the tail lines (from the
    // cut, in-block Y 20, to the text bottom) on the right by the width difference
    expect(specs[1].widthPx).toBeCloseTo(g.widths[0], 3)
    expect(specs[1].dx).toBe(0)
    expect(specs[1].split).toEqual({
      floatPx: g.widths[0] - g.widths[1],
      heightPx: 80,
      insetPx: 21,
    })
    expect(specs[2].split).toBeUndefined()
    // a block whole inside the narrow column keeps that column's width; one whole
    // inside a wide column is not narrowed by columns it never reaches
    expect(specs[2].widthPx).toBeCloseTo(g.widths[1], 3)
    const whole = columnLayoutSpecs(blocks, [page(0, 200, 150)], secs)
    expect(whole[1].widthPx).toBeCloseTo(g.widths[0], 3)
    // only the trailing space after crosses into the narrow column: the lines
    // themselves are whole in the wide one, which keeps its width
    const spaced: BlockBox[] = [
      blocks[0],
      { ...blocks[1], height: 80 + 30, spaceAfterPx: 30 },
      { ...blocks[2], top: 130 },
    ]
    const tail = columnLayoutSpecs(spaced, [page(0, 200, 110)], secs)
    expect(tail[1].widthPx).toBeCloseTo(g.widths[0], 3)
  })

  it('widthPassGate: unequal-column widths re-pass until the landing columns settle, bounded', () => {
    const el = document.createElement('p')
    const at = (w: number) => [{ el, widthPx: w, dx: 0, dy: 0 }]
    const state = newWidthPassState()
    // first decorated pass: blocks got their column widths, line breaks moved
    expect(widthPassGate(state, at(166))).toBe(true)
    // a block moved from the narrow to the wide column: once more
    expect(widthPassGate(state, at(360))).toBe(true)
    // fixed point: same widths, no further pass, run counter cleared
    expect(widthPassGate(state, at(360))).toBe(false)
    expect(state.runs).toBe(0)
    expect(state.forced.size).toBe(0)
    // a run of moving widths stops after maxRuns re-passes
    const drift = newWidthPassState()
    let granted = 0
    for (let i = 0; i < 10; i++) if (widthPassGate(drift, at(100 + i), 3)) granted++
    expect(granted).toBe(3)
    // and recovers once a pass repeats its predecessor
    widthPassGate(drift, at(360), 3)
    expect(widthPassGate(drift, at(360), 3)).toBe(false)
    expect(widthPassGate(drift, at(166), 3)).toBe(true)
    // no regioned pages at all: nothing to converge
    expect(widthPassGate(newWidthPassState(), [])).toBe(false)
  })

  it('widthPassGate: a block ping-ponging between columns is pinned to the narrower width', () => {
    const straddler = document.createElement('p')
    const steady = document.createElement('p')
    const pass = (w: number) => [
      { el: steady, widthPx: 360, dx: 0, dy: 0 },
      { el: straddler, widthPx: w, dx: 0, dy: 0 },
    ]
    const state = newWidthPassState()
    expect(widthPassGate(state, pass(360))).toBe(true)
    expect(widthPassGate(state, pass(166))).toBe(true)
    // wide again = the pass before last: the cycle is caught and the applied
    // specs already carry the narrow width
    const cyc = pass(360)
    expect(widthPassGate(state, cyc)).toBe(true)
    expect(cyc[1].widthPx).toBe(166)
    expect(cyc[0].widthPx).toBe(360)
    expect(state.forced.get(straddler)).toBe(166)
    // the pin holds while the engine keeps wanting the wide column: settled
    const next = pass(360)
    expect(widthPassGate(state, next)).toBe(false)
    expect(next[1].widthPx).toBe(166)
    // a narrower column than the pin still wins (never widen past the landing column)
    const narrower = pass(120)
    widthPassGate(state, narrower)
    expect(narrower[1].widthPx).toBe(120)
    // a document edit releases the pins and forgets the cycle history: the
    // widths of the edited flow are not compared with the old flow's
    resetWidthPassHistory(state)
    const free = pass(360)
    expect(widthPassGate(state, free)).toBe(true)
    expect(free[1].widthPx).toBe(360)
    expect(state.forced.size).toBe(0)
    // a gate exhausted by a long drift is re-armed by an edit
    const drift = newWidthPassState()
    for (let i = 0; i < 10; i++) widthPassGate(drift, pass(100 + i), 3)
    expect(widthPassGate(drift, pass(200), 3)).toBe(false)
    resetWidthPassHistory(drift)
    expect(widthPassGate(drift, pass(300), 3)).toBe(true)
  })

  it('columnLayoutSpecs: an indented paragraph gets the column width minus its margins', () => {
    const indented = document.createElement('p')
    indented.style.marginLeft = '11.33px'
    indented.style.marginRight = '4px'
    const table = document.createElement('table')
    const blocks: BlockBox[] = [
      { ...lineBlock(0, [20]), section: 0, el: indented },
      { ...lineBlock(20, [20]), section: 0, el: table },
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
    const secs = [sec({ columns: 2, colSpace: 720 })]
    const specs = columnLayoutSpecs(blocks, slices, secs)
    const g = sectionColGeom(secs[0])
    expect(specs[0].widthPx).toBeCloseTo(g.widths[0] - 11.33 - 4, 2)
    // tables size themselves: the column width stays the cap
    expect(specs[1].widthPx).toBeCloseTo(g.widths[1], 3)
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

describe('fillLineBoxes — empty paragraphs above a cell heading are row cut points', () => {
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

  it('cuts between leading empty marks and before the heading, not after a trailing empty mark', () => {
    const tbl = document.createElement('table')
    tbl.innerHTML = '<tbody><tr><td></td></tr></tbody>'
    tbl.getBoundingClientRect = () => rectOf(0, 100)
    const td = tbl.querySelector('td')!
    const addPara = (top: number, html: string) => {
      const p = document.createElement('p')
      p.innerHTML = html
      p.getBoundingClientRect = () => rectOf(top, 20)
      td.appendChild(p)
      return p
    }
    addPara(0, '<br class="ProseMirror-trailingBreak">')
    addPara(20, '<br class="ProseMirror-trailingBreak">')
    addPara(40, '<br class="ProseMirror-trailingBreak">')
    const heading = addPara(60, 'Heading')
    addPara(80, '<br class="ProseMirror-trailingBreak">')
    const orig = Range.prototype.getClientRects
    Range.prototype.getClientRects = function (this: Range) {
      const p = this.startContainer.parentElement
      return (p === heading ? [rectOf(60, 20)] : []) as unknown as DOMRectList
    }
    const block: BlockBox = { top: 0, height: 100, el: tbl }
    try {
      expect(fillLineBoxes([block], [{ contentHeight: 50, forceBreak: false }], 1)).toBe(true)
    } finally {
      Range.prototype.getClientRects = orig
    }
    expect(block.tableRows).toHaveLength(1)
    expect(block.tableRows![0].cutYs).toEqual([20, 40, 60])
    expect(block.tableRows![0].contentBottom).toBe(80)
  })
})

describe('computeSectionedSlicesF2 — multi-cell rows split per cell (Word row break)', () => {
  // Word 2013+ layout: widow/orphan control applies inside cells (legacy mode cuts after any line)
  const makeTableBlock = (top: number, rows: TableRowBox[]): BlockBox => ({
    top,
    height: rows.reduce((s, r) => s + r.height, 0),
    tableRows: rows,
    modernTableHeaders: true,
  })
  /** n lines of `pitch` px starting at `top`, all in child `child`, grouped into paragraphs of `perPara` lines */
  const linesOf = (n: number, pitch: number, perPara: number, child = 0, top = 0) => ({
    lines: Array.from({ length: n }, (_, i): [number, number] => [
      top + i * pitch,
      top + (i + 1) * pitch,
    ]),
    childOf: Array.from({ length: n }, () => child),
    paraOf: Array.from({ length: n }, (_, i) => Math.floor(i / perPara)),
  })
  const page = [{ contentHeight: 200, forceBreak: false }]

  it('the first fragment fills the page and each cell breaks at its own line boundary', () => {
    // page 200; a 40px row leaves 160. Cell 0: 15 lines x 20px in 3 paragraphs
    // (8 lines fit, the cut leaves 3+2 lines of paragraph 2). Cell 1: 4 lines x
    // 60px, centered (canvas offset 30): 2 lines fit; on page 2 its remaining lines
    // start flush at the top (shift 40 - 30) with the first two clipped away
    const rows: TableRowBox[] = [
      { height: 40 },
      {
        height: 300,
        contentBottom: 300,
        cells: [
          { ...linesOf(15, 20, 5), alignDy: 0, alignFrac: 0 },
          { ...linesOf(4, 60, 4), alignDy: 30, alignFrac: 0.5 },
        ],
      },
    ]
    const out: SliceOutputs = { rowFills: [], rowSplits: [] }
    const slices = computeSectionedSlicesF2([makeTableBlock(0, rows)], page, 340, out)
    expect(slices.map((s) => [s.start, s.end])).toEqual([
      [0, 200],
      [200, 340],
    ])
    expect(out.rowFills).toEqual([])
    expect(out.rowSplits).toEqual([
      {
        blockTop: 0,
        row: 1,
        rules: [
          { from: 40, cell: 0, child: 0, clipBottom: 140 },
          { from: 40, cell: 0, child: 1, tail: true, hide: true },
          { from: 40, cell: 1, child: 0, tail: true, dy: -30 },
          { from: 40, cell: 1, child: 0, clipBottom: 120 },
          { from: 40, cell: 1, child: 1, tail: true, hide: true },
          { from: 200, cell: 0, child: 0, clipTop: 160 },
          { from: 200, cell: 1, child: 0, tail: true, dy: -30 },
          { from: 200, cell: 1, child: 0, tail: true, dy: 10 },
          { from: 200, cell: 1, child: 0, clipTop: 120 },
        ],
      },
    ])
  })

  it('a centered cell that fits the first fragment whole is centered within the fragment', () => {
    // cell 1 holds one 20px line the canvas centered in the 300px row (offset 140);
    // Word centers it in the 160px fragment instead: (160 - 20) / 2 - 140
    const rows: TableRowBox[] = [
      { height: 40 },
      {
        height: 300,
        contentBottom: 300,
        cells: [
          { ...linesOf(15, 20, 5), alignDy: 0, alignFrac: 0 },
          { ...linesOf(1, 20, 1), alignDy: 140, alignFrac: 0.5 },
        ],
      },
    ]
    const out: SliceOutputs = { rowSplits: [] }
    computeSectionedSlicesF2([makeTableBlock(0, rows)], page, 340, out)
    expect(out.rowSplits![0].rules.filter((r) => r.cell === 1)).toEqual([
      { from: 40, cell: 1, child: 0, tail: true, dy: -70 },
      { from: 200, cell: 1, child: 0, tail: true, hide: true },
    ])
  })

  it('a widow/orphan cut moves the paragraph whole; a continuation shift past the row bottom patches the row height idempotently', () => {
    // cell 1: 4 lines x 75px, one paragraph in child 2: 2 fit (150), the cut leaves
    // 2+2 -> legal; the continuation starts at 160 so the last line ends at 310 > 300
    const cells = [
      { ...linesOf(15, 20, 5), alignDy: 0, alignFrac: 0 },
      { ...linesOf(4, 75, 4, 2), alignDy: 0, alignFrac: 0 },
    ]
    const out: SliceOutputs = { rowFills: [], rowSplits: [] }
    const slices = computeSectionedSlicesF2(
      [makeTableBlock(0, [{ height: 40 }, { height: 300, contentBottom: 300, cells }])],
      page,
      340,
      out,
    )
    // the flow total (340) clamps the last slice; the row itself grows to 310
    expect(slices.map((s) => [s.start, s.end])).toEqual([
      [0, 200],
      [200, 340],
    ])
    expect(out.rowFills).toEqual([{ blockTop: 0, row: 1, targetPx: 310, extraPx: 10 }])
    expect(out.rowSplits![0].rules.filter((r) => r.cell === 1)).toEqual([
      { from: 40, cell: 1, child: 2, clipBottom: 150 },
      { from: 40, cell: 1, child: 3, tail: true, hide: true },
      { from: 200, cell: 1, child: 2, tail: true, dy: 10 },
      { from: 200, cell: 1, child: 2, clipTop: 150 },
    ])
    // the patched DOM row (310px, data-split-extra 10) yields the same plan
    const again: SliceOutputs = { rowFills: [], rowSplits: [] }
    computeSectionedSlicesF2(
      [
        makeTableBlock(0, [
          { height: 40 },
          { height: 310, contentBottom: 300, cells, splitExtra: 10 },
        ]),
      ],
      page,
      350,
      again,
    )
    expect(again.rowFills).toEqual([{ blockTop: 0, row: 1, targetPx: 310, extraPx: 10 }])
    expect(again.rowSplits).toEqual(out.rowSplits)
  })

  it('a fractional row target round-trips through the integer CSS height without drifting', () => {
    // 4 x 75.1px lines: the continuation ends at 310.2 over a 300.4 natural row.
    // The DOM renders height 310 and stores extra 9.6 (310 - 300.4), so the
    // remeasure recovers the same natural height and plan (extra 9.8 again)
    const cells = [
      { ...linesOf(15, 20, 5), alignDy: 0, alignFrac: 0 },
      { ...linesOf(4, 75.1, 4, 2), alignDy: 0, alignFrac: 0 },
    ]
    const plan = (row: TableRowBox) => {
      const out: SliceOutputs = { rowFills: [], rowSplits: [] }
      computeSectionedSlicesF2([makeTableBlock(0, [{ height: 40 }, row])], page, 360, out)
      return out
    }
    const first = plan({ height: 300.4, contentBottom: 300.4, cells })
    expect(first.rowFills![0].targetPx).toBeCloseTo(310.2)
    expect(first.rowFills![0].extraPx).toBeCloseTo(9.8)
    const attrs = rowFillAttrs(first.rowFills![0].targetPx, first.rowFills![0].extraPx)
    expect(attrs).toEqual({ style: 'height:310px', 'data-split-extra': '9.6' })
    const again = plan({ height: 310, contentBottom: 300.4, cells, splitExtra: 9.6 })
    expect(again.rowFills![0].targetPx).toBeCloseTo(310.2)
    expect(again.rowFills![0].extraPx).toBeCloseTo(9.8)
    expect(again.rowSplits).toEqual(first.rowSplits)
  })

  it('a continuation that fits only one more line of a running paragraph holds it back (widow)', () => {
    // one paragraph of 5 x 20px lines; 3 fit the first page (3+2 legal). The next
    // page holds a single line: Word carries the last two lines over together
    // instead of leaving the final line alone on a third page
    const row: TableRowBox = {
      height: 100,
      contentBottom: 100,
      cells: [{ ...linesOf(5, 20, 5), alignDy: 0, alignFrac: 0 }],
    }
    const pages = [20, 200]
    let n = -1
    const plan = planRowSplit(row, 60, 200, () => pages[++n])
    expect(plan).toEqual({
      lastFragment: 40,
      target: 120,
      rules: [
        { from: 0, cell: 0, child: 0, clipBottom: 40 },
        { from: 0, cell: 0, child: 1, tail: true, hide: true },
        { from: 60, cell: 0, child: 0, clipTop: 60, clipBottom: 40 },
        { from: 60, cell: 0, child: 1, tail: true, hide: true },
        { from: 80, cell: 0, child: 0, tail: true, dy: 20 },
        { from: 80, cell: 0, child: 0, clipTop: 60 },
      ],
    })
    // pages that never hold two lines: the rule yields after one held page (no stall)
    const tiny = planRowSplit(row, 60, 200, () => 20)
    expect(tiny).toMatchObject({ lastFragment: 20, target: 120 })
  })

  it('a paragraph opening in the fragment carries two lines over instead of moving whole', () => {
    // 5 x 20px lines, 4 fit: Word cuts after 3 lines (widow), not before the paragraph
    const row: TableRowBox = {
      height: 100,
      contentBottom: 100,
      cells: [{ ...linesOf(5, 20, 5), alignDy: 0, alignFrac: 0 }],
    }
    const plan = planRowSplit(row, 80, 200, () => 200)
    expect(plan).toMatchObject({ lastFragment: 40, target: 120 })
    expect(plan !== 'nofit' && plan?.rules[0]).toEqual({
      from: 0,
      cell: 0,
      child: 0,
      clipBottom: 40,
    })
  })

  it('a 2-line paragraph cannot split: with nothing else fitting the row pushes whole', () => {
    const rows: TableRowBox[] = [
      { height: 150 },
      {
        height: 100,
        contentBottom: 100,
        cells: [
          { ...linesOf(2, 50, 2), alignDy: 0, alignFrac: 0 },
          { ...linesOf(2, 50, 2), alignDy: 0, alignFrac: 0 },
        ],
      },
    ]
    const out: SliceOutputs = { rowSplits: [] }
    const slices = computeSectionedSlicesF2([makeTableBlock(0, rows)], page, 250, out)
    expect(slices.map((s) => [s.start, s.end])).toEqual([
      [0, 150],
      [150, 250],
    ])
    expect(out.rowSplits).toEqual([])
  })

  it('an over-page row keeps splitting page by page; a child cut on both sides gets one combined clip', () => {
    // 500px row on 200px pages: cell 0 lines of 20px tile the pages; cell 1 lines
    // of 30px in 5-line paragraphs: 6 fit but the cut would orphan one line of
    // paragraph 1, so 5 stay (150) and the continuation restarts flush (shift 50);
    // page 2 likewise keeps paragraph 2 whole (top 300 -> shift 100) and its child
    // is clipped at both ends; the last line then ends at 550, growing the row by 50
    const rows: TableRowBox[] = [
      {
        height: 500,
        contentBottom: 500,
        cells: [
          { ...linesOf(25, 20, 5), alignDy: 0, alignFrac: 0 },
          { ...linesOf(15, 30, 5), alignDy: 0, alignFrac: 0 },
        ],
      },
    ]
    const out: SliceOutputs = { rowFills: [], rowSplits: [] }
    const slices = computeSectionedSlicesF2([makeTableBlock(0, rows)], page, 500, out)
    expect(slices.map((s) => [s.start, s.end])).toEqual([
      [0, 200],
      [200, 400],
      [400, 500],
    ])
    expect(out.rowSplits![0].rules.filter((r) => r.cell === 1)).toEqual([
      { from: 0, cell: 1, child: 0, clipBottom: 300 },
      { from: 0, cell: 1, child: 1, tail: true, hide: true },
      { from: 200, cell: 1, child: 0, tail: true, dy: 50 },
      { from: 200, cell: 1, child: 0, clipTop: 150, clipBottom: 150 },
      { from: 200, cell: 1, child: 1, tail: true, hide: true },
      { from: 400, cell: 1, child: 0, tail: true, dy: 100 },
      { from: 400, cell: 1, child: 0, clipTop: 300 },
    ])
    expect(out.rowFills).toEqual([{ blockTop: 0, row: 0, targetPx: 550, extraPx: 50 }])
  })

  it('a row whose lines all fit but whose bottom cell margin hangs over stays on the page', () => {
    // page 200; 40px row, then a 165px row whose last line ends at 158: only the
    // 7px margin overflows, so no empty continuation fragment and no page turn
    const rows: TableRowBox[] = [
      { height: 40 },
      {
        height: 165,
        contentBottom: 158,
        cutYs: [20, 40],
        cells: [
          { ...linesOf(2, 20, 2, 0, 2), alignDy: 0, alignFrac: 0 },
          { ...linesOf(2, 78, 1, 0, 2), alignDy: 0, alignFrac: 0 },
        ],
      },
      { height: 50 },
    ]
    const out: SliceOutputs = { rowFills: [], rowSplits: [] }
    const slices = computeSectionedSlicesF2([makeTableBlock(0, rows)], page, 255, out)
    expect(slices.map((s) => [s.start, s.end])).toEqual([
      [0, 205],
      [205, 255],
    ])
    expect(out.rowSplits).toEqual([])
    expect(out.rowFills).toEqual([])
  })

  it('modernTableHeaders: a header row never stays alone at the page bottom, but a legal first-row split is fine', () => {
    // 200px pages, paragraph 130 leaves 70: the 20px header fits and the first
    // body row has 50px of room. Cell 0 is a 3-line paragraph (2 lines fit -> a
    // 2+1 cut is illegal): Word 2013+ moves the row whole and takes the header
    // with it (probe 2026-09-17, cases 1/18); legacy mode leaves the header and
    // cuts the cell after any line.
    const held = () => ({
      height: 90,
      contentBottom: 90,
      cells: [
        { ...linesOf(3, 20, 3), alignDy: 0, alignFrac: 0 },
        { ...linesOf(5, 18, 5), alignDy: 0, alignFrac: 0 },
      ],
    })
    const table = (rows: TableRowBox[], modern: boolean) => {
      const b = makeTableBlock(130, rows)
      b.modernTableHeaders = modern
      return b
    }
    const page = [{ contentHeight: 200, forceBreak: false }]
    const run = (rows: TableRowBox[], modern: boolean) => {
      const out: SliceOutputs = { rowSplits: [] }
      const slices = computeSectionedSlicesF2(
        [{ top: 0, height: 130 }, table(rows, modern)],
        page,
        300,
        out,
      )
      return { slices, out }
    }
    const body = Array.from({ length: 3 }, () => ({ height: 20 }))
    const modern = run([{ height: 20, isHeader: true }, held(), ...body], true)
    expect(modern.slices.map((s) => [s.start, s.end])).toEqual([
      [0, 130],
      [130, 300],
    ])
    expect(modern.slices[1].repeatHeader).toBeUndefined()
    expect(modern.out.rowSplits).toEqual([])
    const legacy = run([{ height: 20, isHeader: true }, held(), ...body], false)
    expect(legacy.slices.map((s) => [s.start, s.end])).toEqual([
      [0, 200],
      [200, 300],
    ])
    expect(legacy.slices[1].repeatHeader).toEqual({ top: 130, height: 20 })
    // a first row whose cells all break legally splits under the header in modern mode too
    const legal = run(
      [
        { height: 20, isHeader: true },
        {
          height: 90,
          contentBottom: 90,
          cells: [{ ...linesOf(5, 18, 5), alignDy: 0, alignFrac: 0 }],
        },
        ...body,
      ],
      true,
    )
    expect(legal.slices.map((s) => [s.start, s.end])).toEqual([
      [0, 200],
      [200, 300],
    ])
    expect(legal.slices[1].repeatHeader).toEqual({ top: 130, height: 20 })
    // the header is only retracted when the first body row turns the page: a
    // cantSplit first row does the same, a later row leaves the header in place
    const cant = run(
      [{ height: 20, isHeader: true }, { height: 60, cantSplit: true }, ...body],
      true,
    )
    expect(cant.slices[0].end).toBe(130)
    const later = run(
      [{ height: 20, isHeader: true }, { height: 20 }, { height: 60, cantSplit: true }, ...body],
      true,
    )
    expect(later.slices[0].end).toBe(170)
    expect(later.slices[1].repeatHeader).toEqual({ top: 130, height: 20 })
  })

  it('a cell that cannot break legally moves the whole row; a fresh page still pushes it onto the continuation centered', () => {
    // cell 1: one 2-line paragraph (60px) with room for one line in the 40px
    // remainder: the widow rule holds it, and Word moves the row to page 2
    // instead of leaving the cell empty (probe 2026-09-17, 2-line cell / 1.5 lines)
    const rows: TableRowBox[] = [
      { height: 160 },
      {
        height: 300,
        contentBottom: 300,
        cells: [
          { ...linesOf(15, 20, 5), alignDy: 0, alignFrac: 0 },
          { ...linesOf(2, 30, 2), alignDy: 120, alignFrac: 0.5 },
        ],
      },
    ]
    const out: SliceOutputs = { rowSplits: [] }
    const slices = computeSectionedSlicesF2([makeTableBlock(0, rows)], page, 460, out)
    expect(slices.map((s) => [s.start, s.end])).toEqual([
      [0, 160],
      [160, 360],
      [360, 460],
    ])
    // page 2 holds the whole 2-line cell centered in the page-filling fragment
    // (cell 0 runs on to page 3): (200 - 60) / 2 minus the canvas centering offset
    expect(out.rowSplits![0].rules.filter((r) => r.cell === 1)).toEqual([
      { from: 160, cell: 1, child: 0, tail: true, dy: 70 - 120 },
      { from: 360, cell: 1, child: 0, tail: true, hide: true },
    ])
    // on a fresh page the held cell cannot move the row any further: it lands
    // whole on the continuation, flush at the page top (shift 40) plus centering
    const fresh: TableRowBox[] = [
      {
        height: 300,
        contentBottom: 300,
        cells: [
          { ...linesOf(15, 20, 5), alignDy: 0, alignFrac: 0 },
          { ...linesOf(2, 30, 2, 0, 160), alignDy: 40, alignFrac: 0.5 },
        ],
      },
    ]
    const out2: SliceOutputs = { rowSplits: [] }
    computeSectionedSlicesF2([makeTableBlock(0, fresh)], page, 300, out2)
    // the 100px closing fragment centers the 60px cell: (100 - 60) / 2
    expect(out2.rowSplits![0].rules.filter((r) => r.cell === 1)).toEqual([
      { from: 0, cell: 1, child: 0, tail: true, dy: -40 },
      { from: 0, cell: 1, child: 0, tail: true, hide: true },
      { from: 200, cell: 1, child: 0, tail: true, dy: 20 },
    ])
  })

  it('rowSplitCss emits each rule on the page its fragment lands on', () => {
    const tbl = document.createElement('table')
    tbl.innerHTML = '<tbody><tr><td><p>a</p></td><td><p>b</p><p>c</p></td></tr></tbody>'
    const block: BlockBox = { top: 100, height: 300, el: tbl, tableRows: [{ height: 300 }] }
    const css = rowSplitCss(
      [
        {
          blockTop: 100,
          row: 0,
          rules: [
            { from: 100, cell: 0, child: 2, tail: true, hide: true },
            { from: 200, cell: 1, child: 1, tail: true, dy: 20 },
            { from: 400, cell: 0, child: 0, clipTop: 10, clipBottom: 5 },
          ],
        },
      ],
      [block],
      [{ start: 0 }, { start: 200 }, { start: 400 }, { start: 600 }],
    )
    expect(tbl.querySelector('tr')!.dataset.pvSplit).toBe('0')
    expect(css.split('\n')).toEqual([
      '.pv-page[data-pv-page="0"] .pv-content tr[data-pv-split="0"] > :nth-child(1) > :nth-child(n+3){visibility:hidden}',
      '.pv-page[data-pv-page="1"] .pv-content tr[data-pv-split="0"] > :nth-child(2) > :nth-child(n+2){transform:translateY(20.0px)}',
      '.pv-page[data-pv-page="2"] .pv-content tr[data-pv-split="0"] > :nth-child(1) > :nth-child(1){clip-path:inset(10.0px 0 5.0px 0)}',
    ])
  })

  it('cellLinesOf clusters same-line rects and keeps the topmost rect child/paragraph ids', () => {
    const res = cellLinesOf([
      { band: [20, 40], child: 1, para: 1 },
      { band: [0, 20], child: 0, para: 0 },
      { band: [2, 19], child: 0, para: 0 },
      { band: [41, 60], child: 1, para: 1 },
    ])
    expect(res.lines).toEqual([
      [0, 20],
      [20, 40],
      [41, 60],
    ])
    expect(res.childOf).toEqual([0, 1, 1])
    expect(res.paraOf).toEqual([0, 1, 1])
  })
})

describe('vertical-text sections (sectPr w:textDirection)', () => {
  // landscape A4, 1800 twips top/bottom, 1440 left/right
  const vert = (over: Partial<SectionInfo['settings']> = {}) =>
    sec({
      pageWidth: 16838,
      pageHeight: 11906,
      orientation: 'landscape',
      marginTop: 1800,
      marginBottom: 1800,
      textDirection: 'tbRl',
      ...over,
    })
  const px = (twips: number) => (twips / 1440) * 96

  it('sectionGeoms swaps the axes: capacity is the side-margin width, line length the body height', () => {
    const [g] = sectionGeoms([vert()])
    expect(g.vertical).toBe('tbRl')
    expect(g.contentHeight).toBeCloseTo(px(16838 - 2880), 3)
    expect(g.contentWidth).toBeCloseTo(px(11906 - 3600), 3)
    expect(g.firstContentHeight).toBeUndefined()
    const [h] = sectionGeoms([vert({ textDirection: 'lrTb' })])
    expect(h.vertical).toBeUndefined()
    expect(h.contentHeight).toBeCloseTo(px(11906 - 3600), 3)
  })

  it('verticalTextSpecs places blocks from the right edge and trims their flow footprint', () => {
    const els = Array.from({ length: 3 }, () => document.createElement('p'))
    const blocks: BlockBox[] = [
      { ...lineBlock(0, [40]), section: 0, el: els[0] },
      { ...lineBlock(50, [100]), section: 0, el: els[1] },
      { ...lineBlock(150, [30]), section: 0, el: els[2] },
    ]
    const slices: PageSlice[] = [{ start: 0, end: 180, section: 0 }]
    const secs = [vert()]
    const geoms = sectionGeoms(secs)
    const specs = verticalTextSpecs(blocks, slices, secs, geoms)
    const W = geoms[0].contentHeight
    const L = geoms[0].contentWidth!
    expect(specs).toHaveLength(3)
    expect(specs[0]).toMatchObject({ el: els[0], dy: 0, widthPx: L })
    expect(specs[0].dx).toBeCloseTo(W - 40, 3)
    expect(specs[0].vertical).toMatchObject({ mode: 'tbRl', firstOnPage: true })
    // next block starts 50 below: the box is L tall, so margin-bottom pulls 50 - L
    expect(specs[0].vertical!.marginBottomPx).toBeCloseTo(50 - L, 3)
    expect(specs[1].dx).toBeCloseTo(W - 50 - 100, 3)
    expect(specs[1].dy).toBe(-50)
    expect(specs[1].vertical!.marginBottomPx).toBeCloseTo(100 - L, 3)
    expect(specs[1].vertical!.firstOnPage).toBe(false)
    // last block: the page's flow footprint becomes min(used, L) = 180
    expect(specs[2].vertical!.marginBottomPx).toBeCloseTo(180 - 150 - L, 3)
  })

  it('measured inline extras win over the live margins when trimming the line length', () => {
    const el = document.createElement('p')
    el.style.marginLeft = '30px'
    const secs = [vert()]
    const geoms = sectionGeoms(secs)
    const L = geoms[0].contentWidth!
    const slices: PageSlice[] = [{ start: 0, end: 40, section: 0 }]
    const live = verticalTextSpecs([{ ...lineBlock(0, [40]), section: 0, el }], slices, secs, geoms)
    expect(live[0].widthPx).toBeCloseTo(L - 30, 3)
    const captured = verticalTextSpecs(
      [{ ...lineBlock(0, [40]), section: 0, el, inlineExtraPx: 12 }],
      slices,
      secs,
      geoms,
    )
    expect(captured[0].widthPx).toBeCloseTo(L - 12, 3)
  })

  it('a page filled past the body height keeps a body-height footprint; btLr fills from the left', () => {
    const els = Array.from({ length: 2 }, () => document.createElement('p'))
    const secs = [vert({ textDirection: 'btLr' })]
    const geoms = sectionGeoms(secs)
    const L = geoms[0].contentWidth!
    const blocks: BlockBox[] = [
      { ...lineBlock(0, [700]), section: 0, el: els[0] },
      { ...lineBlock(700, [100]), section: 0, el: els[1] },
    ]
    const specs = verticalTextSpecs(blocks, [{ start: 0, end: 800, section: 0 }], secs, geoms)
    expect(specs[0].dx).toBe(0)
    expect(specs[1].dx).toBe(700)
    expect(specs[1].vertical!.marginBottomPx).toBeCloseTo(L - 700 - L, 3)
    expect(specs[0].vertical!.mode).toBe('btLr')
  })

  it('vAlign shifts skip vertical-text pages (their contentHeight is the sideways fill)', () => {
    const el = document.createElement('p')
    const secs = [vert({ vAlign: 'center' })]
    const blocks: BlockBox[] = [{ ...lineBlock(0, [40]), section: 0, el }]
    expect(
      vAlignShiftSpecs(blocks, [{ start: 0, end: 40, section: 0 }], secs, sectionGeoms(secs)),
    ).toEqual([])
  })

  it('horizontal sections and region pages get no vertical specs', () => {
    const el = document.createElement('p')
    const blocks: BlockBox[] = [{ ...lineBlock(0, [40]), section: 0, el }]
    expect(
      verticalTextSpecs(
        blocks,
        [{ start: 0, end: 40, section: 0 }],
        [sec({})],
        sectionGeoms([sec({})]),
      ),
    ).toEqual([])
  })

  it('verticalPageCss hides gap chrome siblings and reveals only the blocks on each page', () => {
    const mk = (idx: number, vert: boolean) => {
      const el = document.createElement('p')
      el.dataset.idx = String(idx)
      if (vert) el.classList.add('doc-vert-block')
      return el
    }
    const blocks: BlockBox[] = [
      { ...lineBlock(0, [40]), section: 0, el: mk(0, true) },
      { ...lineBlock(40, [60]), section: 0, el: mk(1, false) },
      { ...lineBlock(100, [30]), section: 0, el: mk(2, true) },
    ]
    const slices: PageSlice[] = [
      { start: 0, end: 100, section: 0 },
      { start: 100, end: 130, section: 0 },
    ]
    const secs = [vert()]
    const rules = verticalPageCss(blocks, slices, secs, sectionGeoms(secs)).split('\n')
    const hide = (i: number) =>
      `.pv-page[data-pv-page="${i}"] .pv-content > [data-idx],` +
      `.pv-page[data-pv-page="${i}"] .pv-content > :is([class*="page-gap"],.page-float-host){visibility:hidden;}`
    expect(rules[0]).toBe(hide(0))
    expect(rules[1]).toMatch(
      /^\.pv-page\[data-pv-page="0"\] \.pv-content > \[data-idx="0"\]\{visibility:visible;--pv-vdx:/,
    )
    expect(rules[2]).toBe(
      '.pv-page[data-pv-page="0"] .pv-content > [data-idx="1"]{visibility:visible;}',
    )
    expect(rules[3]).toBe(hide(1))
    expect(rules[4]).toMatch(
      /^\.pv-page\[data-pv-page="1"\] \.pv-content > \[data-idx="2"\]\{visibility:visible;--pv-vdx:/,
    )
    expect(rules).toHaveLength(5)
  })
})

describe('row splits cut between line boxes (box-relative clips)', () => {
  // 20px line pitch with 14px ink bands: each line's ink sits 3px inside its box
  const inkLines = (n: number, child = 0) => ({
    lines: Array.from({ length: n }, (_, i): [number, number] => [i * 20 + 3, i * 20 + 17]),
    childOf: Array.from({ length: n }, () => child),
    paraOf: Array.from({ length: n }, () => child),
  })
  const cells = () => [
    { ...inkLines(10), childBox: [[0, 200]] as Array<[number, number]>, alignDy: 0, alignFrac: 0 },
    { ...inkLines(3), childBox: [[0, 60]] as Array<[number, number]>, alignDy: 0, alignFrac: 0 },
  ]
  const bodyRow = (): TableRowBox => ({ height: 200, contentBottom: 197, cells: cells() })
  const page = [{ contentHeight: 120, forceBreak: false }]
  const tableBlock = (rows: TableRowBox[]): BlockBox => ({
    top: 0,
    height: rows.reduce((s, r) => s + r.height, 0),
    tableRows: rows,
  })

  it('clips the outgoing fragment in the leading gap, measured from the box bottom', () => {
    // 90px available: 4 lines (ink bottoms 17..77) fit, line 5 (83..97) does not.
    // The boundary is the gap middle (80) and the inset counts from the 200px box
    // bottom: 120. The old ink-relative inset (197 - 83 = 114) left the top 3px
    // of line 5's glyphs on the page.
    const plan = planRowSplit(bodyRow(), 90, 120, () => 120)
    const rules = plan && plan !== 'nofit' ? plan.rules : []
    expect(rules).toContainEqual({ from: 0, cell: 0, child: 0, clipBottom: 120 })
    // the continuation opens at that boundary: line 5's box top lands on the page top
    expect(rules).toContainEqual({ from: 90, cell: 0, child: 0, tail: true, dy: 10 })
    expect(rules).toContainEqual({ from: 90, cell: 0, child: 0, clipTop: 80 })
    expect(rules).toContainEqual({ from: 90, cell: 1, child: 0, tail: true, hide: true })
  })

  it('a row under a tblHeader row continues below the repeated header from its first unseen line', () => {
    const out: SliceOutputs = { rowFills: [], rowSplits: [] }
    const slices = computeSectionedSlicesF2(
      [tableBlock([{ height: 30, isHeader: true }, bodyRow()])],
      page,
      230,
      out,
    )
    expect(slices.map((s) => [s.start, s.repeatHeader?.height])).toEqual([
      [0, undefined],
      [120, 30],
      [210, 30],
    ])
    expect(out.rowSplits![0].rules.filter((r) => r.cell === 0)).toEqual([
      { from: 30, cell: 0, child: 0, clipBottom: 120 },
      { from: 30, cell: 0, child: 1, tail: true, hide: true },
      { from: 120, cell: 0, child: 0, tail: true, dy: 10 },
      { from: 120, cell: 0, child: 0, clipTop: 80, clipBottom: 40 },
      { from: 120, cell: 0, child: 1, tail: true, hide: true },
      { from: 210, cell: 0, child: 0, tail: true, dy: 20 },
      { from: 210, cell: 0, child: 0, clipTop: 160 },
    ])
    expect(out.rowFills).toEqual([{ blockTop: 0, row: 1, targetPx: 220, extraPx: 20 }])
  })

  it('without a header the boundary at the page bottom leaves the continuation a plain clip', () => {
    // 6 lines fit (ink bottom 117); the boundary (120) is the page bottom itself
    const out: SliceOutputs = { rowFills: [], rowSplits: [] }
    const slices = computeSectionedSlicesF2([tableBlock([bodyRow()])], page, 200, out)
    expect(slices.map((s) => s.start)).toEqual([0, 120])
    expect(slices[1].repeatHeader).toBeUndefined()
    expect(out.rowSplits![0].rules.filter((r) => r.cell === 0)).toEqual([
      { from: 0, cell: 0, child: 0, clipBottom: 80 },
      { from: 0, cell: 0, child: 1, tail: true, hide: true },
      { from: 120, cell: 0, child: 0, clipTop: 120 },
    ])
    expect(out.rowFills).toEqual([])
  })

  it('fillLineBoxes records each cell child box next to its ink lines', () => {
    const wrap = document.createElement('div')
    wrap.innerHTML =
      '<table><tbody><tr><td><p>long</p></td><td><p>short</p></td></tr></tbody></table>'
    const [longP, shortP] = Array.from(wrap.querySelectorAll('p'))
    const rect = (top: number, bottom: number) =>
      ({ top, bottom, height: bottom - top, left: 0, right: 100, width: 100 }) as DOMRect
    const elProto = Element.prototype
    const origBox = elProto.getBoundingClientRect
    const rangeProto = Range.prototype as unknown as { getClientRects?: () => DOMRect[] }
    const origRects = rangeProto.getClientRects
    elProto.getBoundingClientRect = function (this: Element) {
      if (this === longP) return rect(3, 203)
      if (this === shortP) return rect(3, 63)
      return rect(0, 200)
    }
    rangeProto.getClientRects = function (this: Range) {
      const n = this.startContainer.parentElement === longP ? 10 : 3
      return Array.from({ length: n }, (_, i) => rect(i * 20 + 6, i * 20 + 20))
    }
    try {
      const table: BlockBox = { top: 0, height: 200, el: wrap }
      expect(fillLineBoxes([table], [{ contentHeight: 120, forceBreak: false }], 1)).toBe(true)
      const [c0, c1] = table.tableRows![0].cells!
      expect(c0.childBox).toEqual([[3, 203]])
      expect(c0.lines[0]).toEqual([6, 20])
      expect(c1.childBox).toEqual([[3, 63]])
    } finally {
      elProto.getBoundingClientRect = origBox
      rangeProto.getClientRects = origRects
    }
  })
})
