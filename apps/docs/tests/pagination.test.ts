import { describe, expect, it } from 'vitest'
import type { SectionInfo } from '@genoffice/docx-engine'
import {
  appendEndnotesBlock,
  assignSections,
  cellCutYs,
  computePageSlices,
  computeSectionedSlices,
  computeSectionedSlicesF2,
  effectiveHfRefs,
  hasPrintableHeaderFooter,
  insertParityBlanks,
  lineBreakBoundaries,
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

  it('breakAfter at the end of the document does not create an empty page', () => {
    const slices = computePageSlices([block(0, 100, { breakAfter: true })], 800, 100)
    expect(slices).toEqual([{ start: 0, end: 100, section: 0 }])
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
  it('ignores the first glyph top and only returns later line starts', () => {
    expect(lineBreakBoundaries([3.25, 22.5, 41.75])).toEqual([22.5, 41.75])
  })

  it('returns no boundary for a single visual line', () => {
    expect(lineBreakBoundaries([2.75])).toEqual([])
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

  it('counts parity-blank pairs once so NUMPAGES matches the drawn pages', () => {
    expect(visiblePageCount(withBlank)).toBe(3)
    expect(visiblePageCount([{ start: 0, end: 800, section: 0 }])).toBe(1)
    expect(visiblePageCount([])).toBe(0)
  })

  it('maps a physical pageAt index to its visible page number', () => {
    // y=900 lands on physical slice 3 (the real page after the blank) = visible page 2
    expect(visiblePageCount(withBlank, pageAt(withBlank, 900))).toBe(2)
    expect(visiblePageCount(withBlank, pageAt(withBlank, 0))).toBe(1)
    expect(visiblePageCount(withBlank, pageAt(withBlank, 1700))).toBe(3)
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

  it('sectionGeoms: continuous with identical geometry keeps the flow; differing geometry promotes to a page break', () => {
    const a = sec({})
    const cont = sec({}, { startType: 'continuous' })
    const contWide = sec({ pageWidth: 16838, pageHeight: 11906 }, { startType: 'continuous' })
    const next = sec({})
    const geoms = sectionGeoms([a, cont, contWide, next])
    expect(geoms.map((g) => g.forceBreak)).toEqual([false, false, true, true])
    expect(Math.round(geoms[0].contentHeight)).toBe(Math.round(((16838 - 2880) / 1440) * 96))
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

describe('computeSectionedSlicesF2 — line-level pagination', () => {
  it('content shorter than a page does not break', () => {
    const b = lineBlock(0, [50, 50, 50])
    const slices = computeSectionedSlicesF2([b], geoms1, 150)
    expect(slices.length).toBe(1)
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
  const breakDoc = (fillerH: number) => {
    const pm = document.createElement('div')
    const addPara = (top: number, height: number, html: string) => {
      const el = document.createElement('p')
      el.innerHTML = html
      el.getBoundingClientRect = () => rectOf(top, height)
      pm.appendChild(el)
    }
    addPara(0, fillerH, 'filler text')
    addPara(fillerH, 44, '<br class="doc-page-br"><br class="ProseMirror-trailingBreak">')
    addPara(fillerH + 44, 100, 'after the break')
    return measureBlocks(pm, 0, 1)
  }
  const geoms = [{ contentHeight: 200, forceBreak: false }]

  it('absorbs at the page bottom when the break line fits (no blank page)', () => {
    const { blocks, totalHeight } = breakDoc(173) // 27px left
    expect(blocks[1].breakAfter).toBe(true)
    expect(blocks[1].breakOnlyLineH).toBe(44)
    const slices = computeSectionedSlicesF2(blocks, geoms, totalHeight)
    // break paragraph overflows the bottom margin; page 2 starts at the following block
    expect(slices.map((s) => s.start)).toEqual([0, 217])
  })

  it('opens a Word-style blank page when the previous paragraph exactly fills the page', () => {
    const { blocks, totalHeight } = breakDoc(200)
    const slices = computeSectionedSlicesF2(blocks, geoms, totalHeight)
    // page 2 holds only the break paragraph (blank), which then pushes the rest to page 3
    expect(slices.map((s) => s.start)).toEqual([0, 200, 244])
    expect(slices[1]).toMatchObject({ start: 200, end: 244 })
  })

  it('still absorbs within the drift tolerance (half the block height)', () => {
    const { blocks, totalHeight } = breakDoc(177) // 23px left, just above the 22px floor
    const slices = computeSectionedSlicesF2(blocks, geoms, totalHeight)
    expect(slices.map((s) => s.start)).toEqual([0, 221])
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
