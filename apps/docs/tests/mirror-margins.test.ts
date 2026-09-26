import { describe, expect, it } from 'vitest'
import type { SectionInfo, SectionSettings } from '@genoffice/docx-engine'
import { mirrorShiftPx, pageMargins } from '../src/renderer/page-margins'
import {
  insertParityBlanks,
  mirrorMarginSpecs,
  pageNumbers,
  type BlockBox,
  type PageSlice,
  type SectionGeom,
} from '../src/renderer/pagination'

/** inside 1.25in (w:left 1440 + w:gutter 360, folded by the parser), outside 1in */
const settings = (over: Partial<SectionSettings> = {}): SectionSettings =>
  ({
    pageWidth: 12240,
    pageHeight: 15840,
    marginTop: 1440,
    marginBottom: 1440,
    marginLeft: 1800,
    marginRight: 1440,
    gutter: 360,
    ...over,
  }) as SectionSettings

const section = (over: Partial<SectionInfo> = {}): SectionInfo =>
  ({ settings: settings(), startType: 'nextPage', firstBlockIndex: 0, ...over }) as SectionInfo

const block = (top: number, height = 100): BlockBox =>
  ({ el: document.createElement('p'), top, height }) as BlockBox

describe('pageMargins', () => {
  it('reads left/right as authored without mirroring', () => {
    expect(pageMargins(settings(), 2, false)).toEqual({ left: 1800, right: 1440 })
  })

  it('keeps inside (gutter included) on the left of odd pages and swaps on even pages', () => {
    expect(pageMargins(settings(), 1, true)).toEqual({ left: 1800, right: 1440 })
    expect(pageMargins(settings(), 2, true)).toEqual({ left: 1440, right: 1800 })
    expect(pageMargins(settings(), 3, true)).toEqual({ left: 1800, right: 1440 })
  })

  it('shifts an even page by outside minus inside in px', () => {
    expect(mirrorShiftPx(settings(), 1, true)).toBe(0)
    expect(mirrorShiftPx(settings(), 2, true)).toBe(-24)
    expect(mirrorShiftPx(settings(), 2, false)).toBe(0)
  })
})

describe('mirrorMarginSpecs', () => {
  const slices: PageSlice[] = [
    { start: 0, end: 500, section: 0 },
    { start: 500, end: 1000, section: 0 },
    { start: 1000, end: 1500, section: 0 },
  ]
  const blocks = [block(0), block(200), block(500), block(700), block(1000)]

  it('translates only the blocks of even pages', () => {
    const specs = mirrorMarginSpecs(blocks, slices, [section()], undefined, [1, 2, 3])
    expect(specs.map((s) => s.el)).toEqual([blocks[2].el, blocks[3].el])
    expect(specs.map((s) => s.dx)).toEqual([-24, -24])
    expect(specs.every((s) => s.dy === 0)).toBe(true)
  })

  it('follows displayed page numbers, so a section numbered from 2 mirrors its first page', () => {
    const secs = [section({ pageNumberStart: 2 })]
    const nums = pageNumbers(slices, secs)
    expect(nums).toEqual([2, 3, 4])
    const specs = mirrorMarginSpecs(blocks, slices, secs, undefined, nums)
    expect(specs.map((s) => s.el)).toEqual([blocks[0].el, blocks[1].el, blocks[4].el])
  })

  it('mirrors the pages of even/odd section starts by their physical parity', () => {
    const secs = [
      section({ lastBlockIndex: 1 }),
      section({ startType: 'evenPage', firstBlockIndex: 2, lastBlockIndex: 9 }),
    ]
    const geoms: SectionGeom[] = [
      { contentHeight: 500, forceBreak: false },
      { contentHeight: 500, forceBreak: true, startType: 'evenPage' },
    ]
    const raw: PageSlice[] = [
      { start: 0, end: 500, section: 0 },
      { start: 500, end: 1000, section: 1 },
    ]
    // evenPage after page 1: page 2 already has the right parity, the section is mirrored
    const even = insertParityBlanks(raw, geoms)
    expect(even.length).toBe(2)
    const evenNums = pageNumbers(even, secs)
    expect(evenNums).toEqual([1, 2])
    const evenSpecs = mirrorMarginSpecs(blocks.slice(0, 4), even, secs, undefined, evenNums)
    expect(evenSpecs.map((s) => s.el)).toEqual([blocks[2].el, blocks[3].el])

    // oddPage after page 1: a blank even page is inserted, the section opens on page 3 unshifted
    const oddSecs = [secs[0], { ...secs[1], startType: 'oddPage' as const }]
    const odd = insertParityBlanks(raw, [geoms[0], { ...geoms[1], startType: 'oddPage' }])
    expect(odd.map((s) => s.end - s.start)).toEqual([500, 0, 500])
    const oddNums = pageNumbers(odd, oddSecs)
    expect(oddNums).toEqual([1, 2, 3])
    expect(mirrorMarginSpecs(blocks.slice(0, 4), odd, oddSecs, undefined, oddNums)).toEqual([])
  })

  it('leaves floated blocks and equal margins alone', () => {
    const floated = { ...block(500), floated: true } as BlockBox
    expect(mirrorMarginSpecs([floated], slices, [section()], undefined, [1, 2, 3])).toEqual([])
    const even = section({ settings: settings({ marginLeft: 1440 }) })
    expect(mirrorMarginSpecs(blocks, slices, [even], undefined, [1, 2, 3])).toEqual([])
  })

  it('falls back to the canvas section for single-section documents', () => {
    const specs = mirrorMarginSpecs(blocks, slices, [], settings(), [1, 2, 3])
    expect(specs.map((s) => s.dx)).toEqual([-24, -24])
  })
})
