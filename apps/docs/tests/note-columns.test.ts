import { afterAll, beforeAll, describe, expect, it, onTestFinished } from 'vitest'
import { noteSeparatorRow, setNoteSeparatorRows } from '../src/renderer/line-metrics'
import {
  noteAreaHeight,
  noteAreaIncrements,
  noteAreaPx,
  noteColumnLayout,
  noteColumnWidth,
  noteColumnsClass,
  noteColumnsOf,
  noteColumnsStyle,
  noteLines,
  type NoteColumns,
} from '../src/renderer/note-columns'
import {
  TEXT_COLUMN_ROUNDS,
  noteColumnsPass,
  type NoteColumnsStats,
} from '../src/renderer/pagination-note-columns'
import {
  computeSectionedSlicesF2,
  noteReserveKey,
  type BlockBox,
  type PageSlice,
  type SliceOutputs,
} from '../src/renderer/pagination'

/** notes of the given line counts, every line 10 tall */
const notes = (...counts: number[]) => counts.map((n) => Array.from({ length: n }, () => 10))

describe('columns of a footnote area', () => {
  const sec = (columns: number, footnoteColumns?: number, colSpace?: number, colSep?: boolean) => ({
    columns,
    ...(colSpace === undefined ? {} : { colSpace }),
    ...(colSep ? { colSep } : {}),
    ...(footnoteColumns === undefined ? {} : { footnoteColumns }),
  })

  it('takes the count the section states', () => {
    expect(noteColumnsOf(sec(1, 2), 15)).toBe(2)
    expect(noteColumnsOf(sec(1, 3), 15)).toBe(3)
    expect(noteColumnsOf(sec(3, 2), 15)).toBe(2)
    expect(noteColumnsOf(sec(2, 1), 15)).toBe(1)
  })

  it('matches the section when the section states none', () => {
    expect(noteColumnsOf(sec(1), 15)).toBe(1)
    expect(noteColumnsOf(sec(2), 15)).toBe(2)
    expect(noteColumnsOf(sec(3), 15)).toBe(3)
  })

  it('stays one column in a document laid out as Word 2010 or older', () => {
    expect(noteColumnsOf(sec(2), 14)).toBe(1)
    expect(noteColumnsOf(sec(2), 12)).toBe(1)
    expect(noteColumnsOf(sec(2), undefined)).toBe(1)
    expect(noteColumnLayout(sec(2), 14)).toBeNull()
  })

  it('has no layout for one column: the area stays as it was', () => {
    expect(noteColumnLayout(sec(1), 15)).toBeNull()
    expect(noteColumnLayout(sec(2, 1), 15)).toBeNull()
    expect(noteColumnWidth(624, null)).toBe(624)
  })

  it('puts half an inch between stated columns and the section gap between matched ones', () => {
    expect(noteColumnLayout(sec(1, 2), 15)).toEqual({ count: 2, gapPx: 48 })
    expect(noteColumnLayout(sec(2, undefined, 360), 15)).toEqual({ count: 2, gapPx: 24 })
    // the line between a section's columns runs between the notes that match them
    expect(noteColumnLayout(sec(3, undefined, 720, true), 15)).toEqual({
      count: 3,
      gapPx: 48,
      rule: true,
    })
    expect(noteColumnLayout(sec(3, 2, 720, true), 15)).toEqual({ count: 2, gapPx: 48 })
    expect(noteColumnsClass({ count: 3, gapPx: 48, rule: true })).toBe(
      'note-columns note-columns-rule',
    )
    // Letter with 1in margins: 468pt of text, two columns of 216pt
    expect(noteColumnWidth(624, { count: 2, gapPx: 48 })).toBe(288)
    expect(noteColumnWidth(624, { count: 4, gapPx: 48 })).toBe(120)
  })
})

describe('lines of a note', () => {
  it('is whole line boxes when the height is a multiple of the line height', () => {
    expect(noteLines(45, 15)).toEqual([15, 15, 15])
  })

  it('gives the first line what is left over: the raised number makes it taller', () => {
    expect(noteLines(48, 15)).toEqual([18, 15, 15])
    expect(noteLines(15.3, 15.33)[0]).toBeCloseTo(15.3)
  })

  it('is one piece when the height is far from whole lines', () => {
    expect(noteLines(52, 15)).toEqual([52])
    expect(noteLines(40, undefined)).toEqual([40])
  })
})

describe('height of a footnote area in columns', () => {
  it('is the notes one under the other in one column', () => {
    expect(noteAreaHeight(notes(1, 1, 3), 1)).toBe(50)
  })

  it('fills column after column to the least height that holds every note', () => {
    // seven one-line notes: 4 + 3 in two columns, 3 + 3 + 1 in three
    expect(noteAreaHeight(notes(1, 1, 1, 1, 1, 1, 1), 2)).toBe(40)
    expect(noteAreaHeight(notes(1, 1, 1, 1, 1, 1, 1), 3)).toBe(30)
    // twelve three-line notes: 18 lines a side, 12 lines in three
    expect(noteAreaHeight(notes(...Array(12).fill(3)), 2)).toBe(180)
    expect(noteAreaHeight(notes(...Array(12).fill(4)), 3)).toBe(160)
  })

  it('does not divide a note of two or three lines', () => {
    // seven two-line notes in four columns: 2 + 2 + 2 + 1 notes
    expect(noteAreaHeight(notes(2, 2, 2, 2, 2, 2, 2), 4)).toBe(40)
    // five three-line notes in two columns: 3 + 2 notes, not 8 + 7 lines
    expect(noteAreaHeight(notes(3, 3, 3, 3, 3), 2)).toBe(90)
  })

  it('runs a long note on into the next column between two of its lines', () => {
    // eleven lines then four short notes: 8 lines, then 3 lines and the four notes
    expect(noteAreaHeight(notes(11, 1, 1, 1, 1), 2)).toBe(80)
    expect(noteAreaHeight(notes(1, 1, 1, 1, 11), 2)).toBe(80)
    // the 21 lines a page carried over: 11 + 10
    expect(noteAreaHeight(notes(21), 2)).toBe(110)
  })

  it('keeps two lines of a divided note on either side', () => {
    // four lines in two columns: 2 + 2; five lines: 3 + 2
    expect(noteAreaHeight(notes(4), 2)).toBe(20)
    expect(noteAreaHeight(notes(5), 2)).toBe(30)
    // three one-line notes and a four-line one: one line of it may not stay behind alone
    expect(noteAreaHeight(notes(1, 1, 1, 4), 2)).toBe(40)
  })

  it('is as tall as a note that cannot be divided', () => {
    expect(noteAreaHeight([[52], [10]], 2)).toBe(52)
    expect(noteAreaHeight([[10]], 3)).toBe(10)
  })

  it('reads the lines of page items', () => {
    const items = [
      { height: 30, lineHeightPx: 10 },
      { height: 30, lineHeightPx: 10 },
    ]
    expect(noteAreaPx(items, null)).toBe(60)
    expect(noteAreaPx(items, { count: 2, gapPx: 48 })).toBe(30)
    expect(noteColumnsStyle(items, { count: 2, gapPx: 48 })).toEqual({
      columnCount: '2',
      columnGap: '48px',
      height: '31px',
    })
  })
})

describe('what a note adds to its area', () => {
  it('sums to the height of the area', () => {
    for (const [list, count] of [
      [notes(1, 1, 1, 1, 1, 1, 1), 2],
      [notes(11, 1, 1, 1, 1), 2],
      [notes(3, 3, 3, 3, 3), 3],
      [notes(2, 5, 1, 7, 3, 1), 4],
    ] as const) {
      const adds = noteAreaIncrements(list, count)
      expect(adds.reduce((a, b) => a + b, 0)).toBeCloseTo(noteAreaHeight(list, count))
      expect(adds.every((h) => h >= 0)).toBe(true)
    }
  })

  it('is nothing for a note that fits beside the others', () => {
    expect(noteAreaIncrements(notes(1, 1, 1, 1), 2)).toEqual([10, 0, 10, 0])
    expect(noteAreaIncrements(notes(1, 1, 1), 1)).toEqual([10, 10, 10])
  })
})

describe('charging the notes page by page', () => {
  // the pages below are counted without the separator row
  beforeAll(() => {
    const none = { heightPx: 0, rulePx: 0 }
    setNoteSeparatorRows({ footnote: none, endnote: none })
  })
  afterAll(() => {
    const row = noteSeparatorRow(undefined)
    setNoteSeparatorRows({ footnote: row, endnote: row })
  })
  const two = (): NoteColumns | null => ({ count: 2, gapPx: 48 })
  /** one-line paragraphs 20 tall, each with a three-line note (30) */
  const noted = (count: number): BlockBox[] =>
    Array.from({ length: count }, (_, i) => ({
      top: i * 20,
      height: 20 + 30,
      docxIndex: i,
      footnoteExtraPx: 30,
      noteBands: [{ offset: 0, height: 30, lines: [10, 10, 10] }],
    }))
  const run = (
    blocks: BlockBox[],
    contentHeight = 200,
    cols: (section: number) => NoteColumns | null = two,
    separatorPx = 0,
  ) => {
    const geoms = [{ contentHeight, forceBreak: false }]
    const out: SliceOutputs = {}
    const stats: NoteColumnsStats = { rounds: 0, stop: 'settled' }
    const slice = (): PageSlice[] => {
      const s = computeSectionedSlicesF2(blocks, geoms, blocks.length * 20, out)
      out.preParity = s
      return s
    }
    const settle = noteColumnsPass(blocks, out, cols, separatorPx, contentHeight, stats)!
    const slices = settle(slice(), slice)
    return { slices, stats, out }
  }
  const bands = (blocks: BlockBox[]) => blocks.map((b) => b.noteBands![0].height)

  it('charges each note what it adds, so a page takes more lines than with the notes one under the other', () => {
    const geoms = [{ contentHeight: 200, forceBreak: false }]
    // 50 a paragraph: four to a page
    expect(computeSectionedSlicesF2(noted(12), geoms, 240).map((s) => s.start)).toEqual([
      0, 80, 160,
    ])
    const blocks = noted(12)
    const { slices, stats } = run(blocks)
    // five paragraphs (100) over an area of three notes and two (90) on every page
    expect(slices.map((s) => s.start)).toEqual([0, 100, 200])
    // the sixth was tested against page 1 with what it would add there (nothing: three notes
    // a side) and did not fit; opening page 2 its note costs 30, which the page gives up
    expect(bands(blocks)).toEqual([30, 0, 30, 0, 30, 0, 0, 30, 0, 30, 0, 0])
    expect(blocks.map((b) => b.height).slice(0, 8)).toEqual([50, 20, 50, 20, 50, 20, 20, 50])
    expect(stats).toEqual({ rounds: 0, stop: 'settled' })
  })

  it('keeps a line on the page when its note fits beside the others', () => {
    // 215 of text: five paragraphs and their area take 190, the sixth adds a line of text
    // and nothing to the area, so it stays on page 1
    const blocks = noted(8)
    const { slices } = run(blocks, 215)
    expect(slices.map((s) => s.start)).toEqual([0, 120])
    expect(bands(blocks)).toEqual([30, 0, 30, 0, 30, 0, 30, 0])
  })

  it('gives a page that opens with a note charged nothing before its separator, once', () => {
    const row = { heightPx: 16, rulePx: 11 }
    setNoteSeparatorRows({ footnote: row, endnote: row })
    onTestFinished(() => {
      const none = { heightPx: 0, rulePx: 0 }
      setNoteSeparatorRows({ footnote: none, endnote: none })
    })
    // 210 of text: five paragraphs, their area (90) and the separator fill page 1. The sixth
    // adds nothing to that area and opens page 2: 30 for its note and one separator leave
    // room for five paragraphs again (two separators would leave room for four)
    const { slices } = run(noted(12), 210, two, 16)
    expect(slices.map((s) => s.start)).toEqual([0, 100, 200])
  })

  it('is the notes one under the other where a page has one column', () => {
    const blocks = noted(8)
    const { slices } = run(blocks, 200, () => null)
    expect(slices.map((s) => s.start)).toEqual([0, 80])
    expect(bands(blocks)).toEqual(Array(8).fill(30))
  })

  it('lays a page out as the section of its first note says, whatever sections follow on it', () => {
    // sections 0 (one column) and 1 (two): four notes of each on one page
    const mixed = (firstSection: number) =>
      noted(8).map((b, i) => ({ ...b, section: i < 4 ? firstSection : 1 - firstSection }))
    const bySection = (s: number): NoteColumns | null => (s === 1 ? two() : null)
    const plainFirst = mixed(0)
    run(plainFirst, 1000, bySection)
    expect(bands(plainFirst)).toEqual(Array(8).fill(30))
    const columnsFirst = mixed(1)
    run(columnsFirst, 1000, bySection)
    expect(bands(columnsFirst)).toEqual([30, 0, 30, 0, 30, 0, 30, 0])
  })

  it('gives the same pages when the blocks are sliced again', () => {
    const blocks = noted(12)
    const first = run(blocks)
    const charged = bands(blocks)
    const again = run(blocks)
    expect(again.slices.map((s) => s.start)).toEqual(first.slices.map((s) => s.start))
    expect(bands(blocks)).toEqual(charged)
  })

  it('has nothing to do for blocks without note lines', () => {
    const blocks: BlockBox[] = [
      { top: 0, height: 80, footnoteExtraPx: 30, noteBands: [{ offset: 0, height: 30 }] },
    ]
    const out: SliceOutputs = {}
    expect(noteColumnsPass(blocks, out, two, 0, 200)).toBeNull()
    expect(out.notePage).toBeUndefined()
  })

  it('keeps a table row charged with its note', () => {
    const blocks = noted(2)
    blocks[1].noteBands![0].row = 0
    blocks[1].tableRows = [{ height: 20, notesPx: 30 }]
    run(blocks)
    expect(blocks[1].tableRows[0].notesPx).toBe(0)
  })

  it('charges the notes of a paragraph cut by the page on the page their lines land on', () => {
    // a filler of 140, then four lines of 20 with a note on the first and on the last
    const filler: BlockBox = { top: 0, height: 140 }
    const para: BlockBox = {
      top: 140,
      height: 80 + 60,
      docxIndex: 1,
      lineBoxes: [0, 20, 40, 60].map((off) => ({ offsetInBlock: off, height: 20 })),
      footnoteExtraPx: 60,
      noteBands: [
        { offset: 5, height: 30, lines: [10, 10, 10] },
        { offset: 65, height: 30, lines: [10, 10, 10] },
      ],
    }
    const tail = noted(1)
    tail[0].top = 220
    const blocks = [filler, para, ...tail]
    const out: SliceOutputs = {}
    const settle = noteColumnsPass(blocks, out, two, 0, 200)!
    const slices = settle(
      computeSectionedSlicesF2(blocks, [{ contentHeight: 200, forceBreak: false }], 240, out),
      () => [],
    )
    // two lines and the first note stay (140 + 40 + 30 > 200 leaves room for one line only,
    // which widow control does not allow): the paragraph opens page 2 whole
    expect(slices.map((s) => s.start)).toEqual([0, 140])
    // on page 2 its two notes stand side by side (30); the tail's makes it two and one (60)
    expect(tail[0].noteBands![0].height).toBe(30)
  })

  it('reserves in each text column what the notes of the other columns add to the area', () => {
    const blocks = noted(2)
    blocks[1].top = 100
    const out: SliceOutputs = {}
    const stats: NoteColumnsStats = { rounds: 0, stop: 'settled' }
    const settle = noteColumnsPass(blocks, out, two, 16, 200, stats)!
    out.notePage!(blocks[0], 0, 0, true)
    // one note a column, side by side: the area is 30 tall and the second adds nothing
    expect(bands(blocks)).toEqual([30, 0])
    const page = (): PageSlice[] => [
      {
        start: 0,
        end: 200,
        section: 0,
        regions: [
          {
            top: 0,
            height: 100,
            section: 0,
            columns: [
              { start: 0, end: 100 },
              { start: 100, end: 200 },
            ],
          },
        ],
      },
    ]
    settle(page(), page)
    // the second text column gives up the 30 of the first column's note
    expect([...out.noteReserve!]).toEqual([[noteReserveKey(0, 1), 30]])
    expect(stats).toEqual({ rounds: 1, stop: 'settled' })
    expect(TEXT_COLUMN_ROUNDS).toBeGreaterThan(1)
  })

  it('the slicer takes the reserve off the text column it was made for', () => {
    // two columns of 100; the first holds five 20s, the second starts at 100
    const geoms = [{ contentHeight: 100, forceBreak: false, cols: 2 }]
    const blocks: BlockBox[] = Array.from({ length: 10 }, (_, i) => ({ top: i * 20, height: 20 }))
    expect(computeSectionedSlicesF2(blocks, geoms, 200, {})).toHaveLength(1)
    const reserved = computeSectionedSlicesF2(blocks, geoms, 200, {
      noteReserve: new Map([[noteReserveKey(0, 1), 40]]),
    })
    // the second column holds three instead of five: the last two open page 2
    expect(reserved.map((s) => s.start)).toEqual([0, 160])
  })

  it('the slicer takes what a page opens with off every column of that page', () => {
    const geoms = [{ contentHeight: 100, forceBreak: false }]
    const blocks: BlockBox[] = Array.from({ length: 12 }, (_, i) => ({ top: i * 20, height: 20 }))
    const opened: Array<[number, boolean]> = []
    const slices = computeSectionedSlicesF2(blocks, geoms, 240, {
      notePage: (block, y, _section, first) => {
        opened.push([blocks.indexOf(block!), first])
        return { shiftPx: y === 100 ? 40 : 0, withNotes: false }
      },
    })
    // page 2 holds three lines instead of five
    expect(slices.map((s) => s.start)).toEqual([0, 100, 160])
    expect(opened).toEqual([
      [0, true],
      [5, false],
      [8, false],
    ])
  })
})
