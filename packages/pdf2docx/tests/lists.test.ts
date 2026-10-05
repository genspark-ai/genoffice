import { describe, expect, it } from 'vitest'
import {
  analyzeChars,
  analyzePage,
  detectListBlocks,
  groupIntoBlocks,
  parseListMarker,
} from '../src/analyze'
import type { ExtractedPage } from '../src/extract'
import type { IrPage, TextBlock } from '../src/ir'
import { pagesToSaveBlocks } from '../src/rebuild'
import { mkText } from './helpers/chars'
import type { PdfChar } from '../src/ir'

/** lay out lines of text (top → bottom, 14pt pitch) and run the block pipeline */
function blocksOf(lines: Array<{ text: string; x: number }>, seq = { next: 0 }): TextBlock[] {
  const chars: PdfChar[] = []
  lines.forEach((l, i) => {
    chars.push(...mkText(l.text, l.x, { y: 700 - i * 14, fontSize: 10 }).chars)
  })
  const body = { bodyLeft: 72, bodyRight: 540 }
  return detectListBlocks(groupIntoBlocks(analyzeChars(chars), body), seq)
}

const textOf = (b: TextBlock): string =>
  b.lines.map((l) => l.spans.map((s) => s.text).join('')).join(' ')

describe('parseListMarker', () => {
  it('recognizes bullets, numbers and parenthesized numbers', () => {
    const line = (text: string) => analyzeChars(mkText(text, 72, { fontSize: 10 }).chars)[0]!
    expect(parseListMarker(line('• item'))?.kind).toBe('bullet')
    expect(parseListMarker(line('3. item'))).toMatchObject({
      kind: 'ordered',
      value: 3,
      style: 'dot',
    })
    expect(parseListMarker(line('7) item'))).toMatchObject({
      kind: 'ordered',
      value: 7,
      style: 'paren',
    })
    expect(parseListMarker(line('(2) item'))).toMatchObject({
      kind: 'ordered',
      value: 2,
      style: 'parens',
    })
    expect(parseListMarker(line('– item'))).toMatchObject({ kind: 'bullet', weak: true })
    expect(parseListMarker(line('plain text'))).toBeNull()
    expect(parseListMarker(line('3.14 is pi'))).toBeNull()
  })

  it('recognizes multi-level outline markers (trailing dot required)', () => {
    const line = (text: string) => analyzeChars(mkText(text, 72, { fontSize: 10 }).chars)[0]!
    expect(parseListMarker(line('3.1.15. Authorize remote'))).toMatchObject({
      kind: 'ordered',
      style: 'multi',
      values: [3, 1, 15],
    })
    expect(parseListMarker(line('2.4. Section title'))).toMatchObject({
      style: 'multi',
      values: [2, 4],
    })
  })
})

describe('detectListBlocks: multi-level outline numbers', () => {
  it('accepts a same-prefix run counting up by one, keyed to the marker depth', () => {
    const blocks = blocksOf([
      { text: '3.1.15. Authorize remote execution', x: 72 },
      { text: '3.1.16. Authorize wireless access', x: 72 },
      { text: '3.1.17. Protect wireless access', x: 72 },
    ])
    expect(blocks).toHaveLength(3)
    for (const b of blocks) {
      expect(b.list).toMatchObject({
        kind: 'ordered',
        style: 'multi',
        level: 2,
        startValues: [3, 1, 15],
      })
    }
    expect(textOf(blocks[0]!)).toBe('Authorize remote execution')
  })

  it('rejects a lone outline number and prefix jumps', () => {
    const lone = blocksOf([{ text: '3.1.15. Only one item', x: 72 }])
    expect(lone.every((b) => !b.list)).toBe(true)
    const jump = blocksOf([
      { text: '3.1.15. First', x: 72 },
      { text: '3.2.16. Prefix changed AND not +1', x: 72 },
    ])
    expect(jump.every((b) => !b.list)).toBe(true)
  })
})

describe('detectListBlocks: bullets', () => {
  it('turns sibling bullet items into list blocks with the marker stripped', () => {
    const blocks = blocksOf([
      { text: '• First item text', x: 72 },
      { text: '• Second item text', x: 72 },
    ])
    expect(blocks).toHaveLength(2)
    for (const b of blocks) {
      expect(b.list).toMatchObject({ kind: 'bullet', level: 0 })
    }
    expect(textOf(blocks[0]!)).toBe('First item text')
    expect(textOf(blocks[1]!)).toBe('Second item text')
  })

  it('accepts a single bullet with a hanging continuation line', () => {
    const blocks = blocksOf([
      { text: '• A rather long item that', x: 72 },
      { text: 'wraps onto a second line', x: 82 },
    ])
    const item = blocks.find((b) => b.list)
    expect(item).toBeDefined()
    expect(item!.lines).toHaveLength(2)
  })

  it('does NOT treat dash-opened dialogue as a list', () => {
    const blocks = blocksOf([
      { text: '– Hello there, said one.', x: 72 },
      { text: '– Goodbye now, said two.', x: 72 },
    ])
    expect(blocks.every((b) => b.list === undefined)).toBe(true)
    // the dashes stay in the text
    expect(blocks.map(textOf).join(' ')).toContain('–')
  })

  it('accepts dash items when they show hanging-indent structure', () => {
    const blocks = blocksOf([
      { text: '– first entry of the agenda', x: 72 },
      { text: 'continued on a second line', x: 82 },
      { text: '– second entry', x: 72 },
    ])
    expect(blocks.filter((b) => b.list?.kind === 'bullet')).toHaveLength(2)
  })

  it('accepts single-line dash sub-bullets indented past the body text (P20)', () => {
    // slide layout: unmarked parent lines at the margin, dash sub-bullets
    // indented 2.4em, block gaps so parents and sub-bullets stay separate
    const chars: PdfChar[] = [
      ...mkText('Use it to define your venture', 72, { y: 700, fontSize: 10 }).chars,
      ...mkText('– To answer questions early', 96, { y: 672, fontSize: 10 }).chars,
      ...mkText('– In the order they ask them', 96, { y: 658, fontSize: 10 }).chars,
      ...mkText('Goal is to communicate fast', 72, { y: 630, fontSize: 10 }).chars,
    ]
    const body = { bodyLeft: 72, bodyRight: 540 }
    const blocks = detectListBlocks(groupIntoBlocks(analyzeChars(chars), body), { next: 0 })
    const items = blocks.filter((b) => b.list?.kind === 'bullet')
    expect(items).toHaveLength(2)
    expect(textOf(items[0]!)).toBe('To answer questions early')
    expect(textOf(items[1]!)).toBe('In the order they ask them')
  })

  it('dash dialogue at the prose margin stays plain even among plain text (P20)', () => {
    const blocks = blocksOf([
      { text: 'The travellers met at dawn today.', x: 72 },
      { text: '– Hello there, said the first.', x: 72 },
      { text: '– Goodbye now, said the second.', x: 72 },
      { text: 'And so they parted separate ways.', x: 72 },
    ])
    expect(blocks.every((b) => b.list === undefined)).toBe(true)
  })

  it('nests by marker indent', () => {
    const blocks = blocksOf([
      { text: '• outer one', x: 72 },
      { text: '• inner one', x: 90 },
      { text: '• inner two', x: 90 },
      { text: '• outer two', x: 72 },
    ])
    const levels = blocks.map((b) => b.list?.level)
    expect(levels).toEqual([0, 1, 1, 0])
  })
})

describe('detectListBlocks: ordered lists', () => {
  it('accepts increasing sequences and records the start ordinal', () => {
    const blocks = blocksOf([
      { text: '3. gamma comes first here', x: 72 },
      { text: '4. delta follows', x: 72 },
      { text: '5. epsilon closes', x: 72 },
    ])
    expect(blocks).toHaveLength(3)
    for (const b of blocks) {
      expect(b.list).toMatchObject({ kind: 'ordered', level: 0, start: 3, style: 'dot' })
    }
    expect(blocks[0]!.list!.seqId).toBe(blocks[2]!.list!.seqId)
    expect(textOf(blocks[0]!)).toBe('gamma comes first here')
  })

  it('rejects a lone numbered line (heading, not a list)', () => {
    const blocks = blocksOf([
      { text: '1. Introduction', x: 72 },
      { text: 'Body text follows the heading over here.', x: 72 },
    ])
    expect(blocks.every((b) => b.list === undefined)).toBe(true)
    expect(blocks.map(textOf).join(' ')).toContain('1. Introduction')
  })

  it('keeps flush-continuation numbered paragraphs out of lists (P18 C)', () => {
    // continuation lines return to the margin (Chinese official style) — a
    // hanging numbering indent would push them right of the source
    const blocks = blocksOf([
      { text: '5. first item body wraps to the following line and', x: 72 },
      { text: 'returns flush to the very left margin here', x: 72 },
      { text: '6. second item wraps too with a longer body text', x: 72 },
      { text: 'again flush at the left margin of the page', x: 72 },
    ])
    expect(blocks.every((b) => b.list === undefined)).toBe(true)
  })

  it('accepts multi-line ordered items with true hanging bodies', () => {
    const blocks = blocksOf([
      { text: '5. first item body wraps and the continuation is', x: 72 },
      { text: 'indented under the text, not the number', x: 87 },
      { text: '6. second item follows the same hanging shape', x: 72 },
      { text: 'with its continuation indented as well', x: 87 },
    ])
    const items = blocks.filter((b) => b.list)
    expect(items).toHaveLength(2)
  })

  it('breaks the run when numbers do not increment', () => {
    const blocks = blocksOf([
      { text: '1. one', x: 72 },
      { text: '2. two', x: 72 },
      { text: '7. seven', x: 72 },
    ])
    const items = blocks.filter((b) => b.list)
    expect(items).toHaveLength(2)
    expect(blocks.map(textOf).join(' ')).toContain('7. seven')
  })

  it('continues a run broken by a page break (single item on the next page)', () => {
    // page 1 ends the run, page 2 carries only the next ordinal: a lone item
    // is a heading unless it continues the run the shared state ended on
    const seq = { next: 0 }
    const first = blocksOf(
      [
        { text: '1. first item on the opening page', x: 72 },
        { text: '2. second item on the opening page', x: 72 },
      ],
      seq,
    )
    expect(first.map((b) => b.list?.seqId)).toEqual([0, 0])

    const second = blocksOf([{ text: '3. the only item on the next page', x: 72 }], seq)
    expect(second[0]!.list).toMatchObject({ kind: 'ordered', level: 0, start: 1, style: 'dot' })
    expect(second[0]!.list!.seqId).toBe(0)
    expect(textOf(second[0]!)).toBe('the only item on the next page')
  })

  it('still rejects a lone numbered line that continues nothing', () => {
    const seq = { next: 0 }
    blocksOf(
      [
        { text: '1. one', x: 72 },
        { text: '2. two', x: 72 },
      ],
      seq,
    )
    // a different ordinal, not the next one in the run
    const stray = blocksOf([{ text: '9. unrelated lone paragraph', x: 72 }], seq)
    expect(stray[0]!.list).toBeUndefined()
    expect(textOf(stray[0]!)).toContain('9. unrelated lone paragraph')
  })
})

describe('rebuild: list items become real docx numbering', () => {
  it('emits listItem blocks with numbering definitions', () => {
    const seq = { next: 0 }
    const blocks = blocksOf(
      [
        { text: '• bullet one', x: 72 },
        { text: '• bullet two', x: 72 },
      ],
      seq,
    )
    const ordered = blocksOf(
      [
        { text: '3. third', x: 72 },
        { text: '4. fourth', x: 72 },
      ],
      seq,
    )
    const page: IrPage = {
      index: 0,
      widthPt: 612,
      heightPt: 792,
      rotation: 0,
      blocks: [...blocks, ...ordered],
      degraded: false,
      scanned: false,
      hasStructTree: false,
    }
    const result = pagesToSaveBlocks([page])
    const generated = result.blocks.filter(
      (b): b is { kind: 'generated'; block: import('@genoffice/docx-engine').GeneratedBlock } =>
        b.kind === 'generated',
    )
    const listItems = generated.filter((g) => g.block.type === 'listItem')
    expect(listItems).toHaveLength(4)
    expect(listItems[0]!.block.list).toMatchObject({ kind: 'bullet', numId: '1', ilvl: 0 })
    const orderedItems = listItems.filter((g) => g.block.list!.kind === 'ordered')
    expect(orderedItems).toHaveLength(2)
    const numId = orderedItems[0]!.block.list!.numId
    expect(orderedItems[1]!.block.list!.numId).toBe(numId)
    expect(result.numbering?.restartNums).toEqual([
      expect.objectContaining({ numId, startOverrides: expect.objectContaining({ 0: 3 }) }),
    ])
  })
})

describe('detectListBlocks: RTL lists', () => {
  /** lay out an RTL bullet item with the bullet at a fixed right-edge position */
  function rtlBulletItem(text: string, rightEdge: number, y: number): PdfChar[] {
    const logical = `\u2022 ${text}`
    const reversed = [...logical].reverse().join('')
    const { chars } = mkText(reversed, 0, { y, fontSize: 10 })
    const width = chars[chars.length - 1]!.box.x1 - chars[0]!.box.x0
    const offset = rightEdge - width
    return chars.map((c) => ({
      ...c,
      box: { ...c.box, x0: c.box.x0 + offset, x1: c.box.x1 + offset },
      looseBox: { ...c.looseBox, x0: c.looseBox.x0 + offset, x1: c.looseBox.x1 + offset },
      originX: c.originX + offset,
    }))
  }

  it('detects Hebrew bullet items as a list with markers stripped', () => {
    const chars = [
      ...rtlBulletItem('\u05e4\u05e8\u05d9\u05d8 \u05e8\u05d0\u05e9\u05d5\u05df', 540, 700),
      ...rtlBulletItem('\u05e4\u05e8\u05d9\u05d8 \u05e9\u05e0\u05d9', 540, 672),
    ]
    const body = { bodyLeft: 72, bodyRight: 540 }
    const blocks = detectListBlocks(groupIntoBlocks(analyzeChars(chars), body), { next: 0 })
    expect(blocks).toHaveLength(2)
    for (const b of blocks) {
      expect(b.dir).toBe('rtl')
      expect(b.list).toMatchObject({ kind: 'bullet', level: 0 })
    }
    expect(textOf(blocks[0]!)).toBe('\u05e4\u05e8\u05d9\u05d8 \u05e8\u05d0\u05e9\u05d5\u05df')
    expect(textOf(blocks[1]!)).toBe('\u05e4\u05e8\u05d9\u05d8 \u05e9\u05e0\u05d9')
  })
})

describe('analyzePage: the weak-bullet indent origin on a multi-column page', () => {
  const pageOf = (chars: PdfChar[]): ExtractedPage => ({
    index: 0,
    widthPt: 612,
    heightPt: 792,
    rotation: 0,
    chars,
    images: [],
    paths: [],
    degraded: false,
    scanned: false,
    hasStructTree: false,
    vectorRegions: [],
    badUnicodeRatio: 0,
  })
  /** the page's text blocks, column by column (floats and tables dropped) */
  const columnBlocks = (page: IrPage): TextBlock[][] =>
    (page.sections ?? []).flatMap((s) =>
      s.columns.map((c) => c.blocks.filter((b): b is TextBlock => b.kind === 'text')),
    )

  it('judges each bullet column against its own left edge, not the page median', () => {
    // two identical columns of dash bullets, each flush with its own edge, so
    // neither has the plain neighbour the indent evidence is judged against.
    // They must get the SAME verdict: median([72, 320]) is 196, which is no
    // column's edge, and it read the right column's flush dashes as a
    // 124pt indent — a bulleted list minted out of column position alone
    const chars: PdfChar[] = []
    for (const [x, tag] of [
      [72, 'left column bullet'],
      [320, 'right column bullet'],
    ] as const) {
      for (let i = 0; i < 3; i++) {
        chars.push(...mkText(`– ${tag} number ${i}`, x, { y: 700 - i * 20, fontSize: 10 }).chars)
      }
    }
    const columns = columnBlocks(analyzePage(pageOf(chars)))
    expect(columns).toHaveLength(2)
    for (const blocks of columns) {
      expect(blocks.every((b) => b.list === undefined)).toBe(true)
    }
  })

  it('still lists dash sub-bullets indented on a single-column page (P20)', () => {
    // the page-level fallback the fix must not lose: a slide pins the dash
    // group in a column of its own, so only the page's left edge can say
    // it is indented
    const chars: PdfChar[] = [
      ...mkText('Use it to define your venture', 60, { y: 730, fontSize: 10 }).chars,
      ...mkText('– To answer questions early', 110, { y: 700, fontSize: 10 }).chars,
      ...mkText('– In the order they ask them', 110, { y: 686, fontSize: 10 }).chars,
    ]
    const listed = columnBlocks(analyzePage(pageOf(chars)))
      .flat()
      .filter((b) => b.list?.kind === 'bullet')
    expect(listed).toHaveLength(2)
  })
})
