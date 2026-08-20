/** Rebuild layer round-trip: IR → docx bytes → docx-engine parseDocx (no wasm needed). */
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { encodeRgbaPng } from '../src/extract'
import type { ImageBlock, IrPage, Line, PageSection, Span, TableBlock, TextBlock } from '../src/ir'
import { rebuildDocx, bytesToBase64 } from '../src/rebuild'

function span(text: string, over: Partial<Span> = {}): Span {
  return {
    text,
    box: { x0: 72, y0: 690, x1: 300, y1: 700 },
    fontSize: 12,
    fontFamily: 'Helvetica',
    bold: false,
    italic: false,
    color: '000000',
    dir: 'ltr',
    script: 'latin',
    ...over,
  }
}

function line(spans: Span[], top = 700, endsWithHyphen = false): Line {
  const x0 = Math.min(...spans.map((s) => s.box.x0))
  const x1 = Math.max(...spans.map((s) => s.box.x1))
  return {
    spans,
    box: { x0, x1, y0: top - 12, y1: top },
    baseline: top - 10,
    endsWithHyphen,
  }
}

function textBlock(lines: Line[], over: Partial<TextBlock> = {}): TextBlock {
  return {
    kind: 'text',
    lines,
    box: {
      x0: Math.min(...lines.map((l) => l.box.x0)),
      y0: Math.min(...lines.map((l) => l.box.y0)),
      x1: Math.max(...lines.map((l) => l.box.x1)),
      y1: Math.max(...lines.map((l) => l.box.y1)),
    },
    align: 'left',
    firstLineIndentPt: 0,
    dir: 'ltr',
    ...over,
  }
}

function page(blocks: IrPage['blocks'], over: Partial<IrPage> = {}): IrPage {
  return {
    index: 0,
    widthPt: 612,
    heightPt: 792,
    rotation: 0,
    blocks,
    degraded: false,
    scanned: false,
    hasStructTree: false,
    ...over,
  }
}

const paraTexts = (parsed: Awaited<ReturnType<typeof parseDocx>>) =>
  parsed.blocks
    .filter((b) => !b.hidden && b.type !== 'image')
    .map((b) => (b.runs ?? []).map((r) => r.text).join(''))

describe('rebuildDocx', () => {
  it('emits paragraphs with runs, styles and CJK eastAsia font slot', async () => {
    const zh = span('中文内容', {
      script: 'cjk',
      fontFamily: 'SimSun',
      box: { x0: 72, y0: 690, x1: 120, y1: 700 },
    })
    const en = span('English', {
      fontFamily: 'Times',
      bold: true,
      box: { x0: 120, y0: 690, x1: 200, y1: 700 },
    })
    const docx = await rebuildDocx([page([textBlock([line([zh, en])])])])
    const parsed = await parseDocx(docx)

    expect(paraTexts(parsed)).toEqual(['中文内容English'])
    const runs = parsed.blocks.find((b) => b.runs?.length)!.runs!
    expect(runs).toHaveLength(2)
    // CJK run fills the eastAsia slot (docx-engine primary font), Latin run the ascii slot
    expect(runs[0]!.font).toBe('SimSun')
    // 'Times' resolves to itself when installed (macOS) or its first installed
    // metric-compatible stand-in otherwise (Liberation Serif on CI Linux, etc.),
    // so accept the whole Times metric family — the run just needs the ascii slot.
    expect([
      'Times New Roman',
      'Liberation Serif',
      'Tinos',
      'Nimbus Roman',
      'Nimbus Roman No9 L',
      'Times',
      'FreeSerif',
      'DejaVu Serif',
    ]).toContain(runs[1]!.fontAscii)
    expect(runs[1]!.bold).toBe(true)
    expect(runs[0]!.sizeHalfPoints).toBe(24)
  })

  it('joins latin lines with a space and CJK lines without one', async () => {
    const latin = textBlock([line([span('first line')], 700), line([span('second line')], 688)])
    const cjk = textBlock([
      line([span('第一行', { script: 'cjk' })], 650),
      line([span('第二行', { script: 'cjk' })], 638),
    ])
    const docx = await rebuildDocx([page([latin, cjk])])
    const parsed = await parseDocx(docx)
    expect(paraTexts(parsed)).toEqual(['first line second line', '第一行第二行'])
  })

  it('rebuilds a marked hard break as w:br instead of a soft join (P7)', async () => {
    const first = line([span('T E C H N I C A L')], 700)
    const second: Line = { ...line([span('R E V I E W')], 688), hardBreakBefore: true }
    const docx = await rebuildDocx([page([textBlock([first, second])])])
    const parsed = await parseDocx(docx)
    // parseDocx maps <w:br/> back to '\n' in run text
    expect(paraTexts(parsed)).toEqual(['T E C H N I C A L\nR E V I E W'])
  })

  it('rebuilds a decorative rule as a paragraph border with a right inset (P7)', async () => {
    const b = textBlock([line([span('T E C H N I C A L')], 782)], {
      border: { side: 'top', color: '0A3C61', widthPt: 1.5, spacePt: 12, indentRightPt: 341 },
    })
    const docx = await rebuildDocx([page([b])])
    const parsed = await parseDocx(docx)
    const blk = parsed.blocks.find((x) => x.runs?.length)!
    expect(blk.format?.borders).toBe('t')
    expect(blk.format?.indentRight).toBe(341 * 20)
  })

  it('funds the border gap + line width (+1pt slack) out of the measured spacing-before', async () => {
    // 30pt measured spacing kept whole (roomy page: floor surplus not
    // deducted, P18 B); 12pt gap, 1.5pt rule → w:space=12, spaceBefore = 30 − 12 −
    // 1.5 − 1 (slack) = 15.5pt; the paragraph's net height is unchanged
    // (LibreOffice adds border space + width to the box)
    const b = textBlock([line([span('ruled heading')], 700)], {
      spacingBeforePt: 30,
      border: { side: 'top', color: '0A3C61', widthPt: 1.5, spacePt: 12 },
    })
    const docx = await rebuildDocx([page([b, textBlock([line([span('body')], 650)])])])
    const parsed = await parseDocx(docx)
    const blk = parsed.blocks.find((x) => x.format?.borders)!
    expect(blk.format?.spaceBefore).toBe(Math.round(15.5 * 20))
  })

  it('hugs a top rule to its text on a page-break paragraph (P18)', async () => {
    // LibreOffice honours w:space above a w:pageBreakBefore paragraph (there
    // is nothing above to fund it), shifting the whole page down — the gap
    // is dropped instead
    const p1 = textBlock([line([span('page one')], 700)])
    const ruled = textBlock([line([span('page two heading')], 700)], {
      border: { side: 'top', color: '000000', widthPt: 0.5, spacePt: 13 },
    })
    const docx = await rebuildDocx([page([p1]), page([ruled], { index: 1 })])
    const parsed = await parseDocx(docx)
    const xml = parsed.internal.documentXml
    const ruledPara = xml.slice(
      xml.indexOf('<w:pageBreakBefore/>'),
      xml.indexOf('page two heading'),
    )
    expect(ruledPara).toContain('w:space="0"')
    expect(ruledPara).not.toContain('w:space="13"')
  })

  it('clamps the border gap to what the spacing-before can fund', async () => {
    const b = textBlock([line([span('tight heading')], 700)], {
      spacingBeforePt: 5,
      border: { side: 'top', color: '000000', widthPt: 1, spacePt: 20 },
    })
    const docx = await rebuildDocx([page([b, textBlock([line([span('body')], 650)])])])
    const parsed = await parseDocx(docx)
    const blk = parsed.blocks.find((x) => x.format?.borders)!
    // gap clamped to 5 − 1 = 4pt, spacing fully spent
    expect(blk.format?.spaceBefore).toBeUndefined()
  })

  it('a standalone bar block (no lines) renders as an empty bordered paragraph', async () => {
    const bar = textBlock([line([span('placeholder')], 700)], {
      border: { side: 'top', color: 'FF0000', widthPt: 1, spacePt: 0, indentLeftPt: 100 },
    })
    bar.lines = []
    bar.box = { x0: 154, y0: 500, x1: 430, y1: 501.5 }
    const docx = await rebuildDocx([page([bar])])
    const parsed = await parseDocx(docx)
    const blk = parsed.blocks.find((x) => x.format?.borders)!
    expect(blk.format?.borders).toBe('t')
    expect(blk.format?.indentLeft).toBe(100 * 20)
    expect((blk.runs ?? []).map((r) => r.text).join('')).toBe('')
  })

  it('dehyphenates across line joins when the line ends in a hyphenation hyphen', async () => {
    const block = textBlock([line([span('convert-')], 700, true), line([span('ible')], 688)])
    const docx = await rebuildDocx([page([block])])
    const parsed = await parseDocx(docx)
    expect(paraTexts(parsed)).toEqual(['convertible'])
  })

  it('writes alignment and first-line indent to paragraph format', async () => {
    const centered = textBlock([line([span('Title')])], { align: 'center' })
    const indented = textBlock([line([span('body')], 650)], { firstLineIndentPt: 21 })
    const docx = await rebuildDocx([page([centered, indented])])
    const parsed = await parseDocx(docx)
    const visible = parsed.blocks.filter((b) => !b.hidden)
    expect(visible[0]!.format?.align).toBe('center')
    expect(visible[1]!.format?.indentFirstLine).toBe(420) // 21pt × 20
  })

  it('embeds image blocks as pictures', async () => {
    const pngBytes = encodeRgbaPng(new Uint8Array(4 * 4 * 4).fill(200), 4, 4)
    const image: ImageBlock = {
      kind: 'image',
      box: { x0: 200, y0: 500, x1: 412, y1: 600 },
      data: pngBytes,
      mime: 'image/png',
      pixelWidth: 4,
      pixelHeight: 4,
    }
    const docx = await rebuildDocx([page([textBlock([line([span('before')])]), image])])
    const parsed = await parseDocx(docx)
    const imageBlocks = parsed.blocks.filter((b) => b.type === 'image')
    expect(imageBlocks).toHaveLength(1)
    expect(imageBlocks[0]!.imageDataUrl).toMatch(/^data:image\/png;base64,/)
  })

  it('renders scanned pages as one full-width picture and breaks pages', async () => {
    const pngBytes = encodeRgbaPng(new Uint8Array(8 * 8 * 4).fill(255), 8, 8)
    const scanned = page([], {
      index: 1,
      scanned: true,
      render: { data: pngBytes, mime: 'image/png', pixelWidth: 8, pixelHeight: 8 },
    })
    const first = page([textBlock([line([span('text page')])])])
    const docx = await rebuildDocx([first, scanned])
    const parsed = await parseDocx(docx)
    expect(parsed.blocks.filter((b) => b.type === 'image')).toHaveLength(1)
    // the scanned page starts on its own page (break paragraph precedes it)
    const breakPara = parsed.blocks.find((b) => b.format?.pageBreakBefore)
    expect(breakPara).toBeDefined()
  })

  it('sets section page size from page 1 geometry', async () => {
    const docx = await rebuildDocx([
      page([textBlock([line([span('a4')])])], { widthPt: 595, heightPt: 842 }),
    ])
    const parsed = await parseDocx(docx)
    const { readSectionSettings } = await import('@genoffice/docx-engine')
    const section = readSectionSettings(parsed)
    expect(section.pageWidth).toBe(595 * 20)
    expect(section.pageHeight).toBe(842 * 20)
    expect(section.orientation).toBe('portrait')
  })

  it('merges adjacent same-style runs', async () => {
    const a = span('one ')
    const b = span('two')
    const docx = await rebuildDocx([page([textBlock([line([a, b])])])])
    const parsed = await parseDocx(docx)
    const runs = parsed.blocks.find((bl) => bl.runs?.length)!.runs!
    expect(runs).toHaveLength(1)
    expect(runs[0]!.text).toBe('one two')
  })
})

describe('rebuildDocx: tables (P2)', () => {
  const cellOf = (
    text: string,
    box: TableBlock['rows'][0][0]['box'],
  ): TableBlock['rows'][0][0] => ({
    box,
    gridSpan: 1,
    blocks: text ? [textBlock([line([span(text, { box: { ...box } })])])] : [],
  })

  const sampleTable = (): TableBlock => {
    const b = (x0: number, y0: number, x1: number, y1: number) => ({ x0, y0, x1, y1 })
    return {
      kind: 'table',
      box: b(72, 620, 456, 700),
      colWidthsPt: [128, 128, 128],
      rows: [
        [
          { ...cellOf('Name', b(72, 660, 200, 700)), fill: 'FFCC00' },
          cellOf('Qty', b(200, 660, 328, 700)),
          { ...cellOf('Keep', b(328, 620, 456, 700)), vMerge: 'restart' as const },
        ],
        [
          cellOf('Total', b(72, 620, 200, 660)),
          { ...cellOf('9', b(200, 620, 328, 660)), gridSpan: 1 },
          { ...cellOf('', b(328, 620, 456, 700)), vMerge: 'continue' as const },
        ],
      ],
    }
  }

  it('squeezes a set-solid table back to its measured box (P24 B)', async () => {
    // 20 single-line rows at 9pt pitch (180pt total): the single-line font
    // floor (12pt font × 1.16 ≈ 279 twips ≈ 14pt) would render the table
    // ~1.55× its measured height — the squeeze scales rows back to the box
    const b = (x0: number, y0: number, x1: number, y1: number) => ({ x0, y0, x1, y1 })
    const rows: TableBlock['rows'] = []
    for (let i = 0; i < 20; i++) {
      const y1 = 700 - i * 9
      const y0 = y1 - 5.5 // ink height below the pitch (set solid)
      rows.push([
        {
          box: b(72, y0, 200, y1),
          gridSpan: 1,
          blocks: [textBlock([line([span(`r${i}`, { box: b(72, y0, 200, y1) })], y1)])],
        },
        { box: b(200, y0, 328, y1), gridSpan: 1, blocks: [] },
      ])
    }
    const table: TableBlock = {
      kind: 'table',
      box: b(72, 700 - 20 * 9, 328, 700),
      colWidthsPt: [128, 128],
      rows,
    }
    const parsed = await parseDocx(await rebuildDocx([page([table])]))
    const model = parsed.blocks.find((bl) => bl.type === 'table')!.table!
    // every cell paragraph's exact line is scaled so the sum of the emitted
    // rows tracks the measured 180pt box (180pt = 3600 twips)
    const emitted = model.rows.reduce((sum, row) => {
      const tallest = Math.max(
        ...row.map((c) => (c.richParas ?? []).reduce((t, p) => t + (p.lineRawTwips ?? 0), 0)),
      )
      return sum + tallest
    }, 0)
    expect(emitted).toBeLessThanOrEqual(3600 + 60)
    // a padded table (pitch comfortably above the floor) is untouched
    const roomy = sampleTable()
    const parsed2 = await parseDocx(await rebuildDocx([page([roomy])]))
    const model2 = parsed2.blocks.find((bl) => bl.type === 'table')!.table!
    const para = model2.rows[0]![0]!.richParas![0]!
    expect(para.lineRawTwips).toBeGreaterThanOrEqual(Math.round(12 * 1.16 * 20) - 2)
  })

  it('keeps an inset table at its measured left offset (w:tblInd, P17)', async () => {
    const table = sampleTable()
    const shift = 65
    table.box = { ...table.box, x0: table.box.x0 + shift, x1: table.box.x1 + shift }
    // a margin-flush text block pins the derived page margin at x0=72
    const anchor = textBlock([line([span('heading')], 730)])
    const docx = await rebuildDocx([page([anchor, table])])
    const parsed = await parseDocx(docx)
    const model = parsed.blocks.find((b) => b.type === 'table')!.table!
    expect(model.indentTwips).toBe(65 * 20)
  })

  it('shaves a hair off page-filling tables so the bottom border cannot spill (P17)', async () => {
    const table = sampleTable()
    // stretch the sample to fill the page: 2×280pt rows = 560pt of a 792pt
    // page (usable 584pt after clamped margins + slack) → ratio 0.96
    table.box = { x0: 72, y0: 100, x1: 456, y1: 660 }
    table.rows[0]!.forEach((c) => (c.box = { ...c.box, y0: 380, y1: 660 }))
    table.rows[1]!.forEach((c) => (c.box = { ...c.box, y0: 100, y1: 380 }))
    const parsed = await parseDocx(await rebuildDocx([page([table])]))
    const model = parsed.blocks.find((b) => b.type === 'table')!.table!
    // measured 280pt rows −10 border eat −(12pt shave / 2 rows = 120 twips)
    expect(model.rowHeightsTwips).toEqual([280 * 20 - 10 - 120, 280 * 20 - 10 - 120])
  })

  it('emits a real docx table with fills, spans and vMerge', async () => {
    const docx = await rebuildDocx([page([sampleTable()])])
    const parsed = await parseDocx(docx)
    const tableBlock = parsed.blocks.find((b) => b.type === 'table')!
    expect(tableBlock).toBeDefined()
    const model = tableBlock.table!
    expect(model.rows).toHaveLength(2)
    expect(model.rows[0]!.map((c) => c.paras[0])).toEqual(['Name', 'Qty', 'Keep'])
    expect(model.rows[0]![0]!.fill).toBe('FFCC00')
    expect(model.rows[0]![2]!.vMerge).toBe('restart')
    expect(model.rows[1]![2]!.vMerge).toBe('continue')
    expect(model.colWidthsTwips).toEqual([2560, 2560, 2560])
  })

  it('a non-black lattice grid keeps its ruling color (w:tblBorders, P20)', async () => {
    const table = sampleTable()
    table.borderColor = 'FFFFFF'
    const parsed = await parseDocx(await rebuildDocx([page([table])]))
    const model = parsed.blocks.find((b) => b.type === 'table')!.table!
    expect(model.borders?.top?.color).toBe('FFFFFF')
    expect(model.borders?.insideH?.color).toBe('FFFFFF')
    expect(model.borders?.insideV?.style).toBe('single')
  })

  it('a black-ruled lattice table keeps the default auto-color borders', async () => {
    const parsed = await parseDocx(await rebuildDocx([page([sampleTable()])]))
    const model = parsed.blocks.find((b) => b.type === 'table')!.table!
    expect(model.borders?.top?.color ?? 'auto').toBe('auto')
    expect(model.borders?.insideH?.color ?? 'auto').toBe('auto')
  })

  it('gridSpan cells survive the round trip', async () => {
    const table = sampleTable()
    table.rows[1] = [
      { ...cellOf('merged wide', { x0: 72, y0: 620, x1: 328, y1: 660 }), gridSpan: 2 },
      { ...cellOf('', { x0: 328, y0: 620, x1: 456, y1: 700 }), vMerge: 'continue' as const },
    ]
    const docx = await rebuildDocx([page([table])])
    const parsed = await parseDocx(docx)
    const model = parsed.blocks.find((b) => b.type === 'table')!.table!
    expect(model.rows[1]![0]!.colSpan).toBe(2)
    expect(model.rows[1]![0]!.paras[0]).toBe('merged wide')
  })

  it('RTL text inside cells gets bidi paragraphs and rtl runs', async () => {
    const table = sampleTable()
    const he = span('שלום', { script: 'hebrew', dir: 'rtl', fontFamily: '' })
    table.rows[0]![1]!.blocks = [textBlock([line([he])], { dir: 'rtl', align: 'right' })]
    const docx = await rebuildDocx([page([table])])
    const parsed = await parseDocx(docx)
    const model = parsed.blocks.find((b) => b.type === 'table')!.table!
    const para = model.rows[0]![1]!.richParas![0]!
    expect(para.bidi).toBe(true)
    expect(para.runs[0]!.rtl).toBe(true)
    expect(para.runs[0]!.fontCs).toBe('David')
  })
})

describe('rebuildDocx: RTL paragraphs (P2)', () => {
  it('writes w:bidi + run-level w:rtl + cs font slot', async () => {
    const ar = span('مرحبا بالعالم', {
      script: 'arabic',
      dir: 'rtl',
      fontFamily: 'Amiri',
      box: { x0: 300, y0: 690, x1: 540, y1: 700 },
    })
    const block = textBlock([line([ar])], { dir: 'rtl', align: 'right' })
    const docx = await rebuildDocx([page([block])])
    const parsed = await parseDocx(docx)
    const para = parsed.blocks.find((b) => (b.runs ?? []).length > 0)!
    expect(para.format?.bidi).toBe(true)
    // visual right = the bidi default → no explicit align stored
    expect(para.format?.align).toBeUndefined()
    const run = para.runs![0]!
    expect(run.rtl).toBe(true)
    expect(run.fontCs).toBe('Amiri')
    expect(run.fontAscii).toBe('Amiri')
  })

  it('falls back to Traditional Arabic when the PDF has no usable font name', async () => {
    const ar = span('سلام', { script: 'arabic', dir: 'rtl', fontFamily: '' })
    const docx = await rebuildDocx([page([textBlock([line([ar])], { dir: 'rtl' })])])
    const parsed = await parseDocx(docx)
    const run = parsed.blocks.find((b) => (b.runs ?? []).length > 0)!.runs![0]!
    expect(run.fontCs).toBe('Traditional Arabic')
  })

  it('keeps visual center alignment and drops the left-measured indent for RTL', async () => {
    const he = span('כותרת', { script: 'hebrew', dir: 'rtl' })
    const block = textBlock([line([he])], { dir: 'rtl', align: 'center', firstLineIndentPt: 20 })
    const docx = await rebuildDocx([page([block])])
    const parsed = await parseDocx(docx)
    const para = parsed.blocks.find((b) => (b.runs ?? []).length > 0)!
    expect(para.format?.align).toBe('center')
    expect(para.format?.indentFirstLine).toBeUndefined()
  })
})

describe('rebuildDocx: sections, spacing, floats (P3)', () => {
  const box = (x0: number, y0: number, x1: number, y1: number) => ({ x0, y0, x1, y1 })

  const columnBlock = (text: string, x0: number, x1: number, top: number): TextBlock =>
    textBlock([line([span(text, { box: box(x0, top - 12, x1, top) })], top)], {
      box: box(x0, top - 12, x1, top),
    })

  it('emits a continuous section break + w:cols for a title-above-columns page', async () => {
    const title = columnBlock('Title', 200, 400, 740)
    const l1 = columnBlock('left one', 72, 292, 700)
    const l2 = columnBlock('left two', 72, 292, 660)
    const r1 = columnBlock('right one', 320, 540, 700)
    const r2 = columnBlock('right two', 320, 540, 660)
    const p = page([title, l1, l2, r1, r2], {
      sections: [
        {
          box: box(200, 728, 400, 740),
          columns: [{ box: box(200, 728, 400, 740), blocks: [title] }],
          gutterWidthsPt: [],
          dir: 'ltr',
        },
        {
          box: box(72, 648, 540, 700),
          columns: [
            { box: box(72, 648, 292, 700), blocks: [l1, l2] },
            { box: box(320, 648, 540, 700), blocks: [r1, r2] },
          ],
          gutterWidthsPt: [28],
          dir: 'ltr',
        },
      ],
    })
    const docx = await rebuildDocx([p])
    const parsed = await parseDocx(docx)
    const { readSections } = await import('@genoffice/docx-engine')
    const sections = readSections(parsed)
    expect(sections).toHaveLength(2)
    expect(sections[0]!.settings.columns).toBe(1)
    expect(sections[1]!.settings.columns).toBe(2)
    expect(sections[1]!.startType).toBe('continuous')
    // measured 28pt gutter scaled into the content width → 560 twips
    expect(sections[1]!.settings.colSpace).toBe(560)
    // reading order: title, left column, explicit column break (P10 A pins
    // the measured split against Word's column balancing), right column
    expect(paraTexts(parsed).filter((t) => t.trim().length > 0)).toEqual([
      'Title',
      'left one',
      'left two',
      'right one',
      'right two',
    ])
    expect(parsed.internal.documentXml).toContain('<w:br w:type="column"/>')
  })

  it('unequal columns produce explicit w:col children', async () => {
    const a = columnBlock('narrow a', 72, 172, 700)
    const b = columnBlock('narrow b', 72, 172, 660)
    const c = columnBlock('wide c', 220, 540, 700)
    const d = columnBlock('wide d', 220, 540, 660)
    const p = page([a, b, c, d], {
      sections: [
        {
          box: box(72, 648, 540, 700),
          columns: [
            { box: box(72, 648, 172, 700), blocks: [a, b] },
            { box: box(220, 648, 540, 700), blocks: [c, d] },
          ],
          gutterWidthsPt: [48],
          dir: 'ltr',
        },
      ],
    })
    const parsed = await parseDocx(await rebuildDocx([p]))
    const { readSectionSettings } = await import('@genoffice/docx-engine')
    const settings = readSectionSettings(parsed)
    expect(settings.columns).toBe(2)
    expect(settings.colWidths).toHaveLength(2)
    expect(settings.colWidths![1]!).toBeGreaterThan(2 * settings.colWidths![0]!)
  })

  it('a TOC entry in the RIGHT column measures from the column base (P23)', async () => {
    // the old emitter measured indent from the page margin and the leader tab
    // to the full content width: a right-column checklist entry got its column
    // x double-counted as indent and dots running past the column, so LO
    // wrapped it word-by-word (childAttachments src4)
    const l1 = columnBlock('left one', 72, 292, 700)
    const l2 = columnBlock('left two', 72, 292, 660)
    const r1 = columnBlock('right head', 320, 540, 700)
    const entry = textBlock(
      [line([span('File Form 940', { box: box(320, 648, 420, 660) })], 660)],
      {
        box: box(320, 648, 540, 660),
        tocEntry: { level: 1, pageNumber: '8' },
      },
    )
    const p = page([l1, l2, r1, entry], {
      sections: [
        {
          box: box(72, 648, 540, 700),
          columns: [
            { box: box(72, 648, 292, 700), blocks: [l1, l2] },
            { box: box(320, 648, 540, 700), blocks: [r1, entry] },
          ],
          gutterWidthsPt: [28],
          dir: 'ltr',
        },
      ],
    })
    const parsed = await parseDocx(await rebuildDocx([p]))
    const tocPara = parsed.internal.documentXml.match(
      /<w:p><w:pPr><w:pStyle w:val="TOC1"\/>.*?<\/w:p>/,
    )![0]
    const pos = Number(tocPara.match(/w:leader="dot" w:pos="(\d+)"/)![1])
    // clamped inside the ~half-content-width column, not the full 9360-twip width
    expect(pos).toBeLessThanOrEqual(5000)
    // no double-counted column indent (block starts at the column base)
    expect(tocPara).not.toContain('<w:ind')
  })

  it('an RTL column section carries w:bidi', async () => {
    const r = columnBlock('يمين', 320, 540, 700)
    const r2 = columnBlock('يمين اثنان', 320, 540, 660)
    const l = columnBlock('يسار', 72, 292, 700)
    const l2 = columnBlock('يسار اثنان', 72, 292, 660)
    const p = page([r, r2, l, l2], {
      sections: [
        {
          box: box(72, 648, 540, 700),
          // reading order: right column first
          columns: [
            { box: box(320, 648, 540, 700), blocks: [r, r2] },
            { box: box(72, 648, 292, 700), blocks: [l, l2] },
          ],
          gutterWidthsPt: [28],
          dir: 'rtl',
        },
      ],
    })
    const parsed = await parseDocx(await rebuildDocx([p]))
    const { readSectionSettings } = await import('@genoffice/docx-engine')
    expect(readSectionSettings(parsed).bidi).toBe(true)
    expect(readSectionSettings(parsed).columns).toBe(2)
  })

  it('writes spacingBeforePt as w:spacing w:before (twips)', async () => {
    const a = textBlock([line([span('first')], 700)])
    const b = textBlock([line([span('second')], 600)], { spacingBeforePt: 30 })
    const parsed = await parseDocx(await rebuildDocx([page([a, b])]))
    const second = parsed.blocks.find((bl) => (bl.runs ?? []).some((r) => r.text === 'second'))!
    // 30pt measured gap kept verbatim: the page's slack dwarfs the floor
    // surpluses, so none is deducted (P18 B)
    expect(second.format?.spaceBefore).toBe(30 * 20)
  })

  it('keeps the floor surplus on a roomy page (P17 deduction gated by need, P18 B)', async () => {
    // dotted-leader-like line: ink box 9pt but 12pt font → exact line floors
    // to 12 × 1.16 = 13.92pt. The 4.9pt the line renders TALLER than its ink
    // re-spends the measured gap, so the emitted before shrinks by it —
    // pitch (before + exact line) then matches the source pitch.
    const a = textBlock([line([span('first')], 700)])
    const shortInk: Line = {
      spans: [span('a run of low text', { box: { x0: 72, y0: 691, x1: 300, y1: 700 } })],
      box: { x0: 72, y0: 691, x1: 300, y1: 700 },
      baseline: 693,
      endsWithHyphen: false,
    }
    const b = textBlock([shortInk], { spacingBeforePt: 20 })
    const parsed = await parseDocx(await rebuildDocx([page([a, b])]))
    const second = parsed.blocks.find((bl) =>
      (bl.runs ?? []).some((r) => r.text.includes('low text')),
    )!
    // line = 278 twips (13.9pt), ink 9pt → surplus 4.9pt, but the roomy
    // page keeps its measured 20pt gap (deduction gated by need, P18 B)
    expect(second.format?.spaceBefore).toBe(20 * 20)
  })

  it('rebuilds a full-line dot ruling as a wrap-proof tab leader (P17)', async () => {
    // answer rulings ('.....' / '………') re-render a hair wider under the
    // substituted font and wrap a lone dot onto a second line — every ruling
    // then grows the page one pitch. A right-tab with a leader ends exactly
    // at the measured x1 and cannot wrap.
    const dots = span('.'.repeat(60), { box: { x0: 72, y0: 691, x1: 520, y1: 700 } })
    const ruling = textBlock([
      { spans: [dots], box: dots.box, baseline: 693, endsWithHyphen: false },
    ])
    const docx = await rebuildDocx([page([textBlock([line([span('Q1')], 720)]), ruling])])
    const { default: JSZip } = await import('jszip')
    const zip = await JSZip.loadAsync(docx)
    const xml = await zip.file('word/document.xml')!.async('string')
    expect(xml).toMatch(/<w:tab w:val="right" w:leader="dot" w:pos="\d+"\/>/)
    expect(xml).not.toContain('.'.repeat(60))
    // underscore rulings keep their leader style
    const bars = span('_'.repeat(40), { box: { x0: 72, y0: 671, x1: 400, y1: 680 } })
    const barBlock = textBlock([
      { spans: [bars], box: bars.box, baseline: 673, endsWithHyphen: false },
    ])
    const docx2 = await rebuildDocx([page([textBlock([line([span('Q2')], 720)]), barBlock])])
    const zip2 = await JSZip.loadAsync(docx2)
    const xml2 = await zip2.file('word/document.xml')!.async('string')
    expect(xml2).toMatch(/w:leader="underscore"/)
  })

  it('invisible spans emit w:vanish (hidden formatting marks stay unseen, P20)', async () => {
    const hidden = span('分节符', { invisible: true, script: 'cjk' })
    const block = textBlock([
      { spans: [hidden], box: hidden.box, baseline: 693, endsWithHyphen: false },
    ])
    const docx = await rebuildDocx([page([block])])
    const { default: JSZip } = await import('jszip')
    const zip = await JSZip.loadAsync(docx)
    const xml = await zip.file('word/document.xml')!.async('string')
    expect(xml).toContain('<w:vanish/>')
    expect(xml).toContain('分节符')
  })

  it('leaves mixed text lines out of the leader conversion', async () => {
    const mixed = textBlock([line([span('Name: ________')], 650)])
    const parsed = await parseDocx(await rebuildDocx([page([mixed])]))
    expect(paraTexts(parsed)).toContain('Name: ________')
  })

  it('multi-line blocks keep their full measured spacing (pitch already trusted)', async () => {
    const a = textBlock([line([span('first')], 700)])
    const b = textBlock([line([span('one')], 600), line([span('two')], 588)], {
      spacingBeforePt: 20,
    })
    const parsed = await parseDocx(await rebuildDocx([page([a, b])]))
    const second = parsed.blocks.find((bl) => (bl.runs ?? []).some((r) => r.text === 'one two'))!
    expect(second.format?.spaceBefore).toBe(400)
  })

  it('adds page-top leading beyond the margin to the first paragraph', async () => {
    // content top gap 792−600 = 192pt → margin clamps to 108 → 84pt leading
    const a = textBlock([line([span('low title')], 600)])
    const parsed = await parseDocx(await rebuildDocx([page([a])]))
    const first = parsed.blocks.find((bl) => (bl.runs ?? []).length > 0)!
    expect(first.format?.spaceBefore).toBe(84 * 20)
  })

  it('float images become anchored pictures pinned to their page position', async () => {
    const pngBytes = encodeRgbaPng(new Uint8Array(4 * 4 * 4).fill(120), 4, 4)
    const img: ImageBlock = {
      kind: 'image',
      box: { x0: 400, y0: 600, x1: 520, y1: 700 },
      data: pngBytes,
      mime: 'image/png',
      pixelWidth: 4,
      pixelHeight: 4,
      float: { wrap: 'square-right', xOffsetPt: 328 },
    }
    const text = textBlock([line([span('wrapped text')], 700)])
    const parsed = await parseDocx(await rebuildDocx([page([text, img])]))
    const xml = parsed.internal.documentXml
    expect(xml).toContain('<wp:anchor')
    // page-relative pin at the measured box (P9 C): x = x0, y = pageH − y1
    expect(xml).toContain(
      `<wp:positionH relativeFrom="page"><wp:posOffset>${400 * 12700}</wp:posOffset></wp:positionH>`,
    )
    expect(xml).toContain(
      `<wp:positionV relativeFrom="page"><wp:posOffset>${(792 - 700) * 12700}</wp:posOffset></wp:positionV>`,
    )
    expect(xml).toContain('<wp:wrapSquare')
    expect(parsed.blocks.filter((b) => b.type === 'image')).toHaveLength(1)
  })
})

describe('rebuildDocx: explicit paragraph spacing (P5)', () => {
  it('writes measured exact line height and zero space-after on every paragraph', async () => {
    // two 12pt-ink lines at 700/686: block box 674..700 → 26pt / 2 lines =
    // 13pt. A multi-line box spans its own inter-line leading, so the
    // measured pitch is trusted as-is — no font-based floor
    const block = textBlock([line([span('first')], 700), line([span('second')], 686)])
    const parsed = await parseDocx(await rebuildDocx([page([block])]))
    const para = parsed.blocks.find((b) => (b.runs ?? []).length > 0)!
    expect(para.format?.lineRule).toBe('exact')
    expect(para.format?.lineRawTwips).toBe(260) // 13pt × 20
    expect(para.format?.spaceAfter).toBe(0)
  })

  it('floors a SINGLE line to 1.16× its tallest font size (ink understates the em)', async () => {
    // one 12pt line: ink box is 12pt (240 twips) but the glyphs need ~1.16 em
    // of exact line box or their ascenders/descenders clip (P8 B)
    const block = textBlock([line([span('lone heading')], 700)])
    const parsed = await parseDocx(await rebuildDocx([page([block])]))
    const para = parsed.blocks.find((b) => (b.runs ?? []).length > 0)!
    expect(para.format?.lineRule).toBe('exact')
    expect(para.format?.lineRawTwips).toBe(278) // 12pt × 1.16 × 20
  })

  it('keeps table-cell paragraphs off the template docDefaults too', async () => {
    const cellText = textBlock([
      line([span('cell', { box: { x0: 76, y0: 668, x1: 130, y1: 680 } })], 680),
    ])
    const table: TableBlock = {
      kind: 'table',
      box: { x0: 72, y0: 640, x1: 456, y1: 700 },
      colWidthsPt: [384],
      rows: [[{ box: { x0: 72, y0: 640, x1: 456, y1: 700 }, gridSpan: 1, blocks: [cellText] }]],
    }
    const parsed = await parseDocx(await rebuildDocx([page([table])]))
    const xml = parsed.internal.documentXml
    // the cell paragraph carries its own exact line + zero after
    expect(xml).toMatch(/<w:tc>(?:(?!<\/w:tc>).)*w:lineRule="exact"/s)
    expect(xml).toMatch(/<w:tc>(?:(?!<\/w:tc>).)*w:after="0"/s)
  })

  it('writes measured row heights as atLeast trHeight', async () => {
    const table: TableBlock = {
      kind: 'table',
      box: { x0: 72, y0: 640, x1: 456, y1: 700 },
      colWidthsPt: [384],
      rows: [
        [{ box: { x0: 72, y0: 660, x1: 456, y1: 700 }, gridSpan: 1, blocks: [] }],
        [{ box: { x0: 72, y0: 640, x1: 456, y1: 660 }, gridSpan: 1, blocks: [] }],
      ],
    }
    const parsed = await parseDocx(await rebuildDocx([page([table])]))
    const xml = parsed.internal.documentXml
    // measured heights minus the 0.5pt horizontal border LibreOffice adds
    // onto every bordered row (P16 H)
    expect(xml).toContain('<w:trHeight w:val="790" w:hRule="atLeast"/>') // 40pt row
    expect(xml).toContain('<w:trHeight w:val="390" w:hRule="atLeast"/>') // 20pt row
  })

  it('page-break carrier paragraphs are near-zero height (plus any page-top lead)', async () => {
    const pngBytes = encodeRgbaPng(new Uint8Array(4 * 4 * 4).fill(200), 4, 4)
    // image top (700) = the measured content top → no page-top lead
    const image: ImageBlock = {
      kind: 'image',
      box: { x0: 200, y0: 600, x1: 412, y1: 700 },
      data: pngBytes,
      mime: 'image/png',
      pixelWidth: 4,
      pixelHeight: 4,
    }
    const first = page([textBlock([line([span('text page')])])])
    const second = page([image], { index: 1 })
    const parsed = await parseDocx(await rebuildDocx([first, second]))
    const carrier = parsed.blocks.find((b) => b.format?.pageBreakBefore)!
    expect(carrier.format?.lineRule).toBe('exact')
    expect(carrier.format?.lineRawTwips).toBe(20)
    expect(carrier.format?.spaceAfter).toBe(0)
  })

  it('inline images carry explicit holder-paragraph spacing including spacingBefore', async () => {
    const pngBytes = encodeRgbaPng(new Uint8Array(4 * 4 * 4).fill(200), 4, 4)
    const image: ImageBlock = {
      kind: 'image',
      box: { x0: 200, y0: 500, x1: 412, y1: 600 },
      data: pngBytes,
      mime: 'image/png',
      pixelWidth: 4,
      pixelHeight: 4,
      spacingBeforePt: 30,
    }
    const text = textBlock([line([span('above')], 700)])
    const parsed = await parseDocx(await rebuildDocx([page([text, image])]))
    expect(parsed.internal.documentXml).toContain(
      '<w:spacing w:before="600" w:after="0" w:line="20" w:lineRule="atLeast"/>',
    )
  })

  it('restores character compression as w:w + w:spacing alongside model props', async () => {
    const squeezed = span('压缩文本', {
      script: 'cjk',
      fontFamily: 'Noto Sans SC',
      bold: true,
      charScale: 0.9,
      charSpacingPt: -0.4,
    })
    const parsed = await parseDocx(await rebuildDocx([page([textBlock([line([squeezed])])])]))
    const xml = parsed.internal.documentXml
    expect(xml).toContain('<w:w w:val="90"/>')
    expect(xml).toContain('<w:spacing w:val="-8"/>')
    // model-driven props still land next to the raw compression props
    const run = parsed.blocks.find((b) => (b.runs ?? []).length > 0)!.runs![0]!
    expect(run.bold).toBe(true)
    expect(run.font).toBe('Noto Sans SC')
    expect(run.sizeHalfPoints).toBe(24)
  })

  it('tables get a spacer paragraph carrying their measured spacingBefore', async () => {
    const text = textBlock([line([span('above')], 700)])
    const table: TableBlock = {
      kind: 'table',
      box: { x0: 72, y0: 560, x1: 456, y1: 640 },
      colWidthsPt: [384],
      rows: [[{ box: { x0: 72, y0: 560, x1: 456, y1: 640 }, gridSpan: 1, blocks: [] }]],
      spacingBeforePt: 48,
    }
    const parsed = await parseDocx(await rebuildDocx([page([text, table])]))
    const spacer = parsed.blocks.find(
      (b) => (b.runs ?? []).length === 0 && b.format?.lineRule === 'exact' && b.type !== 'table',
    )!
    expect(spacer).toBeDefined()
    expect(spacer.format?.lineRawTwips).toBe(960) // 48pt × 20
    expect(spacer.format?.spaceAfter).toBe(0)
  })
})

describe('rebuildDocx: content-hugging margins (P16 C)', () => {
  it('content touching the page top means a MINIMAL margin, not the default', async () => {
    // Chromium-print PDFs run body text to the very page edge; y1 === heightPt
    const s = span('touches the top', { box: { x0: 72, y0: 780, x1: 300, y1: 792 } })
    const l: Line = { spans: [s], box: s.box, baseline: 782, endsWithHyphen: false }
    const parsed = await parseDocx(await rebuildDocx([page([textBlock([l])])]))
    // 12pt = 240 twips (MARGIN_MIN), not the 72pt default
    expect(parsed.internal.documentXml).toContain('w:top="240"')
  })

  it('a deep measured left margin is honored instead of capped at 1.5in', async () => {
    // every block starts at x=278 on a 596pt page (right-half column layout)
    const s = span('right half body', { box: { x0: 278, y0: 690, x1: 550, y1: 700 } })
    const l: Line = { spans: [s], box: s.box, baseline: 692, endsWithHyphen: false }
    const parsed = await parseDocx(
      await rebuildDocx([page([textBlock([l])], { widthPt: 596, heightPt: 843 })]),
    )
    expect(parsed.internal.documentXml).toContain(`w:left="${278 * 20}"`)
  })
})

describe('rebuildDocx: page vertical budget (P8)', () => {
  it("clamps a page-bottom image's spacing-before into the page budget", async () => {
    // top paragraph at 780, decorative image hugging the page bottom: the
    // faithful chain gap (768 − 40 = 728pt) plus margins and line heights
    // exceeds the page — unclamped, Word pushes the image to the next page
    const pngBytes = encodeRgbaPng(new Uint8Array(4 * 4 * 4).fill(200), 4, 4)
    const image: ImageBlock = {
      kind: 'image',
      box: { x0: 72, y0: 10, x1: 272, y1: 40 },
      data: pngBytes,
      mime: 'image/png',
      pixelWidth: 4,
      pixelHeight: 4,
      spacingBeforePt: 728,
    }
    const text = textBlock([line([span('top text')], 780)])
    const parsed = await parseDocx(await rebuildDocx([page([text, image])]))
    const xml = parsed.internal.documentXml
    // budget: usable = 792 − marginTop 12 − marginBottom 20 − slack 48 =
    // 712pt, plus the half-slack whitespace rebate (P16 C) = 736pt; block
    // heights = text 278tw (13.9pt) + image 30pt → the gap shrinks to what
    // is left: 736 − 43.9 = 692.1pt = 13842tw
    expect(xml).not.toContain('w:before="14560"') // the unclamped 728pt
    expect(xml).toContain('w:before="13842"')
    // nothing on the page may carry spacing beyond the spacing budget
    for (const m of xml.matchAll(/w:before="(\d+)"/g)) {
      expect(Number(m[1])).toBeLessThanOrEqual((712 + 24) * 20)
    }
  })

  it('leaves spacing alone when the page has room', async () => {
    const a = textBlock([line([span('first')], 700)])
    const b = textBlock([line([span('second')], 600)], { spacingBeforePt: 88 })
    const parsed = await parseDocx(await rebuildDocx([page([a, b])]))
    const second = parsed.blocks.find((bl) => (bl.runs ?? []).some((r) => r.text === 'second'))!
    // the page has slack far beyond the floor surpluses, so the measured gap
    // is kept verbatim — deducting on roomy pages only lifts the ink (P18 B)
    expect(second.format?.spaceBefore).toBe(88 * 20)
  })

  it('deducts the floor surplus only when the page is close to its budget', async () => {
    // 46 single-line blocks at 15pt pitch with ~1.9pt floor surplus each fill
    // the page: the un-deducted surpluses (~87pt) would overflow the budget,
    // so spacing-before pays them (P17 rule, now gated on need)
    const blocks = Array.from({ length: 46 }, (_, i) =>
      textBlock([line([span('row ' + i)], 770 - 15 * i)], {
        spacingBeforePt: i === 0 ? 0 : 3,
      }),
    )
    const parsed = await parseDocx(await rebuildDocx([page(blocks)]))
    const withSpace = parsed.blocks.filter((bl) => (bl.format?.spaceBefore ?? 0) > 0)
    expect(withSpace.length).toBeGreaterThan(0)
    for (const bl of withSpace) {
      expect(bl.format!.spaceBefore!).toBeLessThan(3 * 20)
    }
  })
})

describe('rebuildDocx: body wrap-risk headroom (P18 A)', () => {
  // tall body blocks of edge-filling lines + a bottom marker; the measured
  // gaps fit the budget — until the near-full page charges wrap headroom
  const build = async (bBlocks: number) => {
    const mk = (top: number) =>
      line([span('正文', { box: { x0: 72, y0: top - 12, x1: 540, y1: top } })], top)
    // short 2-line blocks: a long paragraph re-flows and swallows metric
    // drift, so only short blocks carry the wrap charge
    const pair = (top: number, sb: number) =>
      textBlock([mk(top), mk(top - 16)], { spacingBeforePt: sb })
    const a = Array.from({ length: 10 }, (_, i) => pair(770 - 32 * i, i === 0 ? 0 : 0))
    const b = Array.from({ length: bBlocks }, (_, i) => pair(430 - 32 * i, i === 0 ? 40 : 0))
    const c = textBlock([mk(40)], { spacingBeforePt: 40 })
    const parsed = await parseDocx(await rebuildDocx([page([...a, ...b, c])]))
    return parsed.blocks
      .map((bl) => bl.format?.spaceBefore ?? 0)
      .filter((v) => v > 0)
      .reduce((s, v) => s + v, 0)
  }

  it('squeezes the gaps of a near-full page whose body lines fill to the edge', async () => {
    const nearFull = await build(20) // heights ~644pt: fits only without wrap growth
    const roomy = await build(10) // same edge-filling lines, ~160pt of slack
    expect(roomy).toBeGreaterThanOrEqual(2 * 700) // both gaps essentially unsqueezed
    expect(nearFull).toBeLessThan(roomy - 300) // headroom charged, gaps squeezed
  })
})

describe('rebuildDocx: micro-section chrome charge (P18 A)', () => {
  // a near-full page of narrow (non-edge) body text, with signature-ruling
  // micro-sections (two short side-by-side columns) mid-page: LibreOffice
  // renders each continuous multi-column transition taller than the measured
  // model, so the budget reserves for them and squeezes the gaps
  const mk = (top: number, x0 = 72, x1 = 350) =>
    line([span('narrow text', { box: { x0, y0: top - 12, x1, y1: top } })], top)
  const stack = (top: number, n: number) => Array.from({ length: n }, (_, i) => mk(top - 16 * i))
  const build = async (withMicro: boolean) => {
    const a = textBlock(stack(770, 20))
    const b = textBlock(stack(430, 20), { spacingBeforePt: 40 })
    const c = textBlock([mk(40)], { spacingBeforePt: 40 })
    const label = textBlock([mk(450, 72, 200)])
    const ruling = textBlock([mk(450, 380, 540)])
    const label2 = textBlock([mk(444, 72, 200)])
    const ruling2 = textBlock([mk(444, 380, 540)])
    const sections: PageSection[] = [
      {
        box: { x0: 72, y0: 454, x1: 540, y1: 782 },
        dir: 'ltr',
        gutterWidthsPt: [],
        columns: [{ box: { x0: 72, y0: 454, x1: 540, y1: 782 }, blocks: [a] }],
      },
      ...(withMicro
        ? ([
            {
              box: { x0: 72, y0: 446, x1: 540, y1: 454 },
              dir: 'ltr',
              gutterWidthsPt: [180],
              columns: [
                { box: { x0: 72, y0: 446, x1: 200, y1: 454 }, blocks: [label] },
                { box: { x0: 380, y0: 446, x1: 540, y1: 454 }, blocks: [ruling] },
              ],
            },
            {
              box: { x0: 72, y0: 438, x1: 540, y1: 446 },
              dir: 'ltr',
              gutterWidthsPt: [180],
              columns: [
                { box: { x0: 72, y0: 438, x1: 200, y1: 446 }, blocks: [label2] },
                { box: { x0: 380, y0: 438, x1: 540, y1: 446 }, blocks: [ruling2] },
              ],
            },
          ] as PageSection[])
        : []),
      {
        box: { x0: 72, y0: 28, x1: 540, y1: 438 },
        dir: 'ltr',
        gutterWidthsPt: [],
        columns: [{ box: { x0: 72, y0: 28, x1: 540, y1: 438 }, blocks: [b, c] }],
      },
    ]
    const parsed = await parseDocx(
      await rebuildDocx([
        page([a, ...(withMicro ? [label, ruling, label2, ruling2] : []), b, c], { sections }),
      ]),
    )
    return parsed.blocks
      .map((bl) => bl.format?.spaceBefore ?? 0)
      .filter((v) => v > 0)
      .reduce((s, v) => s + v, 0)
  }

  it('reserves budget for short multi-column sections on a near-full page', async () => {
    const plain = await build(false)
    const withMicro = await build(true)
    expect(plain).toBeGreaterThanOrEqual(2 * 700) // fits, gaps unsqueezed
    expect(withMicro).toBeLessThan(plain - 300) // chrome charged, gaps squeezed
  })
})

describe('rebuildDocx: mixed page sizes (P5)', () => {
  it('a page-size change opens a new section with its own pgSz', async () => {
    const first = page([textBlock([line([span('a4 page')])])], { widthPt: 595, heightPt: 842 })
    const tall = page([textBlock([line([span('tall page')], 1500)])], {
      index: 1,
      widthPt: 405,
      heightPt: 1584,
    })
    const parsed = await parseDocx(await rebuildDocx([first, tall]))
    const { readSections } = await import('@genoffice/docx-engine')
    const sections = readSections(parsed)
    expect(sections).toHaveLength(2)
    expect(sections[0]!.settings.pageWidth).toBe(595 * 20)
    expect(sections[0]!.settings.pageHeight).toBe(842 * 20)
    expect(sections[1]!.settings.pageWidth).toBe(405 * 20)
    expect(sections[1]!.settings.pageHeight).toBe(1584 * 20)
    expect(sections[1]!.startType).toBe('nextPage')
  })
})

describe('rebuildDocx: blank pages (P24 A)', () => {
  it('a blank FIRST page gets its own placeholder paragraph, not just breaks', async () => {
    // Word/LO ignore w:pageBreakBefore on the document's first paragraph: a
    // doc of N blank pages emitted as N-1 break paragraphs renders N-1 pages
    const pages = [0, 1, 2, 3].map((i) => page([], { index: i }))
    const parsed = await parseDocx(await rebuildDocx(pages))
    const paras = parsed.blocks.filter((b) => b.type === 'paragraph')
    expect(paras).toHaveLength(4)
    // first paragraph occupies page 1 WITHOUT a page break; the other three
    // each break to a fresh page
    expect(paras[0]!.format?.pageBreakBefore).toBeFalsy()
    for (const p of paras.slice(1)) expect(p.format?.pageBreakBefore).toBe(true)
  })

  it('a blank first page keeps its own pgSz when later pages differ', async () => {
    const blank = page([], { widthPt: 612, heightPt: 792 })
    const landscape = page([textBlock([line([span('content')])])], {
      index: 1,
      widthPt: 792,
      heightPt: 612,
    })
    const parsed = await parseDocx(await rebuildDocx([blank, landscape]))
    const { readSections } = await import('@genoffice/docx-engine')
    const sections = readSections(parsed)
    expect(sections).toHaveLength(2)
    expect(sections[0]!.settings.pageWidth).toBe(612 * 20)
    expect(sections[0]!.settings.pageHeight).toBe(792 * 20)
    expect(sections[1]!.settings.pageWidth).toBe(792 * 20)
    expect(sections[1]!.settings.pageHeight).toBe(612 * 20)
  })

  it('a blank page after content still emits its break paragraph', async () => {
    const first = page([textBlock([line([span('hello')])])])
    const blank = page([], { index: 1 })
    const parsed = await parseDocx(await rebuildDocx([first, blank]))
    const breaks = parsed.blocks.filter((b) => b.type === 'paragraph' && b.format?.pageBreakBefore)
    expect(breaks).toHaveLength(1)
  })
})

describe('rebuildDocx: page-fit details (P6 C)', () => {
  it('borderless (confidence-marked) tables emit none-borders and tight cell margins', async () => {
    const table: TableBlock = {
      kind: 'table',
      box: { x0: 72, y0: 660, x1: 328, y1: 700 },
      colWidthsPt: [128, 128],
      confidence: 0.9,
      rows: [
        [
          {
            box: { x0: 72, y0: 660, x1: 200, y1: 700 },
            gridSpan: 1,
            blocks: [textBlock([line([span('A')])])],
          },
          { box: { x0: 200, y0: 660, x1: 328, y1: 700 }, gridSpan: 1, blocks: [] },
        ],
      ],
    }
    const parsed = await parseDocx(await rebuildDocx([page([table])]))
    const model = parsed.blocks.find((b) => b.type === 'table')!.table!
    expect(model.borders?.top?.style).toBe('none')
    expect(model.borders?.insideV?.style).toBe('none')
    expect(model.cellMarTwips?.left).toBe(15)
    // the EMPTY cell's paragraph is explicitly near-zero, not docDefaults-sized
    const empty = model.rows[0]![1]!
    expect(empty.richParas?.[0]?.lineRule).toBe('exact')
    expect(empty.richParas?.[0]?.lineRawTwips).toBeLessThanOrEqual(20)
  })

  it('gives the bottom margin overflow slack below the measured value', async () => {
    // content bottom at y=200 → measured bottom margin 108 (clamped), slack −24
    const block = textBlock([line([span('x')], 212)])
    const parsed = await parseDocx(await rebuildDocx([page([block])]))
    const { readSectionSettings } = await import('@genoffice/docx-engine')
    const settings = readSectionSettings(parsed)
    expect(settings.marginBottom).toBe((108 - 24) * 20)
  })
})

describe('rebuildDocx: re-emitted page furniture (P17)', () => {
  it('creates a header part with a PAGE-field footer and references it from the sectPr', async () => {
    const docx = await rebuildDocx([page([textBlock([line([span('body')])])])], {
      furnitureHf: [
        {
          band: 'top',
          text: 'ACME REPORT',
          pageNo: false,
          fontSizePt: 9,
          fontFamily: 'Arial',
          bold: false,
          italic: false,
          color: '444444',
          x0: 200,
          x1: 412,
          edgeDistPt: 30,
          coversFirstPage: true,
        },
        {
          band: 'bottom',
          text: `Page  of 4`,
          pageNo: true,
          fontSizePt: 9,
          fontFamily: 'Arial',
          bold: false,
          italic: false,
          color: '000000',
          x0: 280,
          x1: 330,
          edgeDistPt: 30,
          coversFirstPage: true,
        },
      ],
    })
    const { default: JSZip } = await import('jszip')
    const zip = await JSZip.loadAsync(docx)
    const doc = await zip.file('word/document.xml')!.async('string')
    expect(doc).toMatch(/<w:headerReference[^>]*\/>/)
    expect(doc).toMatch(/<w:footerReference[^>]*\/>/)
    const headerFile = zip.file(/word\/header\d+\.xml/)[0]
    expect(headerFile).toBeDefined()
    const headerXml = await headerFile!.async('string')
    expect(headerXml).toContain('ACME REPORT')
    expect(headerXml).toContain('<w:sz w:val="18"/>')
    const footerXml = await zip.file(/word\/footer\d+\.xml/)[0]!.async('string')
    expect(footerXml).toContain(' PAGE ')
    expect(footerXml).toContain('of 4')
    // parse round-trip keeps the header text visible to the editor
    const parsed = await parseDocx(docx)
    expect(parsed.headerText).toContain('ACME REPORT')
  })
})

describe('rebuildDocx: footnotes (P6)', () => {
  it('emits word/footnotes.xml entries and a w:footnoteReference at the anchor', async () => {
    const anchor = span('', {
      noteRef: '101',
      box: { x0: 300, y0: 690, x1: 300, y1: 700 },
    })
    const body = textBlock([line([span('The purpose'), anchor, span(' of this system')])])
    const noteBlock = textBlock([line([span('State the final goal here.')])])
    const parsed = await parseDocx(
      await rebuildDocx([page([body], { footnotes: [{ id: '101', blocks: [noteBlock] }] })]),
    )
    expect(parsed.footnotes).toHaveLength(1)
    expect(parsed.footnotes[0]!.id).toBe('101')
    expect(parsed.footnotes[0]!.text).toContain('State the final goal here.')
    const runs = parsed.blocks.flatMap((b) => b.runs ?? [])
    const ref = runs.find((r) => r.noteRef)
    expect(ref?.noteRef).toMatchObject({ kind: 'footnote', id: '101' })
    // note text lives ONLY in the footnotes part, not in the body flow
    const bodyText = runs.map((r) => r.text).join('')
    expect(bodyText).not.toContain('final goal')
  })

  it('carries the source font size into the footnotes part (P17)', async () => {
    // footnote body is 8pt in the source; emitted at the template's body
    // size it re-renders ~40% taller and spills the host page
    const anchor = span('', { noteRef: '101', box: { x0: 300, y0: 690, x1: 300, y1: 700 } })
    const body = textBlock([line([span('The purpose'), anchor])])
    const noteBlock = textBlock([
      line([span('Small print note.', { fontSize: 8, fontFamily: 'Arial' })]),
    ])
    const parsed = await parseDocx(
      await rebuildDocx([page([body], { footnotes: [{ id: '101', blocks: [noteBlock] }] })]),
    )
    expect(parsed.footnotes[0]!.text).toBe('Small print note.')
    expect(parsed.footnotes[0]!.richParas?.[0]).toEqual([
      { text: 'Small print note.', sizeHalfPoints: 16 },
    ])
  })
})

describe('rebuildDocx: TOC entries (P6)', () => {
  it('emits a TOC-styled paragraph with a dot-leader tab and the page number', async () => {
    const entry = textBlock([line([span('Introduction')])], {
      tocEntry: { level: 1, pageNumber: '42' },
    })
    const parsed = await parseDocx(await rebuildDocx([page([entry])]))
    const toc = parsed.blocks.find((b) => b.fieldDisplay?.kind === 'tocLine')
    expect(toc).toBeDefined()
    expect(toc!.fieldDisplay).toMatchObject({ left: 'Introduction', right: '42' })
  })
})

describe('bytesToBase64', () => {
  it('matches the platform encoder', () => {
    const bytes = Uint8Array.from({ length: 300 }, (_, i) => i % 256)
    expect(bytesToBase64(bytes)).toBe(Buffer.from(bytes).toString('base64'))
  })
})

describe('document background (pageColor)', () => {
  const bgPage = (idx: number, bgColor?: string) =>
    page([textBlock([line([span(`page ${idx}`)])])], {
      index: idx,
      ...(bgColor ? { bgColor } : {}),
    })

  it('emits w:background when most pages share a background color', async () => {
    const docx = await rebuildDocx([bgPage(0, 'F0E8D8'), bgPage(1, 'F0E8D8'), bgPage(2)])
    const parsed = await parseDocx(docx)
    expect(parsed.internal.documentXml).toContain('<w:background w:color="F0E8D8"')
  })

  it('stays white when only a minority page is colored', async () => {
    const docx = await rebuildDocx([bgPage(0, '2A9D8F'), bgPage(1), bgPage(2)])
    const parsed = await parseDocx(docx)
    expect(parsed.internal.documentXml).not.toContain('<w:background')
  })

  it('clusters near-identical washes so per-page tone drift still wins', async () => {
    const docx = await rebuildDocx([
      bgPage(0, 'EFE4D2'),
      bgPage(1, 'F5EDE0'),
      bgPage(2, 'FDF1E4'),
      bgPage(3, '2A0A0F'),
    ])
    const parsed = await parseDocx(docx)
    expect(parsed.internal.documentXml).toMatch(/<w:background w:color="(EFE4D2|F5EDE0|FDF1E4)"/)
  })
})
