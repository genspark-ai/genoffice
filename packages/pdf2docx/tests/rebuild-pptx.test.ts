/** rebuild-pptx unit tests (P25): hand-built IR pages, verified via openPptx. */
import { describe, expect, it } from 'vitest'
import { openPptx } from '../../pptx-engine/src/index'
import type { Rect } from '../src/geometry'
import type { IrPage, Line, Span, TableBlock, TextBlock } from '../src/ir'
import { rebuildPptx } from '../src/rebuild-pptx'

const EMU_PER_PT = 12700

const span = (text: string, over: Partial<Span> = {}): Span => ({
  text,
  box: { x0: 0, y0: 0, x1: 10, y1: 10 },
  fontSize: 18,
  fontFamily: 'Helvetica',
  bold: false,
  italic: false,
  color: '112233',
  dir: 'ltr',
  script: 'latin',
  ...over,
})

const line = (text: string, box: Rect, over: Partial<Line> = {}): Line => ({
  spans: [span(text, { box })],
  box,
  baseline: box.y0 + 2,
  endsWithHyphen: false,
  ...over,
})

const textBlock = (lines: Line[], box: Rect, over: Partial<TextBlock> = {}): TextBlock => ({
  kind: 'text',
  lines,
  box,
  align: 'left',
  firstLineIndentPt: 0,
  dir: 'ltr',
  ...over,
})

const page = (over: Partial<IrPage> = {}): IrPage => ({
  index: 0,
  widthPt: 720,
  heightPt: 405,
  rotation: 0,
  blocks: [],
  degraded: false,
  scanned: false,
  hasStructTree: false,
  ...over,
})

// tiny valid 1×1 png (89 bytes)
const PNG_1PX = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  ),
  (c) => c.charCodeAt(0),
)

describe('rebuildPptx', () => {
  it('slide size mirrors the PDF page size in EMU (16:9 landscape page)', async () => {
    const opened = await openPptx(await rebuildPptx([page()]))
    expect(opened.deck.size).toEqual({ cx: 720 * EMU_PER_PT, cy: 405 * EMU_PER_PT })
    expect(opened.deck.slides.length).toBe(1)
  })

  it('page count = slide count, portrait stays portrait', async () => {
    const pages = [0, 1, 2].map((i) => page({ index: i, widthPt: 595, heightPt: 842 }))
    const opened = await openPptx(await rebuildPptx(pages))
    expect(opened.deck.slides.length).toBe(3)
    expect(opened.deck.size.cy).toBeGreaterThan(opened.deck.size.cx)
  })

  it('one text block = one textbox with merged runs at measured coordinates', async () => {
    const box: Rect = { x0: 72, y0: 355, x1: 360, y1: 391 } // two 18pt lines
    const block = textBlock(
      [
        line('Title line one', { x0: 72, y0: 373, x1: 300, y1: 391 }),
        line('and line two', { x0: 72, y0: 355, x1: 260, y1: 373 }),
      ],
      box,
    )
    const opened = await openPptx(await rebuildPptx([page({ blocks: [block] })]))
    const texts = opened.deck.slides[0]!.elements.filter((e) => e.type === 'text') as any[]
    expect(texts.length).toBe(1)
    const paras = texts[0].text.paragraphs
    expect(paras.length).toBe(1)
    const joined = paras[0].runs.map((r: any) => r.text).join('')
    expect(joined).toBe('Title line one and line two')
    // same style spans merge into a single run
    expect(paras[0].runs.length).toBe(1)
    expect(paras[0].runs[0].fontSize).toBe(18)
    expect(paras[0].runs[0].color).toBe('#112233')
    // measured position: x = 72pt, y = pageH - y1 = 405-391 = 14pt
    const off = texts[0].transform.offset
    expect(off.x).toBe(72 * EMU_PER_PT)
    expect(off.y).toBe(14 * EMU_PER_PT)
    // exact line pitch = blockH/2 = 18pt
    expect(paras[0].lineExact).toBeCloseTo(18, 5)
  })

  it('two text blocks stay two textboxes (granularity)', async () => {
    const blocks = [
      textBlock([line('A', { x0: 72, y0: 373, x1: 150, y1: 391 })], {
        x0: 72,
        y0: 373,
        x1: 150,
        y1: 391,
      }),
      textBlock([line('B', { x0: 72, y0: 300, x1: 150, y1: 318 })], {
        x0: 72,
        y0: 300,
        x1: 150,
        y1: 318,
      }),
    ]
    const opened = await openPptx(await rebuildPptx([page({ blocks })]))
    expect(opened.deck.slides[0]!.elements.filter((e) => e.type === 'text').length).toBe(2)
  })

  it('table block maps gridSpan/vMerge/fill onto a native pptx table', async () => {
    const cellText = (t: string, box: Rect): TextBlock => textBlock([line(t, box)], box)
    const table: TableBlock = {
      kind: 'table',
      box: { x0: 100, y0: 205, x1: 400, y1: 305 },
      colWidthsPt: [100, 100, 100],
      rows: [
        [
          {
            box: { x0: 100, y0: 255, x1: 300, y1: 305 },
            gridSpan: 2,
            fill: 'DDEEFF',
            blocks: [cellText('head', { x0: 105, y0: 260, x1: 200, y1: 300 })],
          },
          {
            box: { x0: 300, y0: 205, x1: 400, y1: 305 },
            gridSpan: 1,
            vMerge: 'restart',
            vAlign: 'center',
            blocks: [cellText('tall', { x0: 305, y0: 210, x1: 395, y1: 300 })],
          },
        ],
        [
          {
            box: { x0: 100, y0: 205, x1: 200, y1: 255 },
            gridSpan: 1,
            blocks: [cellText('a2', { x0: 105, y0: 210, x1: 195, y1: 250 })],
          },
          { box: { x0: 200, y0: 205, x1: 300, y1: 255 }, gridSpan: 1, blocks: [] },
          {
            box: { x0: 300, y0: 205, x1: 400, y1: 305 },
            gridSpan: 1,
            vMerge: 'continue',
            blocks: [],
          },
        ],
      ],
    }
    const opened = await openPptx(await rebuildPptx([page({ blocks: [table] })]))
    const tbl = opened.deck.slides[0]!.elements.find((e) => e.type === 'table') as any
    expect(tbl).toBeTruthy()
    expect(tbl.colWidths.length).toBe(3)
    expect(tbl.colWidths[0]).toBe(100 * EMU_PER_PT)
    expect(tbl.rows.length).toBe(2)
    // r0c0 spans 2 columns, r0c1 covered
    expect(tbl.rows[0][0].gridSpan).toBe(2)
    expect(tbl.rows[0][1].merged).toBe(true)
    // r0c2 spans 2 rows; r1c2 covered
    expect(tbl.rows[0][2].rowSpan).toBe(2)
    expect(tbl.rows[1][2].merged).toBe(true)
    expect(tbl.rows[0][0].fill).toEqual({ type: 'solid', color: '#DDEEFF' })
    const cellRuns = tbl.rows[0][0].text.paragraphs.flatMap((p: any) => p.runs)
    expect(cellRuns.map((r: any) => r.text).join('')).toBe('head')
    // row heights from measured boundaries: 50pt each
    expect(tbl.rowHeights[0]).toBe(50 * EMU_PER_PT)
    expect(tbl.rowHeights[1]).toBe(50 * EMU_PER_PT)
  })

  it('scanned page ships as one full-slide picture, still 1 page = 1 slide', async () => {
    const scanned = page({
      scanned: true,
      render: { data: PNG_1PX, mime: 'image/png', pixelWidth: 1, pixelHeight: 1 },
    })
    const opened = await openPptx(await rebuildPptx([scanned, page({ index: 1 })]))
    expect(opened.deck.slides.length).toBe(2)
    const els = opened.deck.slides[0]!.elements
    expect(els.length).toBe(1)
    expect(els[0]!.type).toBe('picture')
    const off = els[0]!.transform.offset
    expect(off.cx).toBe(720 * EMU_PER_PT)
    expect(off.cy).toBe(405 * EMU_PER_PT)
  })

  it('differently-sized page letterboxes uniformly onto the deck size', async () => {
    // deck from page 0 (720x405); page 1 is portrait A4 (595x842) → s = 405/842
    const p1 = page({
      index: 1,
      widthPt: 595,
      heightPt: 842,
      blocks: [
        textBlock([line('x', { x0: 0, y0: 824, x1: 100, y1: 842 })], {
          x0: 0,
          y0: 824,
          x1: 100,
          y1: 842,
        }),
      ],
    })
    const opened = await openPptx(await rebuildPptx([page(), p1]))
    const el: any = opened.deck.slides[1]!.elements.find((e) => e.type === 'text')
    const s = 405 / 842
    const ox = (720 - 595 * s) / 2
    expect(el.transform.offset.x).toBe(Math.round(ox * EMU_PER_PT))
    expect(el.transform.offset.y).toBe(0)
    // font size rides the scale (sz persists in 1/100 pt → 2-decimal precision)
    expect(el.text.paragraphs[0].runs[0].fontSize).toBeCloseTo(18 * s, 1)
  })
})
