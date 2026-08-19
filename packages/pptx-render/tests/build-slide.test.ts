import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import { openPptx } from '@genoffice/pptx-engine'
import { buildRenderSlide } from '../src/index'

const here = dirname(fileURLToPath(import.meta.url))
const enginePptx = (name: string) =>
  readFileSync(join(here, '..', '..', 'pptx-engine', 'tests', 'fixtures', name))

describe('buildRenderSlide (end-to-end on real fixture)', () => {
  it('builds a RenderSlide with correct viewport + positioned nodes', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const slide = deck.slides[0]!
    const rs = buildRenderSlide(slide, deck.size, { fitWidthPx: 1280 })

    // 16:9 viewport
    expect(rs.widthPx).toBe(1280)
    expect(rs.heightPx).toBeCloseTo(720, 0)
    expect(rs.nodes.length).toBeGreaterThan(0)

    // Placeholder title/body geometry has been backfilled → render boxes are non-zero and inside the canvas
    const textNodes = rs.nodes.filter((n) => n.type === 'text' || n.type === 'shape')
    expect(textNodes.length).toBeGreaterThan(0)
    for (const n of textNodes) {
      expect(n.box.w).toBeGreaterThan(0)
      expect(n.box.h).toBeGreaterThan(0)
      expect(n.box.x).toBeGreaterThanOrEqual(0)
      expect(n.box.x).toBeLessThan(rs.widthPx)
    }
  })

  it('lays out text for title placeholder into lines with glyph runs', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const rs = buildRenderSlide(deck.slides[0]!, deck.size, { fitWidthPx: 1280 })
    const withText = rs.nodes.find(
      (n) => (n.type === 'text' || n.type === 'shape') && (n as any).text?.lines.length,
    ) as any
    expect(withText).toBeTruthy()
    expect(withText.text.lines.length).toBeGreaterThan(0)
    const allRunText = withText.text.lines
      .flatMap((l: any) => l.runs)
      .map((r: any) => r.text)
      .join('')
    expect(allRunText.trim().length).toBeGreaterThan(0)
    // Every run has a positive font size and baseline
    for (const ln of withText.text.lines) {
      for (const r of ln.runs) {
        expect(r.fontSizePx).toBeGreaterThan(0)
        expect(r.baselineY).toBeGreaterThan(0)
      }
    }
  })

  it('renders CJK/emoji deck without throwing and keeps text', async () => {
    const { deck } = await openPptx(enginePptx('05_unicode_cjk_emoji.pptx'))
    let sawText = false
    for (const slide of deck.slides) {
      const rs = buildRenderSlide(slide, deck.size, { fitWidthPx: 960 })
      for (const n of rs.nodes) {
        if ((n.type === 'text' || n.type === 'shape') && (n as any).text) {
          const t = (n as any).text.lines
            .flatMap((l: any) => l.runs)
            .map((r: any) => r.text)
            .join('')
          // eslint-disable-next-line no-control-regex -- ASCII range check is intentional
          if (/[^\x00-\x7F]/.test(t)) sawText = true
        }
      }
    }
    expect(sawText).toBe(true)
  })

  it('roundRect cornerRadiusPx follows avLst adj (50000 → pill = half min edge)', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const slide = deck.slides[0]!
    // Build a roundRect element (mutating the model directly to avoid depending on fixture content)
    const el: any = {
      id: 'sp_test',
      type: 'shape',
      anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
      transform: {
        offset: { x: 0, y: 0, cx: 914400, cy: 457200 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      presetGeometry: 'roundRect',
      adjust: { adj: 50000 },
      fill: { type: 'solid', color: '#00FF00' },
    }
    const rs = buildRenderSlide({ ...slide, elements: [el], decorations: [] }, deck.size, {
      fitWidthPx: 1280,
    })
    const n = rs.nodes[0] as any
    expect(n.cornerRadiusPx).toBeCloseTo(n.box.h / 2, 1) // short side is the height → half height = pill
    // Without avLst, the default 16.667% applies
    const el2 = { ...el, adjust: undefined }
    const rs2 = buildRenderSlide({ ...slide, elements: [el2], decorations: [] }, deck.size, {
      fitWidthPx: 1280,
    })
    const n2 = rs2.nodes[0] as any
    expect(n2.cornerRadiusPx).toBeCloseTo(Math.min(n2.box.w, n2.box.h) * 0.16667, 1)
  })

  it('resolves shadow to px offsets (dir 270° → straight up)', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const slide = deck.slides[0]!
    const el: any = {
      id: 'sp_shdw',
      type: 'shape',
      anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
      transform: {
        offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      presetGeometry: 'rect',
      fill: { type: 'solid', color: '#FFFFFF' },
      shadow: { color: '#00000033', blurRad: 63500, dist: 25400, dirDeg: 270 },
    }
    const rs = buildRenderSlide({ ...slide, elements: [el], decorations: [] }, deck.size, {
      fitWidthPx: 1280,
    })
    const n = rs.nodes[0] as any
    expect(n.shadow.color).toBe('#00000033')
    expect(n.shadow.blurPx).toBeGreaterThan(0)
    expect(n.shadow.offsetX).toBeCloseTo(0, 5)
    expect(n.shadow.offsetY).toBeLessThan(0) // 270° = upward
  })

  it('passes picture srcRect through to render node', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const slide = deck.slides[0]!
    const el: any = {
      id: 'pic_crop',
      type: 'picture',
      anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
      transform: {
        offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      mediaRef: '',
      srcRect: { l: 0.1, t: 0.2, r: 0.3, b: 0 },
    }
    const rs = buildRenderSlide({ ...slide, elements: [el], decorations: [] }, deck.size, {
      fitWidthPx: 1280,
    })
    expect((rs.nodes[0] as any).srcRect).toEqual({ l: 0.1, t: 0.2, r: 0.3, b: 0 })
  })

  it('picture presetGeometry -> clip three channels (ellipse/roundRect/polygon/unknown fallback)', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const slide = deck.slides[0]!
    const mk = (presetGeometry?: string, adjust?: Record<string, number>): any => ({
      id: 'pic_geom',
      type: 'picture',
      anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
      transform: {
        offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      mediaRef: '',
      ...(presetGeometry ? { presetGeometry } : {}),
      ...(adjust ? { adjust } : {}),
    })
    const build = (el: any) =>
      buildRenderSlide({ ...slide, elements: [el], decorations: [] }, deck.size, {
        fitWidthPx: 1280,
      }).nodes[0] as any

    const ellipse = build(mk('ellipse'))
    expect(ellipse.clip?.pathData).toMatch(/^M 0 .+ A .+ Z$/)

    const round = build(mk('roundRect', { adj: 50000 }))
    expect(round.clip?.cornerRadiusPx).toBeCloseTo(build(mk('ellipse')).box.w / 2, 1)

    const hex = build(mk('hexagon'))
    expect(hex.clip?.polygonPoints?.length).toBeGreaterThanOrEqual(12)

    expect(build(mk()).clip).toBeUndefined()
    expect(build(mk('rect')).clip).toBeUndefined()
    expect(build(mk('unknownPreset9')).clip).toBeUndefined()
  })

  it('slidenum field displays opts.slideNo without mutating the model', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const slide = deck.slides[0]!
    const el: any = {
      id: 'sp_fld',
      type: 'text',
      anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
      transform: {
        offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      text: {
        paragraphs: [
          {
            runs: [
              { text: 'Pg', fontSize: 12 },
              { text: '3', fontSize: 12, field: 'slidenum' },
            ],
          },
        ],
        insets: { l: 0, t: 0, r: 0, b: 0 },
      },
    }
    const custom = { ...slide, elements: [el], decorations: [] }
    const glyphText = (rs: any) =>
      rs.nodes[0].text.lines
        .flatMap((l: any) => l.runs)
        .map((r: any) => r.text)
        .join('')

    const withNo = buildRenderSlide(custom, deck.size, { fitWidthPx: 1280, slideNo: 5 })
    expect(glyphText(withNo)).toBe('Pg5')
    // model run text was not touched
    expect(el.text.paragraphs[0].runs[1].text).toBe('3')

    const withoutNo = buildRenderSlide(custom, deck.size, { fitWidthPx: 1280 })
    expect(glyphText(withoutNo)).toBe('Pg3')
  })

  it('renders tables as table nodes with positioned cells (no more chip)', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    let sawTable = false
    for (const slide of deck.slides) {
      const rs = buildRenderSlide(slide, deck.size, { fitWidthPx: 960 })
      for (const n of rs.nodes) {
        if (n.type === 'table') {
          sawTable = true
          const cells = (n as any).cells
          expect(cells.length).toBeGreaterThan(0)
          // Cell geometry lies within the table box and cells align with each other
          for (const c of cells) {
            expect(c.w).toBeGreaterThan(0)
            expect(c.h).toBeGreaterThan(0)
            expect(c.x + c.w).toBeLessThanOrEqual(n.box.w + 0.01)
            expect(c.y + c.h).toBeLessThanOrEqual(n.box.h + 0.01)
          }
        }
      }
    }
    // The fixture contains a table → at least one table node
    expect(sawTable).toBe(true)
  })

  it('table cell spans expand across columns and merged cells are skipped', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const slide = deck.slides[0]!
    const el: any = {
      id: 'tbl_1',
      type: 'table',
      anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
      transform: {
        offset: { x: 0, y: 0, cx: 1905000, cy: 952500 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      colWidths: [952500, 952500],
      rowHeights: [476250, 476250],
      rows: [
        [{ gridSpan: 2, text: undefined }, { merged: true }],
        [{}, {}],
      ],
    }
    const rs = buildRenderSlide({ ...slide, elements: [el], decorations: [] }, deck.size, {
      fitWidthPx: 1280,
    })
    const node = rs.nodes[0] as any
    expect(node.type).toBe('table')
    // Merged placeholder cells are not emitted: 2 rows × 2 cols - 1 merged = 3
    expect(node.cells.length).toBe(3)
    // A cell spanning two columns is as wide as the whole table
    expect(node.cells[0].w).toBeCloseTo(node.box.w, 1)
    // The two cells in row 2 each take half
    expect(node.cells[1].w).toBeCloseTo(node.box.w / 2, 1)
  })

  it('table renders at its grid width when the frame ext is a stale placeholder', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const slide = deck.slides[0]!
    const el: any = {
      id: 'tbl_1',
      type: 'table',
      anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
      // ext cx=3000000 is stale: the grid below sums to 6593050 EMU (real prod deck shape)
      transform: {
        offset: { x: 0, y: 0, cx: 3000000, cy: 3000000 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      colWidths: [1449850, 5143200],
      rowHeights: [645100, 645100],
      rows: [
        [{}, {}],
        [{}, {}],
      ],
    }
    const rs = buildRenderSlide({ ...slide, elements: [el], decorations: [] }, deck.size, {
      fitWidthPx: 1280,
    })
    const node = rs.nodes[0] as any
    const emuPerPx = deck.size.cx / 1280
    expect(node.box.w).toBeCloseTo(6593050 / emuPerPx, 0)
    expect(node.cells[1].w / node.cells[0].w).toBeCloseTo(5143200 / 1449850, 1)
  })

  it('a table inside a scaled group keeps the group scale on its grid sizes', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const slide = deck.slides[0]!
    const tbl: any = {
      id: 'tbl_g',
      type: 'table',
      anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
      transform: {
        offset: { x: 0, y: 0, cx: 3810000, cy: 1905000 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      colWidths: [1905000, 1905000],
      rowHeights: [952500, 952500],
      rows: [
        [{}, {}],
        [{}, {}],
      ],
    }
    const grp: any = {
      id: 'grp_1',
      type: 'group',
      anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
      // group is half the size of its child coordinate space -> children scale by 0.5
      transform: {
        offset: { x: 0, y: 0, cx: 1905000, cy: 952500 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      childOffset: { x: 0, y: 0, cx: 3810000, cy: 1905000 },
      children: [tbl],
    }
    const rs = buildRenderSlide({ ...slide, elements: [grp], decorations: [] }, deck.size, {
      fitWidthPx: 1280,
    })
    const node = rs.nodes[0] as any
    const table = node.children[0]
    const emuPerPx = deck.size.cx / 1280
    expect(table.box.w).toBeCloseTo(1905000 / emuPerPx, 0)
    expect(table.cells[0].w).toBeCloseTo(952500 / emuPerPx, 0)
  })

  it('cells after a mid-row merge keep their grid position (continuation tc is not double-counted)', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const slide = deck.slides[0]!
    const el: any = {
      id: 'tbl_1',
      type: 'table',
      anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
      transform: {
        offset: { x: 0, y: 0, cx: 3810000, cy: 952500 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      colWidths: [952500, 952500, 952500, 952500],
      rowHeights: [476250, 476250],
      rows: [
        [{}, { gridSpan: 2 }, { merged: true }, {}],
        [{}, {}, {}, {}],
      ],
    }
    const rs = buildRenderSlide({ ...slide, elements: [el], decorations: [] }, deck.size, {
      fitWidthPx: 1280,
    })
    const node = rs.nodes[0] as any
    const colW = node.box.w / 4
    const row0 = node.cells.filter((c: any) => c.row === 0)
    // 3 emitted cells: [col0], [col1-2 merged], [col3]
    expect(row0.length).toBe(3)
    expect(row0[1].x).toBeCloseTo(colW, 1)
    expect(row0[1].w).toBeCloseTo(colW * 2, 1)
    // The cell after the merge sits at column 3, not shifted right off the table
    expect(row0[2].x).toBeCloseTo(colW * 3, 1)
    expect(row0[2].w).toBeCloseTo(colW, 1)
    // Grid lines exposed for the canvas resize grips
    expect(node.gridX.length).toBe(5)
    expect(node.gridY.length).toBe(3)
    expect(node.gridX[4]).toBeCloseTo(node.box.w, 1)
  })

  it('grows table rows to fit wrapped cell text (row height is a minimum)', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const slide = deck.slides[0]!
    const longText = {
      paragraphs: [
        {
          runs: [
            {
              text: 'a long piece of cell text that definitely wraps across many lines '.repeat(4),
              fontSize: 14,
            },
          ],
        },
      ],
    }
    const mkTable = (text?: any): any => ({
      id: 'tbl_grow',
      type: 'table',
      anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
      // Narrow 2-col table with two short rows
      transform: {
        offset: { x: 0, y: 0, cx: 1905000, cy: 952500 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      colWidths: [952500, 952500],
      rowHeights: [476250, 476250],
      rows: [
        [{ text }, {}],
        [{}, {}],
      ],
    })
    const build = (el: any) =>
      buildRenderSlide({ ...slide, elements: [el], decorations: [] }, deck.size, {
        fitWidthPx: 1280,
      }).nodes[0] as any

    const plain = build(mkTable())
    const grown = build(mkTable(longText))
    const cellOf = (n: any, r: number, c: number) =>
      n.cells.find((x: any) => x.row === r && x.col === c)

    // Row 0 grew to fit the wrapped text; row 1 kept its height but moved down
    expect(cellOf(grown, 0, 0).h).toBeGreaterThan(cellOf(plain, 0, 0).h)
    const layout = cellOf(grown, 0, 0).text
    expect(layout.contentHeight + layout.insets.t + layout.insets.b).toBeLessThanOrEqual(
      cellOf(grown, 0, 0).h + 0.01,
    )
    expect(cellOf(grown, 1, 0).h).toBeCloseTo(cellOf(plain, 1, 0).h, 1)
    expect(cellOf(grown, 1, 0).y).toBeCloseTo(cellOf(grown, 0, 0).h, 1)
    // The node box grew with the total so nothing is clipped
    expect(grown.box.h).toBeCloseTo(cellOf(grown, 0, 0).h + cellOf(grown, 1, 0).h, 1)
    expect(grown.box.h).toBeGreaterThan(plain.box.h)
  })

  it('maps group children into group-local coords (chOff subtracted, chExt scale baked into child boxes)', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const slide = deck.slides[0]!
    const anchor = { spIndex: -1, originalXml: '', range: [0, 0] }
    const group: any = {
      id: 'grp1',
      type: 'group',
      anchor,
      // Group box 2×1 inch; child coordinate origin at (5,2.5) inch, size 1×0.5 inch → childScale 2
      transform: {
        offset: { x: 914400, y: 914400, cx: 1828800, cy: 914400 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      childOffset: { x: 4572000, y: 2286000, cx: 914400, cy: 457200 },
      children: [
        {
          id: 'c1',
          type: 'shape',
          anchor,
          // The child coincides with chOff → local (0,0) within the group
          transform: {
            offset: { x: 4572000, y: 2286000, cx: 457200, cy: 228600 },
            rot: 0,
            flipH: false,
            flipV: false,
          },
          presetGeometry: 'rect',
          fill: { type: 'solid', color: '#123456' },
        },
      ],
    }
    const rs = buildRenderSlide({ ...slide, elements: [group], decorations: [] }, deck.size, {
      fitWidthPx: 1280,
    })
    const g = rs.nodes[0] as any
    expect(g.type).toBe('group')
    expect(g.children).toHaveLength(1)
    // ext/chExt = 2x (metadata for edit-pipeline conversion only; render container no longer applies it)
    expect(g.childScaleX).toBeCloseTo(2, 5)
    expect(g.childScaleY).toBeCloseTo(2, 5)
    // child local coords: same point as chOff -> (0,0); width = child EMU x childScale (scaling baked into geometry)
    const c = g.children[0]
    expect(c.box.x).toBeCloseTo(0, 3)
    expect(c.box.y).toBeCloseTo(0, 3)
    expect(c.box.w).toBeCloseTo(g.box.w / 2, 1) // child cx = half of chExt, childScale already included
  })

  it('lays out group child text against the scaled box with unscaled font (non-uniform ext/chExt)', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const slide = deck.slides[0]!
    const anchor = { spIndex: -1, originalXml: '', range: [0, 0] }
    const mkRun = (text: string) => ({ text, fontSize: 18, fontFamily: 'Arial' })
    const child = (id: string, offset: { x: number; y: number; cx: number; cy: number }) => ({
      id,
      type: 'text',
      anchor,
      transform: { offset, rot: 0, flipH: false, flipV: false },
      presetGeometry: 'rect',
      fill: { type: 'none' },
      text: {
        bodyPr: { wrap: true, anchor: 't', lIns: 0, tIns: 0, rIns: 0, bIns: 0 },
        paragraphs: [{ runs: [mkRun('word '.repeat(20).trim())], properties: {} }],
      },
    })
    const offset = { x: 0, y: 0, cx: 457200, cy: 457200 }
    const groupOf = (kids: any[], scaleX: number, scaleY: number): any => ({
      id: `grp_${scaleX}`,
      type: 'group',
      anchor,
      transform: {
        offset: { x: 0, y: 0, cx: 914400 * scaleX, cy: 914400 * scaleY },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      childOffset: { x: 0, y: 0, cx: 914400, cy: 914400 },
      children: kids,
    })
    // same text: 1x group vs 3x horizontal group -- in the 3x group layout width is x3, fewer wrapped lines; font size unchanged
    const rs = buildRenderSlide(
      {
        ...slide,
        elements: [groupOf([child('t1', offset)], 1, 1), groupOf([child('t2', offset)], 3, 1)],
        decorations: [],
      },
      deck.size,
      { fitWidthPx: 1280 },
    )
    const [g1, g3] = rs.nodes as any[]
    const t1 = g1.children[0]
    const t3 = g3.children[0]
    expect(t3.box.w).toBeCloseTo(t1.box.w * 3, 1)
    expect(t3.text.lines.length).toBeLessThan(t1.text.lines.length)
    // font size does not scale with the group (group scaling only affects geometry)
    const size1 = t1.text.lines[0].runs[0].fontSizePx
    const size3 = t3.text.lines[0].runs[0].fontSizePx
    expect(size3).toBeCloseTo(size1, 3)
  })

  it('builds connectors as line render nodes with flip baked into points', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const slide = deck.slides[0]!
    const el: any = {
      id: 'cxn1',
      type: 'shape',
      anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
      transform: {
        offset: { x: 0, y: 0, cx: 914400, cy: 457200 },
        rot: 0,
        flipH: false,
        flipV: true,
      },
      presetGeometry: 'straightConnector1',
      fill: { type: 'none' },
      stroke: {
        fill: { type: 'solid', color: '#FF0000' },
        width: 12700,
        tailEnd: { type: 'triangle' },
      },
    }
    const rs = buildRenderSlide({ ...slide, elements: [el], decorations: [] }, deck.size, {
      fitWidthPx: 1280,
    })
    const n = rs.nodes[0] as any
    expect(n.line).toBeTruthy()
    expect(n.line.tailEnd).toBeTruthy()
    expect(n.line.tailEnd.type).toBe('triangle')
    expect(n.line.headEnd).toBeUndefined()
    // flipV is baked into the points: start at bottom-left, end at top-right; the box carries no flip flags anymore
    expect(n.line.points).toEqual([0, n.box.h, n.box.w, 0])
    expect(n.box.flipV).toBe(false)
  })

  it('builds preset polygons (triangle) and renders decorations as flagged nodes', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const slide = deck.slides[0]!
    const anchor = { spIndex: -1, originalXml: '', range: [0, 0] }
    const tri: any = {
      id: 'tri1',
      type: 'shape',
      anchor,
      transform: {
        offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      presetGeometry: 'triangle',
      fill: { type: 'solid', color: '#00FF00' },
    }
    const deco: any = {
      id: 'logo1',
      type: 'shape',
      anchor,
      transform: {
        offset: { x: 0, y: 0, cx: 914400, cy: 100000 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      presetGeometry: 'rect',
      fill: { type: 'solid', color: '#C43E1C' },
    }
    const rs = buildRenderSlide({ ...slide, elements: [tri], decorations: [deco] }, deck.size, {
      fitWidthPx: 1280,
    })
    // The decoration layer comes first and carries the decoration flag
    expect(rs.nodes[0]!.decoration).toBe(true)
    expect(rs.nodes[0]!.sourceId).toBe('logo1')
    const n = rs.nodes[1] as any
    expect(n.polygonPoints).toBeTruthy()
    expect(n.polygonPoints.length).toBe(6)
    expect(n.polygonPoints[0]).toBeCloseTo(n.box.w / 2, 3)
  })

  it('scales customGeometry unit path to box px (pathData)', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const slide = deck.slides[0]!
    const el: any = {
      id: 'cg1',
      type: 'shape',
      anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
      transform: {
        offset: { x: 0, y: 0, cx: 914400, cy: 457200 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      customGeometry: { path: 'M 0 1 L 0.5 0 L 1 1 Z', strokePath: 'M 0 0 L 1 1' },
      fill: { type: 'solid', color: '#00FF00' },
    }
    const rs = buildRenderSlide({ ...slide, elements: [el], decorations: [] }, deck.size, {
      fitWidthPx: 1280,
    })
    const n = rs.nodes[0] as any
    const toks = n.pathData.split(' ')
    expect(toks[0]).toBe('M')
    expect(Number(toks[1])).toBeCloseTo(0, 2)
    expect(Number(toks[2])).toBeCloseTo(n.box.h, 1)
    expect(Number(toks[4])).toBeCloseTo(n.box.w / 2, 1)
    // strokePath scales along with it
    expect(n.strokePathData.split(' ').map(Number).filter(Number.isFinite).pop()).toBeCloseTo(
      n.box.h,
      1,
    )
    expect(n.polygonPoints).toBeUndefined()
  })

  it('falls back to presetPath channel for curved presets (pie)', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const slide = deck.slides[0]!
    const el: any = {
      id: 'pie1',
      type: 'shape',
      anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
      transform: {
        offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      presetGeometry: 'pie',
      fill: { type: 'solid', color: '#00FF00' },
    }
    const rs = buildRenderSlide({ ...slide, elements: [el], decorations: [] }, deck.size, {
      fitWidthPx: 1280,
    })
    const n = rs.nodes[0] as any
    expect(n.polygonPoints).toBeUndefined()
    expect(n.pathData).toMatch(/^M .* C /)
    expect(n.pathData.endsWith('Z')).toBe(true)
  })

  it('useBgFill shapes take the slide background fill (tdf93868)', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const slide = deck.slides[0]!
    const bg = {
      type: 'gradient',
      stops: [
        { pos: 0, color: '#000000' },
        { pos: 1, color: '#3F3F3F' },
      ],
      angle: 16200000,
    }
    const el: any = {
      id: 'bgsp1',
      type: 'shape',
      anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
      transform: {
        offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      presetGeometry: 'roundRect',
      fill: { type: 'solid', color: '#FFFFFF' },
      useBgFill: true,
    }
    const rs = buildRenderSlide(
      { ...slide, elements: [el], decorations: [], background: bg as any },
      deck.size,
      { fitWidthPx: 1280 },
    )
    const n = rs.nodes[0] as any
    expect(n.fill.kind).toBe('gradient')
    expect(n.fill.stops).toEqual([
      { pos: 0, color: '#000000' },
      { pos: 1, color: '#3F3F3F' },
    ])
    // without a resolvable background the parsed fill stays as fallback
    const rs2 = buildRenderSlide(
      { ...slide, elements: [el], decorations: [], background: undefined },
      deck.size,
      { fitWidthPx: 1280 },
    )
    expect((rs2.nodes[0] as any).fill).toEqual({ kind: 'solid', color: '#FFFFFF' })

    // useBgFill also applies inside groups
    const group: any = {
      id: 'g1',
      type: 'group',
      anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
      transform: el.transform,
      children: [{ ...el, id: 'bgsp2' }],
      childOffset: { x: 0, y: 0, cx: 914400, cy: 914400 },
    }
    const rs3 = buildRenderSlide(
      { ...slide, elements: [group], decorations: [], background: bg as any },
      deck.size,
      { fitWidthPx: 1280 },
    )
    const child = (rs3.nodes[0] as any).children[0]
    expect(child.fill.kind).toBe('gradient')
  })
})

describe('table row heights are absolute (napierone 0032: h="0" rows size to content)', () => {
  it('does not stretch a lone non-zero header row across the whole frame', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const slide = deck.slides[0]!
    const el: any = {
      id: 'tbl_1',
      type: 'table',
      anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
      transform: {
        offset: { x: 0, y: 0, cx: 8208912, cy: 4566280 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      colWidths: [4104456, 4104456],
      // header has an explicit height; the rest are h="0" (auto)
      rowHeights: [360040, 360040, 0],
      rows: [
        [{}, {}],
        [{}, {}],
        [{}, {}],
      ],
    }
    const rs = buildRenderSlide({ ...slide, elements: [el], decorations: [] }, deck.size, {
      fitWidthPx: 1280,
    })
    const node = rs.nodes[0] as any
    expect(node.type).toBe('table')
    // EMU heights convert absolutely: 360040 EMU of a 4566280 EMU frame ≈ 7.9% of box.h,
    // not box.h/2 (the old proportional split over the two non-zero rows)
    const expected = (360040 / 4566280) * node.box.h
    expect(node.gridY[1]).toBeCloseTo(expected, 1)
    expect(node.gridY[2] - node.gridY[1]).toBeCloseTo(expected, 1)
    // empty h="0" row stays at content height (zero here)
    expect(node.gridY[3]).toBeCloseTo(node.gridY[2], 1)
  })
})

describe('picture spPr fill backdrop (napierone 0027)', () => {
  it('threads a solid pic fill through to the render node', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const slide = deck.slides[0]!
    const el: any = {
      id: 'pic_1',
      type: 'picture',
      anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
      transform: {
        offset: { x: 0, y: 0, cx: 1000000, cy: 1000000 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      mediaRef: 'ppt/media/image1.png',
      opacity: 0.26,
      fill: { type: 'solid', color: '#000000' },
      duotone: ['#111111', '#EEEEEE'],
      clrChange: { from: '#000000', to: '#00000000' },
    }
    const rs = buildRenderSlide({ ...slide, elements: [el], decorations: [] }, deck.size, {
      fitWidthPx: 1280,
    })
    const node = rs.nodes[0] as any
    expect(node.type).toBe('picture')
    expect(node.fill).toEqual({ kind: 'solid', color: '#000000' })
    expect(node.opacity).toBeCloseTo(0.26, 5)
    expect(node.duotone).toEqual(['#111111', '#EEEEEE'])
    expect(node.clrChange).toEqual({ from: '#000000', to: '#00000000' })
  })
})

describe('metafile pictures get an opaque white backdrop (PlanS academy banner)', () => {
  it('sets bgColor for EMF data URLs and leaves raster images alone', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const slide = deck.slides[0]!
    const mk = (dataUrl: string): any => ({
      id: 'pic_m',
      type: 'picture',
      anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
      transform: {
        offset: { x: 0, y: 0, cx: 1000000, cy: 1000000 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      mediaRef: '',
      dataUrl,
    })
    const rs = buildRenderSlide(
      {
        ...slide,
        elements: [mk('data:image/x-emf;base64,AAAA'), mk('data:image/png;base64,AAAA')],
        decorations: [],
      },
      deck.size,
      { fitWidthPx: 1280 },
    )
    const [emf, png] = rs.nodes as any[]
    expect(emf.bgColor).toBe('#FFFFFF')
    expect(png.bgColor).toBeUndefined()
  })

  it('a clrChange metafile drops the white DC backing (the recolor sees through it)', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const slide = deck.slides[0]!
    const el: any = {
      id: 'pic_cc',
      type: 'picture',
      anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
      transform: {
        offset: { x: 0, y: 0, cx: 1000000, cy: 1000000 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      mediaRef: '',
      dataUrl: 'data:image/x-emf;base64,AAAA',
      clrChange: { from: '#FFFFFF', to: '#FFFFFF00' },
    }
    const rs = buildRenderSlide({ ...slide, elements: [el], decorations: [] }, deck.size, {
      fitWidthPx: 1280,
    })
    const node = rs.nodes[0] as any
    expect(node.bgColor).toBeUndefined()
    expect(node.clrChange).toEqual({ from: '#FFFFFF', to: '#FFFFFF00' })

    // non-white keys never touch the DC; an opaque white key recolors it
    const build = (clrChange: any) =>
      buildRenderSlide({ ...slide, elements: [{ ...el, clrChange }], decorations: [] }, deck.size, {
        fitWidthPx: 1280,
      }).nodes[0] as any
    expect(build({ from: '#FF0000', to: '#00FF0000' }).bgColor).toBe('#FFFFFF')
    expect(build({ from: '#FFFFFF', to: '#00FF00' }).bgColor).toBe('#00FF00')
    // partial alpha stays on the backing (canvas accepts #RRGGBBAA)
    expect(build({ from: '#FFFFFF', to: '#00FF0080' }).bgColor).toBe('#00FF0080')
  })
})
