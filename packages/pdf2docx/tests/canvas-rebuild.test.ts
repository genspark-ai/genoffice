/** P19 canvas assembly: canvas pages emit framePr containers in per-page sections. */
import { describe, expect, it } from 'vitest'
import type { ImageBlock, IrPage, Line, Span, TableBlock, TextBlock } from '../src/ir'
import { rebuildDocx } from '../src/rebuild'

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

function blockAt(
  text: string,
  x0: number,
  top: number,
  widthPt: number,
  fontSize = 20,
  over: Partial<TextBlock> = {},
): TextBlock {
  const s = span(text, { fontSize, box: { x0, y0: top - fontSize, x1: x0 + widthPt, y1: top } })
  const line: Line = {
    spans: [s],
    box: s.box,
    baseline: top - fontSize * 0.8,
    endsWithHyphen: false,
  }
  return {
    kind: 'text',
    lines: [line],
    box: s.box,
    align: 'left',
    firstLineIndentPt: 0,
    dir: 'ltr',
    ...over,
  }
}

function canvasPage(blocks: IrPage['blocks'], index = 0, over: Partial<IrPage> = {}): IrPage {
  return {
    index,
    widthPt: 960,
    heightPt: 540,
    rotation: 0,
    blocks,
    degraded: false,
    scanned: false,
    hasStructTree: false,
    canvas: true,
    ...over,
  }
}

const docXml = async (docx: Uint8Array): Promise<string> => {
  const { default: JSZip } = await import('jszip')
  const zip = await JSZip.loadAsync(docx)
  return zip.file('word/document.xml')!.async('string')
}

describe('canvas page assembly (P19)', () => {
  it('emits text blocks as page-anchored framePr paragraphs with measured coordinates', async () => {
    // title at (60, top edge y1=500), body at (120, y1=300)
    const title = blockAt('Slide Title', 60, 500, 400, 32)
    const body = blockAt('body text', 120, 300, 300, 16)
    const xml = await docXml(await rebuildDocx([canvasPage([title, body])]))
    const frames = xml.match(/<w:framePr [^>]*\/>/g) ?? []
    expect(frames).toHaveLength(2)
    for (const f of frames) {
      expect(f).toContain('w:wrap="none"')
      expect(f).toContain('w:vAnchor="page"')
      expect(f).toContain('w:hAnchor="page"')
    }
    // title: x = 60pt = 1200 twips, y = (540 - 500)pt = 800 twips
    expect(frames[0]).toContain('w:x="1200"')
    expect(frames[0]).toContain('w:y="800"')
    // body: x = 120pt = 2400 twips, y = (540 - 300)pt = 4800 twips
    expect(frames[1]).toContain('w:x="2400"')
    expect(frames[1]).toContain('w:y="4800"')
    // exact line pitch survives inside the frame
    expect(xml).toContain('w:lineRule="exact"')
    // no spacing-before chain on canvas pages (position is absolute)
    expect(xml).not.toMatch(/<w:spacing [^>]*w:before="[1-9]/)
  })

  it('closes one section per canvas page (consecutive same-size pages do not merge)', async () => {
    const p0 = canvasPage([blockAt('one', 60, 500, 200)], 0)
    const p1 = canvasPage([blockAt('two', 60, 500, 200)], 1)
    const p2 = canvasPage([blockAt('three', 60, 500, 200)], 2)
    const xml = await docXml(await rebuildDocx([p0, p1, p2]))
    // 2 mid-body sectPr (closing pages 1 and 2) + the trailing body sectPr
    const sectPrs = xml.match(/<w:sectPr/g) ?? []
    expect(sectPrs.length).toBe(3)
    // canvas pages must not rely on pageBreakBefore (LO misplaces frames across
    // in-section page breaks — P19 probe)
    expect(xml).not.toContain('<w:pageBreakBefore/>')
  })

  it('gives right/center-aligned frames their wrap padding on the anchored side', async () => {
    const centered = blockAt('centered', 400, 400, 160, 20, { align: 'center' })
    const righted = blockAt('righted', 700, 300, 200, 20, { align: 'right' })
    const xml = await docXml(await rebuildDocx([canvasPage([centered, righted])]))
    const frames = xml.match(/<w:framePr [^>]*\/>/g) ?? []
    // centered: pad splits both sides — x shifts left of 400pt (8000 twips)
    const cx = parseInt(/w:x="(\d+)"/.exec(frames[0]!)![1]!, 10)
    expect(cx).toBeLessThan(8000)
    // right-aligned: pad extends left — x shifts left of 700pt (14000 twips)
    const rx = parseInt(/w:x="(\d+)"/.exec(frames[1]!)![1]!, 10)
    expect(rx).toBeLessThan(14000)
    // both frames are wider than the measured box (wrap headroom)
    const cw = parseInt(/w:w="(\d+)"/.exec(frames[0]!)![1]!, 10)
    expect(cw).toBeGreaterThan(160 * 20)
  })

  it('emits canvas tables as floating tables at their measured position', async () => {
    const cellBlock = blockAt('cell', 100, 395, 80, 12)
    const table: TableBlock = {
      kind: 'table',
      box: { x0: 100, y0: 200, x1: 500, y1: 400 },
      colWidthsPt: [200, 200],
      rows: [
        [
          { box: { x0: 100, y0: 200, x1: 300, y1: 400 }, gridSpan: 1, blocks: [cellBlock] },
          { box: { x0: 300, y0: 200, x1: 500, y1: 400 }, gridSpan: 1, blocks: [] },
        ],
      ],
    }
    const xml = await docXml(await rebuildDocx([canvasPage([blockAt('t', 60, 500, 100), table])]))
    expect(xml).toContain('<w:tblpPr')
    // x = 100pt = 2000 twips; y = (540 - 400)pt = 2800 twips
    expect(xml).toContain('w:tblpX="2000"')
    expect(xml).toContain('w:tblpY="2800"')
    expect(xml).toContain('<w:tblOverlap w:val="overlap"/>')
  })

  it('keeps floats/backgrounds pinned and list markers literal inside frames (P20)', async () => {
    const wallpaper: ImageBlock = {
      kind: 'image',
      box: { x0: 0, y0: 0, x1: 960, y1: 540 },
      data: new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
      mime: 'image/png',
      pixelWidth: 2,
      pixelHeight: 2,
      float: { wrap: 'behind', xOffsetPt: 0 },
      z: 0,
    }
    const item = blockAt('first point', 100, 350, 300, 18, {
      list: { kind: 'bullet', level: 0, marker: '•' },
    })
    const xml = await docXml(await rebuildDocx([canvasPage([wallpaper, item])]))
    expect(xml).toContain('behindDoc="1"')
    // w:numPr's hanging indent would eat into the fixed frame width and wrap
    // the measured line — the marker rides as literal text instead (P20)
    expect(xml).not.toContain('<w:numPr>')
    expect(xml).toContain('<w:t xml:space="preserve">• </w:t>')
    expect(xml).toContain('first point')
    expect(xml).toContain('<w:framePr')
  })

  it('mixes canvas and flow pages: flow pages keep their break/budget path', async () => {
    const flowPage: IrPage = {
      index: 1,
      widthPt: 612,
      heightPt: 792,
      rotation: 0,
      blocks: [blockAt('plain paragraph on a flow page', 72, 700, 300, 12)],
      degraded: false,
      scanned: false,
      hasStructTree: false,
    }
    const xml = await docXml(
      await rebuildDocx([canvasPage([blockAt('slide', 60, 500, 300)], 0), flowPage]),
    )
    // one mid-body sectPr closes the canvas page; the flow page has no framePr
    const afterSect = xml.slice(xml.indexOf('</w:sectPr>'))
    expect(afterSect).not.toContain('<w:framePr')
    // the flow page's own pgSz differs (612pt portrait)
    expect(xml).toContain('w:w="12240"')
  })
})
