/**
 * P9 B: a page-covering background stack (wallpaper image / gradient shading)
 * is rendered to a bitmap at extraction and rebuilt as a full-page behindDoc
 * float pinned to the page box — white text on gradient covers stays visible.
 */
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { convertPdfToDocx } from '../src'
import { encodeRgbaPng, extractPage, withPdfDocument } from '../src/extract'
import type { ImageBlock, IrPage, Line, Span, TextBlock } from '../src/ir'
import { rebuildDocx } from '../src/rebuild'
import {
  buildBgImagePdf,
  buildBuriedTemplatePdf,
  buildLatinPdf,
  buildPhotoCardPdf,
  buildWhiteWashPdf,
} from './helpers/fixtures'
import { loadPdfium } from './helpers/wasm'

describe('background stack extraction (P9 B)', () => {
  it('a full-page wallpaper image becomes bgRender and leaves the image stream', async () => {
    const m = await loadPdfium()
    const page = withPdfDocument(m, await buildBgImagePdf(), (doc) => extractPage(m, doc, 0))

    expect(page.scanned).toBe(false)
    expect(page.degraded).toBe(false)
    expect(page.bgRender).toBeDefined()
    // the wallpaper must not ALSO flow as a page-sized inline/float image
    expect(page.images).toHaveLength(0)
    // full page at the background raster scale
    expect(page.bgRender!.pixelWidth).toBeGreaterThan(600)
    expect(page.chars.length).toBeGreaterThan(10)
  })

  it('a plain white text page carries no bgRender', async () => {
    const m = await loadPdfium()
    const page = withPdfDocument(m, await buildLatinPdf(), (doc) => extractPage(m, doc, 0))
    expect(page.bgRender).toBeUndefined()
  })

  it('a white full-page wash never activates the stack or swallows vectors', async () => {
    const m = await loadPdfium()
    const page = withPdfDocument(m, await buildWhiteWashPdf(), (doc) => extractPage(m, doc, 0))
    expect(page.bgRender).toBeUndefined()
    // the stroked box after the wash must survive path extraction
    expect(page.paths.some((path) => path.stroked)).toBe(true)
    const text = page.chars.map((c) => String.fromCodePoint(c.code)).join('')
    expect(text).toContain('Content over a plain white wash.')
  })

  it('a template background buried mid-stack bakes into bgRender (P16 B)', async () => {
    const m = await loadPdfium()
    const page = withPdfDocument(m, await buildBuriedTemplatePdf(), (doc) => extractPage(m, doc, 0))
    expect(page.scanned).toBe(false)
    expect(page.degraded).toBe(false)
    // wallpaper + wash composite into the background bitmap…
    expect(page.bgRender).toBeDefined()
    expect(page.images).toHaveLength(0)
    // …taking the junk furniture under them along; the real content stays
    const text = page.chars.map((c) => String.fromCodePoint(c.code)).join('')
    expect(text).toContain('Real title above the wash')
    expect(text).toContain('Real body keeps flowing here.')
    expect(text).not.toContain('template junk')
  })

  it('a large-but-not-full-bleed card never extends the stack (P16 B guard)', async () => {
    const m = await loadPdfium()
    const page = withPdfDocument(m, await buildPhotoCardPdf(), (doc) => extractPage(m, doc, 0))
    // the full-bleed photo (and the junk under it) bake…
    expect(page.bgRender).toBeDefined()
    expect(page.images).toHaveLength(0)
    const text = page.chars.map((c) => String.fromCodePoint(c.code)).join('')
    expect(text).not.toContain('buried under the photo')
    // …but the ~92%×90% white card and its text stay in the flow
    expect(text).toContain('Card title stays in the flow')
    expect(text).toContain('Card body line one keeps flowing.')
    expect(page.paths.some((p) => p.filled && p.fillColor === 'FFFFFF')).toBe(true)
  })
})

describe('background rebuild (P9 B)', () => {
  const span = (text: string): Span => ({
    text,
    box: { x0: 72, y0: 690, x1: 300, y1: 700 },
    fontSize: 12,
    fontFamily: 'Helvetica',
    bold: false,
    italic: false,
    color: 'FFFFFF',
    dir: 'ltr',
    script: 'latin',
  })
  const lineOf = (s: Span): Line => ({
    spans: [s],
    box: s.box,
    baseline: s.box.y0 + 2,
    endsWithHyphen: false,
  })
  const blockOf = (text: string): TextBlock => {
    const s = span(text)
    return {
      kind: 'text',
      lines: [lineOf(s)],
      box: s.box,
      align: 'left',
      firstLineIndentPt: 0,
      dir: 'ltr',
    }
  }
  const bgOf = (rgb: [number, number, number]) => {
    const px = new Uint8Array(4 * 4 * 4)
    for (let i = 0; i < 16; i++) px.set([...rgb, 255], i * 4)
    return {
      data: encodeRgbaPng(px, 4, 4),
      mime: 'image/png' as const,
      pixelWidth: 4,
      pixelHeight: 4,
    }
  }
  const irPage = (over: Partial<IrPage>): IrPage => ({
    index: 0,
    widthPt: 960,
    heightPt: 540,
    rotation: 0,
    blocks: [],
    degraded: false,
    scanned: false,
    hasStructTree: false,
    ...over,
  })

  it('bgRender lands as a page-anchored behindDoc float sized to the page', async () => {
    const docx = await rebuildDocx([
      irPage({ blocks: [blockOf('white title')], bgRender: bgOf([40, 90, 200]) }),
    ])
    const parsed = await parseDocx(docx)
    const xml = parsed.internal.documentXml
    expect(xml).toContain('behindDoc="1"')
    expect(xml).toContain(
      '<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>',
    )
    expect(xml).toContain(
      '<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>',
    )
    // full page size: 960pt × 96/72 = 1280px → 1280 × 9525 EMU
    expect(xml).toContain(`cx="${1280 * 9525}"`)
    // the text still flows as a normal paragraph
    const texts = parsed.blocks.map((b) => (b.runs ?? []).map((r) => r.text).join(''))
    expect(texts).toContain('white title')
  })

  it('pages without bgRender emit no page-anchored background', async () => {
    const docx = await rebuildDocx([irPage({ blocks: [blockOf('plain page')] })])
    const parsed = await parseDocx(docx)
    expect(parsed.internal.documentXml).not.toContain('relativeFrom="page"')
  })

  it('per-page: only the cover page gets its background, later pages stay clean', async () => {
    const docx = await rebuildDocx([
      irPage({ blocks: [blockOf('cover')], bgRender: bgOf([40, 90, 200]) }),
      irPage({ index: 1, blocks: [blockOf('body page')] }),
    ])
    const parsed = await parseDocx(docx)
    const xml = parsed.internal.documentXml
    expect(xml.match(/behindDoc="1"/g)).toHaveLength(1)
  })

  it('page-pinned anchors carry z-ordered relativeHeight (P16 A)', async () => {
    // a full-page image float drawn FIRST in the pdf (z=0) must stack UNDER a
    // card panel drawn later (z=5), no matter the emission order — otherwise
    // the page background paints over the card and its light text vanishes
    const imgOf = (z: number): ImageBlock => ({
      kind: 'image',
      box: { x0: 0, y0: 0, x1: 960, y1: 540 },
      data: bgOf([220, 220, 220]).data,
      mime: 'image/png',
      pixelWidth: 4,
      pixelHeight: 4,
      float: { wrap: 'behind', xOffsetPt: 0 },
      z,
    })
    const panel: ImageBlock = {
      kind: 'image',
      box: { x0: 100, y0: 100, x1: 400, y1: 250 },
      data: bgOf([0, 32, 96]).data,
      mime: 'image/png',
      pixelWidth: 2,
      pixelHeight: 2,
      float: { wrap: 'behind', xOffsetPt: 100 },
      z: 5,
    }
    const body = blockOf('white text on the card')
    const docx = await rebuildDocx([
      irPage({
        blocks: [body],
        bgPanels: [panel],
        sections: [
          {
            box: { x0: 0, y0: 0, x1: 960, y1: 540 },
            columns: [{ box: { x0: 0, y0: 0, x1: 960, y1: 540 }, blocks: [imgOf(0), body] }],
            gutterWidthsPt: [],
            dir: 'ltr',
          },
        ],
      }),
    ])
    const parsed = await parseDocx(docx)
    const xml = parsed.internal.documentXml
    const heights = [...xml.matchAll(/relativeHeight="(\d+)"/g)].map((m) => Number(m[1]))
    expect(heights).toHaveLength(2)
    // anchors are emitted in z order (image z=0 first, panel z=5 second) with
    // strictly increasing relativeHeight
    expect(heights[1]).toBeGreaterThan(heights[0])
    // the full-page image (z=0) sits at the lower height: find each embed's
    // anchor order — image extent is page-wide (960pt*96/72*9525 EMU)
    const anchors = [...xml.matchAll(/<wp:anchor.*?<\/wp:anchor>/gs)]
    expect(anchors[0]![0]).toContain(`cx="${Math.round((960 * 96) / 72) * 9525}"`)
  })

  it('end-to-end: wallpaper PDF converts with a behindDoc page background', async () => {
    const m = await loadPdfium()
    const res = await convertPdfToDocx(await buildBgImagePdf(), { pdfium: m })
    const parsed = await parseDocx(res.docx)
    const xml = parsed.internal.documentXml
    expect(xml).toContain('behindDoc="1"')
    expect(xml).toContain('relativeFrom="page"')
    const texts = parsed.blocks.map((b) => (b.runs ?? []).map((r) => r.text).join(''))
    expect(texts.some((t) => t.includes('Title over the wallpaper'))).toBe(true)
  })
})
