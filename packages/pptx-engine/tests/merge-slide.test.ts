/**
 * mergeSlideFromPptx integration tests (mirrors html→pptx per-slide conversion + deck-level merge):
 * each slide is converted to its own single-slide pptx and merged into an existing deck,
 * without re-converting earlier slides. Uses pptxgenjs to generate real single-slide pptx
 * (same library and structure as the html-pipeline convert-worker), running the real
 * openPptx → mergeSlideFromPptx → savePptx → openPptx chain with no mocks.
 */
import { describe, it, expect } from 'vitest'
import PptxGenJS from 'pptxgenjs'
import { openPptx, savePptx, mergeSlideFromPptx, setSlideLayout } from '../src/index'

// 1x1 red-dot PNG (base64)
const RED_DOT =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

async function onePagePptx(text: string, withImage = false): Promise<Uint8Array> {
  const p = new PptxGenJS()
  p.defineLayout({ name: 'W', width: 13.333, height: 7.5 })
  p.layout = 'W'
  const s = p.addSlide()
  s.addText(text, { x: 1, y: 1, w: 8, h: 1, fontSize: 32 })
  if (withImage) s.addImage({ data: 'image/png;base64,' + RED_DOT, x: 1, y: 3, w: 2, h: 2 })
  const buf = (await p.write({ outputType: 'nodebuffer' })) as Buffer
  return new Uint8Array(buf)
}

describe('mergeSlideFromPptx', () => {
  it('merges a single-slide pptx into an existing deck, slide count grows and content is kept', async () => {
    const base = await openPptx(await onePagePptx('PAGE_ONE'))
    expect(base.deck.slides.length).toBe(1)

    const s2 = await mergeSlideFromPptx(base, await onePagePptx('PAGE_TWO'))
    expect(s2).not.toBeNull()
    expect(base.deck.slides.length).toBe(2)

    const s3 = await mergeSlideFromPptx(base, await onePagePptx('PAGE_THREE'))
    expect(s3).not.toBeNull()
    expect(base.deck.slides.length).toBe(3)

    // Reopen after save: all three slides present, in the right order
    const reopened = await openPptx(await savePptx(base))
    expect(reopened.deck.slides.length).toBe(3)
    // pptxgenjs textboxes parse as text-bearing shapes (not pure text); accept both
    const texts = reopened.deck.slides.map((sl) =>
      sl.elements
        .filter((el) => el.type === 'text' || el.type === 'shape')
        .map(
          (el) =>
            (
              el as { text?: { paragraphs: Array<{ runs: Array<{ text: string }> }> } }
            ).text?.paragraphs
              ?.flatMap((pg) => pg.runs.map((r) => r.text))
              .join('') ?? '',
        )
        .join(' '),
    )
    expect(texts[0]).toContain('PAGE_ONE')
    expect(texts[1]).toContain('PAGE_TWO')
    expect(texts[2]).toContain('PAGE_THREE')
  })

  it('merging a slide with an image: media bytes moved in, rIds do not clash', async () => {
    const base = await openPptx(await onePagePptx('COVER', true))
    const before = [...base.archive.entries.keys()].filter((k) => /ppt\/media\//.test(k)).length
    expect(before).toBeGreaterThanOrEqual(1)

    // Merge another slide with an image — both media are originally named image-1-1.png, so remap must avoid overwriting
    await mergeSlideFromPptx(base, await onePagePptx('SECOND', true))
    expect(base.deck.slides.length).toBe(2)

    const reopened = await openPptx(await savePptx(base))
    expect(reopened.deck.slides.length).toBe(2)
    const media = [...reopened.archive.entries.keys()].filter((k) => /ppt\/media\//.test(k))
    // Both images exist independently (neither overwritten)
    expect(media.length).toBeGreaterThanOrEqual(2)
    // The second slide does contain a picture element
    const picCount = reopened.deck.slides[1]!.elements.filter((el) => el.type === 'picture').length
    expect(picCount).toBeGreaterThanOrEqual(1)
  })

  it('saved output opens with a standard zip and has the right number of slide parts', async () => {
    const base = await openPptx(await onePagePptx('A'))
    await mergeSlideFromPptx(base, await onePagePptx('B'))
    const bytes = await savePptx(base)
    const reopened = await openPptx(bytes)
    const slideParts = [...reopened.archive.entries.keys()].filter((k) =>
      /^ppt\/slides\/slide\d+\.xml$/.test(k),
    )
    expect(slideParts.length).toBe(2)
    // presentation.xml's sldId list should also have 2 entries
    const pres = reopened.archive.readText('ppt/presentation.xml') ?? ''
    const sldIds = [...pres.matchAll(/<p:sldId\b/g)].length
    expect(sldIds).toBe(2)
  })
})

describe('mergeSlideFromPptx layout anchor', () => {
  const LAYOUT_TYPE =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'

  /** pptxgenjs ships one layout; clone it so the deck has two to tell apart */
  async function twoLayoutDeck(): Promise<Uint8Array> {
    const JSZip = (await import('jszip')).default
    const zip = await JSZip.loadAsync(await onePagePptx('FIRST'))
    const layout = await zip.file('ppt/slideLayouts/slideLayout1.xml')!.async('string')
    const rels = await zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels')!.async('string')
    zip.file('ppt/slideLayouts/slideLayout2.xml', layout)
    zip.file('ppt/slideLayouts/_rels/slideLayout2.xml.rels', rels)
    const types = await zip.file('[Content_Types].xml')!.async('string')
    zip.file(
      '[Content_Types].xml',
      types.replace(
        '</Types>',
        '<Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/></Types>',
      ),
    )
    return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))
  }

  const layoutOf = (opened: Awaited<ReturnType<typeof openPptx>>, slide: { path: string }) =>
    [...opened.archive.readRels(slide.path).values()].find((r) => r.type === LAYOUT_TYPE)?.target

  it('takes the layout of the slide it is given instead of the last slide', async () => {
    const base = await openPptx(await twoLayoutDeck())
    const second = await mergeSlideFromPptx(base, await onePagePptx('LAST'))
    expect(second).not.toBeNull()
    expect(setSlideLayout(base, 1, 'ppt/slideLayouts/slideLayout2.xml')).not.toBeNull()
    expect(layoutOf(base, base.deck.slides[0])).toBe('../slideLayouts/slideLayout1.xml')
    expect(layoutOf(base, base.deck.slides[1])).toBe('../slideLayouts/slideLayout2.xml')

    const byDefault = await mergeSlideFromPptx(base, await onePagePptx('DEFAULT'))
    expect(layoutOf(base, byDefault!)).toBe('../slideLayouts/slideLayout2.xml')

    const anchored = await mergeSlideFromPptx(base, await onePagePptx('ANCHORED'), {
      layoutFrom: base.deck.slides[0],
    })
    expect(layoutOf(base, anchored!)).toBe('../slideLayouts/slideLayout1.xml')
  })
})
