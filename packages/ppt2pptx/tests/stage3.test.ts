/** Stage 3 fidelity checks: Escher geometry, StyleTextProp runs, embedded
 *  pictures and master shapes. Every input is built by synthetic-ppt, so the
 *  expected OOXML numbers are known exactly. */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { convertPptToPptx, parsePptPresentation } from '../src/index'
import { buildPptContainer, buildPptStreams, type SlideSpec } from './synthetic-ppt'

/** a real 1x1 PNG (the reader checks its signature) */
const PNG = new Uint8Array(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    'base64',
  ),
)

const pptBytes = (
  slides: SlideSpec[],
  options: Parameters<typeof buildPptStreams>[1] = {},
): Uint8Array => {
  const streams = buildPptStreams(slides, options)
  return buildPptContainer(streams.document, streams.currentUser)
}

async function parts(bytes: Uint8Array): Promise<JSZip> {
  return JSZip.loadAsync((await convertPptToPptx(bytes)).bytes)
}

const slideXml = async (bytes: Uint8Array, n = 1): Promise<string> =>
  (await parts(bytes)).file(`ppt/slides/slide${n}.xml`)!.async('string')

describe('stage 3: Escher geometry', () => {
  it('places a shape at its ChildAnchor and keeps its preset geometry and flips', async () => {
    const slides: SlideSpec[] = [
      {
        textSequences: [],
        shapes: [
          {
            rect: [1000, 800, 3000, 1600],
            shapeType: 3, // msosptEllipse
            text: 'Oval',
            flags: 0x40, // fFlipH
          },
        ],
      },
    ]
    const xml = await slideXml(pptBytes(slides))
    // 1000 master units x 1587.5 EMU = 1587500
    expect(xml).toContain('<a:off x="1587500" y="1270000"/>')
    expect(xml).toContain('<a:ext cx="3175000" cy="1270000"/>')
    expect(xml).toContain('<a:prstGeom prst="ellipse">')
    expect(xml).toContain('flipH="1"')
    expect(xml).toContain('<a:t>Oval</a:t>')
  })

  it('a shape with no anchor keeps the stage-1 stacked fallback', async () => {
    const xml = await slideXml(pptBytes([{ textSequences: [{ kind: 4, text: 'Fallback' }] }]))
    expect(xml).toContain('<a:off x="457200" y="457200"/>')
    expect(xml).toContain('<a:t>Fallback</a:t>')
  })

  it('the parse model exposes the anchor, shape type and properties', () => {
    const streams = buildPptStreams([
      {
        textSequences: [],
        shapes: [{ rect: [10, 20, 30, 40], shapeType: 1, fillColor: 0x0000ff, text: 'A' }],
      },
    ])
    const presentation = parsePptPresentation(streams.document, streams.currentUser)
    const shape = presentation.slides[0]!.shapes[0]!
    expect(shape.rect).toEqual({ left: 10, top: 20, right: 30, bottom: 40 })
    expect(shape.shapeType).toBe(1)
    expect(shape.props.get(0x0181)).toBe(0x0000ff)
  })
})

describe('stage 3: StyleTextProp formatting', () => {
  it('carries bold, italic, size, color and paragraph alignment into runs', async () => {
    const slides: SlideSpec[] = [
      {
        textSequences: [],
        shapes: [
          {
            rect: [100, 100, 4000, 1000],
            text: 'Styled title',
            kind: 0,
            paragraphAlign: 1,
            charRuns: [
              { count: 6, bold: true, size: 24, color: [255, 0, 0] },
              { count: 6, italic: true, underline: true, size: 18 },
            ],
          },
        ],
      },
    ]
    const xml = await slideXml(pptBytes(slides))
    expect(xml).toContain('algn="ctr"')
    expect(xml).toContain('sz="2400"')
    expect(xml).toContain('b="1"')
    expect(xml).toContain('<a:srgbClr val="FF0000"/>')
    expect(xml).toContain('sz="1800"')
    expect(xml).toContain('i="1" u="sng"')
    // the run boundaries keep every character, in order
    const texts = [...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]!)
    expect(texts.join('')).toBe('Styled title')
  })

  it('an outline-referenced shape still receives its SlideListWithText text', async () => {
    const slides: SlideSpec[] = [
      {
        textSequences: [{ kind: 0, text: 'From outline' }],
        shapes: [{ rect: [0, 0, 2000, 500], outlineRef: true }],
      },
    ]
    const xml = await slideXml(pptBytes(slides, { outlineTexts: true }))
    expect(xml).toContain('<a:t>From outline</a:t>')
    expect(xml).toContain('<a:off x="0" y="0"/>')
  })
})

describe('stage 3: embedded pictures', () => {
  it('writes the blip bytes, a pic shape, its relationship and content type', async () => {
    const slides: SlideSpec[] = [
      {
        textSequences: [],
        shapes: [
          { rect: [500, 500, 2500, 2000], shapeType: 75, blipIndex: 1 },
          { rect: [3000, 500, 5000, 1000], text: 'Caption' },
        ],
      },
    ]
    const bytes = pptBytes(slides, { images: [PNG] })
    const zip = await parts(bytes)
    const media = zip.file('ppt/media/image1.png')!
    expect(media).not.toBeNull()
    expect(new Uint8Array(await media.async('uint8array'))).toEqual(PNG)
    const xml = await slideXml(bytes)
    expect(xml).toContain('<p:pic>')
    expect(xml).toContain('r:embed="rId2"')
    expect(xml).toContain('<a:t>Caption</a:t>')
    const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string')
    expect(rels).toContain('Target="../media/image1.png"')
    const types = await zip.file('[Content_Types].xml')!.async('string')
    expect(types).toContain('<Default Extension="png" ContentType="image/png"/>')
  })

  it('an unsupported blip is omitted with a warning, never a crash', async () => {
    const streams = buildPptStreams([{ textSequences: [{ kind: 4, text: 'text' }] }], {
      images: [new Uint8Array([0, 1, 2, 3])], // not a PNG signature
    })
    const result = await convertPptToPptx(buildPptContainer(streams.document, streams.currentUser))
    expect(result.slideCount).toBe(1)
    expect(result.warnings.join(' ')).toMatch(/picture/i)
  })
})

describe('stage 3: the main master', () => {
  it('reproduces the master anchored shapes in slideMaster1.xml', async () => {
    const bytes = pptBytes([{ textSequences: [{ kind: 4, text: 'Body' }] }], {
      master: { shapes: [{ rect: [100, 100, 5500, 600], shapeType: 1, text: 'Master logo' }] },
    })
    const zip = await parts(bytes)
    const xml = await zip.file('ppt/slideMasters/slideMaster1.xml')!.async('string')
    expect(xml).toContain('<a:off x="158750" y="158750"/>')
    expect(xml).toContain('<a:t>Master logo</a:t>')
    // the slide keeps its own text, unchanged by the master walk
    expect(await slideXml(bytes)).toContain('<a:t>Body</a:t>')
  })
})

describe('stage 3: hostile bytes stay bounded', () => {
  it('a truncated drawing still converts or refuses cleanly', async () => {
    const bytes = pptBytes([
      {
        textSequences: [],
        shapes: [{ rect: [10, 10, 500, 500], shapeType: 3, text: 'Keep me' }],
      },
    ])
    for (const cut of [0.5, 0.75, 0.9]) {
      const damaged = bytes.slice(0, Math.floor(bytes.length * cut))
      await convertPptToPptx(damaged).catch(() => undefined)
    }
    const healthy = await convertPptToPptx(bytes)
    expect(healthy.slideCount).toBe(1)
  })

  it('garbage inside the drawing records does not poison the next deck', async () => {
    const streams = buildPptStreams([
      {
        textSequences: [],
        shapes: [{ rect: [10, 10, 500, 500], shapeType: 1, text: 'Shape' }],
      },
    ])
    const document = Uint8Array.from(streams.document)
    for (let at = 40; at < Math.min(document.length, 200); at += 3) {
      document[at] = (document[at] ?? 0) ^ 0xff
    }
    await convertPptToPptx(buildPptContainer(document, streams.currentUser)).catch(() => undefined)
    const healthy = await convertPptToPptx(buildPptContainer(streams.document, streams.currentUser))
    expect(healthy.slideCount).toBe(1)
  })
})
