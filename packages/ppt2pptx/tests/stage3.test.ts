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

/** The type code RT_StyleTextPropAtom is 0x0FA1 (0x0FA5 is
 *  TextPFExceptionAtom). A synthetic builder can share a wrong code with the
 *  reader and still pass every test, so the real deck is the only oracle. */
describe('a real PowerPoint-authored deck: style atoms', () => {
  it('reads the sample StyleTextPropAtom run tables', async () => {
    const { readFile } = await import('node:fs/promises')
    const { fileURLToPath } = await import('node:url')
    const CFB = await import('cfb')
    const sample = fileURLToPath(
      new URL('../../file-parse/tests/fixtures/legacy-sample.ppt', import.meta.url),
    )
    const cfb = CFB.read(await readFile(sample), { type: 'buffer' })
    const read = (name: string): Uint8Array => {
      const content = CFB.find(cfb, name)?.content as ArrayLike<number> | undefined
      return content ? Uint8Array.from(Array.from(content)) : new Uint8Array(0)
    }
    const presentation = parsePptPresentation(read('PowerPoint Document'), read('Current User'))
    const styles = presentation.slides.flatMap((slide) =>
      slide.shapes.flatMap((shape) => (shape.text ? [shape.text.style] : [])),
    )
    expect(styles.length).toBeGreaterThan(0)
    expect(styles.reduce((n, style) => n + style.characters.length, 0)).toBeGreaterThan(0)
    expect(styles.reduce((n, style) => n + style.paragraphs.length, 0)).toBeGreaterThan(0)
  })
})

describe('stage 3b: rotation, groups, placeholders, deck defaults and metafiles', () => {
  it('carries OPT rotation as a DrawingML rot value', async () => {
    const slides: SlideSpec[] = [
      { textSequences: [], shapes: [{ rect: [0, 0, 1000, 1000], text: 'Turned', rotation: 45 }] },
    ]
    const xml = await slideXml(pptBytes(slides))
    expect(xml).toContain('rot="2700000"') // 45 degrees x 60000
  })

  it('maps a grouped shape anchor from group coordinates to the slide', async () => {
    const slides: SlideSpec[] = [
      {
        textSequences: [],
        shapes: [
          {
            group: {
              bounds: [0, 0, 1000, 1000],
              rect: [1000, 1000, 3000, 3000],
              children: [{ rect: [0, 0, 500, 500], text: 'Inside group' }],
            },
          },
        ],
      },
    ]
    const xml = await slideXml(pptBytes(slides))
    // child [0,0,500,500] x 2 through outer [1000,1000,3000,3000]
    expect(xml).toContain('<a:off x="1587500" y="1587500"/>')
    expect(xml).toContain('<a:ext cx="1587500" cy="1587500"/>')
    expect(xml).toContain('<a:t>Inside group</a:t>')
    expect(xml).not.toContain('name="group"')
  })

  it('writes the OEPlaceholderAtom id as a DrawingML placeholder', async () => {
    const slides: SlideSpec[] = [
      {
        textSequences: [],
        shapes: [
          { rect: [0, 0, 2000, 500], text: 'Title', placeholderId: 13 },
          { rect: [0, 600, 2000, 1500], text: 'Body', placeholderId: 14 },
        ],
      },
    ]
    const xml = await slideXml(pptBytes(slides))
    expect(xml).toContain('<p:ph type="title"/>')
    expect(xml).toContain('<p:ph type="body"/>')
  })

  it('inherits the master default size and the FontCollection typeface', async () => {
    const slides: SlideSpec[] = [
      { textSequences: [], shapes: [{ rect: [0, 0, 2000, 500], text: 'Plain text' }] },
    ]
    const xml = await slideXml(
      pptBytes(slides, {
        fonts: ['Calibri'],
        textMasterStyles: [{ textType: 1, sizePt: 28, fontRef: 0 }],
      }),
    )
    expect(xml).toContain('sz="2800"')
    expect(xml).toContain('<a:latin typeface="Calibri"/>')
  })

  it('writes an embedded EMF blip as a picture part', async () => {
    const emf = new Uint8Array([1, 0, 0, 0, 0x20, 0x45, 0x4d, 0x46, 0, 0])
    const slides: SlideSpec[] = [
      { textSequences: [], shapes: [{ rect: [0, 0, 2000, 1000], shapeType: 75, blipIndex: 1 }] },
    ]
    const bytes = pptBytes(slides, { metafiles: [{ kind: 'emf', data: emf }] })
    const zip = await parts(bytes)
    const media = zip.file('ppt/media/image1.emf')!
    expect(media).not.toBeNull()
    expect(new Uint8Array(await media.async('uint8array'))).toEqual(emf)
    expect(await slideXml(bytes)).toContain('<p:pic>')
    const types = await zip.file('[Content_Types].xml')!.async('string')
    expect(types).toContain('<Default Extension="emf" ContentType="image/x-emf"/>')
  })
})

describe('a real PowerPoint-authored deck: deck-wide defaults', () => {
  it('reads the FontCollection and TextMasterStyleAtom levels', async () => {
    const { readFile } = await import('node:fs/promises')
    const { fileURLToPath } = await import('node:url')
    const CFB = await import('cfb')
    const sample = fileURLToPath(
      new URL('../../file-parse/tests/fixtures/legacy-sample.ppt', import.meta.url),
    )
    const cfb = CFB.read(await readFile(sample), { type: 'buffer' })
    const read = (name: string): Uint8Array => {
      const content = CFB.find(cfb, name)?.content as ArrayLike<number> | undefined
      return content ? Uint8Array.from(Array.from(content)) : new Uint8Array(0)
    }
    const presentation = parsePptPresentation(read('PowerPoint Document'), read('Current User'))
    expect(presentation.masterStyles.typefaces.length).toBeGreaterThan(0)
    expect(presentation.masterStyles.levels.length).toBeGreaterThan(0)
    const xml = await slideXml(new Uint8Array(await readFile(sample)))
    expect(xml).toContain('<a:latin typeface=')
  })
})
