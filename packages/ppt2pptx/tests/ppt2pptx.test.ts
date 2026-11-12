import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import * as CFB from 'cfb'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { convertPptToPptx, parsePptPresentation } from '../src/index'
import { officeArtColor } from '../src/escher'
import { buildPersistChain } from '../src/ppt'
import {
  type BuildOptions,
  buildPatternDib,
  buildPicturesStream,
  buildPptContainer,
  buildPptStreams,
  type SlideSpec,
  type TextSpec,
} from './synthetic-ppt'

const pptBytesOf = (streams: { document: Uint8Array; currentUser: Uint8Array }): Uint8Array =>
  buildPptContainer(streams.document, streams.currentUser)

/** The in-repo deck saved by a real PowerPoint-compatible producer. */
const LEGACY_SAMPLE = fileURLToPath(
  new URL('../../file-parse/tests/fixtures/legacy-sample.ppt', import.meta.url),
)

/** the presentation streams of a container, for the model-level assertions */
function streamsOf(bytes: Uint8Array): { document: Uint8Array; currentUser: Uint8Array } {
  const container = CFB.read(bytes as unknown as Buffer, { type: 'buffer' })
  const read = (name: string): Uint8Array => {
    const content = CFB.find(container, name)?.content as ArrayLike<number> | undefined
    // cfb hands big streams back as a Buffer and small ones as a number array
    return content ? Uint8Array.from(Array.from(content)) : new Uint8Array(0)
  }
  return { document: read('PowerPoint Document'), currentUser: read('Current User') }
}

const twoSlides: SlideSpec[] = [
  {
    textSequences: [
      { kind: 0, text: 'Quarterly Review' },
      { kind: 1, text: 'Revenue\nUp 12%' },
    ],
  },
  { textSequences: [{ kind: 4, text: 'Next steps' }] },
]

async function slideTextsOf(bytes: Uint8Array): Promise<string[][]> {
  const out = await convertPptToPptx(bytes)
  const zip = await JSZip.loadAsync(out.bytes)
  const texts: string[][] = []
  for (let s = 0; s < out.slideCount; s++) {
    const xml = await zip.file(`ppt/slides/slide${s + 1}.xml`)!.async('string')
    texts.push([...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]!))
  }
  return texts
}

async function presentationXmlOf(bytes: Uint8Array): Promise<string> {
  const out = await convertPptToPptx(bytes)
  const zip = await JSZip.loadAsync(out.bytes)
  return zip.file('ppt/presentation.xml')!.async('string')
}

describe('convertPptToPptx: editable text', () => {
  it('carries every text sequence as an editable run, in slide order', async () => {
    const texts = await slideTextsOf(pptBytesOf(buildPptStreams(twoSlides)))
    expect(texts).toHaveLength(2)
    expect(texts[0]).toEqual(['Quarterly Review', 'Revenue', 'Up 12%'])
    expect(texts[1]).toEqual(['Next steps'])
  })

  it('decodes CJK through the UTF-16 atom and latin through the bytes atom', async () => {
    const specs: SlideSpec[] = [
      {
        textSequences: [
          { kind: 0, text: '\u5b63\u5ea6\u56de\u987e' },
          { kind: 4, text: 'He said \u0092hi\u0092' },
        ],
      },
    ]
    const texts = await slideTextsOf(pptBytesOf(buildPptStreams(specs)))
    // the cp1252 smart quotes in a TextBytesAtom are named on the way out
    expect(texts[0]).toEqual(['\u5b63\u5ea6\u56de\u987e', 'He said ’hi’'])
    expect(await presentationXmlOf(pptBytesOf(buildPptStreams(specs)))).toContain(
      'cx="9144000" cy="6858000"', // 5760 × 4320 master units = 10 × 7.5 in
    )
  })

  it('decodes the whole cp1252 punctuation block of a TextBytesAtom', async () => {
    const specs: SlideSpec[] = [
      {
        textSequences: [
          { kind: 4, atom: 'bytes', text: '\u0080 price \u0085 \u0087 \u0099' },
          { kind: 4, atom: 'bytes', text: 'caf\u00e9 \u009cuvre' },
        ],
      },
    ]
    const texts = await slideTextsOf(pptBytesOf(buildPptStreams(specs)))
    // 0x80 0x85 0x87 0x99 0x9C are C1 controls in latin1, punctuation in cp1252
    expect(texts[0]).toEqual(['\u20ac price \u2026 \u2021 \u2122', 'caf\u00e9 \u0153uvre'])
  })

  it('reads the slide size from the DocumentAtom (576 master units per inch)', async () => {
    const specs: SlideSpec[] = [{ textSequences: [{ kind: 4, text: 'widescreen' }] }]
    const streams = buildPptStreams(specs, { slideSize: [7680, 4320] })
    const xml = await presentationXmlOf(buildPptContainer(streams.document, streams.currentUser))
    expect(xml).toContain('cx="12192000" cy="6858000"') // 13.33 × 7.5 in
  })

  it('gives every shape in a spTree a unique cNvPr id', async () => {
    const out = await convertPptToPptx(pptBytesOf(buildPptStreams(twoSlides)))
    const zip = await JSZip.loadAsync(out.bytes)
    const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
    const ids = [...xml.matchAll(/<p:cNvPr id="(\d+)"/g)].map((m) => Number(m[1]))
    expect(ids).toEqual([1, 2, 3])
    expect(new Set(ids).size).toBe(ids.length)
  })

  it('a second save generation restates the slides and the stale ones vanish', async () => {
    const second: SlideSpec[] = [
      { textSequences: [{ kind: 0, text: 'Rewritten deck' }] },
      { textSequences: [{ kind: 4, text: 'New slide' }] },
    ]
    const texts = await slideTextsOf(
      pptBytesOf(buildPptStreams(twoSlides, { secondGeneration: second })),
    )
    expect(texts).toHaveLength(2)
    expect(texts[0]).toEqual(['Rewritten deck'])
    expect(texts[1]).toEqual(['New slide'])
    expect(texts.join(' ')).not.toContain('Quarterly Review')
  })

  it('reads outline text from SlideListWithText without duplicating it', async () => {
    const texts = await slideTextsOf(pptBytesOf(buildPptStreams(twoSlides, { outlineTexts: true })))
    expect(texts).toHaveLength(2)
    expect(texts[0]).toEqual(['Quarterly Review', 'Revenue', 'Up 12%'])
    expect(texts[1]).toEqual(['Next steps'])
  })

  it('normalizes the PowerPoint runes into plain newlines', async () => {
    const specs: SlideSpec[] = [
      {
        textSequences: [{ kind: 1, text: 'first\u000bsecond\u001ebullet text' }],
      },
    ]
    const texts = await slideTextsOf(pptBytesOf(buildPptStreams(specs)))
    expect(texts[0]).toEqual(['first', 'second bullet text'])
  })

  it('drops the control characters XML cannot carry', async () => {
    const specs: SlideSpec[] = [
      { textSequences: [{ kind: 1, text: 'a\u0001b\u0008c\u000cd\u001be' }] },
    ]
    const texts = await slideTextsOf(pptBytesOf(buildPptStreams(specs)))
    expect(texts[0]).toEqual(['abcde'])
  })
})

describe('a deck without a change directory', () => {
  it('falls back to the containers in stream order', async () => {
    const { document } = buildPptStreams(twoSlides)
    const container = CFB.utils.cfb_new({ root: 'Root Entry' })
    CFB.utils.cfb_add(container, '/PowerPoint Document', Buffer.from(document))
    const out = CFB.write(container, { type: 'buffer' }) as Buffer
    const bytes = new Uint8Array(out.buffer, out.byteOffset, out.byteLength)
    const texts = await slideTextsOf(bytes)
    expect(texts).toEqual([['Quarterly Review', 'Revenue', 'Up 12%'], ['Next steps']])
  })
})

describe('a real PowerPoint-authored deck', () => {
  it('parses the in-repo legacy-sample.ppt', async () => {
    const bytes = new Uint8Array(await readFile(LEGACY_SAMPLE))
    const { document, currentUser } = streamsOf(bytes)
    const presentation = parsePptPresentation(document, currentUser)
    expect(presentation.slideWidth).toBe(6350)
    expect(presentation.slideHeight).toBe(3572)
    expect(presentation.slides.map((slide) => slide.textSequences.map((s) => s.text))).toEqual([
      ['Legacy PPT title', 'First slide body'],
      ['Second legacy slide'],
    ])
  })

  it('resolves the change directory through the explicit persist ids', async () => {
    const bytes = new Uint8Array(await readFile(LEGACY_SAMPLE))
    const { document, currentUser } = streamsOf(bytes)
    const chain = buildPersistChain(document, currentUser)
    // the fixture's UserEditAtom names persist id 1 (the DocumentContainer)
    // and one PersistDirectoryAtom maps its seven objects
    expect(chain.docPersistId).toBe(1)
    expect(chain.directory.size).toBe(7)
    expect(chain.directory.get(1)).toBe(0)
  })

  it('converts it to a deck with the original text and slide size', async () => {
    const bytes = new Uint8Array(await readFile(LEGACY_SAMPLE))
    const out = await convertPptToPptx(bytes)
    expect(out.slideCount).toBe(2)
    const zip = await JSZip.loadAsync(out.bytes)
    const presentation = await zip.file('ppt/presentation.xml')!.async('string')
    // 6350 × 3572 master units × 1587.5 EMU = 10080625 × 5670550 EMU
    expect(presentation).toContain('cx="10080625" cy="5670550"')
    const texts: string[][] = []
    for (let slide = 1; slide <= 2; slide++) {
      const xml = await zip.file(`ppt/slides/slide${slide}.xml`)!.async('string')
      texts.push([...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]!))
    }
    expect(texts).toEqual([['Legacy PPT title', 'First slide body'], ['Second legacy slide']])
  })
})

describe('convertPptToPptx: the failure model', () => {
  it('refuses a file that is not a compound container', async () => {
    await expect(convertPptToPptx(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]))).rejects.toMatchObject({
      code: 'not_cfb',
    })
  })

  it('refuses a compound signature without the 512-byte header', async () => {
    const bytes = new Uint8Array(100)
    bytes.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
    await expect(convertPptToPptx(bytes)).rejects.toMatchObject({ code: 'truncated' })
  })

  it('refuses a container without the presentation stream', async () => {
    const container = CFB.utils.cfb_new({ root: 'Root Entry' })
    const out = CFB.write(container, { type: 'buffer' }) as Buffer
    await expect(
      convertPptToPptx(new Uint8Array(out.buffer, out.byteOffset, out.byteLength)),
    ).rejects.toMatchObject({ code: 'missing_stream' })
  })

  it('refuses an encrypted deck with a dedicated code', async () => {
    const streams = buildPptStreams(twoSlides, { encrypted: true })
    await expect(convertPptToPptx(pptBytesOf(streams))).rejects.toMatchObject({
      code: 'encrypted',
    })
  })
})

describe('parsePptPresentation: the model underneath', () => {
  it('exposes the parsed text sequences for tests and later stages', () => {
    const { document, currentUser } = buildPptStreams(twoSlides)
    const presentation = parsePptPresentation(document, currentUser)
    expect(presentation.slides).toHaveLength(2)
    expect(presentation.slideWidth).toBe(5760)
    expect(presentation.slideHeight).toBe(4320)
    expect(presentation.slides[0]!.textSequences[0]!.text).toBe('Quarterly Review')
  })
})

describe('convertPptToPptx: geometry, fills, styles, pictures', () => {
  const EMU_PER_MU = 914400 / 576
  // a 1×1 PNG
  const PNG = Uint8Array.from(
    Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
      'base64',
    ),
  )

  async function slideXmlOf(bytes: Uint8Array, slide = 1): Promise<string> {
    const out = await convertPptToPptx(bytes)
    const zip = await JSZip.loadAsync(out.bytes)
    return zip.file(`ppt/slides/slide${slide}.xml`)!.async('string')
  }

  it('places a text box at its Escher anchor', async () => {
    const specs: SlideSpec[] = [
      { textSequences: [{ kind: 4, text: 'Anchored', anchor: [576, 288, 2880, 1152] }] },
    ]
    const xml = await slideXmlOf(pptBytesOf(buildPptStreams(specs)))
    const x = Math.round(288 * EMU_PER_MU)
    const y = Math.round(576 * EMU_PER_MU)
    const cx = Math.round((2880 - 288) * EMU_PER_MU)
    const cy = Math.round((1152 - 576) * EMU_PER_MU)
    expect(xml).toContain(`<a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/>`)
    // a text box has no fill or line of its own
    expect(xml).toContain('<a:noFill/><a:ln><a:noFill/></a:ln>')
  })

  it('carries an autoshape preset with its solid fill and line', async () => {
    const specs: SlideSpec[] = [
      {
        textSequences: [
          {
            kind: 4,
            text: 'Oval',
            shapeType: 3,
            anchor: [0, 0, 1000, 1000],
            // fillColor 0xBBGGRR = pure red, line colour blue 2pt, both booleans opted in
            props: {
              0x0181: 0x0000ff,
              0x01bf: 0x00100010,
              0x01c0: 0xff0000,
              0x01cb: 25400,
              0x01ff: 0x00080008,
            },
          },
        ],
      },
    ]
    const xml = await slideXmlOf(pptBytesOf(buildPptStreams(specs)))
    expect(xml).toContain('<a:prstGeom prst="ellipse">')
    expect(xml).toContain('<a:solidFill><a:srgbClr val="FF0000"></a:srgbClr></a:solidFill>')
    expect(xml).toContain('<a:ln w="25400"><a:solidFill><a:srgbClr val="0000FF">')
  })

  it('lays the StyleTextProp runs over the text', async () => {
    const specs: SlideSpec[] = [
      {
        textSequences: [
          {
            kind: 4,
            text: 'Styled',
            anchor: [0, 0, 1000, 1000],
            style: { align: 1, bold: true, size: 40, colorRgb: 0x336699 },
          },
        ],
      },
    ]
    const xml = await slideXmlOf(pptBytesOf(buildPptStreams(specs)))
    expect(xml).toContain('algn="ctr"')
    expect(xml).toMatch(/<a:rPr[^>]* sz="4000" b="1"/)
    expect(xml).toContain('<a:srgbClr val="996633">')
  })

  it("writes the slide background from the drawing's background shape", async () => {
    const specs: SlideSpec[] = [{ textSequences: [], backgroundRgb: 0x663300 }]
    const xml = await slideXmlOf(pptBytesOf(buildPptStreams(specs)))
    expect(xml).toContain('<p:bg><p:bgPr><a:solidFill><a:srgbClr val="003366">')
  })

  it('embeds a picture from the Pictures stream as a media part', async () => {
    const specs: SlideSpec[] = [
      { textSequences: [], pictures: [{ anchor: [100, 100, 600, 600], pib: 1 }] },
    ]
    const streams = buildPptStreams(specs)
    const bytes = buildPptContainer(
      streams.document,
      streams.currentUser,
      buildPicturesStream([PNG]),
    )
    const out = await convertPptToPptx(bytes)
    const zip = await JSZip.loadAsync(out.bytes)
    const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
    expect(xml).toContain('<p:pic>')
    expect(xml).toContain('r:embed="rId2"')
    const media = await zip.file('ppt/media/image1.png')!.async('uint8array')
    expect(Buffer.from(media)).toEqual(Buffer.from(PNG))
    const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string')
    expect(rels).toContain('Target="../media/image1.png"')
  })

  it('a picture the store cannot decode is dropped with a warning, not a crash', async () => {
    const specs: SlideSpec[] = [
      { textSequences: [], pictures: [{ anchor: [100, 100, 600, 600], pib: 7 }] },
    ]
    const out = await convertPptToPptx(pptBytesOf(buildPptStreams(specs)))
    expect(out.warnings.some((w) => w.includes('picture'))).toBe(true)
  })
})

describe('escher: freeform arrays', () => {
  it('reads a vertex array whose declared length leaves out the 6-byte header', async () => {
    const { parseSpContainer, freeformPath } = await import('../src/escher')
    const { walkRecords } = await import('../src/records')
    const u16 = (n: number): number[] => [n & 0xff, (n >> 8) & 0xff]
    const u32 = (n: number): number[] => [...u16(n & 0xffff), ...u16((n >>> 16) & 0xffff)]
    // pVertices (0x145, complex): 3 points of 2 × int16, elemSize 0xFFF0, declared length = 12 (data only)
    const vertices = [
      ...u16(3),
      ...u16(3),
      ...u16(0xfff0),
      ...u16(0),
      ...u16(0),
      ...u16(100),
      ...u16(0),
      ...u16(50),
      ...u16(80),
    ]
    // pSegmentInfo (0x146, complex): moveTo, lineTo ×2, close, end; declared length includes the header
    const segments = [
      ...u16(5),
      ...u16(5),
      ...u16(2),
      ...u16(0x4000),
      ...u16(0x0002),
      ...u16(0x6001),
      ...u16(0x8000),
    ]
    const opt = [
      ...u16(0x0145 | 0x8000),
      ...u32(12),
      ...u16(0x0146 | 0x8000),
      ...u32(segments.length),
      ...u16(0x0142),
      ...u32(100),
      ...u16(0x0143),
      ...u32(80),
      ...vertices,
      ...segments,
    ]
    const optRecord = [...u16((4 << 4) | 3), ...u16(0xf00b), ...u32(opt.length), ...opt]
    const sp = [...u16((0 << 4) | 2), ...u16(0xf00a), ...u32(8), ...u32(1), ...u32(0xa00)]
    const container = Uint8Array.from([
      ...u16(0xf),
      ...u16(0xf004),
      ...u32(sp.length + optRecord.length),
      ...sp,
      ...optRecord,
    ])
    const [record] = [...walkRecords(container)]
    const shape = parseSpContainer(container, record!)
    const path = freeformPath(shape)
    expect(path?.width).toBe(100)
    expect(path?.commands.map((c) => c.kind)).toEqual(['move', 'line', 'close'])
    expect(path?.commands[1]!.points).toEqual([
      [100, 0],
      [50, 80],
    ])
  })
})

describe('text anchors', () => {
  it('maps the baseline anchors onto their DrawingML vertical anchor', async () => {
    const specs: SlideSpec[] = [
      {
        textSequences: [
          { kind: 4, text: 'Bottom baseline', anchor: [0, 0, 1000, 500], props: { 0x0087: 7 } },
          {
            kind: 4,
            text: 'Top centred baseline',
            anchor: [0, 600, 1000, 1000],
            props: { 0x0087: 8 },
          },
        ],
      },
    ]
    const out = await convertPptToPptx(pptBytesOf(buildPptStreams(specs)))
    const zip = await JSZip.loadAsync(out.bytes)
    const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
    const bodies = [...xml.matchAll(/<a:bodyPr[^>]*>/g)].map((m) => m[0])
    expect(bodies[0]).toContain('anchor="b"')
    expect(bodies[0]).not.toContain('anchorCtr')
    expect(bodies[1]).toContain('anchor="t"')
    expect(bodies[1]).toContain('anchorCtr="1"')
  })
})

describe('groups', () => {
  const EMU_PER_MU = 914400 / 576
  async function xfrmsOf(specs: SlideSpec[]): Promise<string[]> {
    const out = await convertPptToPptx(pptBytesOf(buildPptStreams(specs)))
    const zip = await JSZip.loadAsync(out.bytes)
    const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
    return [...xml.matchAll(/<a:xfrm[^>]*>.*?<\/a:xfrm>/g)].map((m) => m[0]).slice(1)
  }

  it('maps a member through the group coordinate space onto the slide', async () => {
    const xfrms = await xfrmsOf([
      {
        textSequences: [],
        groups: [
          {
            anchor: [1000, 1000, 3000, 2000], // 2000 wide, 1000 tall on the slide
            bounds: [0, 0, 200, 100],
            members: [{ childAnchor: [50, 25, 150, 75], text: 'm' }],
          },
        ],
      },
    ])
    expect(xfrms).toHaveLength(1)
    const emu = (mu: number): number => Math.round(mu * EMU_PER_MU)
    expect(xfrms[0]).toContain(
      `<a:off x="${emu(1500)}" y="${emu(1250)}"/><a:ext cx="${emu(1000)}" cy="${emu(500)}"/>`,
    )
  })

  it('a rotated group turns its members about the group centre', async () => {
    const xfrms = await xfrmsOf([
      {
        textSequences: [],
        groups: [
          {
            // stored as the rotated bounding box: 2000 wide × 1000 tall, really 1000 × 2000
            anchor: [1000, 1000, 3000, 2000],
            bounds: [0, 0, 200, 100],
            rotation: 90,
            // the left half of the group space, centred vertically
            members: [{ childAnchor: [0, 25, 100, 75], text: 'm', rotation: 10 }],
          },
        ],
      },
    ])
    const emu = (mu: number): number => Math.round(mu * EMU_PER_MU)
    // unswapped group box 1500..2500 × 500..2500; member 1500..2000 × 1000..2000, centre (1750, 1500)
    // rotated 90° about the group centre (2000, 1500) → (2000, 1250); size stays 500 × 1000
    expect(xfrms[0]).toContain('rot="6000000"')
    expect(xfrms[0]).toContain(
      `<a:off x="${emu(1750)}" y="${emu(750)}"/><a:ext cx="${emu(500)}" cy="${emu(1000)}"/>`,
    )
  })

  it('a flipped group mirrors its members and toggles their flip', async () => {
    const xfrms = await xfrmsOf([
      {
        textSequences: [],
        groups: [
          {
            anchor: [1000, 1000, 3000, 2000],
            bounds: [0, 0, 200, 100],
            flipH: true,
            members: [{ childAnchor: [0, 0, 100, 50], text: 'm' }],
          },
        ],
      },
    ])
    const emu = (mu: number): number => Math.round(mu * EMU_PER_MU)
    expect(xfrms[0]).toContain('flipH="1"')
    expect(xfrms[0]).toContain(`<a:off x="${emu(2000)}" y="${emu(1000)}"/>`)
  })
})

describe('stage 2 follow-ups', () => {
  const EMU_PER_MU = 914400 / 576
  async function slideXml(specs: SlideSpec[], options?: BuildOptions): Promise<string> {
    const out = await convertPptToPptx(pptBytesOf(buildPptStreams(specs, options)))
    const zip = await JSZip.loadAsync(out.bytes)
    return zip.file('ppt/slides/slide1.xml')!.async('string')
  }

  it('a shape turned 90° swaps the sides of its stored bounding box', async () => {
    const xml = await slideXml([
      {
        textSequences: [
          {
            kind: 4,
            text: 'bar',
            shapeType: 1,
            // a wide thin box, rotated 270°: Escher stored the box as it shows
            anchor: [1000, 0, 4000, 1200],
            props: { 0x0004: Math.round(270 * 65536) >>> 0, 0x01bf: 0x00100010 },
          },
        ],
      },
    ])
    const emu = (mu: number): number => Math.round(mu * EMU_PER_MU)
    // centre (2000, 1100), sides swapped to 200 × 4000
    expect(xml).toContain('rot="16200000"')
    expect(xml).toContain(
      `<a:off x="${emu(1900)}" y="${emu(-900)}"/><a:ext cx="${emu(200)}" cy="${emu(4000)}"/>`,
    )
  })

  it('an opted-in fill without a colour is white, not a scheme slot', async () => {
    const xml = await slideXml([
      {
        textSequences: [
          {
            kind: 4,
            text: 'cell',
            shapeType: 1,
            anchor: [0, 0, 500, 500],
            props: { 0x01bf: 0x00100010 },
          },
        ],
      },
    ])
    expect(xml).toContain('<a:solidFill><a:srgbClr val="FFFFFF"></a:srgbClr></a:solidFill>')
  })

  it('a slide-number MetaCharacter atom becomes a slidenum field', async () => {
    const specs: SlideSpec[] = [
      {
        textSequences: [
          {
            kind: 4,
            text: 'Page *',
            anchor: [0, 0, 2000, 500],
            metaChars: [{ type: 0x0fd8, position: 5 }],
          },
        ],
      },
    ]
    const xml = await slideXml(specs)
    expect(xml).toMatch(
      /<a:r>[^]*?<a:t>Page <\/a:t><\/a:r><a:fld id="\{[0-9A-F-]+\}" type="slidenum">[^]*?<a:t>1<\/a:t><\/a:fld>/,
    )
  })
})

describe('gradient fills', () => {
  async function slideXml(specs: SlideSpec[], options?: BuildOptions): Promise<string> {
    const out = await convertPptToPptx(pptBytesOf(buildPptStreams(specs, options)))
    const zip = await JSZip.loadAsync(out.bytes)
    return zip.file('ppt/slides/slide1.xml')!.async('string')
  }

  it('focus 100 runs fill → back; a shape shade is a centre-out path', async () => {
    const xml = await slideXml([
      {
        textSequences: [
          {
            kind: 4,
            text: 'linear',
            shapeType: 1,
            anchor: [0, 0, 1000, 500],
            // fillType shade(4), fill red, back blue, focus 100, angle 0
            props: {
              0x0180: 4,
              0x0181: 0x0000ff,
              0x0183: 0xff0000,
              0x018c: 100,
              0x01bf: 0x00100010,
            },
          },
          {
            kind: 4,
            text: 'shape',
            shapeType: 1,
            anchor: [600, 0, 1000, 1100],
            props: {
              0x0180: 6,
              0x0181: 0xffffff,
              0x0183: 0x00ff00,
              0x018c: 100,
              0x01bf: 0x00100010,
            },
          },
        ],
      },
    ])
    expect(xml).toContain(
      '<a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/>',
    )
    // no fillTo* props: the focus rectangle is the top-left corner
    expect(xml).toContain(
      '<a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="00FF00"/></a:gs></a:gsLst><a:path path="shape"><a:fillToRect l="0" t="0" r="100000" b="100000"/></a:path>',
    )
  })

  it('focus ±50 mirrors fill → back → fill, focus 0 runs back → fill', async () => {
    const shade = (text: string, left: number, focus: number): TextSpec => ({
      kind: 4,
      text,
      shapeType: 1,
      anchor: [left, 0, left + 500, 500],
      props: {
        0x0180: 7,
        0x0181: 0x0000ff,
        0x0183: 0xff0000,
        0x018c: focus >>> 0,
        0x01bf: 0x00100010,
      },
    })
    const xml = await slideXml([
      { textSequences: [shade('half', 0, 50), shade('neg', 600, -50), shade('zero', 1200, 0)] },
    ])
    const mirrored =
      '<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="50000"><a:srgbClr val="0000FF"/></a:gs><a:gs pos="100000"><a:srgbClr val="FF0000"/></a:gs>'
    expect(xml.split(mirrored)).toHaveLength(3)
    expect(xml).toContain(
      '<a:gsLst><a:gs pos="0"><a:srgbClr val="0000FF"/></a:gs><a:gs pos="100000"><a:srgbClr val="FF0000"/></a:gs></a:gsLst>',
    )
  })

  it('a centre shade reads its focus rectangle from fillToLeft/Top/Right/Bottom; an angle past ±180° exchanges the mirrored colours', async () => {
    const xml = await slideXml([
      {
        textSequences: [
          {
            kind: 4,
            text: 'centre',
            shapeType: 1,
            anchor: [0, 0, 1000, 500],
            props: {
              0x0180: 5,
              0x0181: 0x404040,
              0x0183: 0xffffff,
              0x018d: 0x8000,
              0x018e: 0x8000,
              0x018f: 0x8000,
              0x0190: 0x8000,
              0x01bf: 0x00100010,
            },
          },
          {
            kind: 4,
            text: 'turned',
            shapeType: 1,
            anchor: [0, 600, 1000, 1100],
            // fillAngle -210° (16.16), focus 50
            props: {
              0x0180: 7,
              0x0181: 0xff4700,
              0x0183: 0x800031,
              0x018b: 0xff2e0000,
              0x018c: 50,
              0x01bf: 0x00100010,
            },
          },
        ],
      },
    ])
    expect(xml).toContain(
      '<a:path path="rect"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>',
    )
    expect(xml).toContain(
      '<a:gs pos="0"><a:srgbClr val="310080"/></a:gs><a:gs pos="50000"><a:srgbClr val="0047FF"/></a:gs><a:gs pos="100000"><a:srgbClr val="310080"/></a:gs></a:gsLst><a:lin ang="18000000" scaled="0"/>',
    )
  })

  it('a texture fill tiles its picture; a picture fill stretches it', async () => {
    const PNG = Uint8Array.from(
      Buffer.from(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=',
        'base64',
      ),
    )
    const filled = (text: string, left: number, fillType: number): TextSpec => ({
      kind: 4,
      text,
      shapeType: 1,
      anchor: [left, 0, left + 500, 500],
      props: { 0x0180: fillType, 0x0186: 1, 0x01bf: 0x00100010 },
    })
    const streams = buildPptStreams([
      { textSequences: [filled('texture', 0, 2), filled('picture', 600, 3)] },
    ])
    const out = await convertPptToPptx(
      buildPptContainer(streams.document, streams.currentUser, buildPicturesStream([PNG])),
    )
    const zip = await JSZip.loadAsync(out.bytes)
    const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
    expect(xml).toContain(
      '<a:blipFill><a:blip r:embed="rId2"/><a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/></a:blipFill>',
    )
    expect(xml).toContain(
      '<a:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>',
    )
  })
})

describe('quarter-turned groups', () => {
  it('a group rotated 270° swaps its own bounding box before placing its members', async () => {
    const specs: SlideSpec[] = [
      {
        textSequences: [],
        groups: [
          {
            // a wide thin strip along the bottom, stored as it shows; rotation 270°
            anchor: [4000, 0, 6000, 4300],
            bounds: [0, 0, 300, 6000],
            rotation: 270,
            members: [{ childAnchor: [0, 0, 300, 6000], text: 'strip' }],
          },
        ],
      },
    ]
    const out = await convertPptToPptx(pptBytesOf(buildPptStreams(specs)))
    const zip = await JSZip.loadAsync(out.bytes)
    const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
    const EMU = 914400 / 576
    const emu = (mu: number): number => Math.round(mu * EMU)
    // member fills the group: a 300 × 6000 box centred on (3000, 4150), rotated 270° → the visible strip
    expect(xml).toContain('rot="16200000"')
    expect(xml).toContain(
      `<a:off x="${emu(2850)}" y="${emu(1150)}"/><a:ext cx="${emu(300)}" cy="${emu(6000)}"/>`,
    )
  })
})

describe('outline fallback', () => {
  it('an outline entry no shape claims keeps its MetaCharacter fields', async () => {
    // outline texts with no OutlineTextRef shapes: the entries land as stacked boxes
    const specs: SlideSpec[] = [
      { textSequences: [{ kind: 4, text: 'Slide *', metaChars: [{ type: 0x0fd8, position: 6 }] }] },
    ]
    const streams = buildPptStreams(specs, { outlineTexts: true, outlineMetaChars: true })
    const out = await convertPptToPptx(pptBytesOf(streams))
    const zip = await JSZip.loadAsync(out.bytes)
    const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
    expect(xml).toContain('type="slidenum"')
  })
})

describe('round 3: system colours, fill opt-in, master level cascade', () => {
  async function slideXml(specs: SlideSpec[], options?: BuildOptions): Promise<string> {
    const out = await convertPptToPptx(pptBytesOf(buildPptStreams(specs, options)))
    const zip = await JSZip.loadAsync(out.bytes)
    return zip.file('ppt/slides/slide1.xml')!.async('string')
  }

  it("a system-colour back reference darkens the shape's own fill colour", async () => {
    const xml = await slideXml([
      {
        textSequences: [
          {
            kind: 4,
            text: 'x',
            shapeType: 1,
            anchor: [0, 0, 1000, 500],
            // shade, fill navy 0x990000 (BGR), back = fill colour darkened to 118/255
            props: { 0x0180: 7, 0x0181: 0x990000, 0x0183: 0x107601f0, 0x01bf: 0x00100010 },
          },
        ],
      },
    ])
    expect(xml).toContain('<a:srgbClr val="000099"/>')
    expect(xml).toContain('<a:srgbClr val="000047"/>')
  })

  it('a shadow system colour reads the scheme shadow slot; a Windows system colour stays unresolved', () => {
    const scheme = {
      colors: ['FFFFFF', '000000', '808080', '000000', '00CC99', '3333CC', 'CCCCFF', 'B2B2B2'],
    }
    expect(officeArtColor(0x100000f3, scheme, { fill: 0x0000ff })).toBe('808080')
    expect(officeArtColor(0x10000005, scheme, { fill: 0x0000ff })).toBeUndefined()
    expect(officeArtColor(0x100000f0, scheme, { fill: 0x0000ff })).toBe('FF0000')
  })

  it('a rectangle with only fill colours keeps its default fill; a fill type opts in explicitly', async () => {
    const xml = await slideXml([
      {
        textSequences: [
          {
            kind: 4,
            text: 'plain',
            shapeType: 1,
            anchor: [0, 0, 1000, 500],
            props: { 0x0181: 0x08000004, 0x0183: 0x08000000, 0x01bf: 0x00600000 },
          },
          {
            kind: 4,
            text: 'shaded',
            shapeType: 1,
            anchor: [600, 0, 1000, 1100],
            props: { 0x0180: 4, 0x0181: 0x0000ff, 0x0183: 0xff0000, 0x01bf: 0x00200000 },
          },
        ],
      },
    ])
    const shapes = xml.split('<p:sp>').slice(1)
    // scheme slot 4 of the default scheme
    expect(shapes[0]).toContain('<a:solidFill><a:srgbClr val="00CC99">')
    expect(shapes[1]).toContain('<a:gradFill')
  })

  it('a master style level inherits what it leaves unsaid from the level above', async () => {
    const { parseMasterStyle } = await import('../src/text-style')
    const u16 = (n: number): number[] => [n & 0xff, (n >> 8) & 0xff]
    const u32 = (n: number): number[] => [...u16(n & 0xffff), ...u16((n >>> 16) & 0xffff)]
    // two levels: level 0 has bullet on with char 'n' and align centre + size 32; level 1 only a left margin
    const level0 = [
      ...u32(0x0001 | 0x0080 | 0x0800),
      ...u16(1),
      ...u16(0x6e),
      ...u16(1),
      ...u32(0x20000),
      ...u16(32),
    ]
    const level1 = [...u32(0x0100), ...u16(400), ...u32(0)]
    const levels = parseMasterStyle(Uint8Array.from([...u16(2), ...level0, ...level1]), 1)
    expect(levels[1]!.paragraph).toMatchObject({
      level: 1,
      bullet: true,
      bulletChar: 'n',
      align: 1,
      leftMargin: 400,
    })
    expect(levels[1]!.chars.size).toBe(32)
  })
})

describe('round 4: pattern fills, background fill, picture transparency, arcs', () => {
  const EMU_PER_MU = 914400 / 576
  const u16 = (n: number): number[] => [n & 0xff, (n >> 8) & 0xff]

  async function slideXml(specs: SlideSpec[], pictures?: Uint8Array): Promise<string> {
    const streams = buildPptStreams(specs)
    const out = await convertPptToPptx(
      buildPptContainer(streams.document, streams.currentUser, pictures),
    )
    const zip = await JSZip.loadAsync(out.bytes)
    return zip.file('ppt/slides/slide1.xml')!.async('string')
  }

  it('a pattern fill becomes the preset its 8×8 cell draws, fill colour on back colour', async () => {
    // PowerPoint's "Light horizontal": one row of fill-colour pixels every four; black DIB pixels take the back colour
    const dib = buildPatternDib([0x00, 0xff, 0xff, 0xff, 0x00, 0xff, 0xff, 0xff])
    const xml = await slideXml(
      [
        {
          textSequences: [
            {
              kind: 4,
              text: 'hatched',
              shapeType: 1,
              anchor: [0, 0, 1000, 1000],
              props: {
                0x0180: 1,
                0x0181: 0x0000ff,
                0x0183: 0xffff00,
                0x0186: 1,
                0x01bf: 0x00100010,
              },
            },
          ],
        },
      ],
      buildPicturesStream([{ dib }]),
    )
    expect(xml).toContain(
      '<a:pattFill prst="ltHorz"><a:fgClr><a:srgbClr val="FF0000"/></a:fgClr><a:bgClr><a:srgbClr val="00FFFF"/></a:bgClr></a:pattFill>',
    )
  })

  it('a pattern whose fill or back colour is a scheme index keeps the DIB palette colours', async () => {
    const dib = buildPatternDib([0x00, 0xff, 0xff, 0xff, 0x00, 0xff, 0xff, 0xff])
    const xml = await slideXml(
      [
        {
          textSequences: [
            {
              kind: 4,
              text: 'hatched',
              shapeType: 1,
              anchor: [0, 0, 1000, 1000],
              props: {
                0x0180: 1,
                0x0181: 0x0000ff,
                0x0183: 0x08000000,
                0x0186: 1,
                0x01bf: 0x00100010,
              },
            },
          ],
        },
      ],
      buildPicturesStream([{ dib }]),
    )
    expect(xml).toContain(
      '<a:pattFill prst="ltHorz"><a:fgClr><a:srgbClr val="FFFFFF"/></a:fgClr><a:bgClr><a:srgbClr val="000000"/></a:bgClr></a:pattFill>',
    )
  })

  it('a pattern cell is recognised at any phase; an unknown cell takes the percentage of its density', async () => {
    const { patternPreset } = await import('../src/patterns')
    expect(patternPreset([0x44, 0x22, 0x11, 0x88, 0x44, 0x22, 0x11, 0x88])).toBe('ltDnDiag')
    expect(patternPreset([0x00, 0xff, 0x00, 0xff, 0x00, 0xff, 0x00, 0xff])).toBe('narHorz')
    expect(patternPreset([0x00, 0x00, 0x00, 0xff, 0x00, 0x00, 0x00, 0xff])).toBe('ltHorz')
    // one line in eight is no PowerPoint pattern: 8 pixels of 64 is the 20% cell
    expect(patternPreset([0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xff])).toBe('pct20')
    expect(patternPreset([0x55, 0xaa, 0x55, 0xaa, 0x55, 0xaa, 0x55, 0xaa])).toBe('pct50')
    expect(patternPreset([0x07, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00])).toBe('pct5')
    expect(patternPreset([0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xfe])).toBe('pct90')
  })

  it('a shape filled with the slide background is written with useBgFill and no fill of its own', async () => {
    const xml = await slideXml([
      {
        textSequences: [
          {
            kind: 4,
            text: 'hole',
            shapeType: 3,
            anchor: [0, 0, 1000, 1000],
            props: { 0x0180: 9, 0x0181: 0x08000000, 0x01bf: 0x00600000, 0x01ff: 0x00080000 },
          },
        ],
        backgroundRgb: 0xff0000,
      },
    ])
    expect(xml).toContain('<p:sp useBgFill="1">')
    const shape = xml.slice(xml.indexOf('<p:spPr>'), xml.indexOf('</p:spPr>'))
    expect(shape).toContain('<a:prstGeom prst="ellipse">')
    expect(shape).not.toContain('<a:solidFill>')
    expect(shape).not.toContain('<a:noFill/><a:ln>')
    // a background shape drawn with the "background" fill means nothing: the scheme stays
    expect(xml).toContain('<p:bg><p:bgPr><a:solidFill><a:srgbClr val="0000FF">')
  })

  it("a picture's transparent colour is a colour change to alpha 0", async () => {
    const png = Uint8Array.from(
      Buffer.from(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
        'base64',
      ),
    )
    const xml = await slideXml(
      [
        {
          textSequences: [],
          pictures: [{ anchor: [100, 100, 600, 600], pib: 1, props: { 0x0107: 0xffffff } }],
        },
      ],
      buildPicturesStream([png]),
    )
    expect(xml).toContain(
      '<a:blip r:embed="rId2"><a:clrChange><a:clrFrom><a:srgbClr val="FFFFFF"/></a:clrFrom><a:clrTo><a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr></a:clrTo></a:clrChange></a:blip>',
    )
  })

  /** pVertices for an arc: ellipse left/top, right/bottom, start, end — 2 × uint16 per point */
  const arcVertices = (points: Array<[number, number]>): Uint8Array =>
    Uint8Array.from([
      ...u16(points.length),
      ...u16(points.length),
      ...u16(0xfff0),
      ...points.flatMap(([x, y]) => [...u16(x), ...u16(y)]),
    ])

  it('an arc keeps its anchor when the ellipse fills the geometry box, with its clockwise sweep', async () => {
    const xml = await slideXml([
      {
        textSequences: [
          {
            kind: 4,
            text: '',
            shapeType: 19,
            anchor: [1000, 2000, 4000, 2000],
            // 12 o'clock clockwise round to just short of 12 o'clock
            props: { 0x0142: 43200, 0x0143: 43200, 0x0181: 0x00ccff, 0x01bf: 0x00100010 },
            complex: {
              0x0145: arcVertices([
                [0, 0],
                [43200, 43200],
                [21600, 0],
                [21388, 1],
              ]),
            },
          },
        ],
      },
    ])
    const x = Math.round(2000 * EMU_PER_MU)
    const y = Math.round(1000 * EMU_PER_MU)
    expect(xml).toContain(`<a:off x="${x}" y="${y}"/><a:ext cx="${x}" cy="${y}"/>`)
    expect(xml).toContain(
      '<a:prstGeom prst="arc"><a:avLst><a:gd name="adj1" fmla="val 16200000"/><a:gd name="adj2" fmla="val 16166259"/></a:avLst></a:prstGeom>',
    )
  })

  it('a quarter arc grows from its sector anchor to the whole ellipse', async () => {
    const xml = await slideXml([
      {
        textSequences: [
          {
            kind: 4,
            text: '',
            shapeType: 19,
            anchor: [0, 1000, 2000, 1000],
            // the geometry box is the top-right quarter of a 43200 ellipse
            props: { 0x0140: 21600, 0x0141: 0, 0x0142: 43200, 0x0143: 21600, 0x01bf: 0x00100010 },
            complex: {
              0x0145: arcVertices([
                [0, 0],
                [43200, 43200],
                [21600, 0],
                [43200, 21600],
              ]),
            },
          },
        ],
      },
    ])
    const side = Math.round(1000 * EMU_PER_MU)
    expect(xml).toContain(`<a:off x="0" y="0"/><a:ext cx="${side * 2}" cy="${side * 2}"/>`)
    expect(xml).toContain('<a:gd name="adj1" fmla="val 16200000"/><a:gd name="adj2" fmla="val 0"/>')
  })

  it('an arc without vertices is the default top-right quarter', async () => {
    const { arcGeometry } = await import('../src/escher')
    const arc = arcGeometry({
      spid: 1,
      flags: 0,
      shapeType: 19,
      props: new Map(),
      complex: new Map(),
      children: [],
    })
    expect(arc.ellipse).toEqual({ left: -21600, top: 0, right: 21600, bottom: 43200 })
    expect(arc.box).toEqual({ left: 0, top: 0, right: 21600, bottom: 21600 })
  })
})

describe('round 4: paragraph rulers, centre-body style, metro indents', () => {
  const u16 = (n: number): number[] => [n & 0xff, (n >> 8) & 0xff]
  const u32 = (n: number): number[] => [...u16(n & 0xffff), ...u16((n >>> 16) & 0xffff)]

  async function slideXml(specs: SlideSpec[], options?: BuildOptions): Promise<string> {
    const out = await convertPptToPptx(pptBytesOf(buildPptStreams(specs, options)))
    const zip = await JSZip.loadAsync(out.bytes)
    return zip.file('ppt/slides/slide1.xml')!.async('string')
  }

  it('a ruler reads its tab stops and interleaves margin and indent per level', async () => {
    const { parseTextRuler } = await import('../src/text-style')
    // tabs + leftMargin1 + leftMargin2 + indent2
    const payload = Uint8Array.from([
      ...u32(0x0004 | 0x0008 | 0x0010 | 0x0200),
      ...u16(2),
      ...u16(288),
      ...u16(0),
      ...u16(576),
      ...u16(0),
      ...u16(286),
      ...u16(395),
      ...u16(395),
    ])
    const { levels, tabs } = parseTextRuler(payload)
    expect(tabs).toEqual([
      { position: 288, type: 0 },
      { position: 576, type: 0 },
    ])
    expect(levels[0]).toEqual({ leftMargin: 286 })
    expect(levels[1]).toEqual({ leftMargin: 395, indent: 395 })
    expect(levels[2]).toEqual({})
  })

  it('a text ruler sets marL and the hanging indent the master body level lacks', async () => {
    const xml = await slideXml([
      {
        textSequences: [
          {
            kind: 4,
            text: 'first\rsecond',
            anchor: [0, 0, 2000, 1000],
            // leftMargin1 = 288 (half an inch), indent1 = 0
            ruler: Uint8Array.from([...u32(0x0008 | 0x0100), ...u16(288), ...u16(0)]),
          },
        ],
      },
    ])
    expect(xml).toContain('marL="457200" indent="-457200"')
  })

  it('a centre-body text layers the instance-5 style over the body style', async () => {
    const { levelsForKind } = await import('../src/text-style')
    const body = [
      {
        paragraph: { level: 0, bullet: true, align: 0, leftMargin: 216, spaceBefore: 20 },
        chars: { size: 28, bold: true },
      },
    ]
    const centre = [{ paragraph: { level: 0, bullet: false, align: 1, leftMargin: 0 }, chars: {} }]
    const styles = (instance: number) =>
      instance === 1 ? body : instance === 5 ? centre : undefined
    expect(levelsForKind(5, styles)[0]).toEqual({
      paragraph: { level: 0, bullet: false, align: 1, leftMargin: 0, spaceBefore: 20 },
      chars: { size: 28, bold: true },
    })
    expect(levelsForKind(1, styles)).toBe(body)
    expect(levelsForKind(7, styles)).toEqual(body)
  })

  it("a metroBlob's per-paragraph indents override the shared legacy ruler", async () => {
    const pkg = new JSZip()
    pkg.file(
      'drs/shapexml.xml',
      '<?xml version="1.0"?><p:sp xmlns:p="p" xmlns:a="a"><p:txBody><a:bodyPr/>' +
        '<a:p><a:pPr marL="360000" indent="540000"/><a:r><a:t>_____</a:t></a:r></a:p>' +
        '<a:p xmlns:a="a"><a:pPr marL="720000" indent="-360000"/><a:r><a:t>______</a:t></a:r></a:p>' +
        '</p:txBody></p:sp>',
    )
    const blob = await pkg.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
    const xml = await slideXml([
      {
        textSequences: [
          {
            kind: 4,
            text: 'first\rsecond',
            anchor: [0, 0, 2000, 1000],
            ruler: Uint8Array.from([...u32(0x0008 | 0x0100), ...u16(226), ...u16(566)]),
            metroBlob: blob,
          },
        ],
      },
    ])
    expect(xml).toContain('marL="360000" indent="540000"')
    expect(xml).toContain('marL="720000" indent="-360000"')
  })

  it('an empty outline header still counts for the OutlineTextRefAtom index', async () => {
    const specs: SlideSpec[] = [
      {
        textSequences: [
          { kind: 0, text: '' },
          { kind: 1, text: 'Body text', anchor: [1000, 2000, 5000, 3000] },
        ],
      },
    ]
    const xml = await slideXml(specs, { outlineTexts: true })
    const shapes = xml.split('<p:sp>').slice(1)
    expect(shapes).toHaveLength(1)
    expect(shapes[0]).toContain(`<a:off x="${Math.round((2000 * 914400) / 576)}"`)
    expect(shapes[0]).toContain('Body text')
  })

  it('a solid fill type without fFilled does not opt a text box into a fill', async () => {
    const xml = await slideXml([
      {
        textSequences: [
          {
            kind: 4,
            text: 'boilerplate',
            anchor: [0, 0, 1000, 500],
            props: { 0x0180: 0, 0x0181: 0x08000005, 0x0182: 0x10000, 0x01bf: 0x00600000 },
          },
        ],
      },
    ])
    expect(xml).toContain('<a:noFill/>')
    expect(xml).not.toContain('<a:solidFill><a:srgbClr val="3333CC">')
  })

  it('a literal fill colour of its own opts a text box into a fill; a scheme colour does not', async () => {
    const xml = await slideXml([
      {
        textSequences: [
          {
            kind: 4,
            text: 'blue',
            anchor: [0, 0, 1000, 500],
            props: { 0x0181: 0xff6633, 0x01bf: 0x00600000 },
          },
          {
            kind: 4,
            text: 'automatic',
            anchor: [600, 0, 1000, 1100],
            props: { 0x0181: 0x08000004, 0x0183: 0x08000000, 0x01bf: 0x00600000 },
          },
        ],
      },
    ])
    const shapes = xml.split('<p:sp>').slice(1)
    expect(shapes[0]).toContain('<a:solidFill><a:srgbClr val="3366FF">')
    expect(shapes[1]).toContain('<a:noFill/>')
  })

  it('a ruler writes its tab stops into every paragraph', async () => {
    const xml = await slideXml([
      {
        textSequences: [
          {
            kind: 4,
            text: '\ta\r\tb',
            anchor: [0, 0, 2000, 1000],
            ruler: Uint8Array.from([
              ...u32(0x0004),
              ...u16(2),
              ...u16(0),
              ...u16(0),
              ...u16(288),
              ...u16(2),
            ]),
          },
        ],
      },
    ])
    const tabLst = '<a:tabLst><a:tab pos="0" algn="l"/><a:tab pos="457200" algn="r"/></a:tabLst>'
    expect(xml.split(tabLst)).toHaveLength(3)
  })

  it('a metroBlob with one trailing paragraph more than the text still sets the indents', async () => {
    const pkg = new JSZip()
    pkg.file(
      'drs/shapexml.xml',
      '<p:sp><p:txBody><a:p><a:pPr marL="360000" indent="0"/></a:p><a:p><a:pPr marL="360000" indent="0"/></a:p><a:p><a:pPr marL="0" indent="0"/></a:p></p:txBody></p:sp>',
    )
    const blob = await pkg.generateAsync({ type: 'uint8array' })
    const xml = await slideXml([
      {
        textSequences: [
          {
            kind: 4,
            text: 'first\rsecond',
            anchor: [0, 0, 2000, 1000],
            ruler: Uint8Array.from([...u32(0x0008), ...u16(288)]),
            metroBlob: blob,
          },
        ],
      },
    ])
    expect(xml.match(/marL="360000"/g)).toHaveLength(2)
  })

  it('a metroBlob whose paragraph count differs leaves the legacy indents alone', async () => {
    const pkg = new JSZip()
    pkg.file(
      'drs/shapexml.xml',
      '<p:sp><p:txBody><a:p><a:pPr marL="360000" indent="540000"/></a:p></p:txBody></p:sp>',
    )
    const blob = await pkg.generateAsync({ type: 'uint8array' })
    const xml = await slideXml([
      {
        textSequences: [
          {
            kind: 4,
            text: 'first\rsecond',
            anchor: [0, 0, 2000, 1000],
            ruler: Uint8Array.from([...u32(0x0008), ...u16(288)]),
            metroBlob: blob,
          },
        ],
      },
    ])
    expect(xml).not.toContain('marL="360000"')
    expect(xml.match(/marL="457200"/g)).toHaveLength(2)
  })
})

describe('round 5: text box defaults, bullet flags, text shadow', () => {
  async function slideXmlOf(bytes: Uint8Array): Promise<string> {
    const out = await convertPptToPptx(bytes)
    const zip = await JSZip.loadAsync(out.bytes)
    return zip.file('ppt/slides/slide1.xml')!.async('string')
  }
  const styled = (style: NonNullable<SlideSpec['textSequences'][number]['style']>): SlideSpec[] => [
    { textSequences: [{ kind: 4, text: 'Styled', anchor: [0, 0, 1000, 1000], style }] },
  ]

  it('a slide text box reads the document default style, a master shape its other-text style', async () => {
    const { levelsForText } = await import('../src/text-style')
    const docOther = [{ paragraph: { level: 0 }, chars: { size: 9 } }]
    const masterOther = [{ paragraph: { level: 0 }, chars: { size: 18 } }]
    const body = [{ paragraph: { level: 0, bullet: true }, chars: { size: 32 } }]
    const master = (instance: number) =>
      instance === 4 ? masterOther : instance === 1 ? body : undefined
    const doc = (instance: number) => (instance === 4 ? docOther : undefined)
    expect(levelsForText(4, false, master, doc)).toEqual({ levels: docOther, base: docOther })
    expect(levelsForText(4, true, master, doc).levels).toBe(masterOther)
    expect(levelsForText(1, false, master, doc)).toEqual({ levels: body, base: docOther })
    expect(levelsForText(4, false, master, () => undefined).levels).toBe(masterOther)
  })

  it('a bullet colour only applies when fBulletHasColor says the bullet uses it', async () => {
    const withColor = await slideXmlOf(
      pptBytesOf(
        buildPptStreams(styled({ bullet: true, bulletHasColor: true, bulletColorRgb: 0x0000ff })),
      ),
    )
    expect(withColor).toContain('<a:buClr><a:srgbClr val="FF0000"/></a:buClr>')
    const followsText = await slideXmlOf(
      pptBytesOf(
        buildPptStreams(styled({ bullet: true, bulletHasColor: false, bulletColorRgb: 0x0000ff })),
      ),
    )
    expect(followsText).toContain('<a:buChar')
    expect(followsText).not.toContain('<a:buClr>')
  })

  it('a shadowed run casts the scheme shadow colour as an outer shadow', async () => {
    const xml = await slideXmlOf(pptBytesOf(buildPptStreams(styled({ shadow: true, bold: true }))))
    expect(xml).toMatch(
      /<a:rPr[^>]* b="1"><a:solidFill>.*?<\/a:solidFill><a:effectLst><a:outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl"><a:srgbClr val="[0-9A-F]{6}"><a:alpha val="43137"\/><\/a:srgbClr><\/a:outerShdw><\/a:effectLst>/,
    )
    const plain = await slideXmlOf(pptBytesOf(buildPptStreams(styled({ bold: true }))))
    expect(plain).not.toContain('outerShdw')
  })
})

describe('round 6: PP9 auto-numbering, line dashes', () => {
  async function slideXmlOf(bytes: Uint8Array): Promise<string> {
    const out = await convertPptToPptx(bytes)
    const zip = await JSZip.loadAsync(out.bytes)
    return zip.file('ppt/slides/slide1.xml')!.async('string')
  }
  const numbered = (
    autoNumber: NonNullable<TextSpec['autoNumber']>,
    text = 'First\rSecond',
  ): SlideSpec[] => [
    {
      textSequences: [
        { kind: 1, text, anchor: [0, 0, 1000, 1000], style: { bullet: true }, autoNumber },
      ],
    },
  ]

  it("a shape's StyleTextProp9Atom turns its bullet into an auto-number with the scheme and start", async () => {
    const xml = await slideXmlOf(pptBytesOf(buildPptStreams(numbered({ scheme: 1, startAt: 4 }))))
    expect(xml.match(/<a:buAutoNum type="alphaUcPeriod" startAt="4"\/>/g)).toHaveLength(2)
    expect(xml).not.toContain('<a:buChar')
    const plain = await slideXmlOf(pptBytesOf(buildPptStreams(numbered({}))))
    expect(plain).toContain('<a:buAutoNum type="arabicPeriod"/>')
  })

  it('an outline text finds its numbering in the document OutlineTextProps9 entry for its slide and type', async () => {
    const specs: SlideSpec[] = [
      {
        textSequences: [
          { kind: 0, text: 'Title' },
          { kind: 1, text: 'Plain', style: { bullet: true } },
        ],
      },
      ...numbered({ scheme: 3, startAt: 2 }),
    ]
    const out = await convertPptToPptx(pptBytesOf(buildPptStreams(specs, { outlineTexts: true })))
    const zip = await JSZip.loadAsync(out.bytes)
    expect(await zip.file('ppt/slides/slide1.xml')!.async('string')).not.toContain('buAutoNum')
    expect(await zip.file('ppt/slides/slide2.xml')!.async('string')).toContain(
      '<a:buAutoNum type="arabicPeriod" startAt="2"/>',
    )
  })

  it('an empty paragraph is not numbered', async () => {
    const xml = await slideXmlOf(pptBytesOf(buildPptStreams(numbered({}, 'One\r\rThree'))))
    expect(xml.match(/<a:buAutoNum/g)).toHaveLength(2)
    expect(xml).toContain('<a:buNone/>')
  })

  it('a dashed, compound outline writes prstDash and cmpd', async () => {
    const outlined = (props: Record<number, number>): SlideSpec[] => [
      {
        textSequences: [
          {
            kind: 4,
            text: 'Box',
            shapeType: 1,
            anchor: [0, 0, 1000, 1000],
            props: { 0x01cb: 25400, 0x01ff: 0x00080008, ...props },
          },
        ],
      },
    ]
    const dashed = await slideXmlOf(pptBytesOf(buildPptStreams(outlined({ 0x01ce: 6 }))))
    expect(dashed).toMatch(
      /<a:ln w="25400"><a:solidFill>.*?<\/a:solidFill><a:prstDash val="dash"\/><\/a:ln>/,
    )
    const sysDot = await slideXmlOf(pptBytesOf(buildPptStreams(outlined({ 0x01ce: 2, 0x01cd: 1 }))))
    expect(sysDot).toContain('<a:ln w="25400" cmpd="dbl">')
    expect(sysDot).toContain('<a:prstDash val="sysDot"/>')
    const solid = await slideXmlOf(pptBytesOf(buildPptStreams(outlined({ 0x01ce: 0 }))))
    expect(solid).not.toContain('prstDash')
    expect(solid).not.toContain('cmpd')
  })
})
