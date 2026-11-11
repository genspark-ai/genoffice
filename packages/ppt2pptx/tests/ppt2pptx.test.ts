import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import * as CFB from 'cfb'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { convertPptToPptx, parsePptPresentation } from '../src/index'
import { buildPersistChain } from '../src/ppt'
import { buildPptContainer, buildPptStreams, type SlideSpec } from './synthetic-ppt'

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
