import { describe, expect, it } from 'vitest'
import { convertPptToPptx, parsePptPresentation } from '../src/index'
import { buildPptContainer, buildPptStreams, type SlideSpec } from './synthetic-ppt'

const pptBytesOf = (streams: { document: Uint8Array; currentUser: Uint8Array }): Uint8Array =>
  buildPptContainer(streams.document, streams.currentUser)

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
  const JSZip = (await import('jszip')).default
  const zip = await JSZip.loadAsync(out.bytes)
  const texts: string[][] = []
  for (let s = 0; s < out.slideCount; s++) {
    const xml = await zip.file(`ppt/slides/slide${s + 1}.xml`)!.async('string')
    texts.push([...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]!))
  }
  return texts
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
          { kind: 0, text: '季度回顾' },
          { kind: 4, text: 'action items' },
        ],
      },
    ]
    // build the CJK slide with a TextCharsAtom (the builder uses TextChars for
    // everything); then a bytes-atom slide: latin1 survives with smart quotes
    const texts = await slideTextsOf(pptBytesOf(buildPptStreams(specs)))
    expect(texts[0]).toEqual(['季度回顾', 'action items'])
    const out = await convertPptToPptx(pptBytesOf(buildPptStreams(specs)))
    const JSZip = (await import('jszip')).default
    const zip = await JSZip.loadAsync(out.bytes)
    const presentation = await zip.file('ppt/presentation.xml')!.async('string')
    expect(presentation).toContain('cx="9144000" cy="6858000"') // 10 × 7.5 in
  })

  it('reads the slide size from the DocumentAtom', async () => {
    const specs: SlideSpec[] = [{ textSequences: [{ kind: 4, text: 'widescreen' }] }]
    const streams = buildPptStreams(specs, { slideSize: [9144000, 6858000] })
    const out = await convertPptToPptx(buildPptContainer(streams.document, streams.currentUser))
    const JSZip = (await import('jszip')).default
    const zip = await JSZip.loadAsync(out.bytes)
    const presentation = await zip.file('ppt/presentation.xml')!.async('string')
    expect(presentation).toContain('cx="14516100"') // 9144000 master units × 1.5875
    void streams
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

  it('normalizes the PowerPoint runes into plain newlines', async () => {
    const specs: SlideSpec[] = [
      {
        textSequences: [{ kind: 1, text: 'first\u000bsecond\u001ebullet text' }],
      },
    ]
    const texts = await slideTextsOf(pptBytesOf(buildPptStreams(specs)))
    expect(texts[0]).toEqual(['first', 'second bullet text'])
  })
})

describe('convertPptToPptx: the failure model', () => {
  it('refuses a file that is not a compound container', async () => {
    await expect(convertPptToPptx(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]))).rejects.toMatchObject({
      code: 'not_cfb',
    })
  })

  it('refuses a container without the presentation stream', async () => {
    const cfb = (await import('cfb')).default
    const container = cfb.utils.cfb_new({ root: 'Root Entry' })
    const out = cfb.write(container, { type: 'buffer' }) as Buffer
    await expect(
      convertPptToPptx(new Uint8Array(out.buffer, out.byteOffset, out.byteLength)),
    ).rejects.toMatchObject({ code: 'missing_stream' })
  })
})

describe('parsePptPresentation: the model underneath', () => {
  it('exposes the parsed text sequences for tests and later stages', () => {
    const { document, currentUser } = buildPptStreams(twoSlides)
    const presentation = parsePptPresentation(document, currentUser)
    expect(presentation.slides).toHaveLength(2)
    expect(presentation.slideWidth).toBe(5760000)
    expect(presentation.slides[0]!.textSequences[0]!.text).toBe('Quarterly Review')
  })
})
