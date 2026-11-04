import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '@genoffice/docx-engine'
import {
  LegacyDocError,
  convertLegacyDocToDocx,
  openLegacyDoc,
  parseInlineImages,
  readLegacyDocBody,
  readLegacyDocLayout,
} from '../src/index'
import {
  SYNTHETIC_TEXT_OFFSET,
  buildSyntheticDoc,
  encodeChpxFkp,
  encodePapxFkp,
  encodePicf,
  encodePictureChpx,
  encodeSprm,
} from './helpers/synthetic-doc'

function fixtureBytes(name: string): Uint8Array {
  return new Uint8Array(readFileSync(fileURLToPath(new URL('./fixtures/' + name, import.meta.url))))
}

async function zipText(docx: Uint8Array, path: string): Promise<string | undefined> {
  const zip = await JSZip.loadAsync(docx)
  const file = zip.file(path)
  return file ? file.async('string') : undefined
}

async function zipBytes(docx: Uint8Array, path: string): Promise<Uint8Array | undefined> {
  const zip = await JSZip.loadAsync(docx)
  const file = zip.file(path)
  return file ? file.async('uint8array') : undefined
}

function errorCode(run: () => unknown): string {
  try {
    run()
  } catch (error) {
    expect(error).toBeInstanceOf(LegacyDocError)
    return (error as LegacyDocError).code
  }
  throw new Error('expected a LegacyDocError')
}

/** a minimal but structurally valid 1x1 PNG chunk sequence (CRCs are not checked by the reader) */
function minimalPng(): Uint8Array {
  const signature = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
  const ihdr = new Uint8Array(8 + 13 + 4)
  const ihdrView = new DataView(ihdr.buffer)
  ihdrView.setUint32(0, 13, false)
  ihdr.set([0x49, 0x48, 0x44, 0x52], 4)
  ihdrView.setUint32(8, 1, false)
  ihdrView.setUint32(12, 1, false)
  const iend = Uint8Array.from([0, 0, 0, 0, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82])
  const out = new Uint8Array(signature.length + ihdr.length + iend.length)
  out.set(signature, 0)
  out.set(ihdr, signature.length)
  out.set(iend, signature.length + ihdr.length)
  return out
}

/** a minimal JPEG: SOI, an APP0 marker, then EOI */
function minimalJpeg(): Uint8Array {
  return Uint8Array.from([
    0xff, 0xd8, 0xff, 0xe0, 0x00, 0x04, 0x00, 0x00, 0x00, 0x01, 0x02, 0x03, 0xff, 0xd9,
  ])
}

function syntheticImageDoc(options: {
  text: string
  pictures: Array<{ cp: number; fcPic: number; sprms?: Uint8Array }>
  data?: Uint8Array
}): Uint8Array {
  const runs = options.pictures.map((picture) => ({
    fcStart: SYNTHETIC_TEXT_OFFSET + picture.cp * 2,
    fcEnd: SYNTHETIC_TEXT_OFFSET + picture.cp * 2 + 2,
    grpprl: picture.sprms ?? encodePictureChpx(picture.fcPic),
  }))
  return buildSyntheticDoc({
    pieces: [{ text: options.text }],
    chpxFkpPages: [encodeChpxFkp(runs)],
    ...(options.data ? { dataStream: options.data } : {}),
  })
}

describe('image.doc: inline picture reconstruction and writer output', () => {
  const bytes = fixtureBytes('image.doc')
  const data = openLegacyDoc(bytes).stream('Data')!
  const expectedPng = data.slice(247, 317)
  const body = readLegacyDocBody(bytes)

  it('reads the 0.5-inch PNG picture paragraph between the two text paragraphs', () => {
    expect(body.blocks.map((block) => block.type)).toEqual([
      'paragraph',
      'image',
      'paragraph',
      'paragraph',
    ])
    const imageBlock = body.blocks[1]!
    if (imageBlock.type !== 'image') throw new Error('expected an image block')
    expect(imageBlock.image).toMatchObject({
      cp: 13,
      mediaType: 'image/png',
      widthTwips: 720,
      heightTwips: 720,
    })
    expect(imageBlock.image.bytes.length).toBe(70)
    expect([...imageBlock.image.bytes.slice(0, 8)]).toEqual([
      0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
    ])
    expect([...imageBlock.image.bytes.slice(-8)]).toEqual([
      0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
    ])
    expect(body.warnings).toEqual([])
  })

  it('writes media part, content type default, relationship and the inline drawing', async () => {
    const { docx, warnings } = await convertLegacyDocToDocx(bytes)
    expect(warnings).toEqual([])
    const media = (await zipBytes(docx, 'word/media/image1.png'))!
    expect([...media]).toEqual([...expectedPng])
    const contentTypes = (await zipText(docx, '[Content_Types].xml'))!
    expect(contentTypes).toContain('<Default Extension="png" ContentType="image/png"/>')
    const rels = (await zipText(docx, 'word/_rels/document.xml.rels'))!
    expect(rels).toMatch(/Type="[^"]+\/image" Target="media\/image1\.png"/)
    const xml = (await zipText(docx, 'word/document.xml'))!
    expect(xml).toContain('<w:drawing>')
    expect(xml).toContain('<wp:extent cx="457200" cy="457200"/>')
    expect(xml).toContain('<wp:docPr id="1" name="Picture 1"/>')
    expect(xml).toContain('r:embed="rId3"')
    expect(xml.indexOf('Before image')).toBeLessThan(xml.indexOf('<w:drawing>'))
    expect(xml.indexOf('<w:drawing>')).toBeLessThan(xml.indexOf('After image'))
    await expect(parseDocx(docx)).resolves.toBeDefined()
  })
})

describe('rich.doc: one picture, table still renders, field picture ignored', () => {
  const bytes = fixtureBytes('rich.doc')

  it('extracts only the picture outside the hyperlink field', () => {
    const body = readLegacyDocBody(bytes)
    const imageBlocks = body.blocks.filter((block) => block.type === 'image')
    expect(imageBlocks).toHaveLength(1)
    if (imageBlocks[0]!.type !== 'image') throw new Error('expected an image block')
    expect(imageBlocks[0]!.image.cp).toBe(177)
    expect(imageBlocks[0]!.image.bytes.length).toBe(70)
    expect(body.blocks.some((block) => block.type === 'table')).toBe(true)
    expect(body.warnings.some((warning) => warning.includes('images are not converted'))).toBe(
      false,
    )
  })

  it('writes one drawing, keeps the stage 3c-2 table and the honest warnings', async () => {
    const { docx, warnings } = await convertLegacyDocToDocx(bytes)
    const xml = (await zipText(docx, 'word/document.xml'))!
    expect((xml.match(/<w:drawing>/g) ?? []).length).toBe(1)
    expect(xml).toContain('<w:tbl>')
    expect((xml.match(/<w:drawing>/g) ?? []).length).toBe(1)
    expect(warnings.some((warning) => warning.includes('hyperlinks are not converted'))).toBe(false)
    expect(warnings.some((warning) => warning.includes('lists are not converted'))).toBe(false)
    await expect(parseDocx(docx)).resolves.toBeDefined()
  })
})

describe('synthetic pictures', () => {
  it('scales the display size from dxaGoal and mx/my (unsigned 16-bit)', async () => {
    const png = minimalPng()
    const data = encodePicf({ payload: png, dxaGoal: 20, dyaGoal: 10, mx: 48000, my: 24000 })
    const docxBytes = syntheticImageDoc({
      text: 'Before\r\u0001\rAfter\r',
      pictures: [{ cp: 7, fcPic: 0 }],
      data,
    })
    const body = readLegacyDocBody(docxBytes)
    const imageBlock = body.blocks.find((block) => block.type === 'image')!
    if (imageBlock.type !== 'image') throw new Error('expected an image block')
    expect(imageBlock.image.widthTwips).toBe(960)
    expect(imageBlock.image.heightTwips).toBe(240)
    const { docx } = await convertLegacyDocToDocx(docxBytes)
    const xml = (await zipText(docx, 'word/document.xml'))!
    expect(xml).toContain('<wp:extent cx="609600" cy="152400"/>')
    expect([...(await zipBytes(docx, 'word/media/image1.png'))!]).toEqual([...png])
  })

  it('falls back to xExt/yExt and dxaGoal when mx/my are zero', () => {
    const png = minimalPng()
    const data = encodePicf({ payload: png, xExt: 1440, yExt: 720, dxaGoal: 100 })
    const docxBytes = syntheticImageDoc({ text: '\u0001\r', pictures: [{ cp: 0, fcPic: 0 }], data })
    const body = readLegacyDocBody(docxBytes)
    const imageBlock = body.blocks.find((block) => block.type === 'image')!
    if (imageBlock.type !== 'image') throw new Error('expected an image block')
    expect(imageBlock.image.widthTwips).toBe(1440)
    expect(imageBlock.image.heightTwips).toBe(720)
  })

  it('extracts a JPEG and writes the jpeg content type default', async () => {
    const jpeg = minimalJpeg()
    const data = encodePicf({ payload: jpeg, xExt: 100, yExt: 100 })
    const docxBytes = syntheticImageDoc({ text: '\u0001\r', pictures: [{ cp: 0, fcPic: 0 }], data })
    const body = readLegacyDocBody(docxBytes)
    const imageBlock = body.blocks.find((block) => block.type === 'image')!
    if (imageBlock.type !== 'image') throw new Error('expected an image block')
    expect(imageBlock.image.mediaType).toBe('image/jpeg')
    const { docx } = await convertLegacyDocToDocx(docxBytes)
    expect(await zipBytes(docx, 'word/media/image1.jpeg')).toBeDefined()
    const contentTypes = (await zipText(docx, '[Content_Types].xml'))!
    expect(contentTypes).toContain('<Default Extension="jpeg" ContentType="image/jpeg"/>')
  })

  it('falls back to the raw JPEG when a blip-looking header does not decode', () => {
    // 0xF01E (PNG blip) header claiming 4000 bytes, then the real JPEG payload
    const fakeBlip = Uint8Array.from([0x00, 0x6e, 0x1e, 0xf0, 0xa0, 0x0f, 0x00, 0x00])
    const payload = new Uint8Array([...fakeBlip, ...minimalJpeg()])
    const data = encodePicf({ payload, xExt: 100, yExt: 100 })
    const docxBytes = syntheticImageDoc({ text: '\u0001\r', pictures: [{ cp: 0, fcPic: 0 }], data })
    const body = readLegacyDocBody(docxBytes)
    const imageBlock = body.blocks.find((block) => block.type === 'image')
    expect(imageBlock?.type).toBe('image')
    if (imageBlock?.type !== 'image') return
    expect(imageBlock.image.mediaType).toBe('image/jpeg')
    expect(body.warnings.some((warning) => warning.includes('omitted'))).toBe(false)
  })

  it('warns instead of failing when the Data stream is missing', async () => {
    const docxBytes = syntheticImageDoc({ text: '\u0001\r', pictures: [{ cp: 0, fcPic: 0 }] })
    const body = readLegacyDocBody(docxBytes)
    expect(body.blocks.some((block) => block.type === 'image')).toBe(false)
    expect(body.warnings.some((warning) => warning.includes('no Data stream'))).toBe(true)
    const { docx, warnings } = await convertLegacyDocToDocx(docxBytes)
    expect((await zipText(docx, 'word/document.xml'))!).not.toContain('<w:drawing>')
    expect(warnings.some((warning) => warning.includes('no Data stream'))).toBe(true)
  })

  it('warns about unsupported EMF/WMF/DIB pictures', async () => {
    const emf = Uint8Array.from([0x01, 0x00, 0x00, 0x00, 0x6c, 0x00, 0x00, 0x00, 0x00, 0x00])
    const data = encodePicf({ payload: emf, xExt: 100, yExt: 100 })
    const docxBytes = syntheticImageDoc({ text: '\u0001\r', pictures: [{ cp: 0, fcPic: 0 }], data })
    const body = readLegacyDocBody(docxBytes)
    expect(body.blocks.some((block) => block.type === 'image')).toBe(false)
    expect(body.warnings.some((warning) => warning.includes('unsupported format'))).toBe(true)
    const { docx } = await convertLegacyDocToDocx(docxBytes)
    expect((await zipText(docx, 'word/document.xml'))!).not.toContain('<w:drawing>')
  })

  it('rejects corrupt PICF headers with bad_picture', () => {
    const png = minimalPng()
    const tooLong = encodePicf({ payload: png, xExt: 100, yExt: 100, lcb: 10000 })
    const badSize = syntheticImageDoc({
      text: '\u0001\r',
      pictures: [{ cp: 0, fcPic: 0 }],
      data: tooLong,
    })
    expect(errorCode(() => readLegacyDocBody(badSize))).toBe('bad_picture')
    const badHeader = encodePicf({ payload: png, xExt: 100, yExt: 100, cbHeader: 0 })
    const badHeaderDoc = syntheticImageDoc({
      text: '\u0001\r',
      pictures: [{ cp: 0, fcPic: 0 }],
      data: badHeader,
    })
    expect(errorCode(() => readLegacyDocBody(badHeaderDoc))).toBe('bad_picture')
    const badOffset = syntheticImageDoc({
      text: '\u0001\r',
      pictures: [{ cp: 0, fcPic: 5000 }],
      data: encodePicf({ payload: png, xExt: 10, yExt: 10 }),
    })
    expect(errorCode(() => readLegacyDocBody(badOffset))).toBe('bad_picture')
  })

  it('ignores picture characters inside field ranges', () => {
    const png = minimalPng()
    const data = encodePicf({ payload: png, xExt: 100, yExt: 100 })
    const text = 'Link: \u0013 HYPERLINK "x"\u0001\u0014result\u0015\r'
    const cp = text.indexOf('\u0001')
    const docxBytes = syntheticImageDoc({ text, pictures: [{ cp, fcPic: 0 }], data })
    const body = readLegacyDocBody(docxBytes)
    expect(body.blocks.some((block) => block.type === 'image')).toBe(false)
  })

  it('writes two pictures with distinct media parts and relationships', async () => {
    const png = minimalPng()
    const first = encodePicf({ payload: png, xExt: 100, yExt: 100 })
    const second = encodePicf({ payload: png, xExt: 200, yExt: 100 })
    const data = new Uint8Array(first.length + second.length)
    data.set(first, 0)
    data.set(second, first.length)
    const docxBytes = syntheticImageDoc({
      text: '\u0001\rText\r\u0001\r',
      pictures: [
        { cp: 0, fcPic: 0 },
        { cp: 7, fcPic: first.length },
      ],
      data,
    })
    const body = readLegacyDocBody(docxBytes)
    expect(body.blocks.filter((block) => block.type === 'image')).toHaveLength(2)
    const { docx } = await convertLegacyDocToDocx(docxBytes)
    expect(await zipBytes(docx, 'word/media/image1.png')).toBeDefined()
    expect(await zipBytes(docx, 'word/media/image2.png')).toBeDefined()
    const xml = (await zipText(docx, 'word/document.xml'))!
    expect((xml.match(/<w:drawing>/g) ?? []).length).toBe(2)
    expect((xml.match(/r:embed="rId\d+"/g) ?? []).length).toBe(2)
  })

  it('writes a picture that shares its paragraph with text as an inline run', async () => {
    const png = minimalPng()
    const data = encodePicf({ payload: png, xExt: 100, yExt: 100 })
    const docxBytes = syntheticImageDoc({
      text: 'A\u0001B\r',
      pictures: [{ cp: 1, fcPic: 0 }],
      data,
    })
    const body = readLegacyDocBody(docxBytes)
    expect(body.blocks.some((block) => block.type === 'image')).toBe(false)
    const paragraph = body.blocks[0]!
    if (paragraph.type !== 'paragraph') throw new Error('expected a paragraph')
    expect(paragraph.pictures?.inline?.map((entry) => entry.position)).toEqual(['end'])
    expect(body.warnings.some((warning) => warning.includes('omitted'))).toBe(false)
    const { docx } = await convertLegacyDocToDocx(docxBytes)
    const xml = (await zipText(docx, 'word/document.xml'))!
    expect(xml).toMatch(/<w:t[^>]*>AB<\/w:t><\/w:r><w:r><w:drawing><wp:inline/)
  })
})

describe('image parse API', () => {
  it('parseInlineImages exposes the Data stream pictures directly', () => {
    const doc = openLegacyDoc(fixtureBytes('image.doc'))
    const result = parseInlineImages(doc)
    expect(result.warnings).toEqual([])
    expect(result.images).toHaveLength(1)
    expect(result.images[0]!.cp).toBe(13)
  })

  it('keeps a header story that holds only an EMBED picture and writes it inline', async () => {
    const png = minimalPng()
    const body = 'Body text\r'
    const header = '\u0013 EMBED Word.Picture.8  \u0014\u0001\u0015'
    const cp = body.length + header.indexOf('\u0001')
    const bytes = buildSyntheticDoc({
      pieces: [{ text: body }],
      headerStories: [undefined, header, undefined, undefined, undefined, undefined],
      chpxFkpPages: [
        encodeChpxFkp([
          {
            fcStart: SYNTHETIC_TEXT_OFFSET + cp * 2,
            fcEnd: SYNTHETIC_TEXT_OFFSET + cp * 2 + 2,
            grpprl: encodePictureChpx(0),
          },
        ]),
      ],
      dataStream: encodePicf({ payload: png, dxaGoal: 2028, dyaGoal: 1176, mx: 1000, my: 1000 }),
    })
    const layout = readLegacyDocLayout(bytes)
    const story = layout.headerFooter[0]!.oddHeader!
    expect(story.paragraphs).toEqual([''])
    expect(story.images?.get(0)).toMatchObject([{ position: 'start' }])
    expect(readLegacyDocBody(bytes).blocks.some((block) => block.type === 'image')).toBe(false)
    const { docx, warnings } = await convertLegacyDocToDocx(bytes)
    expect(warnings).toEqual([])
    const headerXml = (await zipText(docx, 'word/header1.xml'))!
    expect(headerXml).toContain('<wp:inline')
    expect(headerXml).toContain('<wp:extent cx="1287780" cy="746760"/>')
    expect(await zipText(docx, 'word/_rels/header1.xml.rels')).toMatch(
      /Target="media\/image1\.png"/,
    )
    await expect(parseDocx(docx)).resolves.toBeDefined()
  })

  it('keeps the paragraph properties of a picture-only paragraph', async () => {
    const png = minimalPng()
    const text = 'Title\r\u0001\rBody\r'
    const cp = 6
    const bytes = buildSyntheticDoc({
      pieces: [{ text }],
      chpxFkpPages: [
        encodeChpxFkp([
          {
            fcStart: SYNTHETIC_TEXT_OFFSET + cp * 2,
            fcEnd: SYNTHETIC_TEXT_OFFSET + cp * 2 + 2,
            grpprl: encodePictureChpx(0),
          },
        ]),
      ],
      papxFkpPages: [
        encodePapxFkp([
          {
            fcStart: SYNTHETIC_TEXT_OFFSET + cp * 2,
            fcEnd: SYNTHETIC_TEXT_OFFSET + (cp + 2) * 2,
            grpprl: encodeSprm(0x2403, [1]),
          },
        ]),
      ],
      dataStream: encodePicf({ payload: png, xExt: 1440, yExt: 720 }),
    })
    const block = readLegacyDocBody(bytes).blocks.find((entry) => entry.type === 'image')
    if (!block || block.type !== 'image') throw new Error('expected an image block')
    expect(block.paragraph?.formatting.alignment).toBe('center')
    const { docx } = await convertLegacyDocToDocx(bytes)
    const xml = (await zipText(docx, 'word/document.xml'))!
    expect(xml).toMatch(/<w:p><w:pPr><w:jc w:val="center"\/>.*?<\/w:pPr><w:r><w:drawing>/)
  })
})
