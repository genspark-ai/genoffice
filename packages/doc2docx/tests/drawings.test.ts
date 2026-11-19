import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '@genoffice/docx-engine'
import {
  convertLegacyDocToDocx,
  openLegacyDoc,
  parseFloatingPictures,
  readLegacyDocBody,
  readLegacyDocLayout,
} from '../src/index'
import {
  SYNTHETIC_TEXT_OFFSET,
  buildSyntheticDoc,
  encodeBlipRecord,
  encodeChpxFkp,
  encodePicf,
  encodePictureChpx,
  type SyntheticBlip,
  type SyntheticFloatingShape,
} from './helpers/synthetic-doc'
import type { FloatingPicture, FloatingShape } from '../src/index'

function mediaOf(shape: FloatingShape | undefined): FloatingPicture['media'] {
  if (shape?.kind !== 'picture') throw new Error('expected a picture shape')
  return shape.media
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

/** placeable WMF header + a couple of records; only the bytes matter here */
function fakeWmf(): Uint8Array {
  const out = new Uint8Array(64)
  out.set([0xd7, 0xcd, 0xc6, 0x9a], 0)
  for (let index = 22; index < out.length; index++) out[index] = index & 0xff
  return out
}

/** BITMAPINFOHEADER (40 bytes) for a 1x1 24-bit bitmap plus its padded pixel row */
function fakeDib(): Uint8Array {
  const out = new Uint8Array(44)
  const view = new DataView(out.buffer)
  view.setUint32(0, 40, true)
  view.setInt32(4, 1, true)
  view.setInt32(8, 1, true)
  view.setUint16(12, 1, true)
  view.setUint16(14, 24, true)
  out.set([0x11, 0x22, 0x33, 0x00], 40)
  return out
}

const PNG: SyntheticBlip = { kind: 'png', bytes: minimalPng() }

function floatingDoc(
  shapes: SyntheticFloatingShape[],
  options: { text?: string; headerStories?: Array<string | undefined> } = {},
): Uint8Array {
  return buildSyntheticDoc({
    pieces: [{ text: options.text ?? 'Anchor \u0008here\rNext\r' }],
    floatingShapes: shapes,
    ...(options.headerStories ? { headerStories: options.headerStories } : {}),
  })
}

describe('floating pictures: FSPA anchors + OfficeArt blips', () => {
  it('reads a body picture with its rectangle, relative-to and wrapping flags', () => {
    const bytes = floatingDoc([
      { story: 'body', cp: 7, rect: [1440, 720, 4320, 2160], bx: 1, by: 0, wr: 1, blip: PNG },
    ])
    const doc = openLegacyDoc(bytes)
    const result = parseFloatingPictures(doc)
    expect(result.warnings).toEqual([])
    expect(result.header).toEqual([])
    expect(result.body).toHaveLength(1)
    expect(result.body[0]).toMatchObject({
      cp: 7,
      story: 'body',
      leftTwips: 1440,
      topTwips: 720,
      widthTwips: 2880,
      heightTwips: 1440,
      horizontalRelative: 'page',
      verticalRelative: 'margin',
      wrap: 'topAndBottom',
      behindText: false,
    })
    expect(mediaOf(result.body[0]).mediaType).toBe('image/png')
    expect([...mediaOf(result.body[0]).bytes]).toEqual([...minimalPng()])
  })

  it("takes the shape's posRelH and posRelV over the FSPA's bx and by", async () => {
    const bytes = floatingDoc([
      {
        story: 'body',
        cp: 7,
        rect: [907, 461, 1341, 895],
        bx: 2,
        by: 2,
        wr: 3,
        blip: PNG,
        props: [
          [0x0390, 1],
          [0x0392, 2],
        ],
      },
      {
        story: 'body',
        cp: 7,
        spid: 2000,
        rect: [0, 0, 720, 720],
        bx: 0,
        by: 0,
        blip: PNG,
        props: [
          [0x0390, 3],
          [0x0392, 3],
        ],
      },
    ])
    const result = parseFloatingPictures(openLegacyDoc(bytes))
    expect(result.body[0]).toMatchObject({
      horizontalRelative: 'page',
      verticalRelative: 'paragraph',
    })
    expect(result.body[1]).toMatchObject({
      horizontalRelative: 'character',
      verticalRelative: 'line',
    })
    const { docx } = await convertLegacyDocToDocx(bytes)
    const xml = (await zipText(docx, 'word/document.xml'))!
    expect(xml).toContain('<wp:positionH relativeFrom="page"><wp:posOffset>575945</wp:posOffset>')
    expect(xml).toContain('<wp:positionH relativeFrom="character">')
    expect(xml).toContain('<wp:positionV relativeFrom="line">')
  })

  it('reads cropFromTop/Bottom/Left/Right as a:srcRect thousandths and ignores a full crop', async () => {
    const bytes = floatingDoc([
      {
        story: 'body',
        cp: 7,
        rect: [1440, 720, 4320, 2160],
        blip: PNG,
        props: [
          [0x0100, 0x4000],
          [0x0101, 0x2000],
          [0x0103, 0x0800],
        ],
      },
      {
        story: 'body',
        cp: 7,
        spid: 2000,
        rect: [0, 0, 720, 720],
        blip: PNG,
        props: [
          [0x0102, 0x8000],
          [0x0103, 0x8000],
        ],
      },
    ])
    const result = parseFloatingPictures(openLegacyDoc(bytes))
    expect(result.body).toHaveLength(2)
    expect(result.body[0]).toMatchObject({ crop: { l: 0, t: 25000, r: 3125, b: 12500 } })
    expect(result.body[1]).not.toHaveProperty('crop')
    const { docx } = await convertLegacyDocToDocx(bytes)
    const xml = (await zipText(docx, 'word/document.xml'))!
    expect(xml).toContain('<a:srcRect t="25000" r="3125" b="12500"/>')
  })

  it('attaches the picture to the anchor paragraph and writes a wp:anchor drawing', async () => {
    const bytes = floatingDoc([
      {
        story: 'body',
        cp: 7,
        rect: [1440, 720, 4320, 2160],
        bx: 2,
        by: 2,
        wr: 3,
        belowText: true,
        blip: PNG,
      },
    ])
    const body = readLegacyDocBody(bytes)
    const first = body.blocks[0]!
    if (first.type !== 'paragraph') throw new Error('expected a paragraph')
    expect(first.paragraph.text).toBe('Anchor here')
    expect(first.pictures?.floating).toHaveLength(1)
    const { docx, warnings } = await convertLegacyDocToDocx(bytes)
    expect(warnings.some((warning) => warning.includes('omitted'))).toBe(false)
    const xml = (await zipText(docx, 'word/document.xml'))!
    expect(xml).toContain(
      '<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658241" behindDoc="1"',
    )
    expect(xml).toContain(
      '<wp:positionH relativeFrom="column"><wp:posOffset>914400</wp:posOffset></wp:positionH>',
    )
    expect(xml).toContain(
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>457200</wp:posOffset></wp:positionV>',
    )
    expect(xml).toContain('<wp:extent cx="1828800" cy="914400"/>')
    expect(xml).toContain('<wp:wrapNone/>')
    expect(xml).toMatch(/<\/wp:anchor><\/w:drawing><\/w:r><w:r><w:t[^>]*>Anchor here<\/w:t>/)
    const rels = (await zipText(docx, 'word/_rels/document.xml.rels'))!
    expect(rels).toMatch(/Type="[^"]+\/image" Target="media\/image1\.png"/)
    expect([...(await zipBytes(docx, 'word/media/image1.png'))!]).toEqual([...minimalPng()])
    await expect(parseDocx(docx)).resolves.toBeDefined()
  })

  it('writes a header picture into the header part with its own relationships', async () => {
    const wmf = fakeWmf()
    const bytes = floatingDoc(
      [
        {
          story: 'header',
          cp: 0,
          rect: [-216, 972, 1344, 1874],
          bx: 2,
          by: 2,
          wr: 1,
          blip: { kind: 'wmf', bytes: wmf },
        },
      ],
      { headerStories: [undefined, '\u0008', undefined, 'Footer', undefined, undefined] },
    )
    const layout = readLegacyDocLayout(bytes)
    const header = layout.headerFooter[0]!.oddHeader!
    expect(header.pictures?.get(0)).toHaveLength(1)
    expect(header.pictures!.get(0)![0]!.wrap).toBe('none')
    const { docx, warnings } = await convertLegacyDocToDocx(bytes)
    expect(warnings.some((warning) => warning.includes('omitted'))).toBe(false)
    const headerXml = (await zipText(docx, 'word/header1.xml'))!
    expect(headerXml).toContain('xmlns:wp=')
    expect(headerXml).toContain('<wp:anchor')
    expect(headerXml).toContain('<wp:wrapNone/>')
    expect(headerXml).toContain(
      '<wp:positionV relativeFrom="page"><wp:posOffset>617220</wp:posOffset>',
    )
    expect(headerXml).toContain('<wp:extent cx="990600" cy="572770"/>')
    expect(headerXml).toContain('r:embed="rId1"')
    const headerRels = (await zipText(docx, 'word/_rels/header1.xml.rels'))!
    expect(headerRels).toMatch(/Id="rId1" Type="[^"]+\/image" Target="media\/image1\.wmf"/)
    expect(await zipText(docx, 'word/_rels/footer1.xml.rels')).toBeUndefined()
    const contentTypes = (await zipText(docx, '[Content_Types].xml'))!
    expect(contentTypes).toContain('<Default Extension="wmf" ContentType="image/x-wmf"/>')
    expect([...(await zipBytes(docx, 'word/media/image1.wmf'))!]).toEqual([...wmf])
    await expect(parseDocx(docx)).resolves.toBeDefined()
  })

  it('keeps a header story that holds nothing but a picture anchor', () => {
    const bytes = floatingDoc([{ story: 'header', cp: 0, rect: [0, 0, 100, 100], blip: PNG }], {
      headerStories: [undefined, '\u0008', undefined, undefined, undefined, undefined],
    })
    const layout = readLegacyDocLayout(bytes)
    expect(layout.headerFooter[0]!.oddHeader?.paragraphs).toEqual([''])
    expect(layout.headerFooter[0]!.oddHeader?.pictures?.size).toBe(1)
  })

  it('turns a DIB blip into a .bmp and keeps an uncompressed EMF as is', async () => {
    const emf = Uint8Array.from([
      0x01, 0x00, 0x00, 0x00, 0x6c, 0x00, 0x00, 0x00, 0x20, 0x45, 0x4d, 0x46,
    ])
    const bytes = floatingDoc([
      { story: 'body', cp: 7, rect: [0, 0, 200, 100], blip: { kind: 'dib', bytes: fakeDib() } },
      {
        story: 'body',
        cp: 12,
        rect: [0, 0, 200, 100],
        blip: { kind: 'emf', bytes: emf, uncompressed: true, embedded: true },
      },
    ])
    const result = parseFloatingPictures(openLegacyDoc(bytes))
    expect(result.warnings).toEqual([])
    expect(result.body.map((picture) => mediaOf(picture).mediaType)).toEqual([
      'image/bmp',
      'image/x-emf',
    ])
    const bmp = mediaOf(result.body[0]).bytes
    expect(String.fromCharCode(bmp[0]!, bmp[1]!)).toBe('BM')
    expect(new DataView(bmp.buffer, bmp.byteOffset).getUint32(2, true)).toBe(14 + 44)
    expect(new DataView(bmp.buffer, bmp.byteOffset).getUint32(10, true)).toBe(54)
    expect([...bmp.slice(14)]).toEqual([...fakeDib()])
    expect([...mediaOf(result.body[1]).bytes]).toEqual([...emf])
    const { docx } = await convertLegacyDocToDocx(bytes)
    const contentTypes = (await zipText(docx, '[Content_Types].xml'))!
    expect(contentTypes).toContain('<Default Extension="bmp" ContentType="image/bmp"/>')
    expect(contentTypes).toContain('<Default Extension="emf" ContentType="image/x-emf"/>')
  })

  it('places the pictures of a group through the group coordinate space', () => {
    const bytes = floatingDoc([
      {
        story: 'body',
        cp: 7,
        rect: [1000, 2000, 3000, 3000],
        group: {
          bounds: [0, 0, 200, 100],
          children: [
            { anchor: [0, 0, 100, 100], blip: PNG },
            {
              anchor: [100, 50, 200, 100],
              blip: { kind: 'jpeg', bytes: Uint8Array.from([0xff, 0xd8, 0xff, 0xd9]) },
            },
            { anchor: [0, 0, 50, 50] },
          ],
        },
      },
    ])
    const result = parseFloatingPictures(openLegacyDoc(bytes))
    expect(result.warnings).toEqual([
      'floating shapes other than pictures and text boxes are not converted',
    ])
    expect(
      result.body.map((picture) => [
        picture.leftTwips,
        picture.topTwips,
        picture.widthTwips,
        picture.heightTwips,
      ]),
    ).toEqual([
      [1000, 2000, 1000, 1000],
      [2000, 2500, 1000, 500],
    ])
  })

  it('skips non-picture shapes with one warning and degrades broken blips per picture', () => {
    const bytes = floatingDoc([
      { story: 'body', cp: 7, rect: [0, 0, 100, 100] },
      { story: 'body', cp: 7, rect: [0, 0, 100, 100], blip: { ...PNG, recordLength: 100000 } },
      { story: 'body', cp: 7, rect: [0, 0, 100, 100], pib: 99 },
      {
        story: 'body',
        cp: 7,
        rect: [0, 0, 100, 100],
        blip: { kind: 'pict', bytes: Uint8Array.from([1, 2, 3]) },
      },
      { story: 'body', cp: 7, rect: [0, 0, 100, 100], blip: PNG },
    ])
    const result = parseFloatingPictures(openLegacyDoc(bytes))
    expect(result.body).toHaveLength(1)
    expect(result.warnings).toContain(
      'floating shapes other than pictures and text boxes are not converted',
    )
    expect(result.warnings).toContain('a picture record is truncated and was omitted')
    expect(result.warnings).toContain(
      'a floating picture references a missing image and was omitted',
    )
    expect(result.warnings).toContain('a picture uses an unsupported format (PICT) and was omitted')
    expect(mediaOf(result.body[0]).bytes.length).toBe(minimalPng().length)
  })

  it('ignores documents without shape tables', () => {
    const result = parseFloatingPictures(
      openLegacyDoc(buildSyntheticDoc({ pieces: [{ text: 'x\r' }] })),
    )
    expect(result).toEqual({ body: [], header: [], warnings: [] })
  })
})

describe('inline pictures through OfficeArt blip records', () => {
  function inlineDoc(payload: Uint8Array, text = '\u0001\r'): Uint8Array {
    return buildSyntheticDoc({
      pieces: [{ text }],
      chpxFkpPages: [
        encodeChpxFkp([
          {
            fcStart: SYNTHETIC_TEXT_OFFSET + text.indexOf('\u0001') * 2,
            fcEnd: SYNTHETIC_TEXT_OFFSET + text.indexOf('\u0001') * 2 + 2,
            grpprl: encodePictureChpx(0),
          },
        ]),
      ],
      dataStream: encodePicf({ payload, xExt: 1440, yExt: 720 }),
    })
  }

  it('decodes a compressed WMF blip inside the PICF', async () => {
    const wmf = fakeWmf()
    const bytes = inlineDoc(encodeBlipRecord({ kind: 'wmf', bytes: wmf }))
    const body = readLegacyDocBody(bytes)
    const image = body.blocks.find((block) => block.type === 'image')
    if (image?.type !== 'image') throw new Error('expected an image block')
    expect(image.image.mediaType).toBe('image/x-wmf')
    expect([...image.image.bytes]).toEqual([...wmf])
    const { docx } = await convertLegacyDocToDocx(bytes)
    expect(await zipBytes(docx, 'word/media/image1.wmf')).toBeDefined()
    expect((await zipText(docx, 'word/document.xml'))!).toContain('<wp:inline')
  })

  it('skips a truncated JPEG with a warning instead of failing the import', async () => {
    const bytes = inlineDoc(
      Uint8Array.from([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x04, 0x01, 0x02]),
      'A\u0001\rB\r',
    )
    const body = readLegacyDocBody(bytes)
    expect(body.blocks.some((block) => block.type === 'image')).toBe(false)
    expect(body.warnings).toContain('a JPEG picture is truncated (no EOI marker) and was omitted')
    const { docx } = await convertLegacyDocToDocx(bytes)
    expect((await zipText(docx, 'word/document.xml'))!).toContain('>B<')
  })

  it('keeps picture characters that sit in a field result', () => {
    const text = '\u0013 EMBED Word.Picture.8 \u0014\u0001\u0015\r'
    const bytes = inlineDoc(encodeBlipRecord(PNG), text)
    const body = readLegacyDocBody(bytes)
    const first = body.blocks[0]!
    if (first.type !== 'paragraph') throw new Error('expected a paragraph')
    expect(first.pictures?.inline?.map((entry) => entry.position)).toEqual(['start'])
  })
})
