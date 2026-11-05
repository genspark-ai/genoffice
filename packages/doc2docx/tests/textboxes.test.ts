import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '@genoffice/docx-engine'
import {
  convertLegacyDocToDocx,
  openLegacyDoc,
  parseFloatingPictures,
  parseTextBoxStories,
  readLegacyDocBody,
  readLegacyDocLayout,
  type FloatingTextBox,
} from '../src/index'
import {
  SYNTHETIC_TEXT_OFFSET,
  buildSyntheticDoc,
  encodeChpxFkp,
  type SyntheticDocOptions,
} from './helpers/synthetic-doc'

const BODY = 'Anchor \u0008here\rNext\r'

async function zipText(docx: Uint8Array, path: string): Promise<string | undefined> {
  const zip = await JSZip.loadAsync(docx)
  const file = zip.file(path)
  return file ? file.async('string') : undefined
}

function textBoxDoc(options: Partial<SyntheticDocOptions> = {}): Uint8Array {
  return buildSyntheticDoc({
    pieces: [{ text: BODY }],
    floatingShapes: [{ story: 'body', cp: 7, spid: 1025, rect: [1440, 720, 4320, 2160] }],
    textBoxStories: [{ spid: 1025, text: 'Hello\rWorld' }],
    ...options,
  })
}

function shapesOf(bytes: Uint8Array) {
  const doc = openLegacyDoc(bytes)
  return parseFloatingPictures(doc, parseTextBoxStories(doc))
}

function boxOf(bytes: Uint8Array, index = 0): FloatingTextBox {
  const shape = shapesOf(bytes).body[index]
  if (shape?.kind !== 'textBox') throw new Error('expected a text box')
  return shape
}

describe('text boxes: PlcftxbxTxt stories matched to OfficeArt shapes', () => {
  it('matches a story to its shape by spid and keeps the shape geometry', () => {
    const result = shapesOf(textBoxDoc())
    expect(result.warnings).toEqual([])
    expect(result.body).toHaveLength(1)
    expect(result.body[0]).toMatchObject({
      kind: 'textBox',
      cp: 7,
      spid: 1025,
      leftTwips: 1440,
      topTwips: 720,
      widthTwips: 2880,
      heightTwips: 1440,
      geometry: 'rect',
      fill: 'FFFFFF',
      line: { color: '000000', widthEmu: 9525 },
      insetsEmu: [91440, 45720, 91440, 45720],
      verticalAnchor: 't',
      autoFit: false,
    })
    expect(boxOf(textBoxDoc()).content.paragraphs).toEqual(['Hello', 'World'])
  })

  it('reads fill, line, insets and anchoring from the property table', () => {
    const box = boxOf(
      textBoxDoc({
        floatingShapes: [
          {
            story: 'body',
            cp: 7,
            spid: 1025,
            rect: [0, 0, 1000, 500],
            props: [
              [0x0181, 0x000000ff],
              [0x01c0, 0x0000ff00],
              [0x01cb, 19050],
              [0x0081, 0],
              [0x0082, 0],
              [0x0087, 1],
              [0x00bf, 0x00020002],
            ],
          },
        ],
      }),
    )
    expect(box).toMatchObject({
      fill: 'FF0000',
      line: { color: '00FF00', widthEmu: 19050 },
      insetsEmu: [0, 0, 91440, 45720],
      verticalAnchor: 'ctr',
      autoFit: true,
    })
  })

  it('drops the fill and the line when their use bits turn them off', () => {
    const box = boxOf(
      textBoxDoc({
        floatingShapes: [
          {
            story: 'body',
            cp: 7,
            spid: 1025,
            rect: [0, 0, 1000, 500],
            props: [
              [0x01bf, 0x00100000],
              [0x01ff, 0x00080000],
            ],
          },
        ],
      }),
    )
    expect(box.fill).toBeUndefined()
    expect(box.line).toBeUndefined()
  })

  it('falls back to the lTxid story index and reports linked continuation boxes', () => {
    const result = shapesOf(
      textBoxDoc({
        floatingShapes: [
          {
            story: 'body',
            cp: 7,
            spid: 2000,
            rect: [0, 0, 1000, 500],
            props: [[0x0080, 0x00010000]],
          },
          {
            story: 'body',
            cp: 12,
            spid: 2001,
            rect: [0, 0, 1000, 500],
            props: [[0x0080, 0x00010001]],
          },
        ],
      }),
    )
    expect(result.body.map((shape) => shape.kind)).toEqual(['textBox', 'textBox'])
    expect((result.body[0] as FloatingTextBox).content.paragraphs).toEqual(['Hello', 'World'])
    expect((result.body[1] as FloatingTextBox).content.paragraphs).toEqual([''])
    expect(result.warnings).toEqual([
      'a linked text box continues another box; its text stays in the first box',
    ])
  })

  it('keeps the text in the lid shape even when a continuation is anchored first', () => {
    const result = shapesOf(
      textBoxDoc({
        floatingShapes: [
          {
            story: 'body',
            cp: 7,
            spid: 2001,
            rect: [0, 0, 1000, 500],
            props: [[0x0080, 0x00010001]],
          },
          {
            story: 'body',
            cp: 12,
            spid: 1025,
            rect: [0, 0, 1000, 500],
            props: [[0x0080, 0x00010000]],
          },
        ],
      }),
    )
    expect(
      result.body.map((shape) => [shape.spid, (shape as FloatingTextBox).content.paragraphs]),
    ).toEqual([
      [2001, ['']],
      [1025, ['Hello', 'World']],
    ])
  })

  it('skips reusable slots and warns about stories no shape claims', () => {
    const result = shapesOf(
      textBoxDoc({
        textBoxStories: [
          { spid: 7, text: 'deleted', reusable: true },
          { spid: 1025, text: 'Hello' },
          { spid: 4000, text: 'orphan' },
        ],
      }),
    )
    expect(result.body).toHaveLength(1)
    expect(boxOf(textBoxDoc()).content.paragraphs).toEqual(['Hello', 'World'])
    expect(result.warnings).toEqual(['a text box story has no shape anchor; its text is omitted'])
  })

  it('lets a picture shape own an (empty) story without a warning', () => {
    const png = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0])
    const bytes = textBoxDoc({
      floatingShapes: [
        {
          story: 'body',
          cp: 7,
          spid: 1025,
          rect: [0, 0, 1000, 500],
          blip: { kind: 'png', bytes: png },
        },
      ],
      textBoxStories: [{ spid: 1025, text: '' }],
    })
    const result = shapesOf(bytes)
    expect(result.body.map((shape) => shape.kind)).toEqual(['picture'])
    expect(result.warnings).toEqual([])
  })

  it('degrades to a warning when the text is declared but the table is missing', () => {
    const doc = openLegacyDoc(buildSyntheticDoc({ pieces: [{ text: BODY }], ccpTxbx: 5 }))
    const stories = parseTextBoxStories(doc)
    expect(stories.body).toEqual([])
    expect(stories.warnings).toEqual([
      'the document declares body text box text but has no text box table; those text boxes are omitted',
    ])
  })

  it('resolves run formatting inside the box like a header story', () => {
    const storyCp = BODY.length
    const bytes = textBoxDoc({
      chpxFkpPages: [
        encodeChpxFkp([
          {
            fcStart: SYNTHETIC_TEXT_OFFSET + storyCp * 2,
            fcEnd: SYNTHETIC_TEXT_OFFSET + (storyCp + 5) * 2,
            grpprl: Uint8Array.from([0x35, 0x08, 1]),
          },
        ]),
      ],
    })
    const body = readLegacyDocBody(bytes)
    const block = body.blocks.find((entry) => entry.type === 'paragraph' && entry.pictures)
    const box = block?.type === 'paragraph' ? block.pictures?.floating?.[0] : undefined
    if (box?.kind !== 'textBox') throw new Error('expected a text box in the anchor paragraph')
    expect(box.content.formatted?.map((paragraph) => paragraph.text)).toEqual(['Hello', 'World'])
    expect(box.content.formatted?.[0]?.runs.map((run) => [run.text, run.formatting.bold])).toEqual([
      ['Hello', true],
    ])
  })
})

describe('text boxes: docx output', () => {
  it('writes the box as an anchored wps shape with its paragraphs in w:txbxContent', async () => {
    const { docx, warnings } = await convertLegacyDocToDocx(textBoxDoc())
    expect(warnings).toEqual([])
    const xml = (await zipText(docx, 'word/document.xml'))!
    expect(xml).toMatch(
      /<w:document[^>]* xmlns:wps="http:\/\/schemas\.microsoft\.com\/office\/word\/2010\/wordprocessingShape"/,
    )
    const anchor =
      /<w:r><w:drawing><wp:anchor [^>]*>([\s\S]*?)<\/wp:anchor><\/w:drawing><\/w:r>/.exec(xml)
    expect(anchor).not.toBeNull()
    const body = anchor![1]!
    expect(body).toContain(
      '<wp:positionH relativeFrom="margin"><wp:posOffset>914400</wp:posOffset>',
    )
    expect(body).toContain(
      '<wp:positionV relativeFrom="margin"><wp:posOffset>457200</wp:posOffset>',
    )
    expect(body).toContain('<wp:extent cx="1828800" cy="914400"/>')
    expect(body).toContain('<wp:wrapSquare wrapText="bothSides"/>')
    expect(body).toContain(
      '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>',
    )
    expect(body).toContain('<wps:cNvSpPr txBox="1"/>')
    expect(body).toContain(
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>',
    )
    expect(body).toMatch(
      /<wps:txbx><w:txbxContent><w:p>(?:<w:pPr>.*?<\/w:pPr>)?<w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:t[^>]*>Hello<\/w:t><\/w:r><\/w:p><w:p>.*?World.*?<\/w:p><\/w:txbxContent><\/wps:txbx>/,
    )
    expect(body).toContain(
      '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0" upright="1"><a:noAutofit/></wps:bodyPr>',
    )
    const parsed = await parseDocx(docx)
    const boxes = parsed.blocks.flatMap((block) =>
      'textboxes' in block ? (block.textboxes ?? []) : [],
    )
    expect(boxes).toHaveLength(1)
    expect(JSON.stringify(boxes[0])).toContain('Hello')
    expect(boxes[0]).toMatchObject({ fill: 'FFFFFF', borderColor: '000000' })
  })

  it('writes a:noFill for unfilled boxes and keeps behind-text, page-relative placement', async () => {
    const { docx } = await convertLegacyDocToDocx(
      textBoxDoc({
        floatingShapes: [
          {
            story: 'body',
            cp: 7,
            spid: 1025,
            rect: [0, 0, 1000, 500],
            bx: 1,
            by: 1,
            wr: 3,
            belowText: true,
            props: [
              [0x01bf, 0x00100000],
              [0x01ff, 0x00080000],
            ],
          },
        ],
      }),
    )
    const xml = (await zipText(docx, 'word/document.xml'))!
    expect(xml).toContain('behindDoc="1"')
    expect(xml).toContain('<wp:positionH relativeFrom="page">')
    expect(xml).toContain('<wp:positionV relativeFrom="page">')
    expect(xml).toContain('<wp:wrapNone/>')
    expect(xml).toContain(
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>',
    )
    await expect(parseDocx(docx)).resolves.toBeDefined()
  })

  it('puts a header-anchored text box into the header part', async () => {
    const bytes = buildSyntheticDoc({
      pieces: [{ text: 'Body\r' }],
      headerStories: [undefined, 'Head \u0008er', undefined, undefined, undefined, undefined],
      floatingShapes: [{ story: 'header', cp: 5, spid: 3001, rect: [100, 200, 1100, 700], by: 2 }],
      textBoxStories: [{ story: 'header', spid: 3001, text: 'Boxed' }],
    })
    const layout = readLegacyDocLayout(bytes)
    const pictures = layout.headerFooter[0]!.oddHeader!.pictures!.get(0)!
    expect(pictures[0]).toMatchObject({ kind: 'textBox', story: 'header', wrap: 'none' })
    const { docx, warnings } = await convertLegacyDocToDocx(bytes)
    expect(warnings).toEqual([])
    const headerXml = (await zipText(docx, 'word/header1.xml'))!
    expect(headerXml).toContain(
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
    )
    expect(headerXml).toMatch(/<w:txbxContent><w:p>.*?Boxed.*?<\/w:p><\/w:txbxContent>/)
    expect(headerXml).toContain(
      '<wp:positionV relativeFrom="page"><wp:posOffset>127000</wp:posOffset>',
    )
    expect((await zipText(docx, 'word/document.xml'))!).not.toContain('<wps:')
    await expect(parseDocx(docx)).resolves.toBeDefined()
  })
})
