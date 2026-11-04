import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { convertLegacyDocToDocx, readLegacyDocBody } from '../src/index'
import {
  SYNTHETIC_TEXT_OFFSET,
  buildSyntheticDoc,
  encodeListParagraphPapx,
  encodePapxFkp,
  encodeSprm,
  type SyntheticFont,
  type SyntheticStyle,
} from './helpers/synthetic-doc'

async function zipText(docx: Uint8Array, path: string): Promise<string | undefined> {
  const zip = await JSZip.loadAsync(docx)
  const file = zip.file(path)
  return file ? file.async('string') : undefined
}

function concat(...parts: Uint8Array[]): Uint8Array {
  return Uint8Array.from(parts.flatMap((part) => [...part]))
}

const NORMAL: SyntheticStyle[] = [{ sti: 0, name: 'Normal', kind: 'paragraph' }]

const FONTS: SyntheticFont[] = [
  {
    name: 'Times New Roman',
    family: 1,
    pitch: 2,
    panose: '02020603050405020304',
    trueType: true,
  },
  { name: 'Symbol', family: 1, pitch: 2, charset: 2, trueType: true },
  { name: 'RotisSemiSans', altName: 'Courier New' },
]

describe('numbering levels keep their run properties', () => {
  it('writes the level font and size as w:rPr and keeps a symbol-font bullet code', async () => {
    const text = 'One\r'
    const { docx } = await convertLegacyDocToDocx(
      buildSyntheticDoc({
        pieces: [{ text }],
        stshStyles: NORMAL,
        fonts: FONTS,
        papxFkpPages: [
          encodePapxFkp([
            {
              fcStart: SYNTHETIC_TEXT_OFFSET,
              fcEnd: SYNTHETIC_TEXT_OFFSET + text.length * 2,
              istd: 0,
              grpprl: encodeListParagraphPapx(1),
            },
          ]),
        ],
        lists: [
          {
            lsid: 7,
            levels: [
              {
                numberFormat: 23,
                levelText: '',
                indentLeft: 720,
                indentFirstLine: -360,
                characterSprms: concat(encodeSprm(0x4a4f, [1, 0]), encodeSprm(0x4a43, [28, 0])),
              },
            ],
          },
        ],
      }),
    )
    const numbering = (await zipText(docx, 'word/numbering.xml'))!
    expect(numbering).toContain(
      '<w:lvlText w:val="&#xF0B7;"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:lvl>',
    )
  })
})

describe('list level indents outrank style indents', () => {
  it('uses the level indents unless the paragraph sets its own', async () => {
    const text = 'One\rTwo\r'
    const paragraphStart = (cp: number) => SYNTHETIC_TEXT_OFFSET + cp * 2
    const bytes = buildSyntheticDoc({
      pieces: [{ text }],
      stshStyles: [
        { sti: 0, name: 'Normal', kind: 'paragraph' },
        {
          name: 'Indented',
          kind: 'paragraph',
          basedOn: 0,
          papxSprms: encodeSprm(0x845e, [105, 0]),
        },
      ],
      papxFkpPages: [
        encodePapxFkp([
          {
            fcStart: paragraphStart(0),
            fcEnd: paragraphStart(4),
            istd: 1,
            grpprl: encodeListParagraphPapx(1),
          },
          {
            fcStart: paragraphStart(4),
            fcEnd: paragraphStart(8),
            istd: 1,
            grpprl: concat(encodeListParagraphPapx(1), encodeSprm(0x845e, [200, 0])),
          },
        ]),
      ],
      lists: [
        {
          lsid: 7,
          levels: [
            { numberFormat: 0, levelText: '\u0000.', indentLeft: 720, indentFirstLine: -360 },
          ],
        },
      ],
    })
    const body = readLegacyDocBody(bytes)
    const indents = body.blocks.flatMap((block) =>
      block.type === 'paragraph' && block.paragraph.text !== ''
        ? [
            [
              block.paragraph.formatting.indentLeftTwips,
              block.paragraph.formatting.indentFirstLineTwips,
            ],
          ]
        : [],
    )
    expect(indents).toEqual([
      [720, -360],
      [200, -360],
    ])
    const xml = (await zipText((await convertLegacyDocToDocx(bytes)).docx, 'word/document.xml'))!
    expect(xml).toContain('<w:ind w:left="720" w:hanging="360"/>')
    expect(xml).toContain('<w:ind w:left="200" w:hanging="360"/>')
  })
})

describe('text box text direction', () => {
  it('maps the OfficeArt text flow to wps:bodyPr vert', async () => {
    const body = 'Anchor \u0008here\rNext\r'
    const docxFor = async (props: Array<[number, number]>) =>
      (await zipText(
        (
          await convertLegacyDocToDocx(
            buildSyntheticDoc({
              pieces: [{ text: body }],
              floatingShapes: [
                { story: 'body', cp: 7, spid: 1025, rect: [1440, 720, 2160, 5040], props },
              ],
              textBoxStories: [{ spid: 1025, text: 'Side\r' }],
            }),
          )
        ).docx,
        'word/document.xml',
      ))!
    expect(await docxFor([[0x0088, 2]])).toContain('<wps:bodyPr rot="0" vert="vert270"')
    expect(await docxFor([[0x0088, 3]])).toContain('<wps:bodyPr rot="0" vert="vert"')
    expect(await docxFor([])).toContain('<wps:bodyPr rot="0" vert="horz"')
  })
})

describe('East Asian compatibility options', () => {
  it('writes the FE layout flags only for documents saved by an East Asian Word', async () => {
    const settingsFor = async (flags?: number) =>
      (await zipText(
        (
          await convertLegacyDocToDocx(
            buildSyntheticDoc({ pieces: [{ text: 'x\r' }], ...(flags ? { flags } : {}) }),
          )
        ).docx,
        'word/settings.xml',
      ))!
    const western = await settingsFor()
    expect(western).not.toContain('<w:useFELayout/>')
    const farEast = await settingsFor(0x4000)
    expect(farEast).toContain('<w:useFELayout/><w:useNormalStyleForList/>')
  })
})
