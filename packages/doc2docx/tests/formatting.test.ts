import { describe, expect, it } from 'vitest'
import {
  LegacyDocError,
  openLegacyDoc,
  readFormatting,
  readLegacyDocModel,
  readLegacyDocText,
} from '../src/index'
import { FIXTURE_NAMES, expectedText, fixtureBytes, type FixtureName } from './helpers/fixtures'
import {
  SYNTHETIC_TEXT_OFFSET,
  buildSyntheticDoc,
  encodeChpxFkp,
  encodePapxFkp,
  encodePlcBte,
  type SyntheticPapxEntry,
} from './helpers/synthetic-doc'

/** byte offset of a CP in a single uncompressed (UTF-16LE) synthetic text piece */
function textFc(cp: number): number {
  return SYNTHETIC_TEXT_OFFSET + cp * 2
}

function sprm(opcode: number, operand: number[]): Uint8Array {
  return Uint8Array.from([opcode & 0xff, (opcode >> 8) & 0xff, ...operand])
}

function concat(...parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0))
  let offset = 0
  for (const part of parts) {
    out.set(part, offset)
    offset += part.length
  }
  return out
}

function modelOf(bytes: Uint8Array) {
  return readLegacyDocModel(bytes)
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

describe('real formatting.doc fixture (python-docx document saved as Word 97 by LibreOffice)', () => {
  const name: FixtureName = 'formatting'
  const model = readLegacyDocModel(fixtureBytes(name))

  it('reads the 7 paragraphs plus the trailing empty one, exactly as the .txt export', () => {
    expect(model.paragraphs).toHaveLength(8)
    expect(model.paragraphs.map((paragraph) => paragraph.text)).toEqual(
      expectedText(name).split('\n'),
    )
    expect(model.warnings).toEqual([])
    expect(model.fib.nFib).toBe(0x0101)
    expect(model.paragraphs[7]).toMatchObject({ text: '', startCp: 159, endCp: 159, runs: [] })
  })

  it('P0: bold run, plain run, italic run, plain run', () => {
    expect(model.paragraphs[0]!.runs).toEqual([
      { text: 'Bold text', formatting: { bold: true } },
      { text: ' and ', formatting: {} },
      { text: 'italic text', formatting: { italic: true } },
      { text: ' plain.', formatting: {} },
    ])
  })

  it('P1: 14pt underlined run (sprmCHps 28 half-points, sprmCKul single)', () => {
    expect(model.paragraphs[1]!.runs).toEqual([
      {
        text: 'Large underlined text',
        formatting: { fontSizeHalfPoints: 28, underline: 'single' },
      },
    ])
  })

  it('P2/P3: centered and right-aligned paragraphs (sprmPJc80 + sprmPJc pairs)', () => {
    expect(model.paragraphs[2]!.formatting).toEqual({
      styleIndex: 0,
      alignment: 'center',
      alignmentLogical: true,
    })
    expect(model.paragraphs[3]!.formatting).toEqual({
      styleIndex: 0,
      alignment: 'right',
      alignmentLogical: true,
    })
  })

  it('P4: 720 left indent, 360 first-line indent, 1.5 line spacing', () => {
    expect(model.paragraphs[4]!.formatting).toEqual({
      styleIndex: 0,
      lineSpacing: { dyaLine: 360, multiple: true },
      indentFirstLineTwips: 360,
      indentLeftTwips: 720,
    })
  })

  it('P5: 240 space before, 120 space after, keep with next', () => {
    expect(model.paragraphs[5]!.formatting).toEqual({
      styleIndex: 0,
      spacingBeforeTwips: 240,
      spacingAfterTwips: 120,
      keepNext: true,
    })
  })

  it('P6: bold CJK run, plus the paragraph defaults LibreOffice writes explicitly', () => {
    expect(model.paragraphs[6]!.runs).toEqual([
      { text: '\u4e2d\u6587\u7c97\u4f53\u6bb5\u843d', formatting: { bold: true } },
    ])
    // LibreOffice wrote its 276/1 default line spacing and left alignment as
    // explicit sprms, and 200 twips of space after; all are real file data.
    expect(model.paragraphs[6]!.formatting).toEqual({
      styleIndex: 0,
      lineSpacing: { dyaLine: 276, multiple: true },
      alignment: 'left',
      alignmentLogical: true,
      bidi: false,
      spacingBeforeTwips: 0,
      spacingAfterTwips: 200,
    })
  })

  it('the run texts concatenate back to the paragraph text', () => {
    for (const paragraph of model.paragraphs)
      expect(paragraph.runs.map((run) => run.text).join('')).toBe(paragraph.text)
  })

  it('stage-1 and stage-2 agree on paragraph boundaries and text', () => {
    const text = readLegacyDocText(fixtureBytes(name))
    expect(model.paragraphs.map((paragraph) => paragraph.text)).toEqual(text.paragraphs)
    expect(model.text).toBe(text.text)
  })
})

describe('synthetic CHPX formatting', () => {
  it('applies known sprms after unknown sprms of every spra class', () => {
    const text = 'abcdef\r'
    const grpprl = concat(
      sprm(0x0001, [0xaa]), // spra 0
      sprm(0x2001, [0xbb]), // spra 1
      sprm(0x4001, [0xcc, 0xdd]), // spra 2
      sprm(0x6001, [1, 2, 3, 4]), // spra 3
      sprm(0x8001, [0xee, 0xff]), // spra 4
      sprm(0xa001, [0x11, 0x22]), // spra 5
      sprm(0xc001, [2, 0x33, 0x44]), // spra 6: length byte 2 + two payload bytes
      sprm(0xe001, [0x55, 0x66, 0x77]), // spra 7
      sprm(0x0835, [1]), // bold on
      sprm(0x4a43, [28, 0]), // 14pt
      sprm(0x2a3e, [1]), // single underline
    )
    const bytes = buildSyntheticDoc({
      pieces: [{ text }],
      chpxFkpPages: [encodeChpxFkp([{ fcStart: textFc(0), fcEnd: textFc(7), grpprl }])],
    })
    const paragraph = modelOf(bytes).paragraphs[0]!
    expect(paragraph.runs).toEqual([
      {
        text: 'abcdef',
        formatting: { bold: true, fontSizeHalfPoints: 28, underline: 'single' },
      },
    ])
  })

  it('maps toggle operands 0/1 to false/true and 0x80/0x81 to unset (no style sheet yet)', () => {
    const text = 'abcd\r'
    const page = encodeChpxFkp([
      { fcStart: textFc(0), fcEnd: textFc(1), grpprl: sprm(0x0835, [1]) },
      { fcStart: textFc(1), fcEnd: textFc(2), grpprl: sprm(0x0835, [0]) },
      { fcStart: textFc(2), fcEnd: textFc(3), grpprl: sprm(0x0835, [0x80]) },
      { fcStart: textFc(3), fcEnd: textFc(4), grpprl: sprm(0x0835, [0x81]) },
    ])
    const paragraph = modelOf(buildSyntheticDoc({ pieces: [{ text }], chpxFkpPages: [page] }))
      .paragraphs[0]!
    expect(paragraph.runs).toEqual([
      { text: 'a', formatting: { bold: true } },
      { text: 'b', formatting: { bold: false } },
      { text: 'cd', formatting: {} },
    ])
  })

  it('splits a run across compressed and UTF-16LE pieces', () => {
    const bytes = buildSyntheticDoc({
      pieces: [{ text: 'abcd', compressed: true }, { text: '\u4e2d\u6587x\r' }],
      chpxFkpPages: [
        encodeChpxFkp([
          { fcStart: 2048, fcEnd: 2050 }, // CP 0-1: 'ab'
          { fcStart: 2050, fcEnd: 2060, grpprl: sprm(0x0835, [1]) }, // CP 2-7: the CJK part of the mixed piece
        ]),
      ],
    })
    const model = modelOf(bytes)
    expect(model.paragraphs[0]!.text).toBe('abcd\u4e2d\u6587x')
    expect(model.paragraphs[0]!.runs).toEqual([
      { text: 'ab', formatting: {} },
      { text: 'cd\u4e2d\u6587x', formatting: { bold: true } },
    ])
  })

  it('reads the short and long PapxInFkp forms and maps paragraph start FCs', () => {
    const text = 'one\rtwo\rthree\r'
    const longSprms = sprm(0xa413, [0xf0, 0x00]) // 4 bytes -> istd + 4 = 6 -> long form
    const shortSprms = sprm(0x2406, [1]) // 3 bytes -> istd + 3 = 5 -> short form
    const entries: SyntheticPapxEntry[] = [
      { fcStart: textFc(0), fcEnd: textFc(4), grpprl: longSprms },
      { fcStart: textFc(4), fcEnd: textFc(8), grpprl: shortSprms },
      { fcStart: textFc(8), fcEnd: textFc(15), defaults: true },
    ]
    const bytes = buildSyntheticDoc({ pieces: [{ text }], papxFkpPages: [encodePapxFkp(entries)] })
    const model = modelOf(bytes)
    expect(model.paragraphs.map((paragraph) => paragraph.text)).toEqual(['one', 'two', 'three', ''])
    expect(model.paragraphs[0]!.formatting).toEqual({ styleIndex: 0, spacingBeforeTwips: 240 })
    expect(model.paragraphs[1]!.formatting).toEqual({ styleIndex: 0, keepNext: true })
    // a default paragraph (BxPap.bOffset 0) carries the FKP's implicit style index only
    expect(model.paragraphs[2]!.formatting).toEqual({ styleIndex: 0 })
  })

  it('reads sprmPFBiDi / sprmCFBiDi and tells physical from logical justification', () => {
    const text = 'x\r'
    const bytes = buildSyntheticDoc({
      pieces: [{ text }],
      papxFkpPages: [
        encodePapxFkp([
          {
            fcStart: textFc(0),
            fcEnd: textFc(2),
            grpprl: concat(sprm(0x2403, [2]), sprm(0x2441, [1])),
          },
        ]),
      ],
      chpxFkpPages: [
        encodeChpxFkp([{ fcStart: textFc(0), fcEnd: textFc(2), grpprl: sprm(0x085a, [1]) }]),
      ],
    })
    const paragraph = modelOf(bytes).paragraphs[0]!
    expect(paragraph.formatting).toEqual({
      styleIndex: 0,
      alignment: 'right',
      alignmentLogical: false,
      bidi: true,
    })
    expect(paragraph.runs[0]!.formatting.rtl).toBe(true)
  })

  it('maps paragraph-level sprms (alignment, indents, spacing, keep flags, lists)', () => {
    const text = 'x\r'
    const grpprl = concat(
      sprm(0x2403, [1]), // sprmPJc80: center
      sprm(0x2461, [1]), // sprmPJc: center
      sprm(0x845e, [0xd0, 0x02]), // sprmPDxaLeft: 720
      sprm(0x8460, [0x68, 0x01]), // sprmPDxaLeft1: 360
      sprm(0x845d, [0xb4, 0x00]), // sprmPDxaRight: 180
      sprm(0x6412, [0x68, 0x01, 0x01, 0x00]), // sprmPDyaLine: dyaLine 360, multiple
      sprm(0xa413, [0xf0, 0x00]), // sprmPDyaBefore: 240
      sprm(0xa414, [0x78, 0x00]), // sprmPDyaAfter: 120
      sprm(0x2405, [1]), // sprmPFKeep
      sprm(0x2406, [1]), // sprmPFKeepFollow
      sprm(0x2407, [1]), // sprmPFPageBreakBefore
      sprm(0x2640, [2]), // sprmPOutLvl
      sprm(0x260a, [1]), // sprmPIlvl
      sprm(0x460b, [3, 0]), // sprmPIlfo: 3
      sprm(0x2416, [1]), // sprmPFInTable
      sprm(0x2417, [1]), // sprmPFTtp
    )
    const bytes = buildSyntheticDoc({
      pieces: [{ text }],
      papxFkpPages: [encodePapxFkp([{ fcStart: textFc(0), fcEnd: textFc(2), grpprl }])],
    })
    expect(modelOf(bytes).paragraphs[0]!.formatting).toEqual({
      styleIndex: 0,
      alignment: 'center',
      alignmentLogical: true,
      indentLeftTwips: 720,
      indentFirstLineTwips: 360,
      indentRightTwips: 180,
      lineSpacing: { dyaLine: 360, multiple: true },
      spacingBeforeTwips: 240,
      spacingAfterTwips: 120,
      keepLines: true,
      keepNext: true,
      pageBreakBefore: true,
      outlineLevel: 2,
      listLevel: 1,
      listId: 3,
      inTable: true,
      tableRowEnd: true,
    })
  })

  it('reads the positioned-frame sprms into formatting.frame', () => {
    const grpprl = concat(
      sprm(0x8418, [0x96, 0x03]), // sprmPDxaAbs 918
      sprm(0x8419, [0xfc, 0xff]), // sprmPDyaAbs -4 (top)
      sprm(0x841a, [0xd0, 0x02]), // sprmPDxaWidth 720
      sprm(0x261b, [0x90]), // sprmPPc: vertical page, horizontal page
      sprm(0x2423, [0x02]), // sprmPWr around
      sprm(0x442b, [0x68, 0x81]), // sprmPDyaHeight exact 360
      sprm(0x842e, [0xb4, 0x00]), // sprmPDyaFromText 180
      sprm(0x842f, [0x8d, 0x00]), // sprmPDxaFromText 141
    )
    const bytes = buildSyntheticDoc({
      pieces: [{ text: 'x\r' }],
      papxFkpPages: [encodePapxFkp([{ fcStart: textFc(0), fcEnd: textFc(2), grpprl }])],
    })
    expect(modelOf(bytes).paragraphs[0]!.formatting.frame).toEqual({
      xTwips: 918,
      yTwips: -4,
      widthTwips: 720,
      vAnchor: 'page',
      hAnchor: 'page',
      wrap: 'around',
      heightTwips: 360,
      heightExact: true,
      vSpaceTwips: 180,
      hSpaceTwips: 141,
    })
  })

  it('reads sprmCCv RGB colors, keeping the palette index alongside', () => {
    const text = 'ab\r'
    const bytes = buildSyntheticDoc({
      pieces: [{ text }],
      chpxFkpPages: [
        encodeChpxFkp([
          {
            fcStart: textFc(0),
            fcEnd: textFc(1),
            grpprl: concat(sprm(0x2a42, [6]), sprm(0x6870, [0xcc, 0, 0, 0])),
          },
          { fcStart: textFc(1), fcEnd: textFc(2), grpprl: sprm(0x6870, [0, 0, 0, 0xff]) },
        ]),
      ],
    })
    const runs = modelOf(bytes).paragraphs[0]!.runs
    expect(runs[0]!.formatting).toEqual({ colorIndex: 6, colorRgb: 'CC0000' })
    expect(runs[1]!.formatting).toEqual({ colorRgb: 'auto' })
  })

  it('reads character spacing (sprmCDxaSpace) and horizontal scale (sprmCCharScale)', () => {
    const text = 'ab\r'
    const bytes = buildSyntheticDoc({
      pieces: [{ text }],
      chpxFkpPages: [
        encodeChpxFkp([
          {
            fcStart: textFc(0),
            fcEnd: textFc(1),
            grpprl: concat(sprm(0x8840, [0xec, 0xff]), sprm(0x4852, [0x50, 0x00])),
          },
          { fcStart: textFc(1), fcEnd: textFc(3), grpprl: sprm(0x8840, [0x1e, 0x00]) },
        ]),
      ],
    })
    const runs = modelOf(bytes).paragraphs[0]!.runs
    expect(runs[0]!.formatting).toEqual({ spacingTwips: -20, scalePercent: 80 })
    expect(runs[1]!.formatting).toEqual({ spacingTwips: 30 })
  })

  it('reads tab stops from sprmPChgTabsPapx and sprmPChgTabs (delete, then add)', () => {
    const text = 'a\rb\r'
    // sprmPChgTabsPapx: no deletions; add 720 left and 8640 right-dot (TBD jc 2, tlc 1)
    const papxTabs = sprm(0xc60d, [8, 0, 2, 0xd0, 0x02, 0xc0, 0x21, 0x00, 0x0a])
    // sprmPChgTabs: delete 720 (with close tolerance), add 4320 center
    const chgTabs = sprm(0xc615, [10, 0, 1, 0xd0, 0x02, 0x10, 0x00, 1, 0xe0, 0x10, 0x01])
    const bytes = buildSyntheticDoc({
      pieces: [{ text }],
      papxFkpPages: [
        encodePapxFkp([
          { fcStart: textFc(0), fcEnd: textFc(2), grpprl: papxTabs },
          { fcStart: textFc(2), fcEnd: textFc(4), grpprl: concat(papxTabs, chgTabs) },
        ]),
      ],
    })
    const paragraphs = modelOf(bytes).paragraphs
    expect(paragraphs[0]!.formatting.tabStops).toEqual([
      { positionTwips: 720, alignment: 'left' },
      { positionTwips: 8640, alignment: 'right', leader: 'dot' },
    ])
    expect(paragraphs[1]!.formatting.tabStops).toEqual([
      { positionTwips: 4320, alignment: 'center' },
      { positionTwips: 8640, alignment: 'right', leader: 'dot' },
    ])
  })

  it('reads paragraph borders from the Brc80 and Brc sprms', () => {
    const text = 'a\rb\r'
    const brc80 = concat(
      sprm(0x6424, [4, 1, 8, 1]), // top: 0.5pt single, ico 8 (white), 1pt space
      sprm(0x6426, [16, 1, 2, 4]), // bottom: 2pt single, ico 2 (blue), 4pt space
    )
    const brc = concat(
      brc80,
      sprm(0xc650, [8, 0x2b, 0x6c, 0xb0, 0, 16, 1, 4, 0]), // bottom again with the RGB twin
      sprm(0xc64e, [8, 0, 0, 0, 0, 0, 0, 0, 0]), // top removed (brcType none)
    )
    const bytes = buildSyntheticDoc({
      pieces: [{ text }],
      papxFkpPages: [
        encodePapxFkp([
          { fcStart: textFc(0), fcEnd: textFc(2), grpprl: brc80 },
          { fcStart: textFc(2), fcEnd: textFc(4), grpprl: brc },
        ]),
      ],
    })
    const paragraphs = modelOf(bytes).paragraphs
    expect(paragraphs[0]!.formatting.borders).toEqual({
      top: { widthEighths: 4, spacePoints: 1, color: 'FFFFFF' },
      bottom: { widthEighths: 16, spacePoints: 4, color: '0000FF' },
    })
    expect(paragraphs[1]!.formatting.borders).toEqual({
      bottom: { widthEighths: 16, spacePoints: 4, color: '2B6CB0' },
    })
  })

  it('reads character-unit indents next to their twips twins', () => {
    const text = 'a\r'
    const grpprl = concat(sprm(0x8460, [0x80, 0x02]), sprm(0x4457, [200, 0]), sprm(0x4456, [50, 0]))
    const bytes = buildSyntheticDoc({
      pieces: [{ text }],
      papxFkpPages: [encodePapxFkp([{ fcStart: textFc(0), fcEnd: textFc(2), grpprl }])],
    })
    expect(modelOf(bytes).paragraphs[0]!.formatting).toEqual({
      styleIndex: 0,
      indentFirstLineTwips: 640,
      charIndents: { firstLine: 200, left: 50 },
    })
  })

  it('reads paragraph properties from the Data stream via sprmPHugePapx', () => {
    const text = 'x\r'
    const hugeSprms = concat(sprm(0x6645, [0, 0, 0, 0])) // sprmPHugePapx: fc 0 into the Data stream
    const propertyBlock = concat(sprm(0xa413, [0xf0, 0x00])) // sprmPDyaBefore: 240
    const dataStream = new Uint8Array(2 + propertyBlock.length)
    new DataView(dataStream.buffer).setInt16(0, propertyBlock.length, true)
    dataStream.set(propertyBlock, 2)
    const bytes = buildSyntheticDoc({
      pieces: [{ text }],
      papxFkpPages: [encodePapxFkp([{ fcStart: textFc(0), fcEnd: textFc(2), grpprl: hugeSprms }])],
      dataStream,
    })
    const model = modelOf(bytes)
    expect(model.warnings).toEqual([])
    expect(model.paragraphs[0]!.formatting).toEqual({ styleIndex: 0, spacingBeforeTwips: 240 })
  })

  it('records a warning and keeps the known properties when the Data stream is missing', () => {
    const text = 'x\r'
    const hugeSprms = concat(sprm(0x6645, [0, 0, 0, 0]))
    const bytes = buildSyntheticDoc({
      pieces: [{ text }],
      papxFkpPages: [encodePapxFkp([{ fcStart: textFc(0), fcEnd: textFc(2), grpprl: hugeSprms }])],
    })
    const model = modelOf(bytes)
    expect(model.paragraphs[0]!.formatting).toEqual({ styleIndex: 0 })
    expect(model.warnings).toHaveLength(1)
    expect(model.warnings[0]).toContain('Data stream')
  })

  it('warns about a sprmPHugePapx FC outside the Data stream', () => {
    const text = 'x\r'
    const hugeSprms = concat(sprm(0x6645, [0x10, 0, 0, 0]))
    const bytes = buildSyntheticDoc({
      pieces: [{ text }],
      papxFkpPages: [encodePapxFkp([{ fcStart: textFc(0), fcEnd: textFc(2), grpprl: hugeSprms }])],
      dataStream: new Uint8Array(4),
    })
    const model = modelOf(bytes)
    expect(model.warnings[0]).toContain('outside the Data stream')
  })

  it('warns about a PAPX entry that does not start at a paragraph start', () => {
    const text = 'one\rtwo\r'
    const bytes = buildSyntheticDoc({
      pieces: [{ text }],
      papxFkpPages: [
        encodePapxFkp([{ fcStart: textFc(1), fcEnd: textFc(4), grpprl: sprm(0x2405, [1]) }]),
      ],
    })
    const model = modelOf(bytes)
    expect(model.paragraphs[0]!.formatting).toEqual({})
    expect(model.warnings).toHaveLength(1)
    expect(model.warnings[0]).toContain('not a paragraph start')
  })

  it('combines CHPX runs from two FKP pages', () => {
    const text = 'ab\rcd\r'
    const bytes = buildSyntheticDoc({
      pieces: [{ text }],
      chpxFkpPages: [
        encodeChpxFkp([{ fcStart: textFc(0), fcEnd: textFc(2), grpprl: sprm(0x0835, [1]) }]),
        encodeChpxFkp([{ fcStart: textFc(2), fcEnd: textFc(6), grpprl: sprm(0x0836, [1]) }]),
      ],
    })
    const model = modelOf(bytes)
    expect(model.paragraphs[0]!.runs).toEqual([{ text: 'ab', formatting: { bold: true } }])
    expect(model.paragraphs[1]!.runs).toEqual([{ text: 'cd', formatting: { italic: true } }])
  })

  it('returns plain formatting when the document has no formatting tables', () => {
    const model = modelOf(buildSyntheticDoc({ pieces: [{ text: 'plain\r' }] }))
    expect(model.paragraphs[0]!.formatting).toEqual({})
    expect(model.paragraphs[0]!.runs).toEqual([{ text: 'plain', formatting: {} }])
    expect(model.warnings).toEqual([])
  })

  it('reads an empty document (ccpText 0)', () => {
    const model = modelOf(buildSyntheticDoc({}))
    expect(model.paragraphs).toEqual([{ startCp: 0, endCp: 0, text: '', formatting: {}, runs: [] }])
  })

  it('keeps empty paragraphs between paragraph marks', () => {
    const model = modelOf(buildSyntheticDoc({ pieces: [{ text: 'a\r\rb\r' }] }))
    expect(model.paragraphs.map((paragraph) => paragraph.text)).toEqual(['a', '', 'b', ''])
    expect(model.paragraphs[1]!.runs).toEqual([])
  })

  it('propagates the character style index (sprmCIstd)', () => {
    const bytes = buildSyntheticDoc({
      pieces: [{ text: 'x\r' }],
      chpxFkpPages: [
        encodeChpxFkp([{ fcStart: textFc(0), fcEnd: textFc(1), grpprl: sprm(0x4a30, [2, 0]) }]),
      ],
    })
    expect(modelOf(bytes).paragraphs[0]!.runs[0]!.formatting).toEqual({ styleIndex: 2 })
  })
})

describe('stage-2 error paths', () => {
  it('rejects a truncated sprm in a CHPX grpprl (bad_sprm)', () => {
    const bytes = buildSyntheticDoc({
      pieces: [{ text: 'x\r' }],
      chpxFkpPages: [
        encodeChpxFkp([
          { fcStart: textFc(0), fcEnd: textFc(2), grpprl: Uint8Array.from([0x35, 0x08]) },
        ]),
      ],
    })
    expect(errorCode(() => readLegacyDocModel(bytes))).toBe('bad_sprm')
  })

  it('rejects a CHPX run FC that is not on a character boundary (bad_fkp)', () => {
    const bytes = buildSyntheticDoc({
      pieces: [{ text: 'ab\r' }],
      chpxFkpPages: [
        encodeChpxFkp([{ fcStart: textFc(0) + 1, fcEnd: textFc(3), grpprl: sprm(0x0835, [1]) }]),
      ],
    })
    expect(errorCode(() => readLegacyDocModel(bytes))).toBe('bad_fkp')
  })

  it('rejects a formatting table page outside the WordDocument stream (bad_fkp)', () => {
    const bytes = buildSyntheticDoc({
      pieces: [{ text: 'x\r' }],
      chpxPlcf: encodePlcBte([2048, 4096], [999999]),
    })
    expect(errorCode(() => readLegacyDocModel(bytes))).toBe('bad_fkp')
  })

  it('rejects a formatting table with a length that is not 8n+4 (bad_fkp)', () => {
    const bytes = buildSyntheticDoc({
      pieces: [{ text: 'x\r' }],
      chpxFkpPages: [encodeChpxFkp([{ fcStart: textFc(0), fcEnd: textFc(2) }])],
      chpxPlcLcb: 5,
    })
    expect(errorCode(() => readLegacyDocModel(bytes))).toBe('bad_fkp')
  })

  it('rejects a formatting table range outside the table stream (the FIB layer catches it first)', () => {
    const bytes = buildSyntheticDoc({
      pieces: [{ text: 'x\r' }],
      chpxFkpPages: [encodeChpxFkp([{ fcStart: textFc(0), fcEnd: textFc(2) }])],
      chpxPlcLcb: 9999,
    })
    // stage 1 validates FibRgFcLcb pair ranges against the table stream, so an
    // out-of-range PLCF is reported as bad_fib before the CLI stage runs.
    expect(errorCode(() => readLegacyDocModel(bytes))).toBe('bad_fib')
  })

  it('rejects an FKP with an impossible run count (bad_fkp)', () => {
    const page = encodeChpxFkp([{ fcStart: textFc(0), fcEnd: textFc(2) }])
    page[511] = 200
    const bytes = buildSyntheticDoc({ pieces: [{ text: 'x\r' }], chpxFkpPages: [page] })
    expect(errorCode(() => readLegacyDocModel(bytes))).toBe('bad_fkp')
  })

  it('rejects a corrupt PAPX page (bad_fkp)', () => {
    const page = encodePapxFkp([
      { fcStart: textFc(0), fcEnd: textFc(2), grpprl: sprm(0x2406, [1]) },
    ])
    page[511] = 100
    const bytes = buildSyntheticDoc({ pieces: [{ text: 'x\r' }], papxFkpPages: [page] })
    expect(errorCode(() => readLegacyDocModel(bytes))).toBe('bad_fkp')
  })

  it('accepts every stage-1 fixture and keeps runs consistent with the text', () => {
    for (const fixture of FIXTURE_NAMES) {
      const model = readLegacyDocModel(fixtureBytes(fixture))
      expect(model.paragraphs.length).toBeGreaterThan(0)
      for (const paragraph of model.paragraphs)
        expect(paragraph.runs.map((run) => run.text).join('')).toBe(paragraph.text)
      expect(model.warnings).toEqual([])
    }
  })

  it('keeps readFormatting callable directly on a parsed document', () => {
    const doc = openLegacyDoc(fixtureBytes('text-basic'))
    const paragraphs = readFormatting(doc)
    expect(paragraphs.map((paragraph) => paragraph.text)).toEqual(doc.paragraphs)
  })
})
