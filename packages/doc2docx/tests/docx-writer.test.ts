import JSZip from 'jszip'
import { beforeAll, describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import type { Block } from '@genoffice/docx-engine'
import {
  LegacyDocError,
  convertLegacyDocToDocx,
  readLegacyDoc,
  readLegacyDocBody,
  readLegacyDocLayout,
  readLegacyDocModel,
  readResolvedLegacyDocModel,
  toDocxRun,
  toParaFormat,
  writeLegacyDocx,
  type LegacyDocModel,
} from '../src/index'
import { expectedText, fixtureBytes } from './helpers/fixtures'
import {
  SYNTHETIC_TEXT_OFFSET,
  buildSyntheticDoc,
  encodeChpxFkp,
  encodePapxFkp,
} from './helpers/synthetic-doc'

type Parsed = Awaited<ReturnType<typeof parseDocx>>

/** byte offset of a CP in a single uncompressed (UTF-16LE) synthetic text piece */
function textFc(cp: number): number {
  return SYNTHETIC_TEXT_OFFSET + cp * 2
}

function sprm(opcode: number, operand: number[]): Uint8Array {
  return Uint8Array.from([opcode & 0xff, (opcode >> 8) & 0xff, ...operand])
}

/** the generated-document paragraphs (parseDocx adds one passthrough block for the sectPr) */
function paragraphsOf(parsed: Parsed): Block[] {
  return parsed.blocks.filter((block) => block.type === 'paragraph')
}

function textOf(block: Block): string {
  return (block.runs ?? []).map((run) => run.text).join('')
}

async function rejectionCode(run: () => Promise<unknown>): Promise<string> {
  try {
    await run()
  } catch (error) {
    expect(error).toBeInstanceOf(LegacyDocError)
    return (error as LegacyDocError).code
  }
  throw new Error('expected the conversion to reject with a LegacyDocError')
}

describe('toDocxRun mapping', () => {
  it('maps text and boolean character formatting', () => {
    expect(
      toDocxRun({
        text: 'x',
        formatting: {
          bold: true,
          italic: true,
          strike: true,
          fontSizeHalfPoints: 28,
          verticalAlign: 'superscript',
        },
      }),
    ).toEqual({
      text: 'x',
      bold: true,
      italic: true,
      strike: true,
      sizeHalfPoints: 28,
      vertAlign: 'superscript',
    })
  })

  it('maps underline kinds to a boolean and baseline to nothing', () => {
    for (const underline of [
      'single',
      'words',
      'double',
      'dotted',
      'thick',
      'dash',
      'dotDash',
      'dotDotDash',
      'wave',
    ] as const)
      expect(toDocxRun({ text: 'x', formatting: { underline } }).underline).toBe(true)
    for (const underline of ['none', 'inherit', 'invert', 'other:5'] as const)
      expect(toDocxRun({ text: 'x', formatting: { underline } }).underline).toBeUndefined()
    expect(
      toDocxRun({ text: 'x', formatting: { verticalAlign: 'baseline' } }).vertAlign,
    ).toBeUndefined()
    expect(toDocxRun({ text: 'x', formatting: { verticalAlign: 'subscript' } }).vertAlign).toBe(
      'subscript',
    )
  })

  it('maps Word color indices 1..16 and omits auto/unknown', () => {
    expect(toDocxRun({ text: 'x', formatting: { colorIndex: 0 } }).color).toBeUndefined()
    expect(toDocxRun({ text: 'x', formatting: { colorIndex: 1 } }).color).toBe('000000')
    expect(toDocxRun({ text: 'x', formatting: { colorIndex: 6 } }).color).toBe('FF0000')
    expect(toDocxRun({ text: 'x', formatting: { colorIndex: 9 } }).color).toBe('000080')
    expect(toDocxRun({ text: 'x', formatting: { colorIndex: 16 } }).color).toBe('C0C0C0')
    expect(toDocxRun({ text: 'x', formatting: { colorIndex: 17 } }).color).toBeUndefined()
  })

  it('maps Word highlight indices 1..16 to OOXML names and omits unknown', () => {
    expect(toDocxRun({ text: 'x', formatting: { highlight: 1 } }).highlight).toBe('black')
    expect(toDocxRun({ text: 'x', formatting: { highlight: 7 } }).highlight).toBe('yellow')
    expect(toDocxRun({ text: 'x', formatting: { highlight: 9 } }).highlight).toBe('darkBlue')
    expect(toDocxRun({ text: 'x', formatting: { highlight: 16 } }).highlight).toBe('lightGray')
    expect(toDocxRun({ text: 'x', formatting: { highlight: 17 } }).highlight).toBeUndefined()
  })

  it('drops explicit "off" formatting instead of inventing w:val="0" (blank template default is off)', () => {
    expect(toDocxRun({ text: 'x', formatting: { bold: false, italic: false } })).toEqual({
      text: 'x',
    })
  })
})

describe('toParaFormat mapping', () => {
  it('returns undefined for empty formatting', () => {
    expect(toParaFormat({})).toBeUndefined()
  })

  it('maps alignment, indents, spacing and keep flags', () => {
    expect(
      toParaFormat({
        alignment: 'center',
        indentLeftTwips: 720,
        indentRightTwips: 180,
        indentFirstLineTwips: 360,
        spacingBeforeTwips: 240,
        spacingAfterTwips: 120,
        keepNext: true,
        keepLines: true,
        pageBreakBefore: true,
      }),
    ).toEqual({
      align: 'center',
      indentLeft: 720,
      indentRight: 180,
      indentFirstLine: 360,
      spaceBefore: 240,
      spaceAfter: 120,
      keepNext: true,
      keepLines: true,
      pageBreakBefore: true,
    })
  })

  it('maps multiple line spacing to a multiplier and exact spacing to a raw twips rule', () => {
    expect(toParaFormat({ lineSpacing: { dyaLine: 360, multiple: true } })).toEqual({
      lineSpacing: 1.5,
    })
    expect(toParaFormat({ lineSpacing: { dyaLine: 240, multiple: true } })).toEqual({
      lineSpacing: 1,
    })
    expect(toParaFormat({ lineSpacing: { dyaLine: 360, multiple: false } })).toEqual({
      lineRule: 'atLeast',
      lineRawTwips: 360,
    })
    expect(toParaFormat({ lineSpacing: { dyaLine: -360, multiple: false } })).toEqual({
      lineRule: 'exact',
      lineRawTwips: 360,
    })
  })

  it('maps tab stops, paragraph borders and character-unit first-line indents', () => {
    expect(
      toParaFormat({
        tabStops: [
          { positionTwips: 720, alignment: 'left' },
          { positionTwips: 8640, alignment: 'right', leader: 'dot' },
        ],
        borders: {
          bottom: { widthEighths: 16, spacePoints: 4, color: '2B6CB0' },
          top: { widthEighths: 4, spacePoints: 1 },
        },
        charIndents: { firstLine: 200 },
      }),
    ).toEqual({
      tabStops: [
        { pos: 720, val: 'left' },
        { pos: 8640, val: 'right', leader: 'dot' },
      ],
      borders: 'tb',
      borderLines: { t: { szPt: 0.5, spacePt: 1 }, b: { szPt: 2, spacePt: 4, color: '2B6CB0' } },
      charIndents: { firstLine: 200 },
    })
  })
})

describe('toDocxRun: RGB color and hidden text', () => {
  it('prefers the sprmCCv RGB value over the palette index and writes hidden as vanish', () => {
    expect(toDocxRun({ text: 'x', formatting: { colorIndex: 6, colorRgb: 'CC0000' } })).toEqual({
      text: 'x',
      color: 'CC0000',
    })
    expect(toDocxRun({ text: 'x', formatting: { colorRgb: 'auto' } })).toEqual({
      text: 'x',
      color: 'auto',
    })
    expect(toDocxRun({ text: 'x', formatting: { hidden: true } })).toEqual({
      text: 'x',
      vanishOwn: true,
    })
  })
})

describe('formatting.doc -> docx round trip', () => {
  let model: LegacyDocModel
  let docx: Uint8Array
  let warnings: string[]
  let parsed: Parsed
  let paragraphs: Block[]

  beforeAll(async () => {
    model = readLegacyDocModel(fixtureBytes('formatting'))
    const result = await convertLegacyDocToDocx(fixtureBytes('formatting'))
    docx = result.docx
    warnings = result.warnings
    parsed = await parseDocx(docx)
    paragraphs = paragraphsOf(parsed)
  })

  it('produces a valid docx with every paragraph, including the trailing empty one', () => {
    expect(docx).toBeInstanceOf(Uint8Array)
    expect(docx.length).toBeGreaterThan(1000)
    expect(paragraphs).toHaveLength(model.paragraphs.length)
    expect(paragraphs.map(textOf)).toEqual(model.paragraphs.map((paragraph) => paragraph.text))
    expect(textOf(paragraphs[7]!)).toBe('')
    expect(warnings).toEqual([])
  })

  it('keeps the bold and italic runs of P0', () => {
    const runs = paragraphs[0]!.runs ?? []
    expect(runs.map((run) => [run.text, run.bold ?? false, run.italic ?? false])).toEqual([
      ['Bold text', true, false],
      [' and ', false, false],
      ['italic text', false, true],
      [' plain.', false, false],
    ])
  })

  it('keeps P1 14pt and underline', () => {
    const runs = paragraphs[1]!.runs ?? []
    expect(runs).toHaveLength(1)
    expect(runs[0]).toMatchObject({
      text: 'Large underlined text',
      sizeHalfPoints: 28,
      underline: true,
    })
  })

  it('keeps P2/P3 center and right alignment', () => {
    expect(paragraphs[2]!.format).toMatchObject({ align: 'center' })
    expect(paragraphs[3]!.format).toMatchObject({ align: 'right' })
  })

  it('keeps P4 indents and 1.5 line spacing', () => {
    expect(paragraphs[4]!.format).toMatchObject({
      indentLeft: 720,
      indentFirstLine: 360,
      lineSpacing: 1.5,
      lineRule: 'auto',
      lineRawTwips: 360,
    })
  })

  it('keeps P5 space before/after and keep-with-next', () => {
    expect(paragraphs[5]!.format).toMatchObject({
      spaceBefore: 240,
      spaceAfter: 120,
      keepNext: true,
    })
  })

  it('keeps the P6 bold CJK run and the explicit paragraph formatting LibreOffice wrote', () => {
    expect(paragraphs[6]!.runs ?? []).toHaveLength(1)
    expect((paragraphs[6]!.runs ?? [])[0]).toMatchObject({
      text: '\u4e2d\u6587\u7c97\u4f53\u6bb5\u843d',
      bold: true,
    })
    expect(paragraphs[6]!.format).toMatchObject({
      align: 'left',
      lineSpacing: 1.15,
      spaceAfter: 200,
    })
  })

  it('stage-1 model text and the parsed docx text agree', () => {
    expect(paragraphs.map(textOf).join('\n')).toBe(model.text)
  })

  it('is stable through a docx-engine parse/save round trip', async () => {
    const saved = await saveDocx(
      parsed,
      parsed.blocks.map((block) => ({ kind: 'original' as const, docxIndex: block.docxIndex! })),
    )
    const reparsed = await parseDocx(saved)
    expect(paragraphsOf(reparsed).map(textOf)).toEqual(paragraphs.map(textOf))
  })
})

describe('text-basic.doc -> docx round trip', () => {
  let model: LegacyDocModel
  let parsed: Parsed
  let paragraphs: Block[]
  let documentXml: string

  beforeAll(async () => {
    const bytes = fixtureBytes('text-basic')
    model = readLegacyDocModel(bytes)
    const { docx } = await convertLegacyDocToDocx(bytes)
    parsed = await parseDocx(docx)
    paragraphs = paragraphsOf(parsed)
    const zip = await JSZip.loadAsync(docx)
    documentXml = await zip.file('word/document.xml')!.async('string')
  })

  it('keeps every paragraph text (CJK, emoji, tab, empty paragraph)', () => {
    expect(paragraphs.map(textOf)).toEqual(model.paragraphs.map((paragraph) => paragraph.text))
    expect(paragraphs.map(textOf).join('\n')).toBe(model.text)
    expect(paragraphs.map(textOf)).toContain(
      '\u4e2d\u6587\u6bb5\u843d\uff1a\u6df7\u5408\u4e2d\u82f1\u6587 ABC \u6d4b\u8bd5\u3002',
    )
    expect(paragraphs.map(textOf).some((text) => text.includes('😀'))).toBe(true)
    expect(paragraphs.map(textOf)).toContain('Column A\tColumn B')
    expect(paragraphs.map(textOf)).toContain('')
  })

  it('writes no raw control characters and uses w:tab for the tab', () => {
    // eslint-disable-next-line no-control-regex
    expect(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(documentXml)).toBe(false)
    expect(documentXml).toContain('<w:tab/>')
    expect(documentXml).toContain('\u4e2d\u6587\u6bb5\u843d')
  })
})

describe('minimal and corrupt inputs', () => {
  it('converts an empty document into a docx with one empty paragraph', async () => {
    const { docx, warnings } = await convertLegacyDocToDocx(buildSyntheticDoc({}))
    const parsed = await parseDocx(docx)
    const paragraphs = paragraphsOf(parsed)
    expect(paragraphs).toHaveLength(1)
    expect(textOf(paragraphs[0]!)).toBe('')
    expect(warnings).toEqual([])
  })

  it('rejects an empty buffer with not_cfb', async () => {
    expect(await rejectionCode(() => convertLegacyDocToDocx(new Uint8Array(0)))).toBe('not_cfb')
  })

  it('rejects a truncated .doc with a LegacyDocError', async () => {
    expect(
      await rejectionCode(() => convertLegacyDocToDocx(fixtureBytes('text-basic').slice(0, 200))),
    ).toBe('truncated')
  })
})

describe('generated zip structure', () => {
  it('contains the required OOXML parts and keeps the blank template sectPr', async () => {
    const { docx } = await convertLegacyDocToDocx(fixtureBytes('formatting'))
    const zip = await JSZip.loadAsync(docx)
    for (const name of [
      '[Content_Types].xml',
      '_rels/.rels',
      'word/document.xml',
      'word/_rels/document.xml.rels',
      'word/styles.xml',
      'word/numbering.xml',
    ])
      expect(zip.file(name)).not.toBeNull()
    const documentXml = await zip.file('word/document.xml')!.async('string')
    expect(documentXml).toContain('<?xml')
    expect(documentXml).toContain('<w:document')
    expect(documentXml).toContain('<w:body>')
    expect(documentXml).toContain('<w:sectPr>')
    expect(documentXml).toContain('Bold text')
    expect(documentXml).toContain('<w:p>')
    expect(documentXml.trim().endsWith('</w:document>')).toBe(true)
    // parseDocx parses the part (docx-engine's XML parser rejects malformed XML)
    await expect(parseDocx(docx)).resolves.toBeDefined()
  })

  it('writeLegacyDocx is usable directly on a parsed model', async () => {
    const model = readLegacyDocModel(fixtureBytes('word97-sample'))
    const docx = await writeLegacyDocx(model)
    const parsed = await parseDocx(docx)
    expect(paragraphsOf(parsed).map(textOf)).toEqual(
      model.paragraphs.map((paragraph) => paragraph.text),
    )
  })
})

describe('honest conversion warnings', () => {
  it('warns that styles are not resolved yet for a non-zero paragraph style index', async () => {
    const bytes = buildSyntheticDoc({
      pieces: [{ text: 'styled\r' }],
      papxFkpPages: [encodePapxFkp([{ fcStart: textFc(0), fcEnd: textFc(7), istd: 1 }])],
    })
    const { warnings } = await convertLegacyDocToDocx(bytes)
    expect(warnings.some((warning) => warning.includes('styles are not resolved'))).toBe(true)
  })

  it('warns that styles are not resolved yet for a non-zero character style index', async () => {
    const bytes = buildSyntheticDoc({
      pieces: [{ text: 'styled\r' }],
      chpxFkpPages: [
        encodeChpxFkp([{ fcStart: textFc(0), fcEnd: textFc(7), grpprl: sprm(0x4a30, [2, 0]) }]),
      ],
    })
    const { warnings } = await convertLegacyDocToDocx(bytes)
    expect(warnings.some((warning) => warning.includes('styles are not resolved'))).toBe(true)
  })

  it('warns that tables are flattened to paragraphs for inTable paragraphs', async () => {
    const bytes = buildSyntheticDoc({
      pieces: [{ text: 'cell\r' }],
      papxFkpPages: [
        encodePapxFkp([{ fcStart: textFc(0), fcEnd: textFc(5), grpprl: sprm(0x2416, [1]) }]),
      ],
    })
    const { warnings } = await convertLegacyDocToDocx(bytes)
    expect(warnings.some((warning) => warning.includes('tables are flattened'))).toBe(true)
  })

  it('warns about headers and text boxes from the story counts', async () => {
    const header = await convertLegacyDocToDocx(
      buildSyntheticDoc({ pieces: [{ text: 'x\r' }], ccpHdd: 6 }),
    )
    expect(header.warnings.some((warning) => warning.includes('headers are not converted'))).toBe(
      true,
    )
    const textbox = await convertLegacyDocToDocx(
      buildSyntheticDoc({ pieces: [{ text: 'x\r' }], ccpTxbx: 5, ccpHdrTxbx: 2 }),
    )
    expect(textbox.warnings).toContain(
      'the document declares body text box text but has no text box table; those text boxes are omitted',
    )
    expect(textbox.warnings).toContain(
      'the document declares header text box text but has no text box table; those text boxes are omitted',
    )
  })

  it('keeps the stage-1/2 model warnings', async () => {
    // a paragraph whose properties live in a missing Data stream produces a model warning
    const bytes = buildSyntheticDoc({
      pieces: [{ text: 'x\r' }],
      papxFkpPages: [
        encodePapxFkp([
          { fcStart: textFc(0), fcEnd: textFc(2), grpprl: sprm(0x6645, [0, 0, 0, 0]) },
        ]),
      ],
    })
    const model = readLegacyDocModel(bytes)
    expect(model.warnings).toHaveLength(1)
    const { warnings } = await convertLegacyDocToDocx(bytes)
    expect(warnings[0]).toBe(model.warnings[0])
    expect(warnings).toContain(model.warnings[0])
  })

  it('emits no warnings for a clean synthetic document', async () => {
    const { warnings } = await convertLegacyDocToDocx(
      buildSyntheticDoc({ pieces: [{ text: 'clean\r' }] }),
    )
    expect(warnings).toEqual([])
  })
})

describe('document defaults of a converted file', () => {
  it('writes Word 97 paragraph defaults and the dominant run languages into styles.xml', async () => {
    const text = 'ab\r'
    const bytes = buildSyntheticDoc({
      pieces: [{ text }],
      chpxFkpPages: [
        encodeChpxFkp([
          {
            fcStart: textFc(0),
            fcEnd: textFc(2),
            grpprl: new Uint8Array([...sprm(0x486d, [0x09, 0x04]), ...sprm(0x486e, [0x04, 0x08])]),
          },
        ]),
      ],
    })
    const { docx } = await convertLegacyDocToDocx(bytes)
    const styles = await (await JSZip.loadAsync(docx)).file('word/styles.xml')!.async('string')
    expect(styles).toContain('<w:lang w:val="en-US" w:eastAsia="zh-CN"/></w:rPr></w:rPrDefault>')
    expect(styles).toContain(
      '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>',
    )
    const plain = await convertLegacyDocToDocx(buildSyntheticDoc({ pieces: [{ text }] }))
    const plainStyles = await (
      await JSZip.loadAsync(plain.docx)
    )
      .file('word/styles.xml')!
      .async('string')
    expect(plainStyles).not.toContain('<w:lang')
  })

  it('ignores neutral and unmapped LCIDs when picking the document language', async () => {
    const text = 'abcdefgh\r'
    const bytes = buildSyntheticDoc({
      pieces: [{ text }],
      chpxFkpPages: [
        encodeChpxFkp([
          {
            fcStart: textFc(0),
            fcEnd: textFc(6),
            grpprl: new Uint8Array([...sprm(0x486d, [0x00, 0x04]), ...sprm(0x486e, [0x00, 0x00])]),
          },
          {
            fcStart: textFc(6),
            fcEnd: textFc(8),
            grpprl: new Uint8Array([...sprm(0x486d, [0x09, 0x04]), ...sprm(0x486e, [0x04, 0x08])]),
          },
        ]),
      ],
    })
    const { docx } = await convertLegacyDocToDocx(bytes)
    const styles = await (await JSZip.loadAsync(docx)).file('word/styles.xml')!.async('string')
    expect(styles).toContain('<w:lang w:val="en-US" w:eastAsia="zh-CN"/>')
  })
})

describe('readLegacyDoc (single parse)', () => {
  it('matches the separate model, layout and body readers', () => {
    for (const name of ['page-setup', 'styles', 'header-footer'] as const) {
      const bytes = fixtureBytes(name)
      const all = readLegacyDoc(bytes)
      const model = readResolvedLegacyDocModel(bytes)
      expect(all.model.paragraphs).toEqual(model.paragraphs)
      expect(all.model.warnings).toEqual(model.warnings)
      expect(all.layout).toEqual(readLegacyDocLayout(bytes))
      expect(all.body.blocks).toEqual(readLegacyDocBody(bytes).blocks)
      expect(all.stylesResolved).toBe(true)
    }
  })
})

describe('styles.doc -> docx round trip (resolved stylesheet)', () => {
  let model: LegacyDocModel
  let docx: Uint8Array
  let warnings: string[]
  let parsed: Parsed
  let paragraphs: Block[]

  beforeAll(async () => {
    model = readResolvedLegacyDocModel(fixtureBytes('styles'))
    const result = await convertLegacyDocToDocx(fixtureBytes('styles'))
    docx = result.docx
    warnings = result.warnings
    parsed = await parseDocx(docx)
    paragraphs = paragraphsOf(parsed)
  })

  it('writes every paragraph and no longer warns about unresolved styles', () => {
    expect(paragraphs.map(textOf)).toEqual(expectedText('styles').split('\n'))
    expect(docx.length).toBeGreaterThan(1000)
    expect(warnings.some((warning) => warning.includes('styles are not resolved'))).toBe(false)
    expect(warnings).toEqual([])
  })

  it('P0 Title becomes 26pt dark-blue direct formatting', () => {
    expect(paragraphs[0]!.runs?.[0]).toMatchObject({
      text: 'Document title',
      sizeHalfPoints: 52,
      color: '17365D',
    })
    expect(paragraphs[0]!.format).toMatchObject({ spaceAfter: 300 })
  })

  it('P1 Heading 1 becomes bold 14pt with keep/spacing', () => {
    expect(paragraphs[1]!.runs?.[0]).toMatchObject({
      text: 'Heading one text',
      bold: true,
      sizeHalfPoints: 28,
    })
    expect(paragraphs[1]!.format).toMatchObject({
      keepLines: true,
      keepNext: true,
      spaceBefore: 480,
    })
  })

  it('P2 Body Bold becomes bold 16pt centered with 12/6pt spacing', () => {
    expect(paragraphs[2]!.runs?.[0]).toMatchObject({ bold: true, sizeHalfPoints: 32 })
    expect(paragraphs[2]!.format).toMatchObject({
      align: 'center',
      spaceBefore: 240,
      spaceAfter: 120,
    })
  })

  it('P3 keeps the plain run and the styled run (italic 12pt red palette)', () => {
    const runs = paragraphs[3]!.runs ?? []
    expect(runs.map((run) => run.text)).toEqual(['Plain run and ', 'styled emphasis run'])
    expect(runs[0]!).toMatchObject({ sizeHalfPoints: 22 })
    expect(runs[1]!).toMatchObject({ italic: true, sizeHalfPoints: 24, color: 'CC0000' })
  })

  it('P4 keeps the direct right alignment over the style center alignment', () => {
    expect(paragraphs[4]!.format).toMatchObject({
      align: 'right',
      spaceBefore: 240,
      spaceAfter: 120,
    })
  })

  it('P5 keeps the derived style first-line indent and center alignment', () => {
    expect(paragraphs[5]!.format).toMatchObject({ align: 'center', indentFirstLine: 360 })
    expect(paragraphs[5]!.runs?.[0]).toMatchObject({ bold: true, sizeHalfPoints: 32 })
  })

  it('P6 keeps the bold Chinese paragraph', () => {
    expect(textOf(paragraphs[6]!)).toBe('\u4e2d\u6587\u6837\u5f0f\u6bb5\u843d')
    expect(paragraphs[6]!.runs?.[0]).toMatchObject({ bold: true, sizeHalfPoints: 32 })
    expect(paragraphs[6]!.format).toMatchObject({ align: 'center' })
  })

  it('the resolved model agrees with the written paragraphs', () => {
    expect(paragraphs.map(textOf)).toEqual(model.paragraphs.map((paragraph) => paragraph.text))
    expect(model.warnings).toEqual([])
  })

  it('writes the paragraph mark rPr so empty paragraphs keep their height, and w:caps', async () => {
    const text = 'ab\r\rcd\r'
    const bytes = buildSyntheticDoc({
      pieces: [{ text }],
      chpxFkpPages: [
        encodeChpxFkp([
          { fcStart: textFc(0), fcEnd: textFc(2), grpprl: sprm(0x083b, [1]) },
          { fcStart: textFc(2), fcEnd: textFc(4), grpprl: sprm(0x4a43, [40, 0]) },
        ]),
      ],
    })
    const { docx } = await convertLegacyDocToDocx(bytes)
    const xml = await (await JSZip.loadAsync(docx)).file('word/document.xml')!.async('string')
    const paragraphs = xml.match(/<w:p>.*?<\/w:p>|<w:p\/>/g)!
    expect(paragraphs[0]).toContain('<w:caps/>')
    expect(paragraphs[0]).toContain(
      '<w:pPr><w:rPr><w:sz w:val="40"/><w:szCs w:val="40"/></w:rPr></w:pPr>',
    )
    expect(paragraphs[1]).toBe(
      '<w:p><w:pPr><w:rPr><w:sz w:val="40"/><w:szCs w:val="40"/></w:rPr></w:pPr></w:p>',
    )
    expect(paragraphs[2]).not.toContain('<w:rPr>')
  })
})
