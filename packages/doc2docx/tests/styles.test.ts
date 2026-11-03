import { describe, expect, it } from 'vitest'
import {
  LegacyDocError,
  openLegacyDoc,
  parseStsh,
  readLegacyDocModel,
  readResolvedLegacyDocModel,
} from '../src/index'
import {
  SYNTHETIC_TEXT_OFFSET,
  buildSyntheticDoc,
  encodeChpxFkp,
  encodePapxFkp,
  encodeStsh,
  type SyntheticStyle,
} from './helpers/synthetic-doc'
import { expectedText, fixtureBytes } from './helpers/fixtures'

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

interface SyntheticParagraphSpec {
  istd?: number
  sprms?: Uint8Array
  defaults?: boolean
}

/**
 * A synthetic doc with an optional stylesheet: one PAPX entry per paragraph of
 * the text (splitting on the paragraph marks) and one CHPX run over the whole
 * text when run sprms are given.
 */
function syntheticDoc(options: {
  styles?: SyntheticStyle[]
  text?: string
  paragraphs?: SyntheticParagraphSpec[]
  runSprms?: Uint8Array
  stshCbStdBase?: 0x0a | 0x12
  stsh?: Uint8Array
}): Uint8Array {
  const text = options.text ?? 'styled\r'
  const starts = [0]
  for (let index = 0; index < text.length; index++) if (text[index] === '\r') starts.push(index + 1)
  const paragraphStarts = starts.slice(0, -1)
  const papxPages =
    options.paragraphs === undefined
      ? []
      : [
          encodePapxFkp(
            options.paragraphs.map((paragraph, index) => ({
              fcStart: textFc(paragraphStarts[index]!),
              fcEnd: textFc(paragraphStarts[index + 1] ?? text.length),
              istd: paragraph.istd ?? 0,
              ...(paragraph.sprms ? { grpprl: paragraph.sprms } : {}),
              ...(paragraph.defaults ? { defaults: true } : {}),
            })),
          ),
        ]
  const chpxPages = options.runSprms
    ? [
        encodeChpxFkp([
          { fcStart: textFc(0), fcEnd: textFc(text.length), grpprl: options.runSprms },
        ]),
      ]
    : []
  return buildSyntheticDoc({
    pieces: [{ text }],
    ...(options.stsh ? { stsh: options.stsh } : {}),
    ...(options.styles ? { stshStyles: options.styles } : {}),
    ...(options.stshCbStdBase ? { stshCbStdBase: options.stshCbStdBase } : {}),
    ...(papxPages.length > 0 ? { papxFkpPages: papxPages } : {}),
    ...(chpxPages.length > 0 ? { chpxFkpPages: chpxPages } : {}),
  })
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

describe('real styles.doc: resolved stylesheet formatting', () => {
  const model = readResolvedLegacyDocModel(fixtureBytes('styles'))
  const direct = readLegacyDocModel(fixtureBytes('styles'))

  it('keeps the paragraph texts of styles.txt (7 paragraphs + trailing empty)', () => {
    expect(model.paragraphs.map((paragraph) => paragraph.text)).toEqual(
      expectedText('styles').split('\n'),
    )
    expect(model.warnings).toEqual([])
    for (const paragraph of model.paragraphs)
      expect(paragraph.runs.map((run) => run.text).join('')).toBe(paragraph.text)
  })

  it('P0 applies the Title style chain (Normal + Title)', () => {
    expect(model.paragraphs[0]!.formatting).toMatchObject({
      alignment: 'left',
      lineSpacing: { dyaLine: 240, multiple: true },
      spacingBeforeTwips: 0,
      spacingAfterTwips: 300,
      styleIndex: 61,
    })
    expect(model.paragraphs[0]!.runs[0]!.formatting).toMatchObject({
      colorIndex: 9,
      fontSizeHalfPoints: 52,
      ftcAscii: 7,
      ftcEastAsia: 8,
    })
    expect(model.paragraphs[0]!.runs[0]!.formatting.bold).toBeUndefined()
  })

  it('P1 applies Heading 1 over Normal', () => {
    expect(model.paragraphs[1]!.formatting).toMatchObject({
      alignment: 'left',
      keepLines: true,
      keepNext: true,
      outlineLevel: 0,
      spacingBeforeTwips: 480,
      spacingAfterTwips: 0,
      styleIndex: 1,
    })
    expect(model.paragraphs[1]!.runs[0]!.formatting).toMatchObject({
      bold: true,
      colorIndex: 9,
      fontSizeHalfPoints: 28,
    })
  })

  it('P2 applies the custom Body Bold style (bold, 16pt, centered, 12/6pt spacing)', () => {
    expect(model.paragraphs[2]!.formatting).toMatchObject({
      alignment: 'center',
      spacingBeforeTwips: 240,
      spacingAfterTwips: 120,
      lineSpacing: { dyaLine: 276, multiple: true },
      styleIndex: 82,
    })
    expect(model.paragraphs[2]!.runs[0]!.formatting).toMatchObject({
      bold: true,
      fontSizeHalfPoints: 32,
    })
  })

  it('P3 keeps the plain run and applies the custom character style to the second run', () => {
    const runs = model.paragraphs[3]!.runs
    expect(runs.map((run) => run.text)).toEqual(['Plain run and ', 'styled emphasis run'])
    expect(runs[0]!.formatting).toMatchObject({ fontSizeHalfPoints: 22 })
    expect(runs[0]!.formatting.italic).toBeUndefined()
    expect(runs[1]!.formatting).toMatchObject({
      italic: true,
      fontSizeHalfPoints: 24,
      colorIndex: 6,
      styleIndex: 42,
    })
  })

  it('P4 lets the direct right alignment win over the style center alignment', () => {
    expect(model.paragraphs[4]!.formatting).toMatchObject({
      alignment: 'right',
      spacingBeforeTwips: 240,
      spacingAfterTwips: 120,
      styleIndex: 82,
    })
    expect(model.paragraphs[4]!.runs[0]!.formatting).toMatchObject({ bold: true })
  })

  it('P5 applies the derived Body Bold Center style (basedOn Body Bold + first-line indent)', () => {
    expect(model.paragraphs[5]!.formatting).toMatchObject({
      alignment: 'center',
      indentFirstLineTwips: 360,
      spacingBeforeTwips: 240,
      spacingAfterTwips: 120,
      styleIndex: 83,
    })
    expect(model.paragraphs[5]!.runs[0]!.formatting).toMatchObject({
      bold: true,
      fontSizeHalfPoints: 32,
    })
  })

  it('P6 applies Body Bold to the Chinese paragraph', () => {
    expect(model.paragraphs[6]!.text).toBe('\u4e2d\u6587\u6837\u5f0f\u6bb5\u843d')
    expect(model.paragraphs[6]!.formatting).toMatchObject({ alignment: 'center', styleIndex: 82 })
    expect(model.paragraphs[6]!.runs[0]!.formatting).toMatchObject({ bold: true })
  })

  it('the direct-formatting model (stage 2 read) is unchanged: style indices only', () => {
    expect(direct.paragraphs[2]!.formatting).toEqual({ styleIndex: 82 })
    expect(direct.paragraphs[2]!.runs[0]!.formatting).toEqual({})
    expect(direct.paragraphs[3]!.runs[1]!.formatting).toEqual({ styleIndex: 42 })
  })
})

describe('real styles.doc: STSH decoding', () => {
  const doc = openLegacyDoc(fixtureBytes('styles'))
  const sheet = parseStsh(doc.table, doc.fib.fcStshf, doc.fib.lcbStshf)

  it('decodes the built-in and custom styles used by the fixture', () => {
    expect(doc.fib.lcbStshf).toBeGreaterThan(0)
    expect(sheet.styles.size).toBeGreaterThan(40)
    expect(sheet.styles.get(0)).toMatchObject({ name: 'Normal', kind: 'paragraph', sti: 0 })
    expect(sheet.styles.get(1)).toMatchObject({
      name: 'heading 1',
      kind: 'paragraph',
      sti: 1,
      basedOn: 0,
    })
    expect(sheet.styles.get(42)).toMatchObject({ name: 'Emphasis Custom', kind: 'character' })
    expect(sheet.styles.get(42)!.basedOn).toBeUndefined()
    expect(sheet.styles.get(61)).toMatchObject({ name: 'Title', kind: 'paragraph', basedOn: 0 })
    expect(sheet.styles.get(82)).toMatchObject({ name: 'Body Bold', basedOn: 0 })
    expect(sheet.styles.get(83)).toMatchObject({ name: 'Body Bold Center', basedOn: 82 })
    // the Word 97 default character style slot is an empty STD
    expect(sheet.styles.get(10)).toMatchObject({
      kind: 'paragraph',
      name: '',
      papxSprms: new Uint8Array(0),
    })
    expect(sheet.styles.get(82)!.papxSprms.length).toBeGreaterThan(0)
    expect(sheet.styles.get(42)!.chpxSprms.length).toBeGreaterThan(0)
  })
})

describe('synthetic stylesheets: chains and defaults', () => {
  it('applies a paragraph style chain base -> derived before the direct sprms', () => {
    const bytes = syntheticDoc({
      styles: [
        { name: 'Normal', kind: 'paragraph', sti: 0, papxSprms: sprm(0xa413, [0x64, 0x00]) },
        { name: 'Centered', kind: 'paragraph', basedOn: 0, papxSprms: sprm(0x2461, [0x01]) },
      ],
      text: 'one\rtwo\r',
      paragraphs: [{ istd: 1 }, { istd: 0 }],
    })
    const model = readResolvedLegacyDocModel(bytes)
    expect(model.warnings).toEqual([])
    expect(model.paragraphs[0]!.formatting).toMatchObject({
      spacingBeforeTwips: 100,
      alignment: 'center',
      styleIndex: 1,
    })
    expect(model.paragraphs[1]!.formatting).toMatchObject({
      spacingBeforeTwips: 100,
      styleIndex: 0,
    })
  })

  it('applies character style chains after the paragraph style CHPX', () => {
    const bytes = syntheticDoc({
      styles: [
        { name: 'Normal', kind: 'paragraph', sti: 0, chpxSprms: sprm(0x4a43, [0x16, 0x00]) },
        { name: 'Base', kind: 'character', chpxSprms: sprm(0x0836, [1]) },
        { name: 'Derived', kind: 'character', basedOn: 1, chpxSprms: sprm(0x2a42, [6]) },
      ],
      runSprms: concat(sprm(0x4a30, [2, 0]), sprm(0x4a43, [0x1c, 0x00])),
    })
    const model = readResolvedLegacyDocModel(bytes)
    const formatting = model.paragraphs[0]!.runs[0]!.formatting
    expect(formatting).toMatchObject({
      italic: true,
      colorIndex: 6,
      fontSizeHalfPoints: 28,
      styleIndex: 2,
    })
    expect(model.warnings).toEqual([])
  })

  it('applies the default character style (istd 10) to runs without a character style', () => {
    const styles: SyntheticStyle[] = [
      { name: 'Normal', kind: 'paragraph', sti: 0 },
      { empty: true },
      { empty: true },
      { empty: true },
      { empty: true },
      { empty: true },
      { empty: true },
      { empty: true },
      { empty: true },
      { empty: true },
      { name: 'Default Paragraph Font', kind: 'character', chpxSprms: sprm(0x0837, [1]) },
    ]
    const bytes = syntheticDoc({ styles })
    const formatting = readResolvedLegacyDocModel(bytes).paragraphs[0]!.runs[0]!.formatting
    expect(formatting).toMatchObject({ strike: true })
  })

  it('accepts the Word 2000+ STD base length (0x12)', () => {
    const bytes = syntheticDoc({
      styles: [
        { name: 'Normal', kind: 'paragraph', sti: 0, papxSprms: sprm(0xa413, [0xf0, 0x00]) },
      ],
      stshCbStdBase: 0x12,
    })
    expect(readResolvedLegacyDocModel(bytes).paragraphs[0]!.formatting).toMatchObject({
      spacingBeforeTwips: 240,
    })
  })
})

describe('synthetic stylesheets: toggle layering', () => {
  function toggleDoc(styleBold: number | undefined, directBold: number): Uint8Array {
    const styles: SyntheticStyle[] = [
      {
        name: 'Normal',
        kind: 'paragraph',
        sti: 0,
        ...(styleBold === undefined ? {} : { chpxSprms: sprm(0x0835, [styleBold]) }),
      },
    ]
    return syntheticDoc({ styles, runSprms: sprm(0x0835, [directBold]) })
  }

  function boldOf(bytes: Uint8Array): boolean | undefined {
    return readResolvedLegacyDocModel(bytes).paragraphs[0]!.runs[0]!.formatting.bold
  }

  it('0x80 keeps the inherited value from the style chain', () => {
    expect(boldOf(toggleDoc(1, 0x80))).toBe(true)
    expect(boldOf(toggleDoc(0, 0x80))).toBe(false)
    expect(boldOf(toggleDoc(undefined, 0x80))).toBeUndefined()
  })

  it('0x81 inverts the inherited value (the chain base is off)', () => {
    expect(boldOf(toggleDoc(1, 0x81))).toBe(false)
    expect(boldOf(toggleDoc(0, 0x81))).toBe(true)
    expect(boldOf(toggleDoc(undefined, 0x81))).toBe(true)
  })

  it('explicit 0/1 in the direct layer still wins', () => {
    expect(boldOf(toggleDoc(1, 0))).toBe(false)
    expect(boldOf(toggleDoc(0, 1))).toBe(true)
  })

  it('paragraph toggles use the same layering (keepNext 0x80/0x81)', () => {
    const onInStyle = readResolvedLegacyDocModel(
      syntheticDoc({
        styles: [{ name: 'Normal', kind: 'paragraph', sti: 0, papxSprms: sprm(0x2406, [1]) }],
        paragraphs: [{ istd: 0, sprms: sprm(0x2406, [0x80]) }],
      }),
    )
    expect(onInStyle.paragraphs[0]!.formatting.keepNext).toBe(true)
    const inverted = readResolvedLegacyDocModel(
      syntheticDoc({
        styles: [{ name: 'Normal', kind: 'paragraph', sti: 0, papxSprms: sprm(0x2406, [1]) }],
        paragraphs: [{ istd: 0, sprms: sprm(0x2406, [0x81]) }],
      }),
    )
    expect(inverted.paragraphs[0]!.formatting.keepNext).toBe(false)
  })
})

describe('synthetic stylesheets: broken chains', () => {
  const styles: SyntheticStyle[] = [
    { name: 'Normal', kind: 'paragraph', sti: 0, papxSprms: sprm(0xa413, [0x64, 0x00]) },
  ]

  it('warns about a paragraph style index outside the stylesheet and keeps the direct sprms', () => {
    const bytes = syntheticDoc({
      styles,
      paragraphs: [{ istd: 50, sprms: sprm(0x2461, [0x02]) }],
    })
    const model = readResolvedLegacyDocModel(bytes)
    expect(model.warnings).toHaveLength(1)
    expect(model.warnings[0]).toContain('style index 50')
    expect(model.paragraphs[0]!.formatting).toMatchObject({ alignment: 'right', styleIndex: 50 })
  })

  it('warns once about a run character style outside the stylesheet', () => {
    const bytes = syntheticDoc({
      styles,
      runSprms: sprm(0x4a30, [50, 0]),
      text: 'ab\r',
    })
    const model = readResolvedLegacyDocModel(bytes)
    expect(model.warnings.filter((warning) => warning.includes('style index 50'))).toHaveLength(1)
    expect(model.paragraphs[0]!.runs[0]!.formatting).toMatchObject({ styleIndex: 50 })
  })

  it('stops a style chain at a cycle and reports it once', () => {
    const cycles: SyntheticStyle[] = [
      { name: 'A', kind: 'paragraph', basedOn: 1, papxSprms: sprm(0x2461, [0x01]) },
      { name: 'B', kind: 'paragraph', basedOn: 0, papxSprms: sprm(0xa413, [0xf0, 0x00]) },
    ]
    const bytes = syntheticDoc({ styles: cycles, paragraphs: [{ istd: 1 }], text: 'styled' })
    const model = readResolvedLegacyDocModel(bytes)
    expect(model.warnings).toHaveLength(1)
    expect(model.warnings[0]).toContain('cycle')
    expect(model.paragraphs[0]!.formatting).toMatchObject({
      alignment: 'center',
      spacingBeforeTwips: 240,
    })
  })

  it('keeps direct formatting when the document has no stylesheet at all', () => {
    const bytes = buildSyntheticDoc({
      pieces: [{ text: 'plain\r' }],
      papxFkpPages: [
        encodePapxFkp([{ fcStart: textFc(0), fcEnd: textFc(6), grpprl: sprm(0x2461, [0x01]) }]),
      ],
    })
    const resolved = readResolvedLegacyDocModel(bytes)
    const direct = readLegacyDocModel(bytes)
    expect(resolved.warnings).toEqual([])
    expect(resolved.paragraphs[0]!.formatting).toEqual({ styleIndex: 0, alignment: 'center' })
    expect(resolved.paragraphs).toEqual(direct.paragraphs)
  })
})

describe('stylesheet errors', () => {
  it('rejects a stylesheet too short for its header', () => {
    const bytes = syntheticDoc({ styles: [], stsh: Uint8Array.from([0x01, 0x00]) })
    expect(errorCode(() => readResolvedLegacyDocModel(bytes))).toBe('bad_stsh')
  })

  it('rejects an unsupported STD base length', () => {
    const stsh = encodeStsh([{ name: 'Normal', kind: 'paragraph' }])
    new DataView(stsh.buffer).setUint16(4, 0x0b, true) // cbSTDBaseInFile
    const bytes = syntheticDoc({ stsh })
    expect(errorCode(() => readResolvedLegacyDocModel(bytes))).toBe('bad_stsh')
  })

  it('rejects an unknown style kind (stk)', () => {
    const stsh = encodeStsh([{ name: 'Normal', kind: 'paragraph' }])
    // style 0 STD starts after cbStshi(2) + STSHI(18) + cbStd(2): info2 is 2 bytes in
    new DataView(stsh.buffer).setUint16(2 + 18 + 2 + 2, 5, true)
    const bytes = syntheticDoc({ stsh })
    expect(errorCode(() => readResolvedLegacyDocModel(bytes))).toBe('bad_stsh')
  })

  it('rejects a style entry that overruns the stylesheet', () => {
    const stsh = encodeStsh([{ name: 'Normal', kind: 'paragraph' }])
    new DataView(stsh.buffer).setUint16(2 + 18, 0xffff, true) // cbStd of style 0
    const bytes = syntheticDoc({ stsh })
    expect(errorCode(() => readResolvedLegacyDocModel(bytes))).toBe('bad_stsh')
  })

  it('rejects a truncated name inside a style entry', () => {
    const stsh = encodeStsh([{ name: 'Normal', kind: 'paragraph' }])
    // cch sits right after cbStd(2) + STSHI(18) + cbStd size(2) + StdfBase(10)
    new DataView(stsh.buffer).setUint16(2 + 18 + 2 + 10, 500, true)
    const bytes = syntheticDoc({ stsh })
    expect(errorCode(() => readResolvedLegacyDocModel(bytes))).toBe('bad_stsh')
  })
})
