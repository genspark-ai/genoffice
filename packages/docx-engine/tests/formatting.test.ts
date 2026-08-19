import { describe, expect, it } from 'vitest'
import { generateParagraphXml, parseDocx, type GenerateContext } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const GEN_CTX: GenerateContext = {
  headingStyleIds: new Map([[1, 'Heading1']]),
  allocateHyperlinkRel: () => 'rId999',
}

describe('run-level formatting (font / highlight / vertAlign)', () => {
  it('parses rFonts, highlight and vertAlign from runs', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/></w:rPr><w:t>font</w:t></w:r>' +
        '<w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>hl</w:t></w:r>' +
        '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>sup</w:t></w:r>' +
        '<w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>sub</w:t></w:r></w:p>',
    })
    const doc = await parseDocx(bytes)
    const runs = doc.blocks[0].runs!
    expect(runs[0].font).toBe('Georgia')
    expect(runs[1].highlight).toBe('yellow')
    expect(runs[2].vertAlign).toBe('superscript')
    expect(runs[3].vertAlign).toBe('subscript')
  })

  it('prefers eastAsia font name when present', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="宋体"/></w:rPr><w:t>文字</w:t></w:r></w:p>',
    })
    const doc = await parseDocx(bytes)
    expect(doc.blocks[0].runs![0].font).toBe('宋体')
  })

  it('generates rPr children for the new run props', () => {
    const xml = generateParagraphXml(
      {
        type: 'paragraph',
        runs: [
          { text: 'x', font: '宋体', highlight: 'yellow', vertAlign: 'superscript', bold: true },
        ],
      },
      GEN_CTX,
    )
    expect(xml).toContain('<w:rFonts w:ascii="宋体" w:eastAsia="宋体" w:hAnsi="宋体" w:cs="宋体"/>')
    expect(xml).toContain('<w:highlight w:val="yellow"/>')
    expect(xml).toContain('<w:vertAlign w:val="superscript"/>')
    // schema order: rFonts before b, highlight after sz-position, vertAlign last
    expect(xml.indexOf('<w:rFonts')).toBeLessThan(xml.indexOf('<w:b/>'))
    expect(xml.indexOf('<w:highlight')).toBeLessThan(xml.indexOf('<w:vertAlign'))
  })

  it('round-trips run formatting through generate -> parse', async () => {
    const para = generateParagraphXml(
      {
        type: 'paragraph',
        runs: [{ text: '重点', font: '黑体', highlight: 'green', vertAlign: 'subscript' }],
      },
      GEN_CTX,
    )
    const bytes = await buildDocx({ bodyXml: para })
    const doc = await parseDocx(bytes)
    const run = doc.blocks[0].runs![0]
    expect(run).toMatchObject({
      text: '重点',
      font: '黑体',
      highlight: 'green',
      vertAlign: 'subscript',
    })
  })
})

describe('paragraph-level formatting (align / line spacing / indent)', () => {
  it('parses w:jc, w:spacing line and w:ind left', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>centered</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr><w:r><w:t>1.5x</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:ind w:left="720"/><w:jc w:val="both"/></w:pPr><w:r><w:t>indented justify</w:t></w:r></w:p>' +
        '<w:p><w:r><w:t>plain</w:t></w:r></w:p>',
    })
    const doc = await parseDocx(bytes)
    expect(doc.blocks[0].format).toEqual({ align: 'center' })
    // F1: parser now exposes lineRawTwips + lineRule + lineSpacing for auto-mode spacing
    expect(doc.blocks[1].format).toMatchObject({
      lineSpacing: 1.5,
      lineRawTwips: 360,
      lineRule: 'auto',
    })
    expect(doc.blocks[2].format).toEqual({ align: 'justify', indentLeft: 720 })
    expect(doc.blocks[3].format).toBeUndefined()
  })

  it('parses and generates first-line and hanging indents', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:pPr><w:ind w:firstLine="440"/></w:pPr><w:r><w:t>首行缩进</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:r><w:t>悬挂缩进</w:t></w:r></w:p>',
    })
    const doc = await parseDocx(bytes)
    expect(doc.blocks[0].format).toEqual({ indentFirstLine: 440 })
    expect(doc.blocks[1].format).toEqual({ indentLeft: 720, indentFirstLine: -360 })

    const firstLineXml = generateParagraphXml(
      { type: 'paragraph', format: { indentFirstLine: 440 }, runs: [{ text: 'a' }] },
      GEN_CTX,
    )
    expect(firstLineXml).toContain('<w:ind w:firstLine="440"/>')
    const hangingXml = generateParagraphXml(
      {
        type: 'paragraph',
        format: { indentLeft: 720, indentFirstLine: -360 },
        runs: [{ text: 'a' }],
      },
      GEN_CTX,
    )
    expect(hangingXml).toContain('<w:ind w:left="720" w:hanging="360"/>')
  })

  it('round-trips negative indents (text extending into the margins)', async () => {
    // Word allows signed w:left/w:right; a rebuilt paragraph must not lose
    // them (they used to be dropped, shifting the paragraph on save — r116)
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:pPr><w:ind w:left="-284" w:right="-427"/></w:pPr>' +
        '<w:r><w:t>into the margins</w:t></w:r></w:p>',
    })
    const doc = await parseDocx(bytes)
    expect(doc.blocks[0].format).toEqual({ indentLeft: -284, indentRight: -427 })

    const negativeXml = generateParagraphXml(
      {
        type: 'paragraph',
        format: { indentLeft: -284, indentRight: -427 },
        runs: [{ text: 'a' }],
      },
      GEN_CTX,
    )
    expect(negativeXml).toContain('<w:ind w:left="-284" w:right="-427"/>')
  })

  it('exposes exact/atLeast spacing as lineRawTwips+lineRule (F1)', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:pPr><w:spacing w:line="360" w:lineRule="exact"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
    })
    const doc = await parseDocx(bytes)
    // F1: exact/atLeast spacing is now exposed via lineRawTwips+lineRule instead of being ignored
    expect(doc.blocks[0].format).toMatchObject({ lineRawTwips: 360, lineRule: 'exact' })
  })

  it('generates pPr with schema-ordered children', () => {
    const xml = generateParagraphXml(
      {
        type: 'heading',
        level: 1,
        format: { align: 'justify', lineSpacing: 1.15, indentLeft: 360 },
        runs: [{ text: 'title' }],
      },
      GEN_CTX,
    )
    expect(xml).toContain('<w:pStyle w:val="Heading1"/>')
    expect(xml).toContain('<w:spacing w:line="276" w:lineRule="auto"/>')
    expect(xml).toContain('<w:ind w:left="360"/>')
    expect(xml).toContain('<w:jc w:val="both"/>')
    expect(xml.indexOf('<w:pStyle')).toBeLessThan(xml.indexOf('<w:spacing'))
    expect(xml.indexOf('<w:spacing')).toBeLessThan(xml.indexOf('<w:ind'))
    expect(xml.indexOf('<w:ind')).toBeLessThan(xml.indexOf('<w:jc'))
  })

  it('round-trips paragraph formatting through generate -> parse', async () => {
    const para = generateParagraphXml(
      {
        type: 'paragraph',
        format: { align: 'right', lineSpacing: 2, indentLeft: 720 },
        runs: [{ text: 'formatted' }],
      },
      GEN_CTX,
    )
    const bytes = await buildDocx({ bodyXml: para })
    const doc = await parseDocx(bytes)
    // F1: parser now exposes lineRawTwips+lineRule+lineSpacing for full accuracy
    expect(doc.blocks[0].format).toMatchObject({
      align: 'right',
      lineSpacing: 2,
      lineRawTwips: 480,
      lineRule: 'auto',
      indentLeft: 720,
    })
  })

  it('round-trips exact/atLeast line rule (previously only auto was recognized, so edits lost line spacing)', async () => {
    for (const lineRule of ['exact', 'atLeast'] as const) {
      const para = generateParagraphXml(
        { type: 'paragraph', format: { lineRule, lineRawTwips: 216 }, runs: [{ text: 'x' }] },
        GEN_CTX,
      )
      expect(para).toContain(`<w:spacing w:line="216" w:lineRule="${lineRule}"/>`)
      const doc = await parseDocx(await buildDocx({ bodyXml: para }))
      expect(doc.blocks[0].format).toMatchObject({ lineRule, lineRawTwips: 216 })
    }
  })
})
