import { describe, expect, it } from 'vitest'
import { mergePPrFormat, parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

/**
 * Word merges w:ind per attribute, direct value over numbering over style, and
 * an explicit 0 is a value: `w:firstLine="0"` on a paragraph whose style says
 * firstLine=400 lays out flush (prod-sas 004/006), a child style's
 * `w:ind w:left="0"` cancels Normal's left indent (prod-sas 075).
 */

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const STYLES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}>` +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>' +
  '<w:pPr><w:ind w:left="454" w:right="120"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:ind w:firstLine="400"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="HDR"><w:name w:val="HDR"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:ind w:left="0" w:right="0" w:hanging="0"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Hang"><w:name w:val="Hang"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:ind w:hanging="300"/></w:pPr></w:style>' +
  '</w:styles>'

const para = (pPr: string) => `<w:p><w:pPr>${pPr}</w:pPr><w:r><w:t>text</w:t></w:r></w:p>`

describe('explicit zero indents', () => {
  it('keeps a direct 0 for every w:ind component and leaves absent ones unset', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          para('<w:pStyle w:val="Body"/><w:ind w:left="0" w:right="0" w:firstLine="0"/>') +
          para('<w:pStyle w:val="Hang"/><w:ind w:hanging="0"/>') +
          para(
            '<w:pStyle w:val="Body"/><w:ind w:left="0" w:leftChars="0" w:firstLine="0" w:firstLineChars="0"/>',
          ) +
          para('<w:pStyle w:val="Body"/><w:ind w:right="0"/>') +
          para('<w:pStyle w:val="Body"/>'),
        stylesXml: STYLES,
      }),
    )
    const [zeros, hangZero, charZeros, rightOnly, plain] = doc.blocks.map((b) => b.format)
    expect(zeros).toMatchObject({ indentLeft: 0, indentRight: 0, indentFirstLine: 0 })
    expect(hangZero).toMatchObject({ indentFirstLine: 0 })
    expect(hangZero?.indentLeft).toBeUndefined()
    expect(charZeros).toMatchObject({ indentLeft: 0, indentFirstLine: 0 })
    expect(rightOnly).toMatchObject({ indentRight: 0 })
    expect(rightOnly?.indentFirstLine).toBeUndefined()
    expect(rightOnly?.indentLeft).toBeUndefined()
    expect(plain).toBeUndefined()
  })

  it('a child style zero overrides the basedOn chain, absent components inherit', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: para(''), stylesXml: STYLES }))
    expect(doc.styles.get('Normal')?.display).toMatchObject({
      indentLeftTwips: 454,
      indentRightTwips: 120,
    })
    const hdr = doc.styles.get('HDR')?.display
    expect(hdr).toMatchObject({ indentLeftTwips: 0, indentRightTwips: 0, indentFirstLineTwips: 0 })
    const body = doc.styles.get('Body')?.display
    expect(body).toMatchObject({
      indentLeftTwips: 454,
      indentRightTwips: 120,
      indentFirstLineTwips: 400,
    })
    expect(doc.styles.get('Hang')?.display?.indentFirstLineTwips).toBe(-300)
  })

  it('round-trips explicit zeros through a paragraph rebuild', () => {
    const raw = '<w:pPr><w:ind w:left="0" w:right="0" w:firstLine="0"/><w:jc w:val="both"/></w:pPr>'
    const model = { indentLeft: 0, indentRight: 0, indentFirstLine: 0, align: 'justify' as const }
    expect(mergePPrFormat(raw, model)).toBe(raw)
    expect(mergePPrFormat(raw, { ...model, align: 'center' })).toContain(
      '<w:ind w:left="0" w:right="0" w:firstLine="0"/>',
    )
    // w:hanging="0" is the same value as w:firstLine="0": the untouched group keeps its bytes
    const hang = '<w:pPr><w:ind w:hanging="0"/></w:pPr>'
    expect(mergePPrFormat(hang, { indentFirstLine: 0 })).toBe(hang)
    // dropping the indent altogether still removes w:ind
    expect(mergePPrFormat(raw, { align: 'justify' })).not.toContain('<w:ind')
  })
})
