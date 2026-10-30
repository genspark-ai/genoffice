/**
 * Body and cell paragraphs and their runs read their direct-only model fields
 * off the property resolver (`source === 'direct'`); the former w:pPr / w:rPr
 * readers survive only as the shadow-mode comparison.
 */
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

const styles = (docDefaultsPPr = '', extra = '') =>
  `${XML_DECL}<w:styles ${W}><w:docDefaults>` +
  '<w:rPrDefault><w:rPr><w:sz w:val="21"/></w:rPr></w:rPrDefault>' +
  `<w:pPrDefault><w:pPr>${docDefaultsPPr}</w:pPr></w:pPrDefault></w:docDefaults>` +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Keep"><w:name w:val="Keep"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:keepNext/><w:spacing w:before="240" w:line="360" w:lineRule="auto"/><w:ind w:left="300"/></w:pPr>' +
  '<w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>' +
  `${extra}</w:styles>`

const p = (pPr: string, runs: string) => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`
const r = (rPr: string, text: string) =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`
const table = (cellPs: string) =>
  '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
  `<w:tr><w:tc>${cellPs}</w:tc></w:tr></w:tbl>`

const load = async (bodyXml: string, stylesXml = styles()) =>
  parseDocx(await buildDocx({ bodyXml, stylesXml }), { shadowResolve: true })

describe('paragraph direct keys off the resolver', () => {
  it('keeps only what the paragraph itself declares; the style chain stays off the model', async () => {
    const doc = await load(
      p(
        '<w:pStyle w:val="Keep"/><w:keepNext w:val="0"/><w:spacing w:after="120"/><w:ind w:firstLine="0"/>',
        r('', 'a'),
      ),
    )
    const f = doc.blocks[0].format!
    expect(f.keepNext).toBe(false)
    expect(f.spaceAfter).toBe(120)
    expect(f.spaceBefore).toBeUndefined()
    expect(f.lineRawTwips).toBeUndefined()
    expect(f.indentLeft).toBeUndefined()
    expect(f.indentFirstLine).toBe(0)
    expect(f.effective).toMatchObject({ keepNext: false, spacing: { before: 240, after: 120 } })
    expect(doc.extras.resolveDiff).toEqual([])
  })

  it('reads every ST_OnOff spelling of w:snapToGrid, in the body and in a cell', async () => {
    const doc = await load(
      p('<w:snapToGrid w:val="off"/>', r('', 'a')) +
        table(p('<w:snapToGrid w:val="off"/>', r('', 'c')) + p('<w:snapToGrid/>', r('', 'd'))),
    )
    expect(doc.blocks[0].format!.snapToGrid).toBe(false)
    const paras = doc.blocks[1].table!.rows[0][0].richParas!
    expect(paras[0].snapToGrid).toBe(false)
    expect(paras[1].snapToGrid).toBeUndefined()
    expect(doc.extras.resolveDiff).toEqual([])
  })

  it('a pPr-less paragraph in a docDefaults w:bidi document inherits the direction', async () => {
    const doc = await load(
      p('', r('', 'a')) + p('<w:jc w:val="left"/>', r('', 'b')) + table(p('', r('', 'c'))),
      styles('<w:bidi/>'),
    )
    expect(doc.blocks[0].format).toMatchObject({ bidiInherited: true })
    expect(doc.blocks[0].format!.bidi).toBeUndefined()
    // logical left in a bidi paragraph is the visual right
    expect(doc.blocks[1].format).toMatchObject({ bidiInherited: true, align: 'right' })
    expect(doc.blocks[2].table!.rows[0][0].richParas![0]).toMatchObject({ bidiInherited: true })
    expect(doc.extras.resolveDiff).toEqual([])
  })

  it('the shadow comparison tags its surface', async () => {
    const doc = await load(
      p('<w:spacing w:line="276" w:lineRule="auto"/>', r('', 'a')) +
        table(p('<w:spacing w:before="60"/>', r('', 'c'))),
    )
    expect(doc.blocks[0].format).toMatchObject({ lineRawTwips: 276, lineSpacing: 1.15 })
    expect(doc.blocks[1].table!.rows[0][0].richParas![0]).toMatchObject({ spaceBefore: 60 })
    expect(doc.extras.resolveDiff).toEqual([])
  })
})

describe('run direct keys off the resolver', () => {
  it('direct rPr fields come through, style-chain values stay on effective only', async () => {
    const doc = await load(
      p(
        '<w:pStyle w:val="Keep"/>',
        r('<w:b w:val="0"/><w:i/><w:spacing w:val="20"/><w:kern w:val="2"/><w:caps/>', 'a') +
          r(
            '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="SimSun" w:cs="Tahoma"/><w:vertAlign w:val="superscript"/>',
            'b',
          ) +
          r('', 'c'),
      ),
    )
    const [a, b, c] = doc.blocks[0].runs!
    expect(a).toMatchObject({
      bold: false,
      italic: true,
      charSpacingTwips: 20,
      kernHalfPoints: 2,
      caps: 'all',
    })
    expect(a.sizeHalfPoints).toBeUndefined()
    expect(a.effective!.sizeHalfPoints).toBe(28)
    expect(b).toMatchObject({
      font: 'SimSun',
      fontAscii: 'Arial',
      eastAsiaFont: 'SimSun',
      csFont: 'Tahoma',
      fontCs: 'Tahoma',
      vertAlign: 'superscript',
    })
    expect(c.bold).toBeUndefined()
    expect(c.effective!.bold).toBe(true)
    expect(doc.extras.resolveDiff).toEqual([])
  })

  it('w:u needs a w:val to underline; w:val="none" is an explicit off', async () => {
    const doc = await load(
      p(
        '',
        r('<w:u w:color="FFD966"/>', 'a') +
          r('<w:u w:val="single"/>', 'b') +
          r('<w:u w:val="none"/>', 'c'),
      ),
    )
    const [a, b, c] = doc.blocks[0].runs!
    expect(a.underline).toBeUndefined()
    expect(a.effective!.underline).toBe(false)
    expect(b.underline).toBe(true)
    expect(c.underline).toBe(false)
    expect(doc.extras.resolveDiff).toEqual([])
  })

  it('an rtl run reads the Cs twins', async () => {
    const doc = await load(
      p('', r('<w:rtl/><w:b/><w:bCs w:val="0"/><w:sz w:val="40"/><w:szCs w:val="30"/>', 'שלום')),
    )
    const run = doc.blocks[0].runs![0]
    expect(run).toMatchObject({ cs: true, bold: false, sizeHalfPoints: 30, rtl: true })
    expect(doc.extras.resolveDiff).toEqual([])
  })
})
