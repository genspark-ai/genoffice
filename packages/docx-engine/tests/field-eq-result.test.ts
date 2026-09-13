import { describe, expect, it } from 'vitest'
import { eqFieldToOmml, inlineEqFieldResults } from '../src/eq-field'
import { ommlToMathML } from '../src/math'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const SZ36 = '<w:rPr><w:sz w:val="36"/></w:rPr>'

function field(instr: string, result: string | null, rPr = ''): string {
  return (
    `<w:r>${rPr}<w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r>${rPr}<w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
    (result === null ? '' : `<w:r>${rPr}<w:fldChar w:fldCharType="separate"/></w:r>${result}`) +
    `<w:r>${rPr}<w:fldChar w:fldCharType="end"/></w:r>`
  )
}

function cellTable(content: string): string {
  return (
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
    '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' +
    content +
    '</w:tc></w:tr></w:tbl>'
  )
}

describe('EQ field instructions', () => {
  it('lays out an array as a left-aligned matrix, one element per row', () => {
    const eq = eqFieldToOmml(' EQ \\a \\al (Milk,Eggs,Detergent,Veggies) ')!
    expect(eq.text).toBe('Milk\nEggs\nDetergent\nVeggies')
    expect(eq.omml).toContain('<m:mcJc m:val="left"/>')
    expect(eq.omml.match(/<m:mr>/g)).toHaveLength(4)
    expect(ommlToMathML(eq.omml)).toContain('<mtd style="text-align:left;padding-left:0">')
  })

  it('fills \\co columns row by row and pads the last row', () => {
    const eq = eqFieldToOmml('EQ \\a \\co2 (a,b,c)')!
    expect(eq.text).toBe('a b\nc ')
    expect(eq.omml.match(/<m:mr>/g)).toHaveLength(2)
  })

  it('fraction, radical, scripts, brackets and box map onto OMML structures', () => {
    expect(eqFieldToOmml('EQ \\f(1,2)')!.omml).toContain('<m:f><m:num>')
    expect(eqFieldToOmml('EQ \\r(3,x)')!.omml).toContain('<m:rad><m:deg>')
    expect(eqFieldToOmml('EQ \\r(x)')!.omml).toContain('<m:degHide m:val="1"/>')
    expect(eqFieldToOmml('EQ x\\s\\up6(2)')!.omml).toContain('<m:sSup><m:e/><m:sup>')
    expect(eqFieldToOmml('EQ x\\s\\do4(i)')!.omml).toContain('<m:sSub><m:e/><m:sub>')
    expect(eqFieldToOmml('EQ \\b \\bc\\[ (x)')!.omml).toContain(
      '<m:begChr m:val="["/><m:endChr m:val="]"/>',
    )
    expect(eqFieldToOmml('EQ \\x \\to \\bo (x)')!.omml).toContain(
      '<m:hideLeft m:val="1"/><m:hideRight m:val="1"/>',
    )
    expect(eqFieldToOmml('EQ \\i \\su (i=1,n,x)')!.omml).toContain('<m:chr m:val="∑"/>')
    expect(eqFieldToOmml('EQ \\o(a,b)')!.text).toBe('ab')
  })

  it('nested instructions and escaped separators parse', () => {
    const eq = eqFieldToOmml('EQ \\f(\\r(2),1\\,5)')!
    expect(eq.text).toBe('√(2)/1,5')
    expect(eq.omml).toContain('<m:num><m:rad>')
  })

  it('general switches are not part of the layout', () => {
    const eq = eqFieldToOmml('EQ \\f(1,2) \\* MERGEFORMAT')!
    expect(eq.text).toBe('1/2')
    expect(eq.omml).not.toContain('MERGEFORMAT')
  })

  it('unknown switches and non-EQ instructions return null', () => {
    expect(eqFieldToOmml('EQ \\d \\fo10 ()')).toBeNull()
    expect(eqFieldToOmml('EQ \\f(1,2')).toBeNull()
    expect(eqFieldToOmml(' PAGE ')).toBeNull()
  })

  it('plain single-letter tokens render upright', () => {
    expect(ommlToMathML(eqFieldToOmml('EQ \\f(a,b)')!.omml)).toContain(
      '<mi mathvariant="normal">a</mi>',
    )
  })

  it('a bordered box becomes a CSS-framed row', () => {
    expect(ommlToMathML(eqFieldToOmml('EQ \\x \\to (x)')!.omml)).toContain(
      '<mrow style="border-top:0.06em solid;padding:0.15em">',
    )
  })

  it('inlineEqFieldResults substitutes the layout text in the field code rPr', () => {
    const xml = `<w:p>${field('EQ \\f(1,2)', null, SZ36)}</w:p>`
    expect(inlineEqFieldResults(xml)).toBe(
      `<w:p><w:r>${SZ36}<w:t xml:space="preserve">1/2</w:t></w:r></w:p>`,
    )
  })
})

describe('EQ field paragraphs', () => {
  it('a resultless EQ array paragraph shows the laid-out matrix instead of nothing', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: `<w:p>${field('EQ \\a \\al (Milk,Eggs,Detergent,Veggies)', null)}</w:p>`,
      }),
    )
    const block = doc.blocks[0]
    expect(block.type).toBe('passthrough')
    expect(block.fieldDisplay).toMatchObject({
      kind: 'text',
      left: 'Milk\nEggs\nDetergent\nVeggies',
    })
    expect(block.fieldDisplay?.runs).toHaveLength(1)
    expect(block.fieldDisplay?.runs?.[0].math?.omml).toContain('<m:m>')
  })

  it('an EQ field inside text keeps the surrounding runs and its own size', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          `<w:p><w:r>${SZ36}<w:t xml:space="preserve">Area = </w:t></w:r>` +
          field('EQ \\f(1,2)', null, SZ36) +
          `<w:r>${SZ36}<w:t xml:space="preserve"> m</w:t></w:r></w:p>`,
      }),
    )
    const fd = doc.blocks[0].fieldDisplay!
    expect(fd).toMatchObject({ kind: 'text', left: 'Area = 1/2 m', szHalfPoints: 36 })
    expect(fd.runs?.map((r) => r.text)).toEqual(['Area = ', '1/2', ' m'])
    expect(fd.runs?.[1]).toMatchObject({ sizeHalfPoints: 36 })
    expect(fd.runs?.[1].math?.omml).toContain('<m:f>')
  })

  it('an EQ with an unsupported switch falls back to its cached result', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: `<w:p>${field('EQ \\d \\fo10 () x', '<w:r><w:t>cached</w:t></w:r>')}</w:p>`,
      }),
    )
    expect(doc.blocks[0].fieldDisplay).toMatchObject({ kind: 'text', left: 'cached' })
  })

  it('a cell EQ field becomes an inline math run', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: cellTable(`<w:p>${field('EQ \\r(x)', null)}</w:p>`) }),
    )
    const runs = doc.blocks[0].table!.rows[0][0].richParas?.[0]?.runs ?? []
    expect(runs).toHaveLength(1)
    expect(runs[0].math?.omml).toContain('<m:rad>')
  })
})

describe('inline field result formatting', () => {
  it('a simple field result keeps its own rPr instead of the paragraph default', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:pPr><w:rPr><w:sz w:val="36"/></w:rPr></w:pPr>' +
          field(
            'AUTHOR  Ada',
            '<w:r><w:rPr><w:noProof/><w:sz w:val="36"/></w:rPr><w:t>Zhe</w:t></w:r>',
            SZ36,
          ) +
          '</w:p>',
      }),
    )
    expect(doc.blocks[0].type).toBe('paragraph')
    expect(doc.blocks[0].runs?.[0]).toMatchObject({
      text: 'Zhe',
      instrField: 'AUTHOR  Ada',
      sizeHalfPoints: 36,
    })
  })

  it('a resultless simple field takes the field code rPr', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: `<w:p>${field('PAGE', null, '<w:rPr><w:b/><w:sz w:val="28"/></w:rPr>')}</w:p>`,
      }),
    )
    expect(doc.blocks[0].runs?.[0]).toMatchObject({
      text: ' ',
      instrField: 'PAGE',
      bold: true,
      sizeHalfPoints: 28,
    })
  })

  it('an unknown field with a cached result shows the result runs in a cell', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: cellTable(
          `<w:p>${field('USERNAME', '<w:r><w:rPr><w:i/></w:rPr><w:t>Erika</w:t></w:r>')}</w:p>`,
        ),
      }),
    )
    const runs = doc.blocks[0].table!.rows[0][0].richParas?.[0]?.runs ?? []
    expect(runs).toHaveLength(1)
    expect(runs[0]).toMatchObject({ text: 'Erika', italic: true })
  })
})
