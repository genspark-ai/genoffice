import { describe, expect, it } from 'vitest'
import { mergePPrFormat, parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const W15 = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'
const style = (id: string, pPr: string, basedOn = 'Normal') =>
  `<w:style w:type="paragraph" w:styleId="${id}" ${W15}><w:name w:val="${id}"/>` +
  `<w:basedOn w:val="${basedOn}"/><w:pPr>${pPr}</w:pPr></w:style>`

async function parse(bodyXml: string, extraStylesXml?: string) {
  return parseDocx(await buildDocx({ bodyXml, extraStylesXml, docRootExtraAttrs: W15 }))
}
const p = (pPr: string) => `<w:p><w:pPr>${pPr}</w:pPr><w:r><w:t>x</w:t></w:r></w:p>`

describe('w15:collapsed ("Collapsed by default")', () => {
  it('reads the flag off the paragraph, w:val="0" included', async () => {
    const { blocks } = await parse(
      p('<w:pStyle w:val="Heading1"/><w15:collapsed/>') +
        p('<w:pStyle w:val="Heading1"/><w15:collapsed w:val="0"/>') +
        p('<w:pStyle w:val="Heading1"/>'),
    )
    expect(blocks.slice(0, 3).map((b) => b.format?.collapsed)).toEqual([true, false, undefined])
  })

  it('reads it off a style and through basedOn, a nearer w:val="0" winning', async () => {
    const { styles } = await parse(
      p('<w:pStyle w:val="Fold"/>'),
      style('Fold', '<w:outlineLvl w:val="0"/><w15:collapsed/>') +
        style('FoldChild', '', 'Fold') +
        style('FoldOff', '<w15:collapsed w:val="0"/>', 'Fold'),
    )
    expect(styles.get('Fold')?.collapsed).toBe(true)
    expect(styles.get('FoldChild')?.collapsed).toBe(true)
    expect(styles.get('FoldOff')?.collapsed).toBe(false)
    expect(styles.get('Heading1')?.collapsed).toBeUndefined()
  })

  it('writes the flag after the outline level and before the mark properties', () => {
    const raw =
      '<w:pPr><w:pStyle w:val="Heading1"/><w:outlineLvl w:val="0"/><w:rPr><w:b/></w:rPr></w:pPr>'
    expect(mergePPrFormat(raw, { collapsed: true })).toBe(
      '<w:pPr><w:pStyle w:val="Heading1"/><w:outlineLvl w:val="0"/><w15:collapsed/><w:rPr><w:b/></w:rPr></w:pPr>',
    )
    expect(mergePPrFormat(raw, { collapsed: false })).toContain('<w15:collapsed w:val="0"/>')
  })

  it('keeps an untouched flag byte for byte and drops a cleared one', () => {
    const raw = '<w:pPr><w:pStyle w:val="Heading1"/><w15:collapsed w:val="1"/></w:pPr>'
    expect(mergePPrFormat(raw, { collapsed: true }, { collapsed: true })).toBe(raw)
    expect(mergePPrFormat(raw, {}, { collapsed: true })).toBe(
      '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>',
    )
  })
})
