import { describe, expect, it } from 'vitest'
import { mergePPrFormat, parseDocx, type ParaFormat } from '../src/index'
import { buildDocx } from './helpers/build-docx'

// pPr full of attributes the ParaFormat model does not capture:
// CJK char-unit indents/spacing, autospacing, custom border color/style, shading pattern
const RAW =
  '<w:pPr>' +
  '<w:pBdr><w:top w:val="dashed" w:sz="12" w:space="1" w:color="FF0000"/>' +
  '<w:bottom w:val="dashed" w:sz="12" w:space="1" w:color="FF0000"/></w:pBdr>' +
  '<w:shd w:val="pct10" w:color="auto" w:fill="EEEEEE"/>' +
  '<w:spacing w:beforeLines="100" w:before="240" w:afterLines="50" w:after="120"/>' +
  '<w:ind w:leftChars="0" w:left="0" w:firstLineChars="200" w:firstLine="420"/>' +
  '<w:jc w:val="both"/>' +
  '</w:pPr>'

// what extractParaFormat yields for RAW
const MODEL: ParaFormat = {
  borders: 'tb',
  shadingFill: 'EEEEEE',
  spaceBefore: 240,
  spaceAfter: 120,
  indentFirstLine: 420,
  align: 'justify',
}

describe('mergePPrFormat keeps unedited groups byte-identical', () => {
  it('unchanged model -> byte-identical pPr', () => {
    expect(mergePPrFormat(RAW, MODEL)).toBe(RAW)
  })

  it('model straight from parseDocx round-trips the raw bytes', async () => {
    const bytes = await buildDocx({ bodyXml: `<w:p>${RAW}<w:r><w:t>正文</w:t></w:r></w:p>` })
    const doc = await parseDocx(bytes)
    expect(mergePPrFormat(doc.blocks[0].rawPPr!, doc.blocks[0].format)).toBe(RAW)
  })

  it('changing only the alignment keeps firstLineChars/afterLines/custom borders', () => {
    const out = mergePPrFormat(RAW, { ...MODEL, align: 'center' })
    expect(out).toContain('<w:jc w:val="center"/>')
    expect(out).not.toContain('w:val="both"')
    expect(out).toContain('w:firstLineChars="200"')
    expect(out).toContain('w:afterLines="50"')
    expect(out).toContain('w:beforeLines="100"')
    expect(out).toContain('<w:top w:val="dashed" w:sz="12" w:space="1" w:color="FF0000"/>')
    expect(out).toContain('<w:shd w:val="pct10" w:color="auto" w:fill="EEEEEE"/>')
  })

  it('changing spacing rebuilds only w:spacing; w:ind bytes survive', () => {
    const out = mergePPrFormat(RAW, { ...MODEL, spaceAfter: 240 })
    expect(out).toContain('<w:spacing w:before="240" w:after="240"/>')
    expect(out).not.toContain('w:afterLines')
    expect(out).toContain('w:firstLineChars="200"')
  })

  it('nil border resets do not force a pBdr rebuild', async () => {
    // parse skips w:val="nil" sides, so the raw comparison must skip them too —
    // otherwise every save rebuilds w:pBdr and drops the bottom border's color/size
    const raw =
      '<w:pPr><w:pBdr><w:top w:val="nil"/><w:left w:val="nil"/><w:right w:val="nil"/>' +
      '<w:bottom w:val="single" w:sz="18" w:space="1" w:color="4472C4"/></w:pBdr></w:pPr>'
    const bytes = await buildDocx({ bodyXml: `<w:p>${raw}<w:r><w:t>x</w:t></w:r></w:p>` })
    const doc = await parseDocx(bytes)
    expect(doc.blocks[0].format?.borders).toBe('b')
    expect(mergePPrFormat(raw, doc.blocks[0].format)).toBe(raw)
  })

  it('changing indent rebuilds w:ind and drops the char-unit variants', () => {
    // Word prefers *Chars over the twips attrs, so a stale firstLineChars would
    // override the user's new indent — the rebuilt w:ind must not carry them
    const out = mergePPrFormat(RAW, { ...MODEL, indentFirstLine: 640 })
    expect(out).toContain('<w:ind w:firstLine="640"/>')
    expect(out).not.toContain('firstLineChars')
    expect(out).not.toContain('leftChars')
    expect(out).toContain('w:afterLines="50"')
  })
})

describe('explicit w:after="0"', () => {
  it('parses to spaceAfter 0 and keeps the bytes when unchanged', async () => {
    const raw =
      '<w:pPr><w:spacing w:after="0" w:afterAutospacing="0" w:line="276" w:lineRule="auto"/></w:pPr>'
    const bytes = await buildDocx({ bodyXml: `<w:p>${raw}<w:r><w:t>x</w:t></w:r></w:p>` })
    const doc = await parseDocx(bytes)
    expect(doc.blocks[0].format?.spaceAfter).toBe(0)
    expect(mergePPrFormat(raw, doc.blocks[0].format)).toBe(raw)
  })

  it('writes w:after="0" back when the spacing group is rebuilt', () => {
    const out = mergePPrFormat('<w:pPr><w:spacing w:before="240" w:after="0"/></w:pPr>', {
      spaceAfter: 0,
    })
    expect(out).toBe('<w:pPr><w:spacing w:after="0"/></w:pPr>')
  })
})
