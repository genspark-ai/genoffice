import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { paraMarkColor, parseDocx, saveDocx, stripParaMarkColor } from '../src'
import { buildDocx } from './helpers/build-docx'

describe('paragraph-mark color helpers', () => {
  const pPr =
    '<w:pPr><w:jc w:val="center"/><w:rPr><w:b/><w:color w:val="FF0000"/><w:sz w:val="21"/></w:rPr></w:pPr>'

  it('reads the mark color and ignores the pPrChange history', () => {
    expect(paraMarkColor(pPr)).toBe('FF0000')
    expect(paraMarkColor('<w:pPr><w:rPr><w:color w:val="auto"/></w:rPr></w:pPr>')).toBeUndefined()
    expect(
      paraMarkColor(
        '<w:pPr><w:pPrChange w:id="1" w:author="a"><w:pPr><w:rPr><w:color w:val="00FF00"/></w:rPr></w:pPr></w:pPrChange></w:pPr>',
      ),
    ).toBeUndefined()
  })

  it('reads a hash-prefixed value and the paired tag form like the run parser does', () => {
    expect(paraMarkColor('<w:pPr><w:rPr><w:color w:val="#ff0000"/></w:rPr></w:pPr>')).toBe('FF0000')
    expect(paraMarkColor('<w:pPr><w:rPr><w:color w:val="FF0000"></w:color></w:rPr></w:pPr>')).toBe(
      'FF0000',
    )
    expect(
      stripParaMarkColor('<w:pPr><w:rPr><w:b/><w:color w:val="FF0000"></w:color></w:rPr></w:pPr>'),
    ).toBe('<w:pPr><w:rPr><w:b/></w:rPr></w:pPr>')
  })

  it('strips only the color, keeps the rest of the mark rPr and the pPrChange bytes', () => {
    expect(stripParaMarkColor(pPr)).toBe(
      '<w:pPr><w:jc w:val="center"/><w:rPr><w:b/><w:sz w:val="21"/></w:rPr></w:pPr>',
    )
    const change =
      '<w:pPrChange w:id="1" w:author="a"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange>'
    expect(
      stripParaMarkColor(`<w:pPr><w:rPr><w:color w:val="FF0000"/></w:rPr>${change}</w:pPr>`),
    ).toBe(`<w:pPr>${change}</w:pPr>`)
  })
})

describe('NUMPAGES cached result', () => {
  it('writes the page count given at save time, defaulting to 1', async () => {
    const bytes = await buildDocx({ bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>' })
    const footer = {
      text: 'Page {PAGE} of {NUMPAGES}',
      pageNumber: false,
      paras: [{ runs: [{ text: 'Page  of ' }] }],
    }
    const read = async (saved: Uint8Array) => {
      const zip = await JSZip.loadAsync(saved)
      const name = Object.keys(zip.files).find((n) => /^word\/footer\d*\.xml$/.test(n))!
      return zip.file(name)!.async('string')
    }
    const parsed = await parseDocx(bytes)
    const withCount = await read(await saveDocx(parsed, [], { footer, pageCount: 12 }))
    expect(withCount).toContain(
      '<w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>12</w:t></w:r>',
    )
    const plain = await read(await saveDocx(await parseDocx(bytes), [], { footer }))
    expect(plain).toContain(
      'NUMPAGES </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r>',
    )
  })
})
