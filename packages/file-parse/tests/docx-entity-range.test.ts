import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { docxToText } from '../src/docx'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

async function docx(body: string): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file(
    '_rels/.rels',
    `<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/word/document.xml"/></Relationships>`,
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0"?><w:document ${W}><w:body>${body}</w:body></w:document>`,
  )
  zip.file('word/styles.xml', `<w:styles ${W}></w:styles>`)
  return zip.generateAsync({ type: 'uint8array' })
}

describe('docxToText numeric character references', () => {
  it('skips a code point beyond the Unicode range instead of throwing', async () => {
    // &#x110000; and beyond are out of Unicode range: String.fromCodePoint would
    // throw a RangeError and fail the whole extraction on one hostile fragment
    const text = await docxToText(
      await docx('<w:p><w:r><w:t>a&#x110000;b&#xFFFFFFFF;c</w:t></w:r></w:p>'),
    )
    expect(text).toContain('abc')
  })

  it('still decodes in-range numeric and named entities', async () => {
    const text = await docxToText(
      await docx('<w:p><w:r><w:t>&#65;&#x4e2d;&amp;&quot;</w:t></w:r></w:p>'),
    )
    expect(text).toContain('A中&"')
  })
})
