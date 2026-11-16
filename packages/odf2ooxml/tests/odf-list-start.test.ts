import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { convertOdtToDocx } from '../src/index'
async function convert(body: string, styles = '', part = 'word/document.xml'): Promise<string> {
  const zip = new JSZip()
  zip.file(
    'content.xml',
    `<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"><office:automatic-styles>${styles}</office:automatic-styles><office:body><office:text>${body}</office:text></office:body></office:document-content>`,
  )
  const result = await convertOdtToDocx(await zip.generateAsync({ type: 'uint8array' }))
  return (await JSZip.loadAsync(result.docx)).file(part)!.async('string')
}
it.each([0, 7])('preserves a list style start value of %s', async (start) => {
  const styles = `<text:list-style style:name="L"><text:list-level-style-number text:level="1" style:num-format="1" text:start-value="${start}"/></text:list-style>`
  const xml = await convert(
    '<text:list text:style-name="L"><text:list-item><text:p>Start</text:p></text:list-item></text:list>',
    styles,
    'word/numbering.xml',
  )
  expect(xml).toContain('<w:start w:val="' + start + '"/>')
})
it.each(['-1', '1.5', 'invalid'])('keeps the default start for invalid %s', async (start) => {
  const styles = `<text:list-style style:name="L"><text:list-level-style-number text:level="1" style:num-format="1" text:start-value="${start}"/></text:list-style>`
  const xml = await convert(
    '<text:list text:style-name="L"><text:list-item><text:p>Default</text:p></text:list-item></text:list>',
    styles,
    'word/numbering.xml',
  )
  expect(xml).not.toContain('<w:start w:val="' + start + '"/>')
  expect(xml).toContain('<w:start w:val="1"/>')
})
