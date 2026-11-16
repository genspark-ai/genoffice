import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { convertOdtToDocx } from '../src/index'

async function convert(body: string, styles = '', part = 'word/document.xml'): Promise<string> {
  const zip = new JSZip()
  zip.file('mimetype', 'application/vnd.oasis.opendocument.text')
  zip.file(
    'content.xml',
    `<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:automatic-styles>${styles}</office:automatic-styles><office:body><office:text>${body}</office:text></office:body></office:document-content>`,
  )
  const result = await convertOdtToDocx(await zip.generateAsync({ type: 'uint8array' }))
  const output = await JSZip.loadAsync(result.docx)
  return output.file(part)!.async('string')
}

it('lets a child style turn off inherited bold and italic', async () => {
  const styles =
    '<style:style style:name="Base" style:family="text"><style:text-properties fo:font-weight="bold" fo:font-style="italic"/></style:style><style:style style:name="Plain" style:family="text" style:parent-style-name="Base"><style:text-properties fo:font-weight="normal" fo:font-style="normal"/></style:style>'
  const xml = await convert(
    '<text:p><text:span text:style-name="Plain">plain</text:span></text:p>',
    styles,
  )
  expect(xml).not.toContain('<w:b/>')
  expect(xml).not.toContain('<w:i/>')
  const inherited = await convert(
    '<text:p><text:span text:style-name="Base">bold</text:span></text:p>',
    styles,
  )
  expect(inherited).toContain('<w:b/>')
  expect(inherited).toContain('<w:i/>')
})
