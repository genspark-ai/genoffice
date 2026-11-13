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

it('expands paragraph fo:margin and honors side-specific overrides', async () => {
  const styles =
    '<style:style style:name="P" style:family="paragraph"><style:paragraph-properties fo:margin="0.5in" fo:margin-left="1in"/></style:style>'
  const xml = await convert('<text:p text:style-name="P">content</text:p>', styles)
  expect(xml).toContain('w:before="720"')
  expect(xml).toContain('w:after="720"')
  expect(xml).toContain('w:left="1440"')
  expect(xml).toContain('w:right="720"')
})
