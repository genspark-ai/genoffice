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

it('uses ODF table-column-properties for the exported table grid', async () => {
  const styles =
    '<style:style style:name="Narrow" style:family="table-column"><style:table-column-properties style:column-width="1in"/></style:style><style:style style:name="Wide" style:family="table-column"><style:table-column-properties style:column-width="2in"/></style:style>'
  const xml = await convert(
    '<table:table><table:table-column table:style-name="Narrow"/><table:table-column table:style-name="Wide"/><table:table-row><table:table-cell><text:p>A</text:p></table:table-cell><table:table-cell><text:p>B</text:p></table:table-cell></table:table-row></table:table>',
    styles,
  )
  expect(xml).toContain('<w:gridCol w:w="1440"/>')
  expect(xml).toContain('<w:gridCol w:w="2880"/>')
})
