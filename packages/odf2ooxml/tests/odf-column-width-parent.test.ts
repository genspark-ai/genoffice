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
it('inherits column widths while keeping explicit child widths', async () => {
  const styles =
    '<style:style style:name="Base" style:family="table-column"><style:table-column-properties style:column-width="1in"/></style:style><style:style style:name="Child" style:family="table-column" style:parent-style-name="Base"/><style:style style:name="Override" style:family="table-column" style:parent-style-name="Base"><style:table-column-properties style:column-width="2in"/></style:style>'
  const xml = await convert(
    '<table:table><table:table-column table:style-name="Child"/><table:table-column table:style-name="Override"/><table:table-row><table:table-cell><text:p>A</text:p></table:table-cell><table:table-cell><text:p>B</text:p></table:table-cell></table:table-row></table:table>',
    styles,
  )
  expect(xml).toContain('<w:gridCol w:w="1440"/><w:gridCol w:w="2880"/>')
})
it('terminates cyclic column inheritance and keeps the table readable', async () => {
  const styles =
    '<style:style style:name="A" style:family="table-column" style:parent-style-name="B"/><style:style style:name="B" style:family="table-column" style:parent-style-name="A"/>'
  expect(
    await convert(
      '<table:table><table:table-column table:style-name="A"/><table:table-row><table:table-cell><text:p>Cycle</text:p></table:table-cell></table:table-row></table:table>',
      styles,
    ),
  ).toContain('Cycle')
})
