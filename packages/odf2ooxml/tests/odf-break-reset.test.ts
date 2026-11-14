import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { convertOdtToDocx } from '../src/index'
const ns =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"'
async function documentXml(body: string, styles = '', extra?: string): Promise<string> {
  const z = new JSZip()
  z.file(
    'content.xml',
    `<office:document-content ${ns}><office:automatic-styles>${styles}</office:automatic-styles><office:body><office:text>${body}</office:text></office:body></office:document-content>`,
  )
  if (extra !== undefined) z.file('styles.xml', extra)
  const result = await convertOdtToDocx(await z.generateAsync({ type: 'uint8array' }))
  return (await JSZip.loadAsync(result.docx)).file('word/document.xml')!.async('string')
}
it('allows a paragraph style to reset an inherited page break', async () => {
  const styles =
    '<style:style style:name="Base" style:family="paragraph"><style:paragraph-properties fo:break-before="page"/></style:style><style:style style:name="Child" style:family="paragraph" style:parent-style-name="Base"><style:paragraph-properties fo:break-before="auto"/></style:style>'
  expect(
    await documentXml('<text:p text:style-name="Child">same page</text:p>', styles),
  ).not.toContain('<w:pageBreakBefore')
  expect(await documentXml('<text:p text:style-name="Base">new page</text:p>', styles)).toContain(
    '<w:pageBreakBefore',
  )
})
