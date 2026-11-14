import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { convertOdtToDocx, convertOdpToPptx } from '../src/index'
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
async function slideXml(
  body: string,
  box = 'svg:x="1cm" svg:y="1cm" svg:width="10cm" svg:height="2cm"',
): Promise<string> {
  const z = new JSZip()
  z.file(
    'content.xml',
    `<office:document-content ${ns}><office:body><office:presentation><draw:page><draw:frame ${box}><draw:text-box>${body}</draw:text-box></draw:frame></draw:page></office:presentation></office:body></office:document-content>`,
  )
  const result = await convertOdpToPptx(await z.generateAsync({ type: 'uint8array' }))
  return (await JSZip.loadAsync(result.pptx)).file('ppt/slides/slide1.xml')!.async('string')
}
it('retains an unnumbered list header in presentation text', async () => {
  const xml = await slideXml(
    '<text:list><text:list-header><text:p>Introduction</text:p></text:list-header><text:list-item><text:p>Item</text:p></text:list-item></text:list>',
  )
  expect(xml).toContain('Introduction')
  expect(xml).toContain('Item')
})
it('retains a list header in a table cell', async () => {
  const xml = await documentXml(
    '<table:table><table:table-column/><table:table-row><table:table-cell><text:list><text:list-header><text:p>Introduction</text:p></text:list-header><text:list-item><text:p>Item</text:p></text:list-item></text:list></table:table-cell></table:table-row></table:table>',
  )
  expect(xml).toContain('Introduction')
})
