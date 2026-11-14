import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { convertOdpToPptx } from '../src/index'
async function convert(body: string, styles = '', part = 'ppt/slides/slide1.xml'): Promise<string> {
  const zip = new JSZip()
  zip.file(
    'content.xml',
    `<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"><office:automatic-styles>${styles}</office:automatic-styles><office:body><office:presentation><draw:page><draw:frame svg:x="1cm" svg:y="1cm" svg:width="10cm" svg:height="5cm"><draw:text-box>${body}</draw:text-box></draw:frame></draw:page></office:presentation></office:body></office:document-content>`,
  )
  const result = await convertOdpToPptx(await zip.generateAsync({ type: 'uint8array' }))
  return (await JSZip.loadAsync(result.pptx)).file(part)!.async('string')
}
it('emits one bullet per item while retaining continuation indentation', async () => {
  const xml = await convert(
    '<text:list><text:list-item><text:p>First</text:p><text:p>Continuation</text:p><text:list><text:list-item><text:p>Nested</text:p></text:list-item></text:list></text:list-item><text:list-item><text:p>Second</text:p></text:list-item></text:list>',
  )
  expect(xml.match(/<a:buChar /g)).toHaveLength(3)
  expect(xml).toContain('Continuation')
  expect(xml).toContain('<a:buNone/>')
})
