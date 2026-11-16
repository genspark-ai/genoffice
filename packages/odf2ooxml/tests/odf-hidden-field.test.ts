import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { convertOdtToDocx, convertOdpToPptx } from '../src/index'
async function convert(body: string, styles = '', part = 'word/document.xml'): Promise<string> {
  const zip = new JSZip()
  zip.file(
    'content.xml',
    `<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"><office:automatic-styles>${styles}</office:automatic-styles><office:body><office:text>${body}</office:text></office:body></office:document-content>`,
  )
  const result = await convertOdtToDocx(await zip.generateAsync({ type: 'uint8array' }))
  return (await JSZip.loadAsync(result.docx)).file(part)!.async('string')
}
it.each(['false', '0'])('retains a visible hidden-text field with is-hidden=%s', async (flag) => {
  const xml = await convert(
    `<text:p>Before<text:hidden-text text:is-hidden="${flag}" text:string-value="Visible">Visible</text:hidden-text>After<text:hidden-text text:is-hidden="true" text:condition="ooow:1==1" text:string-value="Secret">Secret</text:hidden-text></text:p>`,
  )
  expect(xml).toContain('BeforeVisibleAfter')
  expect(xml).not.toContain('Secret')
})
it.each(['true', '1'])(
  'omits a hidden field in presentation text with is-hidden=%s',
  async (flag) => {
    const zip = new JSZip()
    zip.file(
      'content.xml',
      `<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"><office:body><office:presentation><draw:page><draw:frame svg:x="1cm" svg:y="1cm" svg:width="10cm" svg:height="5cm"><draw:text-box><text:p>Visible<text:hidden-text text:is-hidden="${flag}" text:condition="ooow:1==1" text:string-value="Secret">Secret</text:hidden-text></text:p></draw:text-box></draw:frame></draw:page></office:presentation></office:body></office:document-content>`,
    )
    const result = await convertOdpToPptx(await zip.generateAsync({ type: 'uint8array' }))
    const xml = await (
      await JSZip.loadAsync(result.pptx)
    )
      .file('ppt/slides/slide1.xml')!
      .async('string')
    expect(xml).toContain('Visible')
    expect(xml).not.toContain('Secret')
  },
)
