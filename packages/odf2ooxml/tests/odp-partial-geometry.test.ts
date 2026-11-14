import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { convertOdpToPptx } from '../src/index'
const ns =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"'
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
it('preserves declared dimensions when only frame origin is missing', async () => {
  const xml = await slideXml(
    '<text:p>Title</text:p>',
    'presentation:class="title" svg:width="10cm" svg:height="2cm"',
  )
  expect(xml).toContain('<a:ext cx="3600000" cy="720000"/>')
  expect(xml).toContain('<a:off x="975360" y="274320"/>')
})
