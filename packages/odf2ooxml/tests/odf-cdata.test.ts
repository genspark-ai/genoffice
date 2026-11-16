import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { convertOdpToPptx, convertOdtToDocx } from '../src/index'

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

it('keeps CDATA interleaved with ordinary text', async () => {
  const xml = await convert('<text:p>before <![CDATA[<data>& value]]> after</text:p>')
  expect(xml).toContain('before &lt;data&gt;&amp; value after')
})

it('keeps CDATA in presentation text boxes', async () => {
  const zip = new JSZip()
  zip.file('mimetype', 'application/vnd.oasis.opendocument.presentation')
  zip.file(
    'content.xml',
    '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"><office:body><office:presentation><draw:page draw:name="Page"><draw:frame svg:x="1cm" svg:y="1cm" svg:width="10cm" svg:height="2cm"><draw:text-box><text:p>before <![CDATA[<data>& value]]> after</text:p></draw:text-box></draw:frame></draw:page></office:presentation></office:body></office:document-content>',
  )
  const result = await convertOdpToPptx(await zip.generateAsync({ type: 'uint8array' }))
  const output = await JSZip.loadAsync(result.pptx)
  expect(await output.file('ppt/slides/slide1.xml')!.async('string')).toContain(
    'before &lt;data&gt;&amp; value after',
  )
})
