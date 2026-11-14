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
it.each([
  ['150%', '360'],
  ['normal', '240'],
])('replaces inherited fixed line height with %s', async (height, line) => {
  const styles = `<style:style style:name="Base" style:family="paragraph"><style:paragraph-properties fo:line-height="24pt"/></style:style><style:style style:name="Child" style:family="paragraph" style:parent-style-name="Base"><style:paragraph-properties fo:line-height="${height}"/></style:style>`
  const xml = await convert('<text:p text:style-name="Child">Lines</text:p>', styles)
  expect(xml).toContain('w:line="' + line + '"')
  expect(xml).toContain('w:lineRule="auto"')
  expect(xml).not.toContain('w:lineRule="exact"')
})
it('normal disables an explicit minimum height and extra spacing', async () => {
  const styles =
    '<style:style style:name="P" style:family="paragraph"><style:paragraph-properties fo:line-height="normal" style:line-height-at-least="30pt" style:line-spacing="6pt"/></style:style>'
  const xml = await convert('<text:p text:style-name="P">Normal</text:p>', styles)
  expect(xml).toContain('w:line="240"')
  expect(xml).toContain('w:lineRule="auto"')
})
it('allows a fixed child height to replace a percentage parent', async () => {
  const styles =
    '<style:style style:name="Base" style:family="paragraph"><style:paragraph-properties fo:line-height="200%"/></style:style><style:style style:name="Child" style:family="paragraph" style:parent-style-name="Base"><style:paragraph-properties fo:line-height="18pt"/></style:style>'
  const xml = await convert('<text:p text:style-name="Child">Fixed</text:p>', styles)
  expect(xml).toContain('w:line="360"')
  expect(xml).toContain('w:lineRule="exact"')
})
