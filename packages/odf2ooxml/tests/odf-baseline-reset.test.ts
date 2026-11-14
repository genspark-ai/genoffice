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
function style(name: string, props: string, parent = '', family = 'text'): string {
  return `<style:style style:name="${name}" style:family="${family}"${parent ? ` style:parent-style-name="${parent}"` : ''}><style:text-properties ${props}/></style:style>`
}
it.each(['super', 'sub'])('restores the baseline after %s', async (position) => {
  const styles =
    style('Base', `style:text-position="${position} 58%"`) +
    style('Child', 'style:text-position="0% 100%"', 'Base')
  expect(
    await documentXml(
      '<text:p><text:span text:style-name="Child">baseline</text:span></text:p>',
      styles,
    ),
  ).not.toContain('<w:vertAlign')
  expect(
    await documentXml(
      '<text:p><text:span text:style-name="Base">offset</text:span></text:p>',
      styles,
    ),
  ).toContain('<w:vertAlign')
})
