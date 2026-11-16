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
it('keeps visible nested section contents in table cells', async () => {
  const xml = await convert(
    '<table:table><table:table-column/><table:table-row><table:table-cell><text:p>Before</text:p><text:section><text:section><text:h text:outline-level="2">Section heading</text:h><text:p>Inside</text:p></text:section><text:section text:display="none"><text:p>Hidden</text:p></text:section></text:section><text:p>After</text:p></table:table-cell></table:table-row></table:table>',
  )
  expect(xml).toContain('Section heading')
  expect(xml).toContain('Inside')
  expect(xml).not.toContain('Hidden')
  expect(xml.indexOf('Before')).toBeLessThan(xml.indexOf('Inside'))
  expect(xml.indexOf('Inside')).toBeLessThan(xml.indexOf('After'))
})
