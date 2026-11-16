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
it.each([false, true])(
  'retains list header headings and item headings inside a cell=%s',
  async (cell) => {
    const list =
      '<text:list><text:list-header><text:h text:outline-level="2">Header</text:h></text:list-header><text:list-item><text:h text:outline-level="3">Item</text:h></text:list-item></text:list>'
    const body = cell
      ? '<table:table><table:table-column/><table:table-row><table:table-cell>' +
        list +
        '</table:table-cell></table:table-row></table:table>'
      : list
    const xml = await convert(body)
    expect(xml).toContain('Header')
    expect(xml).toContain('Item')
    expect(xml).toContain('w:pStyle w:val="Heading2"')
    expect(xml).toContain('w:pStyle w:val="Heading3"')
  },
)
