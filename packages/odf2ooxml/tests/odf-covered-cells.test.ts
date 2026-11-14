import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { convertOdtToDocx } from '../src/index'

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

it.each(['horizontal', 'vertical'])('preserves cells following a %s merge', async (direction) => {
  const start = '<table:table><table:table-column table:number-columns-repeated="3"/>'
  const rows =
    direction === 'horizontal'
      ? '<table:table-row><table:table-cell table:number-columns-spanned="2"><text:p>Merged</text:p></table:table-cell><table:covered-table-cell/><table:table-cell><text:p>Right</text:p></table:table-cell></table:table-row>'
      : '<table:table-row><table:table-cell table:number-rows-spanned="2"><text:p>Merged</text:p></table:table-cell><table:table-cell><text:p>Top</text:p></table:table-cell></table:table-row><table:table-row><table:covered-table-cell/><table:table-cell><text:p>Below</text:p></table:table-cell><table:table-cell><text:p>Right</text:p></table:table-cell></table:table-row>'
  const xml = await convert(start + rows + '</table:table>')
  expect(xml).toContain('Right')
  if (direction === 'vertical') expect(xml).toContain('Below')
  expect(xml).toContain(direction === 'horizontal' ? '<w:gridSpan w:val="2"/>' : '<w:vMerge')
})
