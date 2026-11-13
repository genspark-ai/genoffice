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

it('uses nested list-level-label-alignment indents', async () => {
  const styles =
    '<text:list-style style:name="L"><text:list-level-style-number text:level="1" style:num-format="1"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment fo:margin-left="1.25in" fo:text-indent="-0.3in" text:label-followed-by="listtab"/></style:list-level-properties></text:list-level-style-number></text:list-style>'
  const xml = await convert(
    '<text:list text:style-name="L"><text:list-item><text:p>item</text:p></text:list-item></text:list>',
    styles,
    'word/numbering.xml',
  )
  expect(xml).toContain('w:left="1800"')
  expect(xml).toContain('w:hanging="432"')
})
