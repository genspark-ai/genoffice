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
it('inherits table cell fill and respects an explicit transparent override', async () => {
  const styles =
    '<style:style style:name="Base" style:family="table-cell"><style:table-cell-properties fo:background-color="#336699"/></style:style><style:style style:name="Child" style:family="table-cell" style:parent-style-name="Base"/><style:style style:name="Clear" style:family="table-cell" style:parent-style-name="Base"><style:table-cell-properties fo:background-color="transparent"/></style:style>'
  const body = (name: string) =>
    `<table:table><table:table-column/><table:table-row><table:table-cell table:style-name="${name}"><text:p>Value</text:p></table:table-cell></table:table-row></table:table>`
  expect(await convert(body('Child'), styles)).toContain('w:fill="336699"')
  expect(await convert(body('Clear'), styles)).not.toContain('w:fill="336699"')
})
it('terminates cyclic cell style inheritance', async () => {
  const styles =
    '<style:style style:name="A" style:family="table-cell" style:parent-style-name="B"/><style:style style:name="B" style:family="table-cell" style:parent-style-name="A"/>'
  expect(
    await convert(
      '<table:table><table:table-column/><table:table-row><table:table-cell table:style-name="A"><text:p>Cycle</text:p></table:table-cell></table:table-row></table:table>',
      styles,
    ),
  ).toContain('Cycle')
})
