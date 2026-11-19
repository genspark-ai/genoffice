import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { xlsxToText } from '../src/xlsx'

function workbook(sharedStrings: string, sheetXml: string): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file(
    'xl/workbook.xml',
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/sharedStrings.xml',
    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      sharedStrings +
      '</sst>',
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      `<sheetData>${sheetXml}</sheetData></worksheet>`,
  )
  return zip.generateAsync({ type: 'uint8array' })
}

const sharedCell = (col: string, index: number): string =>
  `<c r="${col}1" t="s"><v>${index}</v></c>`

describe('xlsxToText rich shared strings', () => {
  it('concatenates rich runs and a trailing plain node in document order', async () => {
    const text = await xlsxToText(
      await workbook(
        '<si><r><rPr><b/></rPr><t>Bold</t></r><t> plain</t></si>',
        `<row r="1">${sharedCell('A', 0)}</row>`,
      ),
    )
    expect(text).toContain('Bold plain')
  })

  it('keeps runs and plain nodes in order when the plain node comes first', async () => {
    const text = await xlsxToText(
      await workbook(
        '<si><t>plain </t><r><t>rich</t></r></si>',
        `<row r="1">${sharedCell('A', 0)}</row>`,
      ),
    )
    expect(text).toContain('plain rich')
  })

  it('still reads plain-only, rich-only and multi-run entries', async () => {
    const text = await xlsxToText(
      await workbook(
        '<si><t>plain</t></si>' +
          '<si><r><t>a</t></r><r><t>b</t></r></si>' +
          '<si><r><rPr><i/></rPr><t>x</t></r><r><t>y</t></r></si>',
        `<row r="1">${sharedCell('A', 0)}${sharedCell('B', 1)}${sharedCell('C', 2)}</row>`,
      ),
    )
    expect(text).toContain('plain | ab | xy')
  })

  it('skips phonetic runs', async () => {
    const text = await xlsxToText(
      await workbook(
        '<si><r><t>kan</t><rPh sb="0" eb="2"><t>ka</t></rPh></r><t>ji</t></si>',
        `<row r="1">${sharedCell('A', 0)}</row>`,
      ),
    )
    expect(text).toContain('kanji')
  })

  it('reads a mixed rich/plain inline string', async () => {
    const text = await xlsxToText(
      await workbook(
        '',
        '<row r="1"><c r="A1" t="inlineStr"><is><r><t>In</t></r><t>line</t></is></c></row>',
      ),
    )
    expect(text).toContain('Inline')
  })
})
