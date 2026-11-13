import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { xlsxToText } from '../src/xlsx'
it('does not resolve external worksheet relationships to same-named local parts', async () => {
  const z = new JSZip()
  z.file(
    'xl/workbook.xml',
    '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Local" sheetId="1" r:id="local"/><sheet name="Remote" sheetId="2" r:id="remote"/></sheets></workbook>',
  )
  z.file(
    'xl/_rels/workbook.xml.rels',
    '<Relationships><Relationship Id="local" Target="worksheets/a.xml"/><Relationship Id="remote" Target="worksheets/b.xml" TargetMode="External"/></Relationships>',
  )
  z.file(
    'xl/worksheets/a.xml',
    '<worksheet><sheetData><row><c t="inlineStr"><is><t>Local value</t></is></c></row></sheetData></worksheet>',
  )
  z.file(
    'xl/worksheets/b.xml',
    '<worksheet><sheetData><row><c t="inlineStr"><is><t>Unrelated local value</t></is></c></row></sheetData></worksheet>',
  )
  expect(await xlsxToText(await z.generateAsync({ type: 'uint8array' }))).toBe(
    '# Local\nLocal value',
  )
})
