import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import type { ChangePlan } from '../src/domain/workbook.types'
import { applyPlanToXlsx } from '../src/gateway/xlsx-gateway'

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`

const PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`

const WORKBOOK = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`

const WORKBOOK_RELS = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`

const STYLES = `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf/><xf/></cellXfs>
</styleSheet>`

const WORKSHEET = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"/><c r="B1"><v>5</v></c></row></sheetData>
</worksheet>`

async function fixture(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', CONTENT_TYPES)
  zip.file('_rels/.rels', PACKAGE_RELS)
  zip.file('xl/workbook.xml', WORKBOOK)
  zip.file('xl/_rels/workbook.xml.rels', WORKBOOK_RELS)
  zip.file('xl/styles.xml', STYLES)
  zip.file('xl/worksheets/sheet1.xml', WORKSHEET)
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

function planWith(value: number): ChangePlan {
  return {
    transactionId: 't1',
    baseRevision: 0,
    sheetRenames: [],
    structuralChanges: [],
    formatChanges: [],
    warnings: [],
    cellChanges: [{ sheetId: '1', address: 'A1', before: { value: null }, after: { value } }],
  }
}

describe('non-finite cell values', () => {
  it('never writes NaN or Infinity into the numeric <v> of a cell', async () => {
    for (const value of [Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY]) {
      const mutation = await applyPlanToXlsx(await fixture(), planWith(value), { '1': 'Data' })
      const zip = await JSZip.loadAsync(mutation.buffer)
      const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')

      // The literal NaN/Infinity must not reach a numeric <v>.
      expect(sheet).not.toMatch(/<v>NaN<\/v>/)
      expect(sheet).not.toMatch(/<v>Infinity<\/v>/)
      expect(sheet).not.toMatch(/<v>-Infinity<\/v>/)
      // The cell keeps a representation, and the sibling is untouched.
      expect(sheet).toContain('<c r="B1"><v>5</v></c>')
    }
  })

  it('writes a finite number as a plain numeric cell', async () => {
    const mutation = await applyPlanToXlsx(await fixture(), planWith(42.5), { '1': 'Data' })
    const zip = await JSZip.loadAsync(mutation.buffer)
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheet).toContain('<c r="A1"><v>42.5</v></c>')
  })
})
