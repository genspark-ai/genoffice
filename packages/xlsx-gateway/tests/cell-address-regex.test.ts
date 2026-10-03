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

// A1 is styled but empty and self-closing; B1 is its sibling and holds 5.
const WORKSHEET = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" s="1"/><c r="B1"><v>5</v></c></row></sheetData>
</worksheet>`

async function fixture(worksheet = WORKSHEET): Promise<Buffer> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', CONTENT_TYPES)
  zip.file('_rels/.rels', PACKAGE_RELS)
  zip.file('xl/workbook.xml', WORKBOOK)
  zip.file('xl/_rels/workbook.xml.rels', WORKBOOK_RELS)
  zip.file('xl/styles.xml', STYLES)
  zip.file('xl/worksheets/sheet1.xml', worksheet)
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

// The address is unconstrained on ChangePlan, so the plan can carry anything.
// applyPlanToXlsx takes a ChangePlan directly, which is why the writer itself
// has to reject a non-literal reference before it becomes regex source.
function plan(address: string): ChangePlan {
  return {
    transactionId: 't1',
    baseRevision: 0,
    sheetRenames: [],
    structuralChanges: [],
    formatChanges: [],
    warnings: [],
    cellChanges: [{ sheetId: '1', address, before: { value: null }, after: { value: 1 } }],
  }
}

async function worksheetOf(source: Buffer): Promise<string> {
  const zip = await JSZip.loadAsync(source)
  return zip.file('xl/worksheets/sheet1.xml')!.async('string')
}

describe('cell address validation before regex interpolation', () => {
  it('refuses a plan whose cell address is not a literal A1 reference', async () => {
    // A trailing quote closes the r="..." attribute the patch pattern is
    // built around. Interpolated raw, `r="A1""` is a pattern that matches
    // something other than A1, so the plan has to die before any regex is
    // compiled from it.
    const source = await fixture()

    await expect(applyPlanToXlsx(source, plan('A1"'), { '1': 'Data' })).rejects.toThrow(
      /Invalid cell address in plan/,
    )

    // Fail closed: the workbook keeps both of its original cells, so no
    // sibling was swallowed by a mis-patched match.
    const sheet = await worksheetOf(source)
    expect(sheet).toContain('<c r="A1" s="1"/>')
    expect(sheet).toContain('<c r="B1"><v>5</v></c>')
  })

  it('rejects an out-of-grid address such as XFE1', async () => {
    // XFE is column 16385, one past the last writable column. It is a
    // well-formed A1 reference, so only the grid caps reject it — and the
    // row-append fallback would otherwise have invented the cell.
    await expect(applyPlanToXlsx(await fixture(), plan('XFE1'), { '1': 'Data' })).rejects.toThrow(
      /Invalid cell address in plan/,
    )
  })
})
