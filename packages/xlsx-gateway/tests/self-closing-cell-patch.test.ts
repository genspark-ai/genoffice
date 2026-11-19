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

// Row 1 is an empty custom-height row Excel writes self-closing; row 2 holds A2.
const WORKSHEET_EMPTY_ROW = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1" ht="20" customHeight="1"/><row r="2"><c r="A2"><v>7</v></c></row></sheetData>
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

function plan(address = 'A1'): ChangePlan {
  return {
    transactionId: 't1',
    baseRevision: 0,
    sheetRenames: [],
    structuralChanges: [],
    formatChanges: [],
    warnings: [],
    cellChanges: [{ sheetId: '1', address, before: { value: null }, after: { value: 'hi' } }],
  }
}

describe('self-closing cell patching', () => {
  it('edits a self-closing <c/> without swallowing the next cell', async () => {
    const mutation = await applyPlanToXlsx(await fixture(), plan(), { '1': 'Data' })
    const zip = await JSZip.loadAsync(mutation.buffer)
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')

    // A1 now carries the new value, and its sibling B1 is untouched.
    expect(sheet).toContain('<is><t xml:space="preserve">hi</t></is>')
    expect(sheet).toContain('<c r="B1"><v>5</v></c>')
  })

  it('appends into a self-closing <row/> instead of the next row', async () => {
    const mutation = await applyPlanToXlsx(await fixture(WORKSHEET_EMPTY_ROW), plan('B1'), {
      '1': 'Data',
    })
    const zip = await JSZip.loadAsync(mutation.buffer)
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')

    expect(sheet).toMatch(
      /<row r="1" ht="20" customHeight="1"><c r="B1"[^>]*><is><t xml:space="preserve">hi<\/t><\/is><\/c><\/row>/,
    )
    expect(sheet).toContain('<row r="2"><c r="A2"><v>7</v></c></row>')
  })
})

describe('cell address guard at the plan boundary', () => {
  it('rejects addresses that are not grid cells instead of building a RegExp from them', async () => {
    for (const address of ['A1|B1', 'A1"[^>]*', 'a1', 'XFE1', 'A0', '']) {
      await expect(
        applyPlanToXlsx(await fixture(), plan(address), { '1': 'Data' }),
      ).rejects.toThrow(/Invalid cell address/)
    }
  })
})
