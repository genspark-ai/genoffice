import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  applyCellEditsToXlsx,
  assertOnlyTouchedEntriesChanged,
  inventoryXlsx,
} from '../src/gateway/xlsx-gateway'

/**
 * A workbook's data connections are not part of the grid: Power Query keeps its
 * M code in `customXml/item1.xml` (the DataMashup part) and the query and its
 * destination in `xl/connections.xml`, `xl/queryTables/` and `xl/tables/`.
 * GenOffice does not run those queries, but it must not damage them: the file
 * has to stay refreshable in Excel.
 *
 * So the save path may touch the worksheet and the workbook part, and nothing
 * else. These tests pin that down — without them, a save-path change that
 * rewrites or drops a connection part would look like a successful save and
 * only surface later, as "Excel can no longer refresh this workbook".
 */

const CONNECTIONS_PART = 'xl/connections.xml'
const QUERY_TABLE_PART = 'xl/queryTables/queryTable1.xml'
const TABLE_PART = 'xl/tables/table1.xml'
const DATA_MASHUP_PART = 'customXml/item1.xml'

/** every part a connection-bearing workbook carries besides the grid */
const CONNECTION_PARTS = [CONNECTIONS_PART, QUERY_TABLE_PART, TABLE_PART, DATA_MASHUP_PART]

const CONNECTIONS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <connection id="1" name="Rates" type="5" refreshedVersion="8" background="1" saveData="1">
    <dbPr connection="Provider=Microsoft.Mashup.OleDb;Data Source=$Workbook$;Location=Rates;Extended Properties=&quot;&quot;" command="SELECT * FROM [Rates]"/>
  </connection>
</connections>`

const QUERY_TABLE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="Rates" connectionId="1" autoFormatId="16" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="0">
  <queryTableRefresh nextId="2">
    <queryTableFields count="2">
      <queryTableField id="1" name="Currency" tableColumnId="1"/>
      <queryTableField id="2" name="Rate" tableColumnId="2"/>
    </queryTableFields>
  </queryTableRefresh>
</queryTable>`

const TABLE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Rates" displayName="Rates" ref="A1:B3" totalsRowShown="0">
  <autoFilter ref="A1:B3"/>
  <tableColumns count="2">
    <tableColumn id="1" name="Currency"/>
    <tableColumn id="2" name="Rate"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>`

/** the DataMashup part: Excel stores the M code as a base64 package */
const DATA_MASHUP_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<b xmlns="http://schemas.microsoft.com/office/2006/binary/activeXControl">DATA_MASHUP</b>`

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/>
  <Override PartName="/xl/queryTables/queryTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"/>
  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
</Types>`

const PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`

const WORKBOOK_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" Target="connections.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>
</Relationships>`

const WORKSHEET_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable" Target="../queryTables/queryTable1.xml"/>
</Relationships>`

const CUSTOM_XML_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/>
</Relationships>`

const WORKBOOK_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Rates" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`

const WORKSHEET_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>EUR</t></is></c><c r="B1" t="inlineStr"><is><t>1.09</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>USD</t></is></c><c r="B2" t="inlineStr"><is><t>1.00</t></is></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>JPY</t></is></c><c r="B3" t="inlineStr"><is><t>0.0067</t></is></c></row>
  </sheetData>
  <tableParts count="1"><tablePart r:id="rId1"/></tableParts>
</worksheet>`

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs>
  <cellXfs count="1"><xf xfId="0"/></cellXfs>
</styleSheet>`

async function connectionWorkbook(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', CONTENT_TYPES)
  zip.file('_rels/.rels', PACKAGE_RELS)
  zip.file('xl/workbook.xml', WORKBOOK_XML)
  zip.file('xl/_rels/workbook.xml.rels', WORKBOOK_RELS)
  zip.file('xl/worksheets/sheet1.xml', WORKSHEET_XML)
  zip.file('xl/worksheets/_rels/sheet1.xml.rels', WORKSHEET_RELS)
  zip.file('xl/styles.xml', STYLES_XML)
  zip.file('xl/connections.xml', CONNECTIONS_XML)
  zip.file('xl/queryTables/queryTable1.xml', QUERY_TABLE_XML)
  zip.file('xl/tables/table1.xml', TABLE_XML)
  zip.file('customXml/item1.xml', DATA_MASHUP_XML)
  zip.file(
    'customXml/itemProps1.xml',
    '<?xml version="1.0"?><ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="{1B4A0E2C-6D3F-4C11-9F2A-7E5C0D1F3A44}"><ds:schemaRefs/></ds:datastoreItem>',
  )
  zip.file('customXml/_rels/item1.xml.rels', CUSTOM_XML_RELS)
  return zip.generateAsync({ type: 'nodebuffer' })
}

function shaByPath(entries: readonly { path: string; sha256: string }[]): Map<string, string> {
  return new Map(entries.map((entry) => [entry.path, entry.sha256]))
}

describe('a save leaves the connection parts alone', () => {
  it('keeps connections, the query table, the table and the DataMashup byte-identical', async () => {
    const source = await connectionWorkbook()
    const before = shaByPath(await inventoryXlsx(source))

    const mutation = await applyCellEditsToXlsx(source, [
      {
        sheetName: 'Rates',
        row: 1,
        column: 1,
        writeValue: true,
        cell: { value: 'GBP' },
      },
    ])

    // the guard the save path already runs: nothing outside the touched parts
    // may change, and nothing may be dropped
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()

    const after = shaByPath(await inventoryXlsx(mutation.buffer))
    for (const part of CONNECTION_PARTS) {
      expect(after.has(part), `${part} disappeared from the package`).toBe(true)
      expect(after.get(part), `${part} was rewritten`).toBe(before.get(part))
    }
  })

  it('changed the grid, so the preservation above is not vacuous', async () => {
    const source = await connectionWorkbook()
    const before = shaByPath(await inventoryXlsx(source))

    const mutation = await applyCellEditsToXlsx(source, [
      {
        sheetName: 'Rates',
        row: 1,
        column: 1,
        writeValue: true,
        cell: { value: 'GBP' },
      },
    ])

    const after = shaByPath(await inventoryXlsx(mutation.buffer))
    expect(after.get('xl/worksheets/sheet1.xml')).not.toBe(before.get('xl/worksheets/sheet1.xml'))
    // every connection part still there
    for (const part of CONNECTION_PARTS) {
      expect(after.has(part)).toBe(true)
    }
  })

  it('keeps the connection parts when the edit lands inside the query table range', async () => {
    const source = await connectionWorkbook()
    const before = shaByPath(await inventoryXlsx(source))

    // A cell inside the table's ref: the table part describes the range the
    // query fills, so a careless save would rewrite it and break the binding
    // Excel relies on to refresh the table.
    const mutation = await applyCellEditsToXlsx(source, [
      {
        sheetName: 'Rates',
        row: 2,
        column: 1,
        writeValue: true,
        cell: { value: 'CHF' },
      },
    ])

    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()

    const after = shaByPath(await inventoryXlsx(mutation.buffer))
    for (const part of CONNECTION_PARTS) {
      expect(after.has(part), `${part} disappeared from the package`).toBe(true)
      expect(after.get(part), `${part} was rewritten`).toBe(before.get(part))
    }
  })
})
