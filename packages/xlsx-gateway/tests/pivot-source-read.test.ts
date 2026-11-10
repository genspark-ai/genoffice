import { describe, expect, it } from 'vitest'

import type { MutablePackage } from '../src/gateway/xlsx-drawing-add'
import { applyPivotAdditions, type PivotAddition } from '../src/gateway/xlsx-pivot-add'

const WORKSHEET = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Region</t></is></c><c r="B1" t="inlineStr"><is><t>Product</t></is></c><c r="C1" t="inlineStr"><is><t>Amount</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>East</t></is></c><c r="B2" s="1"/><c r="C2"><v>9</v></c></row>
  </sheetData>
</worksheet>`

const WORKBOOK = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`

const WORKBOOK_RELS = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`

function inMemoryPackage(entries: Record<string, string>): MutablePackage {
  const files = new Map(Object.entries(entries))
  return {
    paths: async () => [...files.keys()],
    has: async (path) => files.has(path),
    readText: async (path) => {
      const content = files.get(path)
      if (content === undefined) throw new Error(`missing part: ${path}`)
      return content
    },
    write: (path, content) => void files.set(path, content),
    add: (path, content) => void files.set(path, content),
    addBinary: (path, bytes) => void files.set(path, new TextDecoder().decode(bytes)),
    remove: (path) => void files.delete(path),
  }
}

function addition(): PivotAddition {
  return {
    worksheetPath: 'xl/worksheets/sheet1.xml',
    sourceSheetName: 'Data',
    sourceArea: { startRow: 0, startColumn: 0, endRow: 1, endColumn: 2 },
    location: { startRow: 0, startColumn: 4, endRow: 2, endColumn: 5 },
    name: 'Pivot1',
    fieldNames: ['Region', 'Product', 'Amount'],
    rowFieldIndices: [0],
    rowItems: ['East'],
    values: [{ fieldIndex: 1, agg: 'sum' }],
  }
}

async function cacheRecords(): Promise<string> {
  const pkg = inMemoryPackage({
    '[Content_Types].xml': CONTENT_TYPES,
    'xl/workbook.xml': WORKBOOK,
    'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
    'xl/worksheets/sheet1.xml': WORKSHEET,
  })
  await applyPivotAdditions(pkg, [addition()], WORKBOOK, new Set())
  return pkg.readText('xl/pivotCache/pivotCacheRecords1.xml')
}

describe('pivot source cell reads', () => {
  it('keeps a self-closing styled blank cell from absorbing the next cell', async () => {
    const records = await cacheRecords()
    expect(records).toContain('count="1"')
    expect(records).toContain('<r><x v="0"/><m/><n v="9"/></r>')
  })
})
