import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import type { EntrySource } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import {
  createBufferEntrySource,
  planCellEditsToXlsx,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'

const workbook = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
    <sheet name="Huge" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`

const relationships = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`

const dataSheet = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>`

async function buildSource(hugeSheet: string): Promise<EntrySource> {
  const zip = new JSZip()
  zip.file('xl/workbook.xml', workbook)
  zip.file('xl/_rels/workbook.xml.rels', relationships)
  zip.file('xl/worksheets/sheet1.xml', dataSheet)
  zip.file('xl/worksheets/sheet2.xml', hugeSheet)
  return createBufferEntrySource(await zip.generateAsync({ type: 'nodebuffer' }))
}

/// Simulates the sidecar source: sheet2 is "too large" to read for patching.
function withOversizedSheet2(source: EntrySource): EntrySource {
  return {
    ...source,
    canPatch: async (path) => path !== 'xl/worksheets/sheet2.xml',
    // The real scanner searches decoded text nodes only — mimic that, or the
    // element name "sheetData" would false-positive on the needle "Data".
    containsText: async (path, needle) =>
      (await source.readText(path)).replace(/<[^>]*>/g, '\u0000').includes(needle),
    readText: async (path) => {
      if (path === 'xl/worksheets/sheet2.xml') {
        throw new Error('sheet2 must not be read — it is beyond the patch limit')
      }
      return source.readText(path)
    },
  }
}

const insertRow = {
  sheetName: 'Data',
  ops: [{ kind: 'insert-rows' as const, index: 0, count: 1 }],
}

describe('planCellEditsToXlsx with oversized entries', () => {
  it('skips an oversized sheet that does not reference the shifted one', async () => {
    const source = withOversizedSheet2(
      await buildSource(
        '<worksheet><sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData></worksheet>',
      ),
    )
    const plan = await planCellEditsToXlsx(source, [], [insertRow])
    expect([...plan.replaced.keys()].sort()).toEqual([
      'xl/workbook.xml',
      'xl/worksheets/sheet1.xml',
    ])
  })

  it('fails closed when the oversized sheet references the shifted one', async () => {
    const source = withOversizedSheet2(
      await buildSource(
        '<worksheet><sheetData><row r="1"><c r="A1"><f>Data!A1</f></c></row></sheetData></worksheet>',
      ),
    )
    await expect(planCellEditsToXlsx(source, [], [insertRow])).rejects.toThrow(
      /too large to rewrite/,
    )
  })
})
