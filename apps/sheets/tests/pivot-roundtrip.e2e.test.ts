import { execFileSync } from 'node:child_process'
import { existsSync } from 'node:fs'
import { mkdtemp, readdir, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  assembleWithJsZip,
  createBufferEntrySource,
  planCellEditsToXlsx,
} from '../src/gateway/xlsx-gateway'
import type { SheetPivotAddition } from '../src/gateway/xlsx-gateway'

/// Round-trips a pivot-carrying save through LibreOffice: a headless convert
/// re-parses the whole package, so a malformed pivot part fails the convert
/// (or drops the parts). Runs only where soffice is installed.

const SOFFICE = ['/opt/homebrew/bin/soffice', '/usr/local/bin/soffice', '/usr/bin/soffice'].find(
  (path) => existsSync(path),
)

async function buildPivotSourceWorkbook(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
  )
  zip.file(
    'xl/styles.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf/></cellXfs>
</styleSheet>`,
  )
  // Source data A1:C4 (inline strings) plus the baked pivot grid at F1:G5.
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Region</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Product</t></is></c>
      <c r="C1" t="inlineStr"><is><t>Amount</t></is></c>
      <c r="F1" t="inlineStr"><is><t>Region</t></is></c>
      <c r="G1" t="inlineStr"><is><t>Sum of Amount</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>East</t></is></c>
      <c r="B2" t="inlineStr"><is><t>A</t></is></c>
      <c r="C2"><v>100</v></c>
      <c r="F2" t="inlineStr"><is><t>East</t></is></c>
      <c r="G2"><v>150</v></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>East</t></is></c>
      <c r="B3" t="inlineStr"><is><t>B</t></is></c>
      <c r="C3"><v>50</v></c>
      <c r="F3" t="inlineStr"><is><t>South</t></is></c>
      <c r="G3"><v>30</v></c>
    </row>
    <row r="4">
      <c r="A4" t="inlineStr"><is><t>South</t></is></c>
      <c r="B4" t="inlineStr"><is><t>A</t></is></c>
      <c r="C4"><v>30</v></c>
      <c r="F4" t="inlineStr"><is><t>Grand Total</t></is></c>
      <c r="G4"><v>180</v></c>
    </row>
  </sheetData>
</worksheet>`,
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

describe.skipIf(!SOFFICE)('pivot LibreOffice round-trip', () => {
  it('survives a headless convert with the pivot parts intact', async () => {
    const sourceBuffer = await buildPivotSourceWorkbook()
    const pivot: SheetPivotAddition = {
      sheetName: 'Data',
      sourceSheetName: 'Data',
      sourceArea: { startRow: 0, startColumn: 0, endRow: 3, endColumn: 2 },
      location: { startRow: 0, startColumn: 5, endRow: 3, endColumn: 6 },
      name: 'PivotRT',
      fieldNames: ['Region', 'Product', 'Amount'],
      rowFieldIndices: [0],
      rowItems: ['East', 'South'],
      values: [{ fieldIndex: 2, agg: 'sum' }],
    }
    const plan = await planCellEditsToXlsx(
      await createBufferEntrySource(sourceBuffer),
      [],
      [],
      [],
      undefined,
      [],
      [],
      [],
      [],
      [],
      null,
      [],
      [],
      [],
      [],
      [pivot],
    )
    const mutation = await assembleWithJsZip(sourceBuffer, plan)

    const workDir = await mkdtemp(join(tmpdir(), 'pivot-rt-'))
    const savedPath = join(workDir, 'pivot.xlsx')
    await writeFile(savedPath, mutation.buffer)
    execFileSync(
      SOFFICE as string,
      ['--headless', '--convert-to', 'xlsx', '--outdir', join(workDir, 'out'), savedPath],
      { timeout: 120_000 },
    )
    const outputs = await readdir(join(workDir, 'out'))
    expect(outputs).toContain('pivot.xlsx')

    // LibreOffice keeps (rewrites) the pivot parts when it understood them.
    const converted = await JSZip.loadAsync(
      await import('node:fs/promises').then((fs) =>
        fs.readFile(join(workDir, 'out', 'pivot.xlsx')),
      ),
    )
    const paths = Object.keys(converted.files)
    expect(paths.some((path) => /pivotTable/i.test(path))).toBe(true)
    expect(paths.some((path) => /pivotCache/i.test(path))).toBe(true)
  }, 180_000)
})
