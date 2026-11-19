// One-off benchmark for the resident-model recalc cache (not part of gate).
// Builds an Excel-style workbook with ROWS data rows + one SUM column, then
// times a cold recalc (file import) against warm recalcs (resident model).
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import JSZip from 'jszip'

import { XlsxSidecarClient } from '../src/main/xlsx-sidecar-client'

const ROWS = Number(process.env.BENCH_ROWS ?? 50_000)

async function buildFixture(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>' +
      '</workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/styles.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
      '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
      '<borders count="1"><border/></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '</styleSheet>',
  )
  const rows: string[] = []
  for (let row = 1; row <= ROWS; row++) {
    rows.push(
      `<row r="${row}">` +
        `<c r="A${row}"><v>${row}</v></c>` +
        `<c r="B${row}"><v>${row * 2}</v></c>` +
        `<c r="C${row}"><f>A${row}+B${row}</f><v>${row * 3}</v></c>` +
        '</row>',
    )
  }
  zip.file(
    'xl/worksheets/sheet1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      `<sheetData>${rows.join('')}</sheetData>` +
      '</worksheet>',
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

async function timeRecalc(
  client: XlsxSidecarClient,
  path: string,
  input: string,
): Promise<{ ms: number; cached: boolean }> {
  const start = performance.now()
  const result = (await client.recalcCells({
    path,
    edits: [{ sheet: 'Data', row: 0, column: 0, input }],
    reads: [{ sheet: 'Data', range: { startRow: 0, endRow: 99, startColumn: 2, endColumn: 2 } }],
  })) as { cached: boolean }
  return { ms: performance.now() - start, cached: result.cached }
}

async function main(): Promise<void> {
  const directory = await mkdtemp(join(tmpdir(), 'bench-recalc-'))
  const path = join(directory, 'bench.xlsx')
  await writeFile(path, await buildFixture())
  const executable = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
  const client = new XlsxSidecarClient(
    fileURLToPath(new URL(`../native/xlsx-engine/target/release/${executable}`, import.meta.url)),
  )
  try {
    const cold = await timeRecalc(client, path, '111')
    const warm1 = await timeRecalc(client, path, '222')
    const warm2 = await timeRecalc(client, path, '333')
    console.log(`rows=${ROWS} (${ROWS} formulas)`)
    console.log(`cold  ${cold.ms.toFixed(0).padStart(6)}ms  cached=${cold.cached}`)
    console.log(`warm1 ${warm1.ms.toFixed(0).padStart(6)}ms  cached=${warm1.cached}`)
    console.log(`warm2 ${warm2.ms.toFixed(0).padStart(6)}ms  cached=${warm2.cached}`)
  } finally {
    client.stop()
    await rm(directory, { recursive: true, force: true })
  }
}

void main()
