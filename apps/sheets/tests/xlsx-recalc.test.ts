import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { afterAll, describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { z } from 'zod'

import { XlsxSidecarClient } from '../src/main/xlsx-sidecar-client'

const recalcResultSchema = z
  .object({
    cells: z.array(
      z
        .object({
          sheet: z.string(),
          row: z.number().int().nonnegative(),
          column: z.number().int().nonnegative(),
          formatted: z.string(),
          number: z.number().optional(),
          isFormula: z.boolean(),
        })
        .strict(),
    ),
    cached: z.boolean(),
  })
  .strict()

// IronCalc's importer is strict: single-line XML (whitespace text nodes crash
// it) and a styles.xml complete with cellStyleXfs/cellStyles. This mirrors
// what Excel itself produces.
async function buildRecalcFixture(): Promise<Buffer> {
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
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '</styleSheet>',
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<dimension ref="A1:B4"/>' +
      '<sheetViews><sheetView workbookViewId="0"/></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="15"/>' +
      '<sheetData>' +
      '<row r="1"><c r="A1"><v>10</v></c></row>' +
      '<row r="2"><c r="A2"><v>20</v></c></row>' +
      '<row r="3"><c r="A3"><f>SUM(A1:A2)</f><v>30</v></c></row>' +
      '<row r="4"><c r="B4" t="str"><f>MID(CELL("filename",A1),FIND("]",CELL("filename",A1))+1,31)</f><v>Data</v></c></row>' +
      '</sheetData>' +
      '</worksheet>',
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

// Excel keeps the cached values of external-workbook references when the
// source is unreachable; IronCalc cannot even parse them. Content types and
// rels carry the externalLink part exactly as Excel writes it.
async function buildExternalLinkFixture(): Promise<Buffer> {
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
      '<Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>' +
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
      '<externalReferences><externalReference r:id="rId3"/></externalReferences>' +
      '</workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/styles.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '</styleSheet>',
  )
  zip.file(
    'xl/externalLinks/externalLink1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<externalBook r:id="rId1"><sheetNames><sheetName val="Sheet1"/></sheetNames>' +
      '<sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="A1"><v>42</v></cell></row></sheetData></sheetDataSet>' +
      '</externalBook></externalLink>',
  )
  zip.file(
    'xl/externalLinks/_rels/externalLink1.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="file:///C:/data/source.xlsx" TargetMode="External"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<dimension ref="A1:C2"/>' +
      '<sheetViews><sheetView workbookViewId="0"/></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="15"/>' +
      '<sheetData>' +
      '<row r="1"><c r="A1"><f>[1]Sheet1!A1*2</f><v>84</v></c><c r="B1"><f>A1+1</f><v>85</v></c>' +
      '<c r="C1" t="str"><f>\'[1]Sheet1\'!A1&amp;" units"</f><v>42 units</v></c></row>' +
      '<row r="2"><c r="A2"><v>5</v></c><c r="B2"><f>SUM(A1:A2)</f><v>89</v></c></row>' +
      '</sheetData>' +
      '</worksheet>',
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

const cleanups: string[] = []

afterAll(async () => {
  await Promise.all(cleanups.map((directory) => rm(directory, { recursive: true, force: true })))
})

describe('sidecar IronCalc recalculation channel', () => {
  it('applies edits and returns recalculated formula values', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'xlsx-recalc-test-'))
    cleanups.push(directory)
    const path = join(directory, 'recalc.xlsx')
    await writeFile(path, await buildRecalcFixture())
    const client = new XlsxSidecarClient(sidecarBinaryPath())
    try {
      const result = recalcResultSchema.parse(
        await client.recalcCells({
          path,
          edits: [{ sheet: 'Data', row: 0, column: 0, input: '100' }],
          reads: [
            {
              sheet: 'Data',
              range: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 },
            },
          ],
        }),
      )
      const sum = result.cells.find((cell) => cell.row === 2)
      expect(sum).toEqual({
        sheet: 'Data',
        row: 2,
        column: 0,
        formatted: '120',
        number: 120,
        isFormula: true,
      })
      expect(result.cached).toBe(false)
    } finally {
      client.stop()
    }
  })

  it('serves repeat requests from the resident model and rebuilds after a revert', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'xlsx-recalc-test-'))
    cleanups.push(directory)
    const path = join(directory, 'recalc.xlsx')
    await writeFile(path, await buildRecalcFixture())
    const client = new XlsxSidecarClient(sidecarBinaryPath())
    const reads = [
      { sheet: 'Data', range: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 } },
    ]
    try {
      const first = recalcResultSchema.parse(
        await client.recalcCells({
          path,
          edits: [{ sheet: 'Data', row: 0, column: 0, input: '100' }],
          reads,
        }),
      )
      expect(first.cached).toBe(false)
      // changed edit: incremental set_user_input on the resident model
      const second = recalcResultSchema.parse(
        await client.recalcCells({
          path,
          edits: [{ sheet: 'Data', row: 0, column: 0, input: '200' }],
          reads,
        }),
      )
      expect(second.cached).toBe(true)
      expect(second.cells.find((cell) => cell.row === 2)?.formatted).toBe('220')
      // undo removed the edit: only the file knows A1's original content
      const third = recalcResultSchema.parse(await client.recalcCells({ path, edits: [], reads }))
      expect(third.cached).toBe(false)
      expect(third.cells.find((cell) => cell.row === 2)?.formatted).toBe('30')
    } finally {
      client.stop()
    }
  })

  it('pins external-workbook references to their cached values (issue 235)', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'xlsx-recalc-test-'))
    cleanups.push(directory)
    const path = join(directory, 'external.xlsx')
    await writeFile(path, await buildExternalLinkFixture())
    const client = new XlsxSidecarClient(sidecarBinaryPath())
    try {
      const result = recalcResultSchema.parse(
        await client.recalcCells({
          path,
          edits: [{ sheet: 'Data', row: 1, column: 0, input: '10' }],
          reads: [
            {
              sheet: 'Data',
              range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 2 },
            },
          ],
        }),
      )
      const at = (row: number, column: number) =>
        result.cells.find((cell) => cell.row === row && cell.column === column)
      expect(result.cells.some((cell) => cell.formatted === '#ERROR!')).toBe(false)
      // the linked cells come back as plain cached values, not formulas
      expect(at(0, 0)).toMatchObject({ formatted: '84', isFormula: false })
      expect(at(0, 2)).toMatchObject({ formatted: '42 units', isFormula: false })
      // dependents compute against those values instead of cascading
      expect(at(0, 1)).toMatchObject({ formatted: '85', isFormula: true })
      expect(at(1, 1)).toMatchObject({ formatted: '94', isFormula: true })
    } finally {
      client.stop()
    }
  })

  // IronCalc has no CELL("filename"); its error must not displace the
  // cached sheet-name text the file already carries.
  it('omits erroring CELL("filename") cells so the cached value survives', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'xlsx-recalc-test-'))
    cleanups.push(directory)
    const path = join(directory, 'recalc.xlsx')
    await writeFile(path, await buildRecalcFixture())
    const client = new XlsxSidecarClient(sidecarBinaryPath())
    try {
      const result = recalcResultSchema.parse(
        await client.recalcCells({
          path,
          edits: [{ sheet: 'Data', row: 0, column: 0, input: '100' }],
          reads: [
            {
              sheet: 'Data',
              range: { startRow: 3, endRow: 3, startColumn: 1, endColumn: 1 },
            },
          ],
        }),
      )
      expect(result.cells).toEqual([])
    } finally {
      client.stop()
    }
  })

  it('fails soft on workbooks the engine cannot import', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'xlsx-recalc-test-'))
    cleanups.push(directory)
    const path = join(directory, 'broken.xlsx')
    await writeFile(path, Buffer.from('not a zip'))
    const client = new XlsxSidecarClient(sidecarBinaryPath())
    try {
      await expect(client.recalcCells({ path, edits: [], reads: [] })).rejects.toThrow()
    } finally {
      client.stop()
    }
  })
})

function sidecarBinaryPath(): string {
  const executable = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
  return fileURLToPath(
    new URL(`../native/xlsx-engine/target/release/${executable}`, import.meta.url),
  )
}
