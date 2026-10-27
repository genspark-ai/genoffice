import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { z } from 'zod'

import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'

import { XlsxSidecarClient } from '../src/main/xlsx-sidecar-client'
import { toUniverDvRule } from '../src/renderer/univer-sync'
import { workbookRangeResultSchema } from '../src/shared/desktop-api'

/// Excel's x14 data-validation extension, spelled out literally: a fixture
/// that imported the constant from the source could not tell it apart from
/// the conditional-formatting extension.
const X14_DV_EXT_URI = '{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}'
/// Excel's x14 conditional-formatting extension — a different one.
const X14_CF_EXT_URI = '{78C0D931-6437-407d-A8EE-F0AAD7539E65}'

const X14_EXT =
  '<extLst>' +
  `<ext uri="${X14_CF_EXT_URI}"` +
  ' xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
  '<x14:conditionalFormattings/>' +
  '</ext>' +
  `<ext uri="${X14_DV_EXT_URI}"` +
  ' xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
  '<x14:dataValidations xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" count="2">' +
  '<x14:dataValidation type="list" allowBlank="1">' +
  '<x14:formula1><xm:f>Lookup!$A$1:$A$5</xm:f></x14:formula1>' +
  '<xm:sqref>C1:C9</xm:sqref></x14:dataValidation>' +
  '<x14:dataValidation type="whole" operator="lessThan">' +
  '<x14:formula1><xm:f>100</xm:f></x14:formula1>' +
  '<xm:sqref>D1:D9</xm:sqref></x14:dataValidation>' +
  '</x14:dataValidations></ext></extLst>'

/// The content of the `<ext>` element carrying `uri`, or null when the
/// worksheet has no such extension.
function extContent(xml: string, uri: string): string | null {
  const open = new RegExp(`<ext\\b[^>]*\\buri="${uri.replace(/[{}]/g, '\\$&')}"[^>]*>`).exec(xml)
  if (open === null) return null
  const close = xml.indexOf('</ext>', open.index)
  return close === -1 ? null : xml.slice(open.index + open[0].length, close)
}

/// The same rules parked in the conditional-formatting extension, which is
/// where Excel does not read them from.
const X14_EXT_IN_CF =
  '<extLst>' +
  `<ext uri="${X14_CF_EXT_URI}"` +
  ' xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
  '<x14:dataValidations xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" count="1">' +
  '<x14:dataValidation type="list" allowBlank="1">' +
  '<x14:formula1><xm:f>Lookup!$A$1:$A$5</xm:f></x14:formula1>' +
  '<xm:sqref>C1:C9</xm:sqref></x14:dataValidation>' +
  '</x14:dataValidations></ext></extLst>'

function homeWorksheet(ext: string): string {
  return (
    '<?xml version="1.0"?>\n' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<dimension ref="A1:J100"/>' +
    '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Home</t></is></c></row></sheetData>' +
    '<dataValidations count="1"><dataValidation type="whole" operator="equal" sqref="B1">' +
    '<formula1>7</formula1></dataValidation></dataValidations>' +
    ext +
    '</worksheet>'
  )
}

const LOOKUP_WORKSHEET =
  '<?xml version="1.0"?>\n' +
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<dimension ref="A1:J100"/>' +
  '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>red</t></is></c></row>' +
  '<row r="2"><c r="A2" t="inlineStr"><is><t>green</t></is></c></row>' +
  '<row r="3"><c r="A3" t="inlineStr"><is><t>blue</t></is></c></row></sheetData></worksheet>'

async function buildFixture(ext: string = X14_EXT): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="Home" sheetId="1" r:id="rId1"/>' +
      '<sheet name="Lookup" sheetId="2" r:id="rId2"/></sheets></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' +
      '</Relationships>',
  )
  zip.file('xl/worksheets/sheet1.xml', homeWorksheet(ext))
  zip.file('xl/worksheets/sheet2.xml', LOOKUP_WORKSHEET)
  return zip.generateAsync({ type: 'nodebuffer' })
}

function sidecarBinaryPath(): string {
  const executable = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
  return fileURLToPath(
    new URL(`../native/xlsx-engine/target/release/${executable}`, import.meta.url),
  )
}

type RangeResult = z.infer<typeof workbookRangeResultSchema>

const openResultSchema = z.object({
  sessionId: z.string().uuid(),
  sheets: z.array(z.object({ name: z.string() })),
})

async function readHomeRules(
  client: XlsxSidecarClient,
  sessionId: string,
): Promise<RangeResult['dataValidations']> {
  return workbookRangeResultSchema.parse(
    await client.readRange({
      sessionId,
      sheetId: 'sheet-1',
      range: { startRow: 0, endRow: 99, startColumn: 0, endColumn: 9 },
    }),
  ).dataValidations
}

describe('x14 data validation round-trip (sidecar read + gateway write)', () => {
  it('loads, edits, saves, and reloads x14 validation rules without loss', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'xlsx-dv-x14-'))
    const client = new XlsxSidecarClient(sidecarBinaryPath())
    let sessionId: string | null = null
    try {
      const path = join(directory, 'fixture.xlsx')
      await writeFile(path, await buildFixture())
      const opened = openResultSchema.parse(await client.open(path))
      sessionId = opened.sessionId
      expect(opened.sheets.map((sheet) => sheet.name)).toEqual(['Home', 'Lookup'])

      // Read side: the x14 rules surface like any other rule.
      const fileRules = await readHomeRules(client, sessionId)
      expect(fileRules).toHaveLength(3)
      const crossSheet = fileRules.find((rule) => rule.ruleType === 'list') ?? null
      expect(crossSheet).toMatchObject({
        ranges: [{ startRow: 0, endRow: 8, startColumn: 2, endColumn: 2 }],
        formulas: ['Lookup!$A$1:$A$5'],
        allowBlank: true,
      })
      const legacyCapable = fileRules.filter((rule) => rule.ruleType === 'whole')
      expect(legacyCapable).toMatchObject([
        { operator: 'equal', formulas: ['7'], ranges: [{ startRow: 0, startColumn: 1 }] },
        { operator: 'lessThan', formulas: ['100'], ranges: [{ startRow: 0, startColumn: 3 }] },
      ])

      // Renderer install: the cross-sheet source keeps its sheet qualifier.
      expect(crossSheet).not.toBeNull()
      const univerRule = toUniverDvRule(
        crossSheet as NonNullable<typeof crossSheet>,
        'file-dv-sheet-1-1',
      )
      expect(univerRule).toMatchObject({ type: 'list', formula1: '=Lookup!$A$1:$A$5' })

      // Edit + save: the user retargets the list to Lookup!$B$2:$B$6 (still
      // cross-sheet, so still x14); the legacy-capable whole rules ride the
      // legacy section now.
      const mutation = await applyCellEditsToXlsx(
        await buildFixture(),
        [],
        [],
        [],
        undefined,
        [],
        [],
        [],
        [
          {
            sheetName: 'Home',
            rules: [
              {
                ranges: [{ startRow: 0, endRow: 8, startColumn: 2, endColumn: 2 }],
                rule: {
                  type: 'list',
                  formula1: '=Lookup!$B$2:$B$6',
                  allowBlank: true,
                  showDropDown: true,
                },
              },
              {
                ranges: [{ startRow: 0, endRow: 0, startColumn: 1, endColumn: 1 }],
                rule: { type: 'whole', operator: 'equal', formula1: '7' },
              },
              {
                ranges: [{ startRow: 0, endRow: 8, startColumn: 3, endColumn: 3 }],
                rule: { type: 'whole', operator: 'lessThan', formula1: '100' },
              },
            ],
          },
        ],
      )

      // Written XML: legacy section plus the x14 block, sibling CF intact.
      const zip = await JSZip.loadAsync(mutation.buffer)
      const worksheet = await zip.file('xl/worksheets/sheet1.xml')?.async('text')
      expect(worksheet).toContain('<dataValidations count="2">')
      expect(worksheet).toContain('<dataValidation type="whole" operator="equal" sqref="B1">')
      expect(worksheet).toContain('<dataValidation type="whole" operator="lessThan" sqref="D1:D9">')
      expect(worksheet).toContain('<x14:conditionalFormattings/>')
      expect(worksheet).toContain(
        '<x14:dataValidations xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"' +
          ' xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" count="1">',
      )
      expect(worksheet).toContain('<xm:f>Lookup!$B$2:$B$6</xm:f>')
      expect(worksheet).toContain('<xm:sqref>C1:C9</xm:sqref>')
      expect(worksheet).not.toContain('Lookup!$A$1:$A$5')
      // The block rides in Excel's x14 data-validation extension; the
      // conditional-formatting extension keeps only its own content.
      expect(extContent(worksheet ?? '', X14_DV_EXT_URI)).toContain('<x14:dataValidations')
      expect(extContent(worksheet ?? '', X14_CF_EXT_URI)).toBe('<x14:conditionalFormattings/>')

      // Reopen: the saved rules come back exactly as written.
      const savedPath = join(directory, 'saved.xlsx')
      await writeFile(savedPath, mutation.buffer)
      const reopened = openResultSchema.parse(await client.open(savedPath))
      const reloaded = await readHomeRules(client, reopened.sessionId)
      expect(reloaded).toHaveLength(3)
      expect(reloaded.find((rule) => rule.ruleType === 'list')).toMatchObject({
        ranges: [{ startRow: 0, endRow: 8, startColumn: 2, endColumn: 2 }],
        formulas: ['Lookup!$B$2:$B$6'],
        allowBlank: true,
      })
      expect(reloaded.filter((rule) => rule.ruleType === 'whole')).toMatchObject([
        { operator: 'equal', formulas: ['7'] },
        { operator: 'lessThan', formulas: ['100'] },
      ])
      await client.close(reopened.sessionId)
    } finally {
      if (sessionId) await client.close(sessionId)
      client.stop()
      await rm(directory, { recursive: true, force: true })
    }
  })

  it('ignores an x14 block parked in the conditional-formatting extension', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'xlsx-dv-x14-wrong-ext-'))
    const client = new XlsxSidecarClient(sidecarBinaryPath())
    let sessionId: string | null = null
    try {
      const path = join(directory, 'fixture.xlsx')
      await writeFile(path, await buildFixture(X14_EXT_IN_CF))
      const opened = openResultSchema.parse(await client.open(path))
      sessionId = opened.sessionId
      // Excel does not read validation rules from that extension, so neither
      // does the sidecar: only the legacy rule of the sheet is a rule.
      const fileRules = await readHomeRules(client, sessionId)
      expect(fileRules).toHaveLength(1)
      expect(fileRules[0]).toMatchObject({
        ruleType: 'whole',
        operator: 'equal',
        formulas: ['7'],
      })
    } finally {
      if (sessionId) await client.close(sessionId)
      client.stop()
      await rm(directory, { recursive: true, force: true })
    }
  })

  it('a legacy-only snapshot never grows an x14 block', async () => {
    const plain = homeWorksheet('').replace(
      /<dataValidations count="1">[\s\S]*?<\/dataValidations>/,
      '',
    )
    const zip = new JSZip()
    zip.file(
      'xl/workbook.xml',
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="Home" sheetId="1" r:id="rId1"/></sheets></workbook>',
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '</Relationships>',
    )
    zip.file('xl/worksheets/sheet1.xml', plain)
    const mutation = await applyCellEditsToXlsx(
      await zip.generateAsync({ type: 'nodebuffer' }),
      [],
      [],
      [],
      undefined,
      [],
      [],
      [],
      [
        {
          sheetName: 'Home',
          rules: [
            {
              ranges: [{ startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 }],
              rule: { type: 'list', formula1: 'Yes,No' },
            },
          ],
        },
      ],
    )
    const saved = await (
      await JSZip.loadAsync(mutation.buffer)
    )
      .file('xl/worksheets/sheet1.xml')
      ?.async('text')
    expect(saved).toContain('<formula1>"Yes,No"</formula1>')
    expect(saved).not.toContain('x14:')
    expect(saved).not.toContain('<extLst')
  })
})
