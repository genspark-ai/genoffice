import { readFileSync } from 'node:fs'
import { runInNewContext } from 'node:vm'
import ts from 'typescript'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { saveWorkbookViaSidecar } from '@genoffice/xlsx-gateway/gateway/xlsx-package-io'
import { XlsxSidecarClient } from '../src/main/xlsx-sidecar-client'
import { createEditJournal, recordSetRangeValues, toSaveEdits } from '../src/renderer/edit-journal'
import { handleRibbonCommand, type RibbonCommandContext } from '../src/renderer/ribbon-actions'
import { cellFontTextStyle, fontTextStyleOf, toUniverStyle } from '../src/renderer/univer-sync'
import { workbookFileSchema, type WorkbookCellStyle } from '../src/shared/desktop-api'

// Execute the actual preload whitelist without starting Electron.
const preloadSource = ts.createSourceFile(
  'preload.ts',
  readFileSync(new URL('../src/preload/index.ts', import.meta.url), 'utf8'),
  ts.ScriptTarget.Latest,
  true,
)
const helperNames = new Set([
  'parseCellStyle',
  'parseBorderEdge',
  'isRecord',
  'isNonnegativeInteger',
  'isOptionalString',
])
const helperSource = preloadSource.statements
  .filter((node) => ts.isFunctionDeclaration(node) && helperNames.has(node.name?.text ?? ''))
  .map((node) => node.getText(preloadSource))
  .join('\n')
const parsePreloadStyle = runInNewContext(
  ts.transpile(helperSource + '\nparseCellStyle', { target: ts.ScriptTarget.ES2022 }),
) as (input: unknown) => WorkbookCellStyle

const SML = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
const OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
async function fixture(path: string) {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OFFICE}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  )
  zip.file(
    'xl/workbook.xml',
    `<workbook xmlns="${SML}" xmlns:r="${OFFICE}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OFFICE}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${OFFICE}/styles" Target="styles.xml"/></Relationships>`,
  )
  zip.file(
    'xl/styles.xml',
    `<styleSheet xmlns="${SML}"><fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><u/></font><font><u val="double"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3">${[0, 1, 2].map((id) => `<xf numFmtId="0" fontId="${id}" fillId="0" borderId="0" xfId="0" applyFont="1"/>`).join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<worksheet xmlns="${SML}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Plain text</t></is></c><c r="B1" s="1" t="inlineStr"><is><t>Single</t></is></c><c r="C1" s="2" t="inlineStr"><is><t>Existing double</t></is></c></row></sheetData></worksheet>`,
  )
  await writeFile(path, await zip.generateAsync({ type: 'nodebuffer' }))
}
type Opened = { sessionId: string; styles: WorkbookCellStyle[] }
const binary =
  process.env.XLSX_SIDECAR_PATH ??
  fileURLToPath(new URL('../native/xlsx-engine/target/release/xlsx-sidecar', import.meta.url))

describe('double underline after saving and reopening', () => {
  it('the preload retains double, accepts legacy styles, and rejects unknown subtypes', () => {
    const plain = {
      bold: false,
      italic: false,
      underline: true,
      strikethrough: false,
      wrapText: false,
      diagonalUp: false,
      diagonalDown: false,
    }
    expect(parsePreloadStyle({ ...plain, underlineStyle: 'double' }).underlineStyle).toBe('double')
    expect(parsePreloadStyle(plain).underlineStyle).toBeUndefined()
    expect(() => parsePreloadStyle({ ...plain, underlineStyle: 'wavy' })).toThrow(
      'Invalid workbook style response.',
    )
  })

  it('the actual ribbon/journal/save/sidecar/import path retains the second line', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'xlsx-double-underline-'))
    const source = join(dir, 'source.xlsx'),
      target = join(dir, 'saved.xlsx')
    const client = new XlsxSidecarClient(binary),
      sessions: string[] = []
    try {
      await fixture(source)
      const before = (await client.open(source)) as Opened
      sessions.push(before.sessionId)
      const journal = createEditJournal()
      let current: Record<string, unknown> = {}
      let mutation: unknown
      const range = {
        getCellStyleData: () => current,
        setValue: (value: { s: Record<string, unknown> }) => {
          mutation = value
          current = value.s
          recordSetRangeValues(journal, 'sheet-1', { 0: { 0: value } })
        },
      }
      const ctx = {
        univerRef: {
          current: {
            univerAPI: {
              getActiveWorkbook: () => ({
                getActiveRange: () => range,
                getActiveSheet: () => null,
              }),
            },
          },
        },
        setMessage: () => {},
      } as unknown as RibbonCommandContext
      handleRibbonCommand(ctx, 'underline:double')
      expect(mutation).toEqual({ s: { ul: { s: 1, t: 10 } } })
      const edits = toSaveEdits(journal)
      expect(edits[0]?.style).toMatchObject({ underline: true, underlineStyle: 'double' })
      await saveWorkbookViaSidecar({
        client,
        sourcePath: source,
        targetPath: target,
        edits: edits.map(({ sheetId: _, value, ...edit }) => ({
          ...edit,
          sheetName: 'Sheet1',
          cell: { value },
        })),
      })
      const saved = await JSZip.loadAsync(await readFile(target))
      const xml = await saved.file('xl/worksheets/sheet1.xml')!.async('string')
      const index = Number(/<c\b(?=[^>]*\br="A1")[^>]*\bs="(\d+)"/.exec(xml)?.[1])
      const stylesXml = await saved.file('xl/styles.xml')!.async('string')
      const xfs = [
        ...(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stylesXml)?.[1] ?? '').matchAll(
          /<xf\b[^>]*\/?>(?:[\s\S]*?<\/xf>)?/g,
        ),
      ]
      const fontId = Number(/\bfontId="(\d+)"/.exec(xfs[index]?.[0] ?? '')?.[1])
      const fonts = [...stylesXml.matchAll(/<font\b[^>]*>[\s\S]*?<\/font>/g)]
      expect(fonts[fontId]?.[0]).toContain('<u val="double"/>')
      const after = (await client.open(target)) as Opened
      sessions.push(after.sessionId)
      const imported = parsePreloadStyle(
        workbookFileSchema.shape.styles.element.parse(after.styles[index]!),
      )
      expect(toUniverStyle(imported).ul).toEqual({ s: 1, t: 10 })
      expect(cellFontTextStyle(imported).ul).toEqual({ s: 1, t: 10 })
      expect(fontTextStyleOf(toUniverStyle(imported)).ul).toEqual({ s: 1, t: 10 })
      expect(toUniverStyle(after.styles[0]!).ul).toBeUndefined()
      expect(toUniverStyle(after.styles[1]!).ul).toEqual({ s: 1 })
      expect(toUniverStyle(after.styles[2]!).ul).toEqual({ s: 1, t: 10 })
      const original = await JSZip.loadAsync(await readFile(source))
      for (const name of ['xl/_rels/workbook.xml.rels', '_rels/.rels', '[Content_Types].xml'])
        expect(await saved.file(name)!.async('nodebuffer')).toEqual(
          await original.file(name)!.async('nodebuffer'),
        )
      // Saving requests recalculation; the rest of the workbook part is untouched.
      const workbookXml = await saved.file('xl/workbook.xml')!.async('string')
      expect(workbookXml.replace(/<calcPr\b[^>]*\/>/g, '')).toBe(
        await original.file('xl/workbook.xml')!.async('string'),
      )
      expect(xml).toContain('Plain text')
      expect(xml).toContain('Existing double')
      current = toUniverStyle(imported) as Record<string, unknown>
      handleRibbonCommand(ctx, 'underline:double')
      expect(toSaveEdits(journal)[0]?.style).toMatchObject({ underline: false })
      const cleared = join(dir, 'cleared.xlsx')
      await saveWorkbookViaSidecar({
        client,
        sourcePath: target,
        targetPath: cleared,
        edits: [
          {
            sheetName: 'Sheet1',
            row: 0,
            column: 0,
            writeValue: false,
            cell: { value: null },
            style: toSaveEdits(journal)[0]!.style,
          },
        ],
      })
      const clearedFile = (await client.open(cleared)) as Opened
      sessions.push(clearedFile.sessionId)
      const clearedXml = await (
        await JSZip.loadAsync(await readFile(cleared))
      )
        .file('xl/worksheets/sheet1.xml')!
        .async('string')
      const clearedIndex = Number(/<c\b(?=[^>]*\br="A1")[^>]*\bs="(\d+)"/.exec(clearedXml)?.[1])
      expect(toUniverStyle(clearedFile.styles[clearedIndex]!).ul).toBeUndefined()
    } finally {
      for (const id of sessions) await client.close(id).catch(() => {})
      client.stop()
      await rm(dir, { recursive: true, force: true })
    }
  })
})
