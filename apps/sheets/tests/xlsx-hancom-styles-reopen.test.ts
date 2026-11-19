import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { saveWorkbookViaSidecar } from '@genoffice/xlsx-gateway/gateway/xlsx-package-io'
import { XlsxSidecarClient } from '../src/main/xlsx-sidecar-client'

// The shape Hancom Cell (HCell 11.0) saves: every spreadsheetml element is
// x:-prefixed, and individual styles.xml entries are mc:AlternateContent
// blocks with xmlns:mc on the wrapper and xmlns:hs on mc:Choice only. A style
// edit used to hoist the Choice entries out of their wrappers, so the save
// failed its own reopen check with "an unknown namespace prefix 'hs'".
const SML = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
const HS = 'http://schemas.haansoft.com/office/spreadsheet/8.0'

const wrap = (choice: string, fallback: string): string =>
  `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:hs="${HS}" Requires="hs">${choice}</mc:Choice>` +
  `<mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`

const STYLES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><x:styleSheet xmlns:r="${R}" xmlns:x="${SML}">` +
  `<x:fonts count="3"><x:font><x:sz val="11"/><x:name val="Calibri"/></x:font>` +
  wrap(
    '<x:font hs:extension="1"><x:b/><x:sz val="18"/><x:name val="Calibri"/><hs:ratio val="90"/></x:font>',
    '<x:font><x:b/><x:sz val="18"/><x:name val="Calibri"/></x:font>',
  ) +
  `<x:font><x:sz val="9"/><x:name val="Calibri"/></x:font></x:fonts>` +
  `<x:fills count="2"><x:fill><x:patternFill patternType="none"/></x:fill><x:fill><x:patternFill patternType="gray125"/></x:fill></x:fills>` +
  `<x:borders count="1"><x:border/></x:borders>` +
  `<x:cellStyleXfs count="1"><x:xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></x:cellStyleXfs>` +
  `<x:cellXfs count="3"><x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
  `<x:xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>` +
  `<x:xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/></x:cellXfs>` +
  `<x:cellStyles count="1"><x:cellStyle name="Normal" xfId="0" builtinId="0"/></x:cellStyles>` +
  `<x:dxfs count="2">` +
  wrap(
    '<x:dxf hs:applyExtension="1"><x:font hs:extension="1"><x:b/><hs:spacing val="0"/></x:font></x:dxf>',
    '<x:dxf><x:font><x:b/></x:font></x:dxf>',
  ) +
  `<x:dxf><x:fill><x:patternFill><x:bgColor rgb="FFFF0000"/></x:patternFill></x:fill></x:dxf></x:dxfs>` +
  `</x:styleSheet>`

async function hancomWorkbook(path: string): Promise<void> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  )
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8"?><x:workbook xmlns:r="${R}" xmlns:x="${SML}"><x:sheets><x:sheet name="Sheet1" sheetId="1" r:id="rId1"/></x:sheets><x:calcPr xmlns:mc="${MC}" xmlns:hs="${HS}" mc:Ignorable="hs" hs:hclCalcId="904"/></x:workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
  )
  zip.file('xl/styles.xml', STYLES)
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8"?><x:worksheet xmlns:r="${R}" xmlns:x="${SML}"><x:sheetData><x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c><x:c r="B1" s="1"><x:v>2</x:v></x:c><x:c r="C1" s="2"><x:v>3</x:v></x:c></x:row></x:sheetData></x:worksheet>`,
  )
  await writeFile(path, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
}

type Opened = {
  sessionId: string
  styles: { bold?: boolean; italic?: boolean; fontSize?: number }[]
  dxfStyles: { bold?: boolean; fillColor?: string }[]
}

describe('Hancom Cell stylesheet across a style edit', () => {
  it('saves a result the sidecar reopens, with every existing id unchanged', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'xlsx-hancom-styles-'))
    const source = join(directory, 'hancom.xlsx')
    const target = join(directory, 'saved.xlsx')
    await hancomWorkbook(source)
    const client = new XlsxSidecarClient(sidecarBinaryPath())
    const sessions: string[] = []
    try {
      const before = (await client.open(source)) as Opened
      sessions.push(before.sessionId)
      await saveWorkbookViaSidecar({
        client,
        sourcePath: source,
        targetPath: target,
        edits: [
          // derives from xf 2 (9pt); a shifted list handed it the 18pt font
          {
            sheetName: 'Sheet1',
            row: 0,
            column: 2,
            writeValue: false,
            cell: { value: null },
            style: { italic: true },
          },
        ],
      })
      const after = (await client.open(target)) as Opened
      sessions.push(after.sessionId)

      expect(after.styles.slice(0, 3).map((style) => [style.bold, style.fontSize])).toEqual(
        before.styles.map((style) => [style.bold, style.fontSize]),
      )
      expect(after.styles).toHaveLength(4)
      expect(after.styles[3]).toMatchObject({ italic: true, fontSize: 9 })
      expect(after.styles[3]?.bold).not.toBe(true)
      expect(after.dxfStyles).toHaveLength(2)
      expect(after.dxfStyles[0]?.bold).toBe(true)
      expect(after.dxfStyles[1]?.fillColor).toBe('#FF0000')

      const saved = await JSZip.loadAsync(await readFile(target))
      const styles = (await saved.file('xl/styles.xml')?.async('string')) ?? ''
      // the Hancom branches survive byte for byte
      expect(styles).toContain(`<mc:Choice xmlns:hs="${HS}" Requires="hs"><font hs:extension="1">`)
      expect(styles).toContain('<dxfs count="2">')
    } finally {
      for (const sessionId of sessions) await client.close(sessionId).catch(() => undefined)
      client.stop()
      await rm(directory, { recursive: true, force: true })
    }
  })
})

function sidecarBinaryPath(): string {
  const executable = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
  return fileURLToPath(
    new URL(`../native/xlsx-engine/target/release/${executable}`, import.meta.url),
  )
}
