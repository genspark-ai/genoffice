import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { applyCellEditsToXlsx, assertOnlyTouchedEntriesChanged } from '../src/gateway/xlsx-gateway'

// Round trip over a workbook whose styles.xml wraps entries the way Hancom
// Cell does (mc:AlternateContent, xmlns:hs bound on mc:Choice only): only the
// parts an edit touches change, and inside styles.xml every existing entry —
// wrapped blocks included — is kept byte for byte ahead of the new ones.
const SML = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
const HS = 'http://schemas.haansoft.com/office/spreadsheet/8.0'

const wrap = (choice: string, fallback: string): string =>
  `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:hs="${HS}" Requires="hs">${choice}</mc:Choice>` +
  `<mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`

const SECTIONS = {
  fonts: [
    '<font><sz val="11"/><name val="Calibri"/></font>',
    wrap(
      '<font hs:extension="1"><b/><sz val="18"/><name val="Calibri"/><hs:ratio val="90"/></font>',
      '<font><b/><sz val="18"/><name val="Calibri"/></font>',
    ),
  ],
  fills: [
    '<fill><patternFill patternType="none"/></fill>',
    '<fill><patternFill patternType="gray125"/></fill>',
  ],
  borders: ['<border/>'],
  cellXfs: [
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
    wrap(
      '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" hs:applyExtension="1"/>',
      '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
    ),
  ],
  dxfs: [
    wrap(
      '<dxf hs:applyExtension="1"><font hs:extension="1"><b/><hs:spacing val="0"/></font></dxf>',
      '<dxf><font><b/></font></dxf>',
    ),
  ],
} as const

const section = (tag: keyof typeof SECTIONS): string =>
  `<${tag} count="${SECTIONS[tag].length}">${SECTIONS[tag].join('')}</${tag}>`

const STYLES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="${SML}">` +
  section('fonts') +
  section('fills') +
  section('borders') +
  `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
  section('cellXfs') +
  `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
  section('dxfs') +
  `</styleSheet>`

async function workbook(): Promise<Buffer> {
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
    `<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="${SML}" xmlns:r="${R}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><calcPr xmlns:mc="${MC}" xmlns:hs="${HS}" mc:Ignorable="hs" hs:hclCalcId="904"/></workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
  )
  zip.file('xl/styles.xml', STYLES)
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${SML}"><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>`,
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

async function stylesOf(buffer: Buffer): Promise<string> {
  return (await (await JSZip.loadAsync(buffer)).file('xl/styles.xml')?.async('string')) ?? ''
}

describe('styles.xml with mc:AlternateContent entries across a save', () => {
  it('keeps styles.xml byte for byte when no style changes', async () => {
    const source = await workbook()
    const mutation = await applyCellEditsToXlsx(source, [
      { sheetName: 'Sheet1', row: 0, column: 1, writeValue: true, cell: { value: 3 } },
    ])
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    expect(await stylesOf(mutation.buffer)).toBe(STYLES)
  })

  it('changes only the touched parts and keeps every existing entry ahead of the new ones', async () => {
    const source = await workbook()
    const mutation = await applyCellEditsToXlsx(source, [
      {
        sheetName: 'Sheet1',
        row: 0,
        column: 0,
        writeValue: false,
        cell: { value: null },
        style: { italic: true },
      },
    ])
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    const styles = await stylesOf(mutation.buffer)
    for (const tag of Object.keys(SECTIONS) as (keyof typeof SECTIONS)[]) {
      const inner = new RegExp(`<${tag}\\b[^>]*>([\\s\\S]*?)</${tag}>`).exec(styles)?.[1] ?? ''
      expect(inner.startsWith(SECTIONS[tag].join('')), tag).toBe(true)
    }
    // one font and one xf appended after the wrapped ones; no dxf added
    expect(styles).toContain('<fonts count="3">')
    expect(styles).toContain('<cellXfs count="3">')
    expect(styles).toContain('<dxfs count="1">')
    // the new entries derive from the Fallback and carry no hs: markup
    expect(styles).toContain('<font><i/><b/><sz val="18"/><name val="Calibri"/></font></fonts>')
    expect(styles).toContain(
      '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>',
    )
  })
})
