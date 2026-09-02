import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import JSZip from 'jszip'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'

import { saveWorkbookViaSidecar } from '../src/gateway/xlsx-package-io'
import { XlsxSidecarClient } from '../src/main/xlsx-sidecar-client'
import { workbookFileSchema } from '../src/shared/desktop-api'

const OLE_PARTS = [
  'xl/worksheets/sheet1.xml',
  'xl/worksheets/_rels/sheet1.xml.rels',
  'xl/embeddings/oleObject1.bin',
  'xl/drawings/vmlDrawing1.vml',
  'xl/drawings/_rels/vmlDrawing1.vml.rels',
  'xl/media/image1.emf',
]

/// Excel 2010+ shape of an embedded Word document: x14 oleObject with
/// objectPr anchor + cached EMF preview, a legacy VML shape, the embedding.
async function buildOleFixture(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
  <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
  <Default Extension="emf" ContentType="image/x-emf"/>
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
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
  <legacyDrawing r:id="rId2"/>
  <oleObjects><mc:AlternateContent><mc:Choice Requires="x14"><oleObject progId="Word.Document.12" shapeId="1025" r:id="rId3"><objectPr defaultSize="0" r:id="rId4"><anchor moveWithCells="1"><from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></from><to><xdr:col>4</xdr:col><xdr:colOff>384810</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>156210</xdr:rowOff></to></anchor></objectPr></oleObject></mc:Choice><mc:Fallback><oleObject progId="Word.Document.12" shapeId="1025" r:id="rId3"/></mc:Fallback></mc:AlternateContent></oleObjects>
</worksheet>`,
  )
  zip.file(
    'xl/worksheets/_rels/sheet1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.emf"/>
</Relationships>`,
  )
  zip.file(
    'xl/drawings/vmlDrawing1.vml',
    `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" filled="f" stroked="f"/><v:shape id="_x0000_s1025" type="#_x0000_t75" style="position:absolute" filled="t" fillcolor="window [65]" stroked="t" strokecolor="windowText [64]"><v:imagedata o:relid="rId1" o:title=""/><x:ClientData ObjectType="Pict"><x:Anchor>1, 0, 1, 0, 4, 40, 3, 16</x:Anchor></x:ClientData></v:shape></xml>`,
  )
  zip.file(
    'xl/drawings/_rels/vmlDrawing1.vml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.emf"/>
</Relationships>`,
  )
  // Opaque payloads: only their byte identity matters here.
  zip.file(
    'xl/embeddings/oleObject1.bin',
    Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 7, 7, 7]),
  )
  zip.file('xl/media/image1.emf', Buffer.from([1, 0, 0, 0, 0x6c, 0, 0, 0, 9, 9, 9]))
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

describe('embedded OLE objects', () => {
  let directory: string
  let client: XlsxSidecarClient

  beforeAll(async () => {
    directory = await mkdtemp(join(tmpdir(), 'xlsx-ole-roundtrip-'))
    client = new XlsxSidecarClient(sidecarBinaryPath())
  })

  afterAll(async () => {
    client.stop()
    await rm(directory, { recursive: true, force: true })
  })

  it('surfaces the object as a read-only ole visual with its frame and preview', async () => {
    const sourcePath = join(directory, 'ole.xlsx')
    await writeFile(sourcePath, await buildOleFixture())
    // The sidecar's open payload is the file schema minus the main-process
    // additions (sha256/readOnly); the visuals array is what matters here.
    const opened = workbookFileSchema
      .pick({ sessionId: true, visuals: true })
      .loose()
      .parse(await client.open(sourcePath))
    try {
      expect(opened.visuals).toHaveLength(1)
      const [ole] = opened.visuals
      expect(ole?.kind).toBe('ole')
      expect(ole?.progId).toBe('Word.Document.12')
      expect(ole?.mediaPath).toBe('xl/media/image1.emf')
      expect(ole?.mediaType).toBe('image/x-emf')
      expect(ole?.lineColor).toBe('#000000')
      expect(ole?.fillColor).toBe('#FFFFFF')
      expect(ole?.anchor.fromColumn).toBe(1)
      expect(ole?.anchor.toColumn).toBe(4)
      // No drawing locator: the visuals editor cannot rewrite it as a picture.
      expect(ole?.drawingPath).toBeUndefined()
      expect(ole?.drawingIndex).toBeUndefined()
    } finally {
      await client.close(opened.sessionId)
    }
  })

  it('keeps oleObjects, the embedding, the VML and the preview byte-identical on save', async () => {
    const sourcePath = join(directory, 'ole-source.xlsx')
    const targetPath = join(directory, 'ole-saved.xlsx')
    const sourceBuffer = await buildOleFixture()
    await writeFile(sourcePath, sourceBuffer)

    const result = await saveWorkbookViaSidecar({
      client,
      sourcePath,
      targetPath,
      edits: [{ sheetName: 'Data', row: 0, column: 0, writeValue: true, cell: { value: 2 } }],
    })
    expect(result.removedEntries).toEqual([])
    expect(result.addedEntries).toEqual([])
    expect(result.touchedEntries).not.toContain('xl/embeddings/oleObject1.bin')
    expect(result.touchedEntries).not.toContain('xl/drawings/vmlDrawing1.vml')
    expect(result.touchedEntries).not.toContain('xl/media/image1.emf')

    const sourceZip = await JSZip.loadAsync(sourceBuffer)
    const savedZip = await JSZip.loadAsync(await readFile(targetPath))
    expect(Object.keys(savedZip.files).sort()).toEqual(Object.keys(sourceZip.files).sort())
    for (const path of OLE_PARTS) {
      const [sourceBytes, savedBytes] = await Promise.all([
        sourceZip.file(path)?.async('nodebuffer'),
        savedZip.file(path)?.async('nodebuffer'),
      ])
      if (path === 'xl/worksheets/sheet1.xml') {
        // The edited sheet is rewritten, but its <oleObjects> block is not.
        const oleBlock = /<oleObjects>.*<\/oleObjects>/s
        expect(savedBytes?.toString('utf8').match(oleBlock)?.[0]).toBe(
          sourceBytes?.toString('utf8').match(oleBlock)?.[0],
        )
        expect(savedBytes?.toString('utf8')).toContain('<legacyDrawing r:id="rId2"/>')
        continue
      }
      expect(savedBytes?.equals(sourceBytes ?? Buffer.of()), path).toBe(true)
    }
  })

  it('moves the saved oleObjects and VML anchors with inserted rows and columns', async () => {
    const sourcePath = join(directory, 'ole-shift-source.xlsx')
    const targetPath = join(directory, 'ole-shift-saved.xlsx')
    await writeFile(sourcePath, await buildOleFixture())

    const result = await saveWorkbookViaSidecar({
      client,
      sourcePath,
      targetPath,
      edits: [],
      structuralOps: [
        {
          sheetName: 'Data',
          ops: [
            { kind: 'insert-rows', index: 0, count: 2 },
            { kind: 'insert-cols', index: 2, count: 1 },
          ],
        },
      ],
    })
    expect(result.touchedEntries).toContain('xl/drawings/vmlDrawing1.vml')
    expect(result.touchedEntries).not.toContain('xl/embeddings/oleObject1.bin')
    expect(result.touchedEntries).not.toContain('xl/media/image1.emf')

    const savedZip = await JSZip.loadAsync(await readFile(targetPath))
    const sheet = await savedZip.file('xl/worksheets/sheet1.xml')?.async('text')
    // Rows 1..3 -> 3..5; the from column (1) sits before the inserted column
    // 2 and stays, the to column (4) moves to 5. Offsets are kept.
    expect(sheet).toContain(
      '<from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></from>',
    )
    expect(sheet).toContain(
      '<to><xdr:col>5</xdr:col><xdr:colOff>384810</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>156210</xdr:rowOff></to>',
    )
    // The embedding relationship and the fallback element are untouched.
    expect(sheet).toContain(
      '<mc:Fallback><oleObject progId="Word.Document.12" shapeId="1025" r:id="rId3"/></mc:Fallback>',
    )
    const vml = await savedZip.file('xl/drawings/vmlDrawing1.vml')?.async('text')
    expect(vml).toContain('<x:Anchor>1, 0, 3, 0, 5, 40, 5, 16</x:Anchor>')
  })
})

function sidecarBinaryPath(): string {
  const executable = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
  return fileURLToPath(
    new URL(`../native/xlsx-engine/target/release/${executable}`, import.meta.url),
  )
}
