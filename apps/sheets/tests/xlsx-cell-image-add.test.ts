import { describe, expect, it } from 'vitest'

import {
  assertOnlyTouchedEntriesChanged,
  assembleWithJsZip,
  createBufferEntrySource,
  planCellEditsToXlsx,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { SheetCellImageAddition } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import {
  applyCellImageAdditions,
  type CellImageAddition,
} from '@genoffice/xlsx-gateway/gateway/xlsx-cell-image-add'
import type { MutablePackage } from '@genoffice/xlsx-gateway/gateway/xlsx-drawing-add'
import { buildEditFixture } from './fixture-builder'

/// 1×1 PNG; the save copies the bytes, so the payload only has to decode.
const PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>`
const WORKBOOK_RELS = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`
const WORKSHEET = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="2"><c r="B2" s="4"><v>1</v></c></row></sheetData>
</worksheet>`

interface MemoryPackage extends MutablePackage {
  readonly files: Map<string, string | Uint8Array>
}

function memoryPackage(files: Record<string, string>): MemoryPackage {
  const entries = new Map<string, string | Uint8Array>(Object.entries(files))
  return {
    files: entries,
    paths: async () => [...entries.keys()],
    has: async (path) => entries.has(path),
    readText: async (path) => {
      const content = entries.get(path)
      if (typeof content !== 'string') throw new Error(`missing text part: ${path}`)
      return content
    },
    write: (path, content) => void entries.set(path, content),
    add: (path, content) => void entries.set(path, content),
    addBinary: (path, bytes) => void entries.set(path, bytes),
    remove: (path) => void entries.delete(path),
  }
}

function text(pkg: MemoryPackage, path: string): string {
  const content = pkg.files.get(path)
  if (typeof content !== 'string') throw new Error(`missing text part: ${path}`)
  return content
}

function addition(overrides: Partial<CellImageAddition> = {}): CellImageAddition {
  return {
    worksheetPath: 'xl/worksheets/sheet1.xml',
    row: 1,
    column: 1,
    mediaType: 'image/png',
    base64: PNG_BASE64,
    ...overrides,
  }
}

async function applyDirect(
  additions: readonly CellImageAddition[],
  files: Record<string, string> = {},
): Promise<MemoryPackage> {
  const pkg = memoryPackage({
    '[Content_Types].xml': CONTENT_TYPES,
    'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
    'xl/worksheets/sheet1.xml': WORKSHEET,
    ...files,
  })
  await applyCellImageAdditions(pkg, additions, new Set())
  return pkg
}

describe('cell image additions: fresh rich-value parts', () => {
  it('writes media, rich-value bookkeeping, and vm-marked cells', async () => {
    const pkg = await applyDirect([addition(), addition({ row: 6, column: 3 })])

    expect(pkg.files.get('xl/media/image1.png')).toBeInstanceOf(Uint8Array)
    expect(pkg.files.get('xl/media/image2.png')).toBeInstanceOf(Uint8Array)

    const metadata = text(pkg, 'xl/metadata.xml')
    expect(metadata).toContain(
      'xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"',
    )
    expect(metadata).toContain('<metadataTypes count="1">')
    expect(metadata).toContain('<metadataType name="XLRICHVALUE"')
    expect(metadata).toContain('<futureMetadata name="XLRICHVALUE" count="2">')
    expect(metadata).toContain('<xlrd:rvb i="0"/>')
    expect(metadata).toContain('<xlrd:rvb i="1"/>')
    expect(metadata).toContain('<valueMetadata count="2">')
    expect(metadata).toContain('<bk><rc t="1" v="0"/></bk><bk><rc t="1" v="1"/></bk>')

    const structures = text(pkg, 'xl/richData/rdrichvaluestructure.xml')
    expect(structures).toContain(
      '<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">',
    )
    expect(structures).toContain(
      '<s t="_localImage"><k n="_rvRel:LocalImageIdentifier" t="i"/><k n="CalcOrigin" t="i"/></s>',
    )

    const richValues = text(pkg, 'xl/richData/rdrichvalue.xml')
    expect(richValues).toContain('count="2"')
    expect(richValues).toContain('<rv s="0"><v>0</v><v>5</v></rv><rv s="0"><v>1</v><v>5</v></rv>')

    const rels = text(pkg, 'xl/richData/richValueRel.xml')
    expect(rels).toContain('<rel r:id="rId1"/><rel r:id="rId2"/>')
    const relParts = text(pkg, 'xl/richData/_rels/richValueRel.xml.rels')
    expect(relParts).toContain('Id="rId1"')
    expect(relParts).toContain('Target="../media/image1.png"')
    expect(relParts).toContain('Target="../media/image2.png"')

    const workbookRels = text(pkg, 'xl/_rels/workbook.xml.rels')
    expect(workbookRels).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"',
    )
    expect(workbookRels).toContain(
      'Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel" Target="richData/richValueRel.xml"',
    )
    expect(workbookRels).toContain(
      'Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue" Target="richData/rdrichvalue.xml"',
    )
    expect(workbookRels).toContain(
      'Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure" Target="richData/rdrichvaluestructure.xml"',
    )

    const contentTypes = text(pkg, '[Content_Types].xml')
    expect(contentTypes).toContain('<Default Extension="png" ContentType="image/png"/>')
    expect(contentTypes).toContain(
      '<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>',
    )
    expect(contentTypes).toContain(
      '<Override PartName="/xl/richData/richValueRel.xml" ContentType="application/vnd.ms-excel.richvaluerel+xml"/>',
    )
    expect(contentTypes).toContain(
      '<Override PartName="/xl/richData/rdrichvalue.xml" ContentType="application/vnd.ms-excel.rdrichvalue+xml"/>',
    )
    expect(contentTypes).toContain(
      '<Override PartName="/xl/richData/rdrichvaluestructure.xml" ContentType="application/vnd.ms-excel.rdrichvaluestructure+xml"/>',
    )

    // The styled host cell keeps its style; the missing row lands in order.
    const worksheet = text(pkg, 'xl/worksheets/sheet1.xml')
    expect(worksheet).toContain('<c r="B2" s="4" t="e" vm="1"><v>#VALUE!</v></c>')
    expect(worksheet).toContain('<row r="7"><c r="D7" t="e" vm="2"><v>#VALUE!</v></c></row>')
    expect(worksheet.indexOf('<row r="2"')).toBeLessThan(worksheet.indexOf('<row r="7"'))
  })

  it('skips parts that already exist and keeps earlier indexes stable', async () => {
    const existing = {
      'xl/metadata.xml':
        `${XML_HEADER}<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">` +
        '<metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000"/></metadataTypes>' +
        '<futureMetadata name="XLRICHVALUE" count="1"><bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="0"/></ext></extLst></bk></futureMetadata>' +
        '<valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata></metadata>',
      'xl/richData/rdrichvaluestructure.xml':
        '<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1"><s t="_localImage"><k n="_rvRel:LocalImageIdentifier" t="i"/><k n="CalcOrigin" t="i"/></s></rvStructures>',
      'xl/richData/rdrichvalue.xml':
        '<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1"><rv s="0"><v>0</v><v>5</v></rv></rvData>',
      'xl/richData/richValueRel.xml':
        '<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><rel r:id="rId1"/></richValueRels>',
      'xl/richData/_rels/richValueRel.xml.rels':
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>',
      // The picture behind rId1; its presence makes the new media image2.png.
      'xl/media/image1.png': 'existing-picture-bytes',
    }
    const pkg = await applyDirect([addition({ row: 2, column: 2 })], existing)

    const metadata = text(pkg, 'xl/metadata.xml')
    expect(metadata).toContain('<futureMetadata name="XLRICHVALUE" count="2">')
    expect(metadata).toContain(
      '<xlrd:rvb i="0"/></ext></extLst></bk><bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="1"/>',
    )
    expect(metadata).toContain('<valueMetadata count="2">')
    expect(metadata).toContain('<bk><rc t="1" v="0"/></bk><bk><rc t="1" v="1"/></bk>')
    expect(metadata.match(/<metadataType\b/g)).toHaveLength(1)

    expect(text(pkg, 'xl/richData/rdrichvalue.xml')).toContain(
      '<rv s="0"><v>0</v><v>5</v></rv><rv s="0"><v>1</v><v>5</v></rv>',
    )
    expect(text(pkg, 'xl/richData/richValueRel.xml')).toContain(
      '<rel r:id="rId1"/><rel r:id="rId2"/>',
    )
    expect(text(pkg, 'xl/richData/_rels/richValueRel.xml.rels')).toContain(
      'Target="../media/image2.png"',
    )
    expect(text(pkg, 'xl/worksheets/sheet1.xml')).toContain('<c r="C3" t="e" vm="2">')
    // The workbook relationships and content types were already complete.
    expect(text(pkg, 'xl/_rels/workbook.xml.rels')).toContain('Target="richData/rdrichvalue.xml"')
    expect(
      text(pkg, 'xl/_rels/workbook.xml.rels').match(/Target="richData\/rdrichvalue\.xml"/g),
    ).toHaveLength(1)
  })

  it('extends a dynamic-array metadata part without disturbing its record', async () => {
    const pkg = await applyDirect([addition()], {
      'xl/metadata.xml':
        `${XML_HEADER}<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">` +
        '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000"/></metadataTypes>' +
        '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>' +
        '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>',
    })

    const metadata = text(pkg, 'xl/metadata.xml')
    expect(metadata).toContain(
      'xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"',
    )
    expect(metadata).toContain('<metadataTypes count="2">')
    expect(metadata).toContain(
      '<metadataType name="XLDAPR" minSupportedVersion="120000"/><metadataType name="XLRICHVALUE"',
    )
    expect(metadata).toContain('<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>')
    expect(metadata).toContain('<futureMetadata name="XLRICHVALUE" count="1">')
    // XLRICHVALUE is the second metadata type, so the value record points at
    // t=2 while the dynamic-array cellMetadata keeps t=1.
    expect(metadata).toContain(
      '<valueMetadata count="1"><bk><rc t="2" v="0"/></bk></valueMetadata>',
    )
    expect(text(pkg, 'xl/worksheets/sheet1.xml')).toContain('<c r="B2" s="4" t="e" vm="1">')
  })
})

function imageAddition(overrides: Partial<SheetCellImageAddition> = {}): SheetCellImageAddition {
  return {
    sheetName: 'Data',
    row: 0,
    column: 0,
    mediaType: 'image/png',
    base64: PNG_BASE64,
    ...overrides,
  }
}

describe('cell image additions through the save plan', () => {
  it('lands the media, parts, and host cell in one plan', async () => {
    const buffer = await buildEditFixture()
    const source = await createBufferEntrySource(buffer)
    const plan = await planCellEditsToXlsx(
      source,
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
      [],
      [],
      [],
      [],
      [],
      [],
      null,
      null,
      [],
      [],
      [],
      [],
      [imageAddition({ row: 2, column: 1 })],
    )

    expect(plan.addedBinary.get('xl/media/image1.png')).toBeDefined()
    expect(plan.added.get('xl/metadata.xml')).toContain('XLRICHVALUE')
    expect(plan.added.get('xl/richData/rdrichvalue.xml')).toContain(
      '<rv s="0"><v>0</v><v>5</v></rv>',
    )
    expect(plan.replaced.get('xl/worksheets/sheet1.xml')).toContain(
      '<c r="B3" t="e" vm="1"><v>#VALUE!</v></c>',
    )
    expect(plan.replaced.get('xl/_rels/workbook.xml.rels')).toContain('Target="metadata.xml"')
    expect(plan.replaced.get('[Content_Types].xml')).toContain(
      'PartName="/xl/richData/richValueRel.xml"',
    )

    const mutation = await assembleWithJsZip(buffer, plan)
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
  })
})
