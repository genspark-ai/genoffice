import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import type { MutablePackage } from '../src/gateway/xlsx-drawing-add'
import {
  applySlicerState,
  parseSlicerState,
  SLICER_CACHE_REL_TYPE,
  SLICER_REL_TYPE,
  SlicerAddError,
  type SlicerPartState,
} from '../src/gateway/xlsx-slicer'
import {
  assembleWithJsZip,
  assertOnlyTouchedEntriesChanged,
  createBufferEntrySource,
  planCellEditsToXlsx,
  type SheetSlicerState,
} from '../src/gateway/xlsx-gateway'

const WORKSHEET = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>`

const WORKBOOK = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
  <pivotCaches><pivotCache cacheId="7" r:id="rId2"/></pivotCaches>
  <calcPr fullCalcOnLoad="1"/>
</workbook>`

const WORKBOOK_RELS = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
</Relationships>`

const PIVOT_TABLE = `<?xml version="1.0" encoding="UTF-8"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="Pivot1" cacheId="7"/>`

const PIVOT_CACHE = `<?xml version="1.0" encoding="UTF-8"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1">
  <cacheSource type="worksheet"><worksheetSource ref="A1:C4" sheet="Data"/></cacheSource>
  <cacheFields count="1"><cacheField name="Region" numFmtId="0"><sharedItems count="3"><s v="East"/><s v="West"/><s v="North"/></sharedItems></cacheField></cacheFields>
</pivotCacheDefinition>`

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`

const WORKBOOK_WITH_SELF_CLOSING_DEFINED_NAMES = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
  <definedNames/>
  <pivotCaches><pivotCache cacheId="7" r:id="rId2"/></pivotCaches>
  <calcPr fullCalcOnLoad="1"/>
</workbook>`

/// A dashboard shape: the pivot lives on Data (sheetId 1) while the control is
/// meant for Dash (sheetId 2).
const DASH_WORKBOOK = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Dash" sheetId="2" r:id="rId3"/></sheets>
  <pivotCaches><pivotCache cacheId="7" r:id="rId2"/></pivotCaches>
  <calcPr fullCalcOnLoad="1"/>
</workbook>`

const DASH_WORKBOOK_RELS = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`

const EXCEL_SLICER_PART = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x"><slicer name="Region" cache="Slicer_Region" caption="Region" columnCount="2" style="SlicerStyleOther1" lockedPosition="1" rowHeight="300000"><extLst><ext uri="{11111111-1111-1111-1111-111111111111}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicer name="Region"/></ext></extLst></slicer></slicers>`

const EXCEL_SLICER_CACHE_PART = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" name="Slicer_Region" sourceName="Region" xr10:uid="{33333333-3333-3333-3333-333333333333}"><pivotTables><pivotTable tabId="1" name="Pivot1"/></pivotTables><data><tabular pivotCacheId="7" sortOrder="descending" customListSort="0" crossFilter="showItemsWithNoData"><items count="2"><i x="0"/><i x="1" s="1"/></items></tabular></data><extLst><ext uri="{22222222-2222-2222-2222-222222222222}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerCacheDefinition name="Slicer_Region"/></ext></extLst></slicerCacheDefinition>`

function dashboardFixture(): MutablePackage {
  return inMemoryPackage({
    '[Content_Types].xml': CONTENT_TYPES,
    'xl/workbook.xml': DASH_WORKBOOK,
    'xl/_rels/workbook.xml.rels': DASH_WORKBOOK_RELS,
    'xl/worksheets/sheet1.xml': WORKSHEET,
    'xl/worksheets/sheet2.xml': WORKSHEET,
    'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="pRel1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>`,
    'xl/pivotTables/pivotTable1.xml': PIVOT_TABLE,
    'xl/pivotCache/pivotCacheDefinition1.xml': PIVOT_CACHE,
  })
}

function inMemoryPackage(entries: Record<string, string>): MutablePackage {
  const files = new Map(Object.entries(entries))
  return {
    paths: async () => [...files.keys()],
    has: async (path) => files.has(path),
    readText: async (path) => {
      const content = files.get(path)
      if (content === undefined) throw new Error(`missing part: ${path}`)
      return content
    },
    write: (path, content) => void files.set(path, content),
    add: (path, content) => void files.set(path, content),
    addBinary: (path, bytes) => void files.set(path, new TextDecoder().decode(bytes)),
    remove: (path) => void files.delete(path),
  }
}

function fixture(): MutablePackage {
  return inMemoryPackage({
    '[Content_Types].xml': CONTENT_TYPES,
    'xl/workbook.xml': WORKBOOK,
    'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
    'xl/worksheets/sheet1.xml': WORKSHEET,
    // the pivot part's host: slicerCache pivotTables/pivotTable binds through
    // the hosting worksheet's sheetId and the pivot table's name
    'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="pRel1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>`,
    'xl/pivotTables/pivotTable1.xml': PIVOT_TABLE,
    'xl/pivotCache/pivotCacheDefinition1.xml': PIVOT_CACHE,
  })
}

function readerOf(pkg: MutablePackage) {
  return { paths: pkg.paths, readText: pkg.readText }
}

const SLICER_INPUT = {
  worksheetPath: 'xl/worksheets/sheet1.xml',
  pivotPath: 'xl/pivotTables/pivotTable1.xml',
  sourceName: 'Region',
  caption: 'Region',
  items: [
    { x: 0, selected: true },
    { x: 1, selected: false },
    { x: 2, selected: true },
  ],
}

async function writeInitial(): Promise<{ pkg: MutablePackage; slicer: SlicerPartState }> {
  const pkg = fixture()
  const workbookXml = await applySlicerState(
    pkg,
    { slicers: [SLICER_INPUT], removals: [] },
    WORKBOOK,
    new Set(),
  )
  pkg.write('xl/workbook.xml', workbookXml)
  const [slicer] = await parseSlicerState(readerOf(pkg))
  if (slicer === undefined) throw new Error('controls did not restore')
  return { pkg, slicer }
}

describe('slicer write + read round trip', () => {
  it('writes every part with rels, exts, and content types, and reads them back', async () => {
    const { pkg, slicer } = await writeInitial()

    // Parts exist with the right root elements.
    const slicerXml = await pkg.readText('xl/slicers/slicer1.xml')
    expect(slicerXml).toContain('<slicer name="Region" cache="Slicer_Region" caption="Region"')
    const cacheXml = await pkg.readText('xl/slicerCaches/slicerCache1.xml')
    expect(cacheXml).toContain('name="Slicer_Region" sourceName="Region"')
    // the pivot binding: tabId is the hosting worksheet's sheetId, name the
    // pivot table's own name
    expect(cacheXml).toContain('<pivotTable tabId="1" name="Pivot1"/>')
    expect(cacheXml).toContain('<tabular pivotCacheId="7"')
    expect(cacheXml).toContain('<i x="0" s="1"/><i x="1"/><i x="2" s="1"/>')

    // The pivot cache definition carries its x14 slicer ext, and the two
    // pivotCacheIds are the same value — Excel drops the whole slicerCache
    // part ('Removed Part: slicerCache') when they disagree. The ext carries
    // no slicerData: that flag is OLAP-only and MUST be absent/false for a
    // tabular cache — slicerData="1" made Excel remove the slicer cache
    // ("Removed Feature: Slicer Cache").
    const pivotCache = await pkg.readText('xl/pivotCache/pivotCacheDefinition1.xml')
    expect(pivotCache).toContain('<x14:pivotCacheDefinition pivotCacheId="7"/>')
    expect(pivotCache).not.toContain('slicerData')
    expect(pivotCache).toContain('uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}"')

    // Worksheet rel + ext wiring.
    const sheetRels = await pkg.readText('xl/worksheets/_rels/sheet1.xml.rels')
    expect(sheetRels).toContain(`Type="${SLICER_REL_TYPE}" Target="../slicers/slicer1.xml"`)
    const worksheet = await pkg.readText('xl/worksheets/sheet1.xml')
    expect(worksheet).toContain('<x14:slicerList><x14:slicer r:id="rId1"/></x14:slicerList>')
    // the visible control: a drawing part with a graphicFrame, wired from the
    // worksheet and wrapped in an mc:AlternateContent with a down-level fallback
    expect(worksheet).toContain('<drawing r:id="rId2"/>')
    expect(sheetRels).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"',
    )
    const drawingPart = 'xl/drawings/drawing1.xml'
    expect(await pkg.has(drawingPart)).toBe(true)
    const drawing = await pkg.readText(drawingPart)
    expect(drawing).toContain(
      '<sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer" name="Region"/>',
    )
    expect(drawing).toContain(
      '<a:graphicData uri="http://schemas.microsoft.com/office/drawing/2010/slicer">',
    )
    expect(drawing).toContain('<mc:AlternateContent')
    expect(drawing).toContain('Requires="sle"')
    expect(drawing).toContain('<mc:Fallback><xdr:sp')
    // cNvPr ids are unique inside the drawing part
    const ids = [...drawing.matchAll(/<xdr:cNvPr\b[^>]*\bid="(\d+)"/g)].map((m) => m[1])
    expect(new Set(ids).size).toBe(ids.length)
    expect((drawing.match(/<xdr:twoCellAnchor\b/g) ?? []).length).toBe(1)
    // clientData is the anchor's child, after the AlternateContent; inside the
    // graphicFrame it is what Excel repairs ("Repaired Records: Drawing").
    expect(drawing).toContain('</mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>')
    expect(drawing).not.toContain('<xdr:clientData/></xdr:graphicFrame>')

    // Workbook rel + ext wiring, plus the hidden Slicer_<name> definedName.
    const workbookRels = await pkg.readText('xl/_rels/workbook.xml.rels')
    expect(workbookRels).toContain(
      `Type="${SLICER_CACHE_REL_TYPE}" Target="slicerCaches/slicerCache1.xml"`,
    )
    const workbook = await pkg.readText('xl/workbook.xml')
    expect(workbook).toContain(
      '<x14:slicerCaches><x14:slicerCache r:id="rId3"/></x14:slicerCaches>',
    )
    // Excel writes #N/A as the value: the name is a marker, not a range.
    expect(workbook).toContain('<definedName name="Slicer_Region" hidden="1">#N/A</definedName>')

    // Content types.
    const contentTypes = await pkg.readText('[Content_Types].xml')
    expect(contentTypes).toContain(
      '<Override PartName="/xl/slicers/slicer1.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>',
    )
    expect(contentTypes).toContain(
      '<Override PartName="/xl/slicerCaches/slicerCache1.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>',
    )
    expect(contentTypes).toContain(
      '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>',
    )

    // Read side restores identity, binding, and selection.
    expect(slicer).toMatchObject({
      sheetId: 'sheet-1',
      worksheetPath: 'xl/worksheets/sheet1.xml',
      partPath: 'xl/slicers/slicer1.xml',
      cachePath: 'xl/slicerCaches/slicerCache1.xml',
      name: 'Region',
      cacheName: 'Slicer_Region',
      sourceName: 'Region',
      caption: 'Region',
      pivotCacheDefinitionPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
      selectedItems: [0, 2],
    })
  })

  it('updates existing parts in place instead of allocating duplicates', async () => {
    const { pkg, slicer } = await writeInitial()
    pkg.write(
      'xl/workbook.xml',
      await applySlicerState(
        pkg,
        {
          slicers: [
            {
              worksheetPath: slicer.worksheetPath,
              pivotPath: 'xl/pivotTables/pivotTable1.xml',
              sourceName: 'Region',
              caption: 'Region',
              name: slicer.name,
              partPath: slicer.partPath,
              cachePath: slicer.cachePath,
              cacheName: slicer.cacheName,
              items: [
                { x: 0, selected: true },
                { x: 1, selected: true },
                { x: 2, selected: false },
              ],
            },
          ],
          removals: [],
        },
        await pkg.readText('xl/workbook.xml'),
        new Set(),
      ),
    )
    const paths = await pkg.paths()
    expect(paths.filter((path) => path.startsWith('xl/slicers/'))).toEqual([
      'xl/slicers/slicer1.xml',
    ])
    const [restoredSlicer] = await parseSlicerState(readerOf(pkg))
    expect(restoredSlicer?.selectedItems).toEqual([0, 1])
    // Names stay stable through the merge.
    expect(restoredSlicer?.name).toBe('Region')
    expect(restoredSlicer?.cacheName).toBe('Slicer_Region')
  })

  it('carries Excel-authored geometry through a part rewrite', async () => {
    const pkg = fixture()
    // A part Excel could have written: two-column slicer with a custom style.
    pkg.add(
      'xl/slicers/slicer1.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
        '<slicer name="Region" cache="Slicer_Region" caption="Region" columnCount="2" ' +
        'style="SlicerStyleOther1" rowHeight="241300"/></slicers>',
    )
    await applySlicerState(
      pkg,
      {
        slicers: [
          {
            ...SLICER_INPUT,
            name: 'Region',
            partPath: 'xl/slicers/slicer1.xml',
            cachePath: 'xl/slicerCaches/slicerCache1.xml',
            cacheName: 'Slicer_Region',
          },
        ],
        removals: [],
      },
      WORKBOOK,
      new Set(),
    )
    const xml = await pkg.readText('xl/slicers/slicer1.xml')
    expect(xml).toContain('columnCount="2"')
    expect(xml).toContain('style="SlicerStyleOther1"')
    expect(xml).toContain('rowHeight="241300"')
  })

  it('opens up a self-closing pivot cache definition for the x14 ext', async () => {
    const pkg = fixture()
    pkg.write(
      'xl/pivotCache/pivotCacheDefinition1.xml',
      '<?xml version="1.0" encoding="UTF-8"?>' +
        '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r:id="rId1" refreshOnLoad="1"/>',
    )
    await applySlicerState(pkg, { slicers: [SLICER_INPUT], removals: [] }, WORKBOOK, new Set())
    const xml = await pkg.readText('xl/pivotCache/pivotCacheDefinition1.xml')
    expect(xml).toContain('refreshOnLoad="1"')
    expect(xml).toContain('<x14:pivotCacheDefinition pivotCacheId="7"/>')
    expect(xml).not.toContain('slicerData')
    expect(xml).toContain('</pivotCacheDefinition>')
  })

  it('removes parts, rels, exts, frames, and content types without leaving orphans', async () => {
    const { pkg, slicer } = await writeInitial()
    pkg.write(
      'xl/workbook.xml',
      await applySlicerState(
        pkg,
        {
          slicers: [],
          removals: [{ partPath: slicer.partPath, cachePath: slicer.cachePath }],
        },
        await pkg.readText('xl/workbook.xml'),
        new Set(),
      ),
    )
    const paths = await pkg.paths()
    expect(paths.some((path) => /slicers?\//.test(path))).toBe(false)
    expect(paths).toContain('xl/worksheets/_rels/sheet1.xml.rels') // still exists, emptied of slicer rels
    const worksheet = await pkg.readText('xl/worksheets/sheet1.xml')
    expect(worksheet).not.toContain('slicerList')
    expect(worksheet).not.toContain('<extLst>')
    // the drawing's graphicFrame went with the control — no orphan sle:slicer
    const drawing = await pkg.readText('xl/drawings/drawing1.xml')
    expect(drawing).not.toContain('sle:slicer')
    expect(drawing).not.toContain('<xdr:twoCellAnchor')
    const workbook = await pkg.readText('xl/workbook.xml')
    expect(workbook).not.toContain('slicerCaches')
    expect(workbook).not.toContain('<extLst>')
    expect(workbook).not.toContain('Slicer_Region')
    const pivotCache = await pkg.readText('xl/pivotCache/pivotCacheDefinition1.xml')
    expect(pivotCache).not.toContain('{725AE2AE-9491-48be-B2B4-4EB974FC3084}')
    const contentTypes = await pkg.readText('[Content_Types].xml')
    expect(contentTypes).not.toContain('slicer')
    expect(await parseSlicerState(readerOf(pkg))).toEqual([])
  })

  it('keeps a foreign slicer untouched when the session knows nothing of it', async () => {
    const { pkg } = await writeInitial()
    // A second, Excel-authored slicer the session never restored.
    pkg.add(
      'xl/slicers/slicer2.xml',
      '<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
        '<slicer name="Product" cache="Slicer_Product"/></slicers>',
    )
    pkg.add(
      'xl/slicerCaches/slicerCache2.xml',
      '<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="Slicer_Product" sourceName="Product"><pivotTables><pivotTable tabId="1"/></pivotTables><data><olap><pivotTables><pivotTable name="Pivot1"/></pivotTables></olap></data></slicerCacheDefinition>',
    )
    // An empty state (no slicers, no removals) is a no-op for the foreign parts.
    await applySlicerState(
      pkg,
      { slicers: [], removals: [] },
      await pkg.readText('xl/workbook.xml'),
      new Set(),
    )
    expect(await pkg.has('xl/slicers/slicer2.xml')).toBe(true)
    expect(await pkg.has('xl/slicerCaches/slicerCache2.xml')).toBe(true)
    // OLAP caches are skipped by the reader (not representable in-session).
    const restored = await parseSlicerState(readerOf(pkg))
    expect(restored.map((entry) => entry.name)).toEqual(['Region'])
  })

  it('leaves timeline parts alone — their persistence is a separate piece of work', async () => {
    const pkg = fixture()
    pkg.add(
      'xl/timelines/timeline1.xml',
      '<timelines xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">' +
        '<timeline name="Date" cache="Timeline_Date" caption="Date" level="2" selectionLevel="2"/></timelines>',
    )
    await applySlicerState(pkg, { slicers: [SLICER_INPUT], removals: [] }, WORKBOOK, new Set())
    expect(await pkg.has('xl/timelines/timeline1.xml')).toBe(true)
    // and the reader does not surface it as a session control
    const restored = await parseSlicerState(readerOf(pkg))
    expect(restored.map((entry) => entry.name)).toEqual(['Region'])
  })

  it('rejects writing and removing the same part in one save', async () => {
    const { pkg, slicer } = await writeInitial()
    await expect(
      applySlicerState(
        pkg,
        {
          slicers: [{ ...SLICER_INPUT, partPath: slicer.partPath, cachePath: slicer.cachePath }],
          removals: [{ partPath: slicer.partPath, cachePath: slicer.cachePath }],
        },
        WORKBOOK,
        new Set(),
      ),
    ).rejects.toThrow(SlicerAddError)
  })

  it('fails closed when the bound pivot part is missing', async () => {
    const pkg = fixture()
    pkg.remove('xl/pivotTables/pivotTable1.xml')
    await expect(
      applySlicerState(pkg, { slicers: [SLICER_INPUT], removals: [] }, WORKBOOK, new Set()),
    ).rejects.toThrow(/missing from the workbook/)
  })
})

it('opens up a self-closing <definedNames/> instead of appending a second container', async () => {
  const pkg = fixture()
  const workbook = await applySlicerState(
    pkg,
    { slicers: [SLICER_INPUT], removals: [] },
    WORKBOOK_WITH_SELF_CLOSING_DEFINED_NAMES,
    new Set(),
  )
  // openpyxl writes <definedNames/>; a second container after </sheets> is a
  // schema violation Excel repairs
  expect(workbook.match(/<definedNames\b/g) ?? []).toHaveLength(1)
  expect(workbook.match(/<\/definedNames>/g) ?? []).toHaveLength(1)
  expect(workbook).not.toContain('<definedNames/>')
  expect(workbook).toContain(
    '<definedNames><definedName name="Slicer_Region" hidden="1">#N/A</definedName></definedNames>',
  )
  // schema position: before pivotCaches
  expect(workbook.indexOf('<definedNames>')).toBeLessThan(workbook.indexOf('<pivotCaches>'))
})

it('keeps one definedName per control across saves and refreshes its value', async () => {
  const { pkg, slicer } = await writeInitial()
  // a previous export wrote the slicer name as the value
  const stale = (await pkg.readText('xl/workbook.xml')).replace(
    '#N/A</definedName>',
    'Region</definedName>',
  )
  pkg.write('xl/workbook.xml', stale)
  const saved = await applySlicerState(
    pkg,
    {
      slicers: [
        {
          ...SLICER_INPUT,
          name: slicer.name,
          partPath: slicer.partPath,
          cachePath: slicer.cachePath,
          cacheName: slicer.cacheName,
        },
      ],
      removals: [],
    },
    stale,
    new Set(),
  )
  expect(saved.match(/<definedName\b/g) ?? []).toHaveLength(1)
  expect(saved).toContain('<definedName name="Slicer_Region" hidden="1">#N/A</definedName>')
})

it('preserves Excel-authored slicer part and slicerCache bits on rewrite', async () => {
  const pkg = fixture()
  pkg.add('xl/slicers/slicer1.xml', EXCEL_SLICER_PART)
  pkg.add('xl/slicerCaches/slicerCache1.xml', EXCEL_SLICER_CACHE_PART)
  await applySlicerState(
    pkg,
    {
      slicers: [
        {
          ...SLICER_INPUT,
          name: 'Region',
          partPath: 'xl/slicers/slicer1.xml',
          cachePath: 'xl/slicerCaches/slicerCache1.xml',
          cacheName: 'Slicer_Region',
        },
      ],
      removals: [],
    },
    WORKBOOK,
    new Set(),
  )
  const slicerXml = await pkg.readText('xl/slicers/slicer1.xml')
  expect(slicerXml).toContain(
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
  )
  expect(slicerXml).toContain('mc:Ignorable="x"')
  expect(slicerXml).toContain('columnCount="2"')
  expect(slicerXml).toContain('style="SlicerStyleOther1"')
  expect(slicerXml).toContain('lockedPosition="1"')
  expect(slicerXml).toContain('rowHeight="300000"')
  expect(slicerXml).toContain('{11111111-1111-1111-1111-111111111111}')
  const cacheXml = await pkg.readText('xl/slicerCaches/slicerCache1.xml')
  expect(cacheXml).toContain('xr10:uid="{33333333-3333-3333-3333-333333333333}"')
  expect(cacheXml).toContain('sortOrder="descending"')
  expect(cacheXml).toContain('customListSort="0"')
  expect(cacheXml).toContain('crossFilter="showItemsWithNoData"')
  expect(cacheXml).toContain('{22222222-2222-2222-2222-222222222222}')
  // the selection still comes from the session, not the file
  expect(cacheXml).toContain('<items count="3"><i x="0" s="1"/><i x="1"/><i x="2" s="1"/></items>')
  // a file crossFilter is not duplicated by the writer's default
  expect(cacheXml.match(/crossFilter=/g) ?? []).toHaveLength(1)
})

it('wires the control to the sheet that hosts it, not the pivot sheet', async () => {
  const pkg = dashboardFixture()
  const firstSave = await applySlicerState(
    pkg,
    { slicers: [{ ...SLICER_INPUT, worksheetPath: 'xl/worksheets/sheet2.xml' }], removals: [] },
    DASH_WORKBOOK,
    new Set(),
  )
  pkg.write('xl/workbook.xml', firstSave)
  // the binding's tabId is the PivotTable view's sheet (Data), not the
  // control's sheet
  const cacheXml = await pkg.readText('xl/slicerCaches/slicerCache1.xml')
  expect(cacheXml).toContain('<pivotTable tabId="1" name="Pivot1"/>')
  // the control rides Dash: rel + ext + drawing; Data keeps none of it
  const dashRels = await pkg.readText('xl/worksheets/_rels/sheet2.xml.rels')
  expect(dashRels).toContain(`Type="${SLICER_REL_TYPE}"`)
  const dash = await pkg.readText('xl/worksheets/sheet2.xml')
  expect(dash).toContain('slicerList')
  expect(dash).toContain('<drawing ')
  const data = await pkg.readText('xl/worksheets/sheet1.xml')
  expect(data).not.toContain('slicerList')
  expect(data).not.toContain('<drawing ')
  const dataRels = (await pkg.has('xl/worksheets/_rels/sheet1.xml.rels'))
    ? await pkg.readText('xl/worksheets/_rels/sheet1.xml.rels')
    : ''
  expect(dataRels).not.toContain(SLICER_REL_TYPE)
  // a restore-then-save keeps it there and does not duplicate anything
  const [restored] = await parseSlicerState(readerOf(pkg))
  expect(restored).toMatchObject({
    sheetId: 'sheet-2',
    worksheetPath: 'xl/worksheets/sheet2.xml',
  })
  const secondSave = await applySlicerState(
    pkg,
    {
      slicers: [
        {
          ...SLICER_INPUT,
          worksheetPath: restored!.worksheetPath,
          name: restored!.name,
          partPath: restored!.partPath,
          cachePath: restored!.cachePath,
          cacheName: restored!.cacheName,
        },
      ],
      removals: [],
    },
    await pkg.readText('xl/workbook.xml'),
    new Set(),
  )
  pkg.write('xl/workbook.xml', secondSave)
  const dashAgain = await pkg.readText('xl/worksheets/sheet2.xml')
  expect(dashAgain.match(/<x14:slicerList>/g) ?? []).toHaveLength(1)
  expect(dashAgain.match(/<drawing /g) ?? []).toHaveLength(1)
  const drawing = await pkg.readText('xl/drawings/drawing1.xml')
  expect(drawing.match(/<xdr:twoCellAnchor\b/g) ?? []).toHaveLength(1)
  expect((await pkg.readText('xl/worksheets/sheet1.xml')).includes('slicerList')).toBe(false)
})

describe('slicer state through the save planner', () => {
  async function plannerFixture(): Promise<Buffer> {
    const zip = new JSZip()
    zip.file('[Content_Types].xml', CONTENT_TYPES)
    zip.file('xl/workbook.xml', WORKBOOK)
    zip.file('xl/_rels/workbook.xml.rels', WORKBOOK_RELS)
    zip.file('xl/worksheets/sheet1.xml', WORKSHEET)
    // the pivot part's host, for the slicerCache binding
    zip.file(
      'xl/worksheets/_rels/sheet1.xml.rels',
      `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="pRel1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>`,
    )
    zip.file('xl/pivotTables/pivotTable1.xml', PIVOT_TABLE)
    zip.file('xl/pivotCache/pivotCacheDefinition1.xml', PIVOT_CACHE)
    return zip.generateAsync({ type: 'nodebuffer' })
  }

  it('round-trips through planCellEditsToXlsx with zero-diff when absent', async () => {
    const buffer = await plannerFixture()

    // No slicer state: the plan must not touch a single entry.
    const cleanPlan = await planCellEditsToXlsx(await createBufferEntrySource(buffer), [])
    expect(cleanPlan.touchedEntries).toEqual([])
    const cleanMutation = await assembleWithJsZip(buffer, cleanPlan)
    expect(() => assertOnlyTouchedEntriesChanged(cleanMutation)).not.toThrow()

    const state: SheetSlicerState = {
      slicers: [{ sheetName: 'Data', ...SLICER_INPUT }],
      removals: [],
    }
    const plan = await planCellEditsToXlsx(
      await createBufferEntrySource(buffer),
      [], // edits
      [], // structuralOps
      [], // chartEdits
      undefined, // sheetPlan
      [], // filterStates
      [], // hyperlinkEdits
      [], // cfStates
      [], // dvStates
      [], // sheetProtections
      null, // definedNamesState
      [], // visualAdditions
      [], // pageSetupStates
      [], // noteStates
      [], // tableAdditions
      [], // pivotAdditions
      [], // pivotCacheRefreshPaths
      [], // pivotRefreshUpdates
      [], // visualEdits
      [], // sparklineAdditions
      [], // formulaValues
      null, // themeState
      null, // workbookProtectionState
      [], // protectedRangeStates
      [], // bulkConstantFills
      [], // tabColorStates
      [], // tableEdits
      [], // cellImageAdditions
      [], // redactionStates
      state,
    )
    const mutation = await assembleWithJsZip(buffer, plan)
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    for (const part of [
      'xl/slicers/slicer1.xml',
      'xl/slicerCaches/slicerCache1.xml',
      'xl/worksheets/_rels/sheet1.xml.rels',
      '[Content_Types].xml',
      'xl/workbook.xml',
      'xl/worksheets/sheet1.xml',
      'xl/pivotCache/pivotCacheDefinition1.xml',
    ]) {
      expect(plan.touchedEntries).toContain(part)
    }

    // The saved archive restores the same session state through the reader.
    const restored = await parseSlicerState(await createBufferEntrySource(mutation.buffer))
    expect(restored).toHaveLength(1)
    expect(restored[0]).toMatchObject({
      sheetId: 'sheet-1',
      pivotCacheDefinitionPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
      selectedItems: [0, 2],
    })

    // Every part/rel/ext/drawing of the binding chain survives the planner
    // and a re-read of the saved package (what Excel needs to keep a slicer).
    const saved = await JSZip.loadAsync(mutation.buffer)
    const savedWorkbook = await saved.file('xl/workbook.xml')!.async('string')
    expect(savedWorkbook).toContain(
      '<x14:slicerCaches><x14:slicerCache r:id="rId3"/></x14:slicerCaches>',
    )
    expect(savedWorkbook.match(/<definedNames\b/g) ?? []).toHaveLength(1)
    expect(savedWorkbook).toContain(
      '<definedName name="Slicer_Region" hidden="1">#N/A</definedName>',
    )
    expect(savedWorkbook.indexOf('<definedNames>')).toBeLessThan(
      savedWorkbook.indexOf('<pivotCaches>'),
    )
    const savedWorksheet = await saved.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(savedWorksheet).toContain('<x14:slicerList><x14:slicer r:id="rId1"/></x14:slicerList>')
    expect(savedWorksheet).toContain('<drawing r:id="rId2"/>')
    const savedSheetRels = await saved.file('xl/worksheets/_rels/sheet1.xml.rels')!.async('string')
    expect(savedSheetRels).toContain('Target="../slicers/slicer1.xml"')
    const savedSlicer = await saved.file('xl/slicers/slicer1.xml')!.async('string')
    expect(savedSlicer).toContain('rowHeight="241300"')
    const savedDrawing = await saved.file('xl/drawings/drawing1.xml')!.async('string')
    expect(savedDrawing).toContain('</mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>')
    expect(savedDrawing).not.toContain('<xdr:clientData/></xdr:graphicFrame>')
    const savedCache = await saved.file('xl/slicerCaches/slicerCache1.xml')!.async('string')
    expect(savedCache).toContain('<tabular pivotCacheId="7"')
    const savedPivotCache = await saved
      .file('xl/pivotCache/pivotCacheDefinition1.xml')!
      .async('string')
    expect(savedPivotCache).toContain('<x14:pivotCacheDefinition pivotCacheId="7"/>')
    expect(savedPivotCache).not.toContain('slicerData')
    const savedContentTypes = await saved.file('[Content_Types].xml')!.async('string')
    expect(savedContentTypes).toContain('application/vnd.ms-excel.slicer+xml')
    expect(savedContentTypes).toContain('application/vnd.ms-excel.slicerCache+xml')
  })
})
