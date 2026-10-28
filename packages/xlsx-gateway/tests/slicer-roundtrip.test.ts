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
    // part ('Removed Part: slicerCache') when they disagree.
    const pivotCache = await pkg.readText('xl/pivotCache/pivotCacheDefinition1.xml')
    expect(pivotCache).toContain('<x14:pivotCacheDefinition pivotCacheId="7" slicerData="1"/>')
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

    // Workbook rel + ext wiring, plus the hidden Slicer_<name> definedName.
    const workbookRels = await pkg.readText('xl/_rels/workbook.xml.rels')
    expect(workbookRels).toContain(
      `Type="${SLICER_CACHE_REL_TYPE}" Target="slicerCaches/slicerCache1.xml"`,
    )
    const workbook = await pkg.readText('xl/workbook.xml')
    expect(workbook).toContain(
      '<x14:slicerCaches><x14:slicerCache r:id="rId3"/></x14:slicerCaches>',
    )
    expect(workbook).toContain('<definedName name="Slicer_Region" hidden="1">Region</definedName>')

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
    expect(xml).toContain('<x14:pivotCacheDefinition pivotCacheId="7" slicerData="1"/>')
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
    expect(pivotCache).not.toContain('slicerData')
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
  })
})
