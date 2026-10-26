import { describe, expect, it } from 'vitest'

import { parsePivotDefinition } from '@genoffice/xlsx-gateway/gateway/xlsx-pivot'
import { restoreSlicerPanels } from '../src/renderer/pivot-actions'
import type { LazyWorkbookState } from '../src/renderer/univer-state'
import type { WorkbookSlicerPart } from '../src/shared/desktop-api'

/// Region (row field: East/West/South) + Date (ungrouped date field) + Sales,
/// mirroring the shape slicers/timelines bind to in real files.
const PIVOT_XML =
  '<pivotTableDefinition name="PivotTable1" cacheId="1">' +
  '<location ref="E1:F5" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>' +
  '<pivotFields count="3">' +
  '<pivotField axis="axisRow" showAll="0"><items count="4">' +
  '<item x="0"/><item x="1"/><item x="2"/><item t="default"/></items></pivotField>' +
  '<pivotField axis="axisRow" showAll="0"><items count="4">' +
  '<item x="0"/><item x="1"/><item x="2"/><item t="default"/></items></pivotField>' +
  '<pivotField dataField="1" showAll="0"/>' +
  '</pivotFields>' +
  '<rowFields count="2"><field x="0"/><field x="1"/></rowFields>' +
  '<dataFields count="1"><dataField name="Sum of Sales" fld="2"/></dataFields>' +
  '</pivotTableDefinition>'

const CACHE_XML =
  '<pivotCacheDefinition>' +
  '<cacheSource type="worksheet"><worksheetSource ref="A1:C7" sheet="Data"/></cacheSource>' +
  '<cacheFields count="3">' +
  '<cacheField name="Region"><sharedItems count="3">' +
  '<s v="East"/><s v="West"/><s v="South"/></sharedItems></cacheField>' +
  '<cacheField name="Date"><sharedItems count="3">' +
  '<s v="2021-01-15"/><s v="2021-03-02"/><s v="2021-06-20"/></sharedItems></cacheField>' +
  '<cacheField name="Sales"><sharedItems containsString="0" containsNumber="1"/></cacheField>' +
  '</cacheFields>' +
  '</pivotCacheDefinition>'

const PIVOT_PATH = 'xl/pivotTables/pivotTable1.xml'
const CACHE_DEF_PATH = 'xl/pivotCache/pivotCacheDefinition1.xml'

function state(): LazyWorkbookState {
  return {
    file: {
      sheets: [
        {
          id: 'sheet-1',
          pivotTables: [{ path: PIVOT_PATH, cachePath: CACHE_DEF_PATH, outputRef: 'E1:F5' }],
        },
      ],
    },
    pivotDefinitions: new Map([[PIVOT_PATH, parsePivotDefinition(PIVOT_XML, CACHE_XML)]]),
  } as unknown as LazyWorkbookState
}

function slicerEntry(overrides: Partial<WorkbookSlicerPart> = {}): WorkbookSlicerPart {
  return {
    kind: 'slicer',
    sheetId: 'sheet-1',
    worksheetPath: 'xl/worksheets/sheet1.xml',
    partPath: 'xl/slicers/slicer1.xml',
    cachePath: 'xl/slicerCaches/slicerCache1.xml',
    name: 'Region',
    cacheName: 'Slicer_Region',
    sourceName: 'Region',
    caption: 'Region',
    pivotCacheDefinitionPath: CACHE_DEF_PATH,
    selectedItems: [0, 2],
    range: null,
    ...overrides,
  }
}

function timelineEntry(overrides: Partial<WorkbookSlicerPart> = {}): WorkbookSlicerPart {
  return {
    ...slicerEntry(),
    kind: 'timeline',
    partPath: 'xl/timelines/timeline1.xml',
    cachePath: 'xl/timelineCaches/timelineCache1.xml',
    name: 'Date',
    cacheName: 'Timeline_Date',
    sourceName: 'Date',
    caption: 'Date',
    selectedItems: null,
    range: { startMonth: 2021 * 12 + 0, endMonth: 2021 * 12 + 2 },
    ...overrides,
  }
}

describe('restoreSlicerPanels', () => {
  it('rebuilds a slicer panel with its file identity and selection', () => {
    const restored = restoreSlicerPanels(state(), [slicerEntry()])
    expect(restored.timelines).toEqual([])
    expect(restored.slicers).toHaveLength(1)
    const slicer = restored.slicers[0]!
    expect(slicer.sheetId).toBe('sheet-1')
    expect(slicer.pivotPath).toBe(PIVOT_PATH)
    expect(slicer.field).toBe(0)
    expect(slicer.fieldName).toBe('Region')
    expect(slicer.members.map((member) => member.label)).toEqual(['East', 'West', 'South'])
    // selectedItems are sharedItems indices; members carry both indices.
    expect(slicer.members.map((member) => member.x)).toEqual([0, 1, 2])
    expect(slicer.selected).toEqual([0, 2])
    expect(slicer.name).toBe('Region')
    expect(slicer.partPath).toBe('xl/slicers/slicer1.xml')
    expect(slicer.cachePath).toBe('xl/slicerCaches/slicerCache1.xml')
    expect(slicer.cacheName).toBe('Slicer_Region')
  })

  it('treats a cache without an item list as unfiltered', () => {
    const restored = restoreSlicerPanels(state(), [slicerEntry({ selectedItems: null })])
    expect(restored.slicers[0]!.selected).toEqual([0, 1, 2])
  })

  it('never restores a fully-deselected slicer (defensive all-selected)', () => {
    const restored = restoreSlicerPanels(state(), [slicerEntry({ selectedItems: [] })])
    expect(restored.slicers[0]!.selected).toEqual([0, 1, 2])
  })

  it('rebuilds a timeline panel with its month range and identity', () => {
    const restored = restoreSlicerPanels(state(), [timelineEntry()])
    expect(restored.slicers).toEqual([])
    expect(restored.timelines).toHaveLength(1)
    const timeline = restored.timelines[0]!
    expect(timeline.field).toBe(1)
    expect(timeline.fieldName).toBe('Date')
    expect(timeline.minMonth).toBe(2021 * 12 + 0)
    expect(timeline.maxMonth).toBe(2021 * 12 + 5)
    expect(timeline.range).toEqual({ start: 2021 * 12 + 0, end: 2021 * 12 + 2 })
    expect(timeline.partPath).toBe('xl/timelines/timeline1.xml')
    expect(timeline.cacheName).toBe('Timeline_Date')
  })

  it('restores an unfiltered timeline as range null', () => {
    const restored = restoreSlicerPanels(state(), [timelineEntry({ range: null })])
    expect(restored.timelines[0]!.range).toBeNull()
  })

  it('skips entries it cannot bind (unknown field, missing definition, unbound cache)', () => {
    const workbook = state()
    const unknownField = restoreSlicerPanels(workbook, [
      slicerEntry({ sourceName: 'Nope' }),
      timelineEntry({ sourceName: 'Nope' }),
    ])
    expect(unknownField).toEqual({ slicers: [], timelines: [] })

    const noDefinition = state()
    noDefinition.pivotDefinitions.delete(PIVOT_PATH)
    expect(restoreSlicerPanels(noDefinition, [slicerEntry()])).toEqual({
      slicers: [],
      timelines: [],
    })

    expect(restoreSlicerPanels(state(), [slicerEntry({ pivotCacheDefinitionPath: null })])).toEqual(
      { slicers: [], timelines: [] },
    )
  })

  it('skips a timeline whose field no longer parses as dates', () => {
    // Field 0 (Region) is not a date field: the domain is null, no panel.
    const restored = restoreSlicerPanels(state(), [
      timelineEntry({ sourceName: 'Region', cacheName: 'Timeline_Region' }),
    ])
    expect(restored.timelines).toEqual([])
  })
})
