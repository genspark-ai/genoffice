/**
 * PivotChart source wiring: a chart created from a pivot output area saves
 * with chartSpace <c:pivotSource> so Excel recognizes it as a PivotChart
 * (field buttons, layout sync) instead of a regular chart over the same
 * cells. Structure is pinned against real Excel-produced PivotChart parts —
 * name + fmtId (there is no fmtType member in CT_PivotSource), pivotSource
 * before c:chart, pivotFmts as c:chart's first child — and against the
 * CT_ChartSpace sequence in the ECMA-376 dml-chart schema shipped at
 * tools/ooxml-validate/schemas/dml-chart.xsd.
 */
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'

import { applyChartEdit } from '@genoffice/xlsx-gateway/gateway/xlsx-chart'
import {
  assembleWithJsZip,
  createBufferEntrySource,
  planCellEditsToXlsx,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { SheetVisualAddition } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { workbookFileSchema } from '../src/shared/desktop-api'
import { XlsxSidecarClient } from '../src/main/xlsx-sidecar-client'
import { buildEditFixture } from './fixture-builder'

const ANCHOR = {
  fromRow: 1,
  fromColumn: 5,
  fromRowOffset: 0,
  fromColumnOffset: 0,
  toRow: 16,
  toColumn: 12,
  toRowOffset: 0,
  toColumnOffset: 0,
}

const PIVOT_TABLE_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
  ' name="PivotTable1" cacheId="1" dataCaption="Values" updatedVersion="8">' +
  '<location ref="A3:B6" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>' +
  '<pivotFields count="2"><pivotField showAll="0"/><pivotField dataField="1" showAll="0"/></pivotFields>' +
  '<rowFields count="1"><field x="0"/></rowFields>' +
  '<rowItems count="2"><i><x v="0"/></i><i t="grand"><x/></i></rowItems>' +
  '<colItems count="1"><i/></colItems>' +
  '<dataFields count="1"><dataField name="Sum of Value" fld="1" baseField="0" baseItem="0"/></dataFields>' +
  '</pivotTableDefinition>'

function chartAddition(
  overrides: Partial<NonNullable<SheetVisualAddition['chart']>> = {},
  sheetName = 'Data',
): SheetVisualAddition {
  return {
    sheetName,
    anchor: ANCHOR,
    chart: {
      chartType: 'column',
      title: 'Users by Region',
      series: [
        {
          name: 'Users',
          categories: ['US', 'KR', 'BR'],
          values: [8, 6, 3],
          valuesRef: `'${sheetName}'!$B$2:$B$4`,
          categoriesRef: `'${sheetName}'!$A$2:$A$4`,
        },
      ],
      pivotTablePath: 'xl/pivotTables/pivotTable1.xml',
      ...overrides,
    },
  }
}

/// buildEditFixture + a pivotTable part (and its content-type override), the
/// minimal package state a PivotChart insertion saves against.
async function buildPivotFixture(sheetName = 'Data'): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await buildEditFixture())
  if (sheetName !== 'Data') {
    const workbook = await zip.file('xl/workbook.xml')!.async('string')
    zip.file(
      'xl/workbook.xml',
      workbook.replace('name="Data"', `name="${sheetName.replace(/&/g, '&amp;')}"`),
    )
  }
  zip.file('xl/pivotTables/pivotTable1.xml', PIVOT_TABLE_XML)
  const contentTypes = await zip.file('[Content_Types].xml')!.async('string')
  zip.file(
    '[Content_Types].xml',
    contentTypes.replace(
      '</Types>',
      '<Override PartName="/xl/pivotTables/pivotTable1.xml" ' +
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/></Types>',
    ),
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

async function planWith(additions: SheetVisualAddition[], buffer?: Buffer) {
  const source = await createBufferEntrySource(buffer ?? (await buildPivotFixture()))
  return planCellEditsToXlsx(source, [], [], [], undefined, [], [], [], [], [], null, additions)
}

/// CT_ChartSpace child order per the ECMA-376 dml-chart schema
/// (tools/ooxml-validate/schemas/dml-chart.xsd). pivotSource sits between
/// clrMapOvr and protection; chart is the only required member.
const CHART_SPACE_ORDER = [
  'date1904',
  'lang',
  'roundedCorners',
  'style',
  'clrMapOvr',
  'pivotSource',
  'protection',
  'chart',
  'spPr',
  'txPr',
  'externalData',
  'printSettings',
  'userShapes',
  'extLst',
] as const

/// Direct children of the first `c:chartSpace` element, in document order
/// (depth-tracked tag scan; the chart parts carry no text content).
function chartSpaceChildren(xml: string): string[] {
  const open = /<c:chartSpace[> ]/.exec(xml)
  if (!open) throw new Error('no chartSpace element')
  const names: string[] = []
  let depth = 0
  for (const tag of xml.slice(open.index).matchAll(/<(\/)?([\w.]+:)?([\w.-]+)([^>?]*)>/g)) {
    const name = tag[3] ?? ''
    if (name === 'chartSpace' && !tag[1]) {
      depth = 1
      continue
    }
    if (tag[1]) {
      depth -= 1
      continue
    }
    if (depth === 1) names.push(name)
    if (!(tag[4] ?? '').trimEnd().endsWith('/')) depth += 1
  }
  return names
}

/// The children sequence must be a subsequence of the schema order (unknown
/// elements are not tolerated — every member is listed).
function assertSchemaOrder(children: readonly string[]): void {
  let at = -1
  for (const name of children) {
    const next = CHART_SPACE_ORDER.indexOf(name as (typeof CHART_SPACE_ORDER)[number])
    expect(next, `<${name}> is not a CT_ChartSpace member`).toBeGreaterThanOrEqual(0)
    expect(next, `<${name}> follows a later CT_ChartSpace member`).toBeGreaterThan(at)
    at = next
  }
}

/// Faithfully ordered skeleton of a real Excel 2007-era PivotChart chart part
/// (captured in openpyxl's chart test data, files.xlsx): date1904, lang,
/// roundedCorners, pivotSource "[files.xlsx]PIVOT!PivotTable1" + fmtId 0,
/// then c:chart opening with pivotFmts before title.
const EXCEL_PIVOT_CHART_SKELETON =
  '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"' +
  ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
  '<c:date1904 val="0"/><c:lang/><c:roundedCorners val="0"/>' +
  '<c:pivotSource><c:name>[files.xlsx]PIVOT!PivotTable1</c:name><c:fmtId val="0"/></c:pivotSource>' +
  '<c:chart><c:pivotFmts><c:pivotFmt><c:idx val="0"/></c:pivotFmt></c:pivotFmts>' +
  '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Website Performance</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>' +
  '<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/></c:ser><c:axId val="1"/><c:axId val="2"/></c:lineChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea>' +
  '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/>' +
  '</c:chart><c:printSettings><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/></c:printSettings>' +
  '</c:chartSpace>'

describe('PivotChart chartSpace pivotSource', () => {
  it('writes pivotSource bound to the pivotTable part name, before c:chart', async () => {
    const plan = await planWith([chartAddition()])
    const chartXml = plan.added.get('xl/charts/chart2.xml')!
    expect(chartXml).toContain(
      '<c:pivotSource><c:name>[]Data!PivotTable1</c:name><c:fmtId val="0"/></c:pivotSource>',
    )
    // CT_PivotSource requires both name and fmtId; the pivotSource block sits
    // directly before c:chart (we write no earlier CT_ChartSpace members).
    expect(chartXml).toMatch(/<\/c:pivotSource><c:chart>/)
    assertSchemaOrder(chartSpaceChildren(chartXml))
  })

  it('writes pivotFmts as the first child of c:chart, matching Excel parts', async () => {
    const plan = await planWith([chartAddition()])
    const chartXml = plan.added.get('xl/charts/chart2.xml')!
    expect(chartXml).toContain(
      '<c:chart><c:pivotFmts><c:pivotFmt><c:idx val="0"/></c:pivotFmt></c:pivotFmts><c:title>',
    )
  })

  it('escapes the sheet name inside c:name', async () => {
    const buffer = await buildPivotFixture('A&B')
    const plan = await planWith([chartAddition(undefined, 'A&B')], buffer)
    const chartXml = plan.added.get('xl/charts/chart2.xml')!
    expect(chartXml).toContain('<c:name>[]A&amp;B!PivotTable1</c:name>')
  })

  it('regular charts stay byte-identical in structure: no pivotSource, no pivotFmts', async () => {
    const plan = await planWith([
      chartAddition({ pivotTablePath: undefined }),
      chartAddition({ pivotTablePath: undefined, chartType: 'pie' }),
    ])
    for (const path of ['xl/charts/chart2.xml', 'xl/charts/chart3.xml']) {
      const chartXml = plan.added.get(path)!
      expect(chartXml).not.toContain('pivotSource')
      expect(chartXml).not.toContain('pivotFmts')
      assertSchemaOrder(chartSpaceChildren(chartXml))
    }
  })

  it.each([
    ['the pivotTable part is missing', 'xl/pivotTables/pivotTable1.xml', null],
    ['the path is not a pivotTable part', 'xl/worksheets/sheet1.xml', 'worksheet'],
    ['the definition carries no name', 'xl/pivotTables/pivotTable1.xml', 'unnamed'],
  ])('degrades to a regular chart when %s', async (_reason, path, fixtureMode) => {
    let buffer = await buildPivotFixture()
    if (fixtureMode === null) {
      const zip = await JSZip.loadAsync(buffer)
      zip.remove('xl/pivotTables/pivotTable1.xml')
      buffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
    } else if (fixtureMode === 'unnamed') {
      const zip = await JSZip.loadAsync(buffer)
      zip.file('xl/pivotTables/pivotTable1.xml', PIVOT_TABLE_XML.replace(' name="PivotTable1"', ''))
      buffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
    }
    const plan = await planWith([chartAddition({ pivotTablePath: path })], buffer)
    const chartXml = plan.added.get('xl/charts/chart2.xml')!
    // A pivotSource Excel cannot resolve would make it refuse the file, so an
    // unresolvable marker saves as the plain chart instead.
    expect(chartXml).not.toContain('pivotSource')
    expect(chartXml).toContain('<c:barDir val="col"/>')
  })

  it('matches the CT_ChartSpace ordering of a real Excel PivotChart part', () => {
    // Self-check of the pinned order list: the real artifact's members (with
    // its [workbook]Sheet!PivotTable name form and fmtId) validate too.
    assertSchemaOrder(chartSpaceChildren(EXCEL_PIVOT_CHART_SKELETON))
    const excel = chartSpaceChildren(EXCEL_PIVOT_CHART_SKELETON)
    expect(excel.indexOf('pivotSource')).toBe(excel.indexOf('chart') - 1)
    expect(EXCEL_PIVOT_CHART_SKELETON).toMatch(
      /<c:pivotSource><c:name>\[[^\]]*\][^!]+!PivotTable1<\/c:name><c:fmtId val="0"\/><\/c:pivotSource>/,
    )
  })

  it('chart edits keep the pivotSource and pivotFmts blocks intact', async () => {
    const plan = await planWith([chartAddition()])
    const chartXml = plan.added.get('xl/charts/chart2.xml')!
    const edited = applyChartEdit(chartXml, { chartPath: 'xl/charts/chart2.xml', title: 'Renamed' })
    expect(edited).toContain('<c:name>[]Data!PivotTable1</c:name>')
    expect(edited).toContain('<c:fmtId val="0"/>')
    expect(edited).toContain('<c:chart><c:pivotFmts>')
    expect(edited).toContain('<a:t>Renamed</a:t>')
    assertSchemaOrder(chartSpaceChildren(edited))
  })
})

describe('PivotChart read-back through the sidecar', () => {
  it('the sidecar parses a chart part with pivotSource without errors', async () => {
    const source = await buildPivotFixture()
    const plan = await planWith([chartAddition()], source)
    const mutation = await assembleWithJsZip(source, plan)

    const executable = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
    const client = new XlsxSidecarClient(
      fileURLToPath(new URL(`../native/xlsx-engine/target/release/${executable}`, import.meta.url)),
    )
    const directory = await mkdtemp(join(tmpdir(), 'pivot-chart-source-'))
    let sessionId: string | undefined
    try {
      const path = join(directory, 'pivot-chart.xlsx')
      await writeFile(path, mutation.buffer)
      const opened = workbookFileSchema
        .pick({ sessionId: true, visuals: true })
        .loose()
        .parse(await client.open(path))
      sessionId = opened.sessionId
      const chart = opened.visuals.find((visual) => visual.kind === 'chart')
      expect(chart?.chart?.chartTypes).toEqual(['barChart'])
      expect(chart?.chart?.series[0]?.values).toEqual([8, 6, 3])
      // A pivot-bound chart is not mistaken for anything exotic on read; the
      // pivotSource block is simply not part of the wire metadata.
      expect(chart?.chart?.title).toBe('Users by Region')
    } finally {
      if (sessionId !== undefined) await client.close(sessionId)
      client.stop()
      await rm(directory, { recursive: true, force: true })
    }
  })
})
