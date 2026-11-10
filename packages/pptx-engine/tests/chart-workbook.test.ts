/**
 * A chart's embedded workbook, which PowerPoint opens for "Edit Data": written on insert,
 * rewritten in place on every data edit, and dropped with the chart.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { validatePptx, xmllintAvailable } from '../../../tools/ooxml-validate/validate-pptx.mjs'
import {
  CHART_WORKBOOK_REL_TYPE,
  addChart,
  createBlankPptx,
  deleteElement,
  editChartElement,
  openPptx,
  savePptx,
  type NewChartOptions,
  type OpenedPptx,
} from '../src/index'
import { relsPathFor } from '../src/zip'

const schemaGate = xmllintAvailable() || !!process.env.CI
const OFF = { x: 0, y: 0, cx: 4000000, cy: 3000000 }
const SALES: Pick<NewChartOptions, 'categories' | 'series'> = {
  categories: ['Q1', 'Q2', 'Q3'],
  series: [
    { name: 'North', values: [10, 20, 30] },
    { name: 'South & East', values: [5, 15, 25] },
  ],
}

const chartPaths = (opened: OpenedPptx) =>
  [...opened.archive.entries.keys()].filter((p) => /^ppt\/charts\/chart\d+\.xml$/.test(p))
const embeddings = (opened: OpenedPptx) =>
  [...opened.archive.entries.keys()].filter((p) => p.startsWith('ppt/embeddings/'))

/** Sheet1's cells, shared strings resolved: { A1: '', B1: 'North', B2: 10, … } */
async function sheetCells(bytes: Uint8Array): Promise<Record<string, string | number>> {
  const zip = await JSZip.loadAsync(bytes)
  const sst = (await zip.file('xl/sharedStrings.xml')!.async('string')).match(
    /<si><t[^>]*>[^<]*<\/t><\/si>/g,
  )!
  const strings = sst.map((si) => /<t[^>]*>([^<]*)</.exec(si)![1]!.replace(/&amp;/g, '&'))
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  const cells: Record<string, string | number> = {}
  for (const [, ref, type, v] of sheet.matchAll(
    /<c r="([A-Z]+\d+)"( t="s")?><v>([^<]*)<\/v><\/c>/g,
  )) {
    cells[ref!] = type ? strings[Number(v)]! : Number(v)
  }
  return cells
}

function workbookOf(opened: OpenedPptx, chartPath: string) {
  const xml = opened.archive.readText(chartPath)!
  const rId =
    /<c:externalData r:id="([^"]+)"><c:autoUpdate val="0"\/><\/c:externalData><\/c:chartSpace>/.exec(
      xml,
    )?.[1]
  const rel = rId ? opened.archive.readRels(chartPath).get(rId) : undefined
  return { rId, rel, path: rel && `ppt/embeddings/${rel.target.split('/').pop()}` }
}

describe('chart workbook', () => {
  it('a new chart carries the workbook PowerPoint edits, with its data laid out as the chart references it', async () => {
    const opened = await openPptx(await createBlankPptx())
    addChart(opened, 0, { kind: 'bar', ...SALES, offset: OFF })
    const saved = await savePptx(opened)
    const back = await openPptx(saved)
    const [chartPath] = chartPaths(back)
    const { rel, path } = workbookOf(back, chartPath!)

    expect(rel?.type).toBe(CHART_WORKBOOK_REL_TYPE)
    expect(rel?.target).toBe('../embeddings/Microsoft_Excel_Worksheet1.xlsx')
    expect(back.archive.readText('[Content_Types].xml')).toContain(
      '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>',
    )
    expect(await sheetCells(back.archive.entries.get(path!)!)).toEqual({
      A1: '',
      B1: 'North',
      C1: 'South & East',
      A2: 'Q1',
      B2: 10,
      C2: 5,
      A3: 'Q2',
      B3: 20,
      C3: 15,
      A4: 'Q3',
      B4: 30,
      C4: 25,
    })
    const chartXml = back.archive.readText(chartPath!)!
    expect(chartXml).toContain('<c:f>Sheet1!$C$1</c:f>')
    expect(chartXml).toContain('<c:f>Sheet1!$B$2:$B$4</c:f>')
    if (schemaGate) expect(await validatePptx(saved)).toEqual([])
  })

  it('a data edit rewrites the same workbook, and a style edit keeps it in step', async () => {
    const opened = await openPptx(await createBlankPptx())
    const added = addChart(opened, 0, { kind: 'bar', ...SALES, offset: OFF })!
    const [chartPath] = chartPaths(opened)
    const before = workbookOf(opened, chartPath!)

    expect(
      editChartElement(opened, 0, added.elementId, {
        categories: ['Jan', 'Feb'],
        series: [{ name: 'Total', values: [7, 8] }],
      }),
    ).toBe(true)
    const reopened = await openPptx(await savePptx(opened))
    const el = reopened.deck.slides[0]!.elements.at(-1)!
    expect(editChartElement(reopened, 0, el.id, { kind: 'line', gridlines: true })).toBe(true)
    const back = await openPptx(await savePptx(reopened))

    const after = workbookOf(back, chartPath!)
    expect(after).toEqual(before)
    expect(embeddings(back)).toEqual([before.path])
    expect(await sheetCells(back.archive.entries.get(after.path!)!)).toEqual({
      A1: '',
      B1: 'Total',
      A2: 'Jan',
      B2: 7,
      A3: 'Feb',
      B3: 8,
    })
  })

  it("a scatter chart's x values are numbers in column A", async () => {
    const opened = await openPptx(await createBlankPptx())
    addChart(opened, 0, {
      kind: 'scatter',
      categories: ['1.5', 'n/a', '4'],
      series: [{ name: 'y', values: [3, 6, 9] }],
      offset: OFF,
    })
    const { path } = workbookOf(opened, chartPaths(opened)[0]!)
    const cells = await sheetCells(opened.archive.entries.get(path!)!)
    expect([cells.A2, cells.A3, cells.A4]).toEqual([1.5, 2, 4])
  })

  it('a chart linked to another data source gets its own workbook, and the old source goes', async () => {
    const opened = await openPptx(await createBlankPptx())
    const added = addChart(opened, 0, { kind: 'bar', ...SALES, offset: OFF })!
    const [chartPath] = chartPaths(opened)
    const { path } = workbookOf(opened, chartPath!)
    const relsPath = relsPathFor(chartPath!)
    opened.archive.entries.delete(path!)
    opened.archive.entries.set('ppt/embeddings/oleObject1.bin', new Uint8Array([1, 2, 3]))
    opened.archive.entries.set(
      relsPath,
      Buffer.from(
        opened.archive
          .readText(relsPath)!
          .replace(
            /Type="[^"]+" Target="[^"]+"/,
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"',
          ),
      ),
    )

    expect(editChartElement(opened, 0, added.elementId, { title: 'Sales' })).toBe(true)
    const after = workbookOf(opened, chartPath!)
    expect(after.rel?.type).toBe(CHART_WORKBOOK_REL_TYPE)
    expect([...opened.archive.readRels(chartPath!).values()].map((r) => r.type)).toEqual([
      CHART_WORKBOOK_REL_TYPE,
    ])
    expect(embeddings(opened)).toEqual([after.path])
  })

  it('deleting the chart drops its workbook', async () => {
    const opened = await openPptx(await createBlankPptx())
    const added = addChart(opened, 0, { kind: 'pie', ...SALES, offset: OFF })!
    expect(embeddings(opened)).toHaveLength(1)
    expect(deleteElement(opened, added.slide, added.elementId)).toBe(true)
    expect(embeddings(opened)).toEqual([])
  })
})
