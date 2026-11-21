/** A new chart of the 2014 extension is saved as a chartEx part and reads back with its layout. */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { buildChartexPartXml, parseChartPartXml, parseDocx, saveDocx } from '../src/index'
import type { ChartexLayout, NewChart } from '../src/types'
import { buildDocx } from './helpers/build-docx'

const DATA = {
  title: 'Cash flow',
  categories: ['Start', 'Costs', 'Sales'],
  series: [{ name: 'Amount', values: [100, -40, 60] }],
}

const LAYOUTS: Array<[ChartexLayout, NewChart['kind']]> = [
  ['waterfall', 'bar'],
  ['funnel', 'bar'],
  ['boxWhisker', 'bar'],
  ['histogram', 'bar'],
  ['pareto', 'bar'],
  ['treemap', 'pie'],
  ['sunburst', 'pie'],
]

describe('buildChartexPartXml', () => {
  it.each(LAYOUTS)('%s reads back as itself over a %s', (layout, kind) => {
    const xml = buildChartexPartXml({ kind, layout, ...DATA }, layout)
    const display = parseChartPartXml(xml, 'word/charts/chartEx1.xml')!
    expect(display.layout).toBe(layout)
    expect(display.kind).toBe(kind)
    expect(display.title).toBe('Cash flow')
    expect(display.categories).toEqual(DATA.categories)
    expect(display.series).toEqual([{ name: 'Amount', values: [100, -40, 60] }])
  })

  it('a treemap and a sunburst size by the value and have no axes; the rest plot it on two', () => {
    const treemap = buildChartexPartXml({ kind: 'pie', ...DATA }, 'treemap')
    expect(treemap).toContain('<cx:numDim type="size">')
    expect(treemap).not.toContain('<cx:axis ')
    const funnel = buildChartexPartXml({ kind: 'bar', ...DATA }, 'funnel')
    expect(funnel).toContain('<cx:numDim type="val">')
    expect(funnel.match(/<cx:axis /g)).toHaveLength(2)
  })

  it('a histogram bins its clustered columns; a Pareto adds the line on a percentage axis', () => {
    const histogram = buildChartexPartXml({ kind: 'bar', ...DATA }, 'histogram')
    expect(histogram).toContain('<cx:series layoutId="clusteredColumn">')
    expect(histogram).toContain('<cx:binning intervalClosed="r"/>')
    expect(histogram).not.toContain('paretoLine')
    const pareto = buildChartexPartXml({ kind: 'bar', ...DATA }, 'pareto')
    expect(pareto).toContain('<cx:series layoutId="paretoLine" ownerIdx="0"><cx:axisId val="2"/>')
    expect(pareto).toContain('<cx:units unit="percentage"/>')
    expect(pareto.match(/<cx:axis /g)).toHaveLength(3)
  })

  it('a gap in the values is a point left out, and text is escaped', () => {
    const xml = buildChartexPartXml(
      {
        kind: 'bar',
        title: 'A & B',
        categories: ['<x>', 'y'],
        series: [{ name: 'S', values: [null, 2] }],
      },
      'waterfall',
    )
    expect(xml).toContain('<a:t>A &amp; B</a:t>')
    expect(xml).toContain('<cx:pt idx="0">&lt;x&gt;</cx:pt>')
    expect(xml).toContain('formatCode="General"><cx:pt idx="1">2</cx:pt></cx:lvl>')
  })
})

describe('saveDocx with a chart of the 2014 extension', () => {
  it('creates the chartEx part with its own relationship, content type and drawing', async () => {
    const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>before</w:t></w:r></w:p>' })
    const parsed = await parseDocx(source)
    const saved = await saveDocx(parsed, [
      { kind: 'original', docxIndex: 0 },
      { kind: 'chart', chart: { kind: 'bar', layout: 'waterfall', ...DATA } },
      { kind: 'chart', chart: { kind: 'bar', ...DATA } },
    ])
    const zip = await JSZip.loadAsync(saved)
    // the two families keep their own part names
    expect(zip.file('word/charts/chartEx1.xml')).toBeTruthy()
    expect(zip.file('word/charts/chart1.xml')).toBeTruthy()
    expect(zip.file('word/charts/_rels/chartEx1.xml.rels')).toBeTruthy()
    const types = await zip.file('[Content_Types].xml')!.async('string')
    expect(types).toMatch(
      /PartName="\/word\/charts\/chartEx1\.xml" ContentType="application\/vnd\.ms-office\.chartex\+xml"/,
    )
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    expect(rels).toMatch(/relationships\/chartEx"[^>]*Target="charts\/chartEx1\.xml"/)
    const doc = await zip.file('word/document.xml')!.async('string')
    expect(doc).toContain('Requires="cx1"')
    expect(doc).toContain('<mc:Fallback/>')
    const part = await zip.file('word/charts/chartEx1.xml')!.async('string')
    expect(part).toContain('<cx:externalData r:id="rId1" cx:autoUpdate="0"/>')

    // Word refuses the document unless the chartEx part has a chart style and a colour style
    const chartRels = await zip.file('word/charts/_rels/chartEx1.xml.rels')!.async('string')
    expect(chartRels).toMatch(/relationships\/chartStyle" Target="styleEx1\.xml"/)
    expect(chartRels).toMatch(/relationships\/chartColorStyle" Target="colorsEx1\.xml"/)
    expect(await zip.file('word/charts/styleEx1.xml')!.async('string')).toContain('<cs:chartStyle ')
    expect(await zip.file('word/charts/colorsEx1.xml')!.async('string')).toContain(
      '<cs:colorStyle ',
    )
    expect(types).toMatch(
      /styleEx1\.xml" ContentType="application\/vnd\.ms-office\.chartstyle\+xml"/,
    )
    expect(types).toMatch(
      /colorsEx1\.xml" ContentType="application\/vnd\.ms-office\.chartcolorstyle\+xml"/,
    )
    // the classic chart needs neither
    const classicRels = await zip.file('word/charts/_rels/chart1.xml.rels')!.async('string')
    expect(classicRels).not.toContain('chartStyle')

    const reparsed = await parseDocx(saved)
    const charts = reparsed.blocks.filter((b) => b.chartDisplay).map((b) => b.chartDisplay!)
    expect(charts.map((c) => [c.layout, c.kind])).toEqual([
      ['waterfall', 'bar'],
      [undefined, 'bar'],
    ])
    expect(charts[0].series[0].values).toEqual([100, -40, 60])
    expect(charts[0].title).toBe('Cash flow')
  })

  it('declares on the root the prefixes its drawing uses when the source document does not', async () => {
    // a document from another generator: only the w: namespace on its root
    const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>before</w:t></w:r></w:p>' })
    const zipIn = await JSZip.loadAsync(source)
    const bare = (await zipIn.file('word/document.xml')!.async('string')).replace(
      /<w:document [^>]*>/,
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    )
    zipIn.file('word/document.xml', bare)
    const parsed = await parseDocx(await zipIn.generateAsync({ type: 'uint8array' }))
    const saved = await saveDocx(parsed, [
      { kind: 'original', docxIndex: 0 },
      { kind: 'chart', chart: { kind: 'bar', ...DATA } },
    ])
    const doc = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    const root = /<w:document [^>]*>/.exec(doc)![0]
    expect(doc).toContain('<wp:inline ')
    expect(root).toContain('xmlns:wp=')
    expect(root).toContain('xmlns:r=')
  })
})
