import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  buildChartPartXml,
  buildChartWorkbookXlsxBase64,
  parseChartPartXml,
  parseDocx,
  patchChartPartXml,
  patchChartWorkbookXlsxBase64,
  saveDocx,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

const SPEC = {
  kind: 'bar' as const,
  title: '季度销售',
  categories: ['Q1', 'Q2', 'Q3'],
  series: [
    { name: '华东', values: [120, 88.5, 96] },
    { name: '华南', values: [70, null, 110] },
  ],
}

describe('buildChartPartXml', () => {
  it('produces a part our own parser reads back', () => {
    const xml = buildChartPartXml(SPEC)
    const display = parseChartPartXml(xml, 'word/charts/chart1.xml')!
    expect(display.kind).toBe('bar')
    expect(display.title).toBe('季度销售')
    expect(display.categories).toEqual(['Q1', 'Q2', 'Q3'])
    expect(display.series).toEqual([
      { name: '华东', values: [120, 88.5, 96] },
      { name: '华南', values: [70, null, 110] },
    ])
  })

  it('emits axes for bar/line and none for pie', () => {
    expect(buildChartPartXml(SPEC)).toContain('<c:catAx>')
    expect(buildChartPartXml({ ...SPEC, kind: 'line' })).toContain('<c:lineChart>')
    const pie = buildChartPartXml({ ...SPEC, kind: 'pie', series: [SPEC.series[0]] })
    expect(pie).toContain('<c:pieChart>')
    expect(pie).not.toContain('<c:catAx>')
  })
})

describe('saveDocx kind:chart', () => {
  it('creates the part, relationship, content type and drawing paragraph', async () => {
    const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>前文</w:t></w:r></w:p>' })
    const parsed = await parseDocx(source)
    const saved = await saveDocx(parsed, [
      { kind: 'original', docxIndex: 0 },
      { kind: 'chart', chart: SPEC },
    ])
    const zip = await JSZip.loadAsync(saved)
    expect(zip.file('word/charts/chart1.xml')).toBeTruthy()
    const contentTypes = await zip.file('[Content_Types].xml')!.async('string')
    expect(contentTypes).toContain('/word/charts/chart1.xml')
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    expect(rels).toContain('charts/chart1.xml')

    const reparsed = await parseDocx(saved)
    const chartBlock = reparsed.blocks.find((b) => b.chartDisplay)
    expect(chartBlock?.chartDisplay?.title).toBe('季度销售')
    expect(chartBlock?.chartDisplay?.series[1].values).toEqual([70, null, 110])
  })

  it('allocates distinct part names for multiple charts in one save', async () => {
    const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' })
    const parsed = await parseDocx(source)
    const saved = await saveDocx(parsed, [
      { kind: 'original', docxIndex: 0 },
      { kind: 'chart', chart: SPEC },
      { kind: 'chart', chart: { ...SPEC, kind: 'pie', title: '占比' } },
    ])
    const zip = await JSZip.loadAsync(saved)
    expect(zip.file('word/charts/chart1.xml')).toBeTruthy()
    expect(zip.file('word/charts/chart2.xml')).toBeTruthy()
    const reparsed = await parseDocx(saved)
    const charts = reparsed.blocks.filter((b) => b.chartDisplay)
    expect(charts).toHaveLength(2)
    expect(charts[1].chartDisplay?.kind).toBe('pie')
  })
})

describe('chart embedded workbook', () => {
  it('buildChartWorkbookXlsxBase64 produces a valid xlsx with correct sheet data', async () => {
    const base64 = await buildChartWorkbookXlsxBase64(
      ['Q1', 'Q2', 'Q3'],
      [
        { name: '华东', values: [120, 88, 96] },
        { name: '华南', values: [70, 80, 110] },
      ],
    )
    expect(typeof base64).toBe('string')
    expect(base64.length).toBeGreaterThan(0)
    // Decode and verify it's a valid zip with the expected parts
    const binaryStr = atob(base64)
    const bytes = new Uint8Array(binaryStr.length)
    for (let i = 0; i < binaryStr.length; i++) bytes[i] = binaryStr.charCodeAt(i)
    const xlsx = await JSZip.loadAsync(bytes)
    expect(xlsx.file('xl/workbook.xml')).toBeTruthy()
    expect(xlsx.file('xl/worksheets/sheet1.xml')).toBeTruthy()
    expect(xlsx.file('xl/sharedStrings.xml')).toBeTruthy()
    const sheet = await xlsx.file('xl/worksheets/sheet1.xml')!.async('string')
    // Sheet uses shared string indices (t="s"), numeric values are inline
    expect(sheet).toContain('<v>120</v>')
    const ss = await xlsx.file('xl/sharedStrings.xml')!.async('string')
    // Category labels Q1/Q2/Q3 live in the shared strings table
    expect(ss).toContain('Q1')
    expect(ss).toContain('华东')
    expect(ss).toContain('华南')
  })

  it('saveDocx kind:chart embeds the workbook file and chart rels', async () => {
    const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>前文</w:t></w:r></w:p>' })
    const parsed = await parseDocx(source)
    const saved = await saveDocx(parsed, [
      { kind: 'original', docxIndex: 0 },
      { kind: 'chart', chart: SPEC },
    ])
    const zip = await JSZip.loadAsync(saved)
    // Workbook file must exist
    expect(zip.file('word/charts/embeddings/workbook1.xlsx')).toBeTruthy()
    // Chart rels must reference it
    const relsXml = await zip.file('word/charts/_rels/chart1.xml.rels')!.async('string')
    expect(relsXml).toContain('embeddings/workbook1.xlsx')
    // Chart XML must have c:externalData reference
    const chartXml = await zip.file('word/charts/chart1.xml')!.async('string')
    expect(chartXml).toContain('c:externalData')
    // Content types must register the workbook
    const ct = await zip.file('[Content_Types].xml')!.async('string')
    expect(ct).toContain('spreadsheetml.sheet')
  })

  it('patchChartWorkbookXlsxBase64 updates cell values', async () => {
    const original = await buildChartWorkbookXlsxBase64(
      ['Q1', 'Q2'],
      [{ name: '华东', values: [100, 200] }],
    )
    const patched = await patchChartWorkbookXlsxBase64(
      original,
      ['A', 'B'],
      [{ name: '华东', values: [999, 888] }],
    )
    expect(patched).toBeTruthy()
    const binaryStr = atob(patched!)
    const bytes = new Uint8Array(binaryStr.length)
    for (let i = 0; i < binaryStr.length; i++) bytes[i] = binaryStr.charCodeAt(i)
    const xlsx = await JSZip.loadAsync(bytes)
    const sheet = await xlsx.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheet).toContain('999')
    expect(sheet).toContain('888')
  })
})

describe('parseChartPartXml direction and auto title', () => {
  const chartSpace = (inner: string) =>
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    `<c:chart>${inner}</c:chart></c:chartSpace>`
  const barPlot = (dir: string) =>
    `<c:plotArea><c:barChart><c:barDir val="${dir}"/><c:ser><c:idx val="0"/>` +
    '<c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val>' +
    '</c:ser></c:barChart></c:plotArea>'

  it('c:barDir="bar" marks the chart as horizontal', () => {
    const display = parseChartPartXml(chartSpace(barPlot('bar')), 'word/charts/chart1.xml')!
    expect(display.kind).toBe('bar')
    expect(display.horizontal).toBe(true)
    expect(parseChartPartXml(chartSpace(barPlot('col')), 'p')!.horizontal).toBeUndefined()
  })

  it('text-less c:title renders the Word auto-title placeholder unless deleted', () => {
    const title = '<c:title><c:overlay val="0"/></c:title>'
    const kept = parseChartPartXml(
      chartSpace(`${title}<c:autoTitleDeleted val="0"/>${barPlot('col')}`),
      'p',
    )!
    expect(kept.title).toBe('Chart Title')
    const deleted = parseChartPartXml(
      chartSpace(`${title}<c:autoTitleDeleted val="1"/>${barPlot('col')}`),
      'p',
    )!
    expect(deleted.title).toBeUndefined()
    // CT_Boolean: a val-less element and val="true" also mean deleted
    for (const form of ['<c:autoTitleDeleted/>', '<c:autoTitleDeleted val="true"/>']) {
      const d = parseChartPartXml(chartSpace(`${title}${form}${barPlot('col')}`), 'p')!
      expect(d.title).toBeUndefined()
    }
  })

  it('title patches reach strRef caches and synthesize rich bodies for auto titles', () => {
    const strRefTitle =
      '<c:title><c:tx><c:strRef><c:f>Sheet1!$A$1</c:f>' +
      '<c:strCache><c:pt idx="0"><c:v>old</c:v></c:pt></c:strCache></c:strRef></c:tx></c:title>'
    const strRefXml = chartSpace(`${strRefTitle}${barPlot('col')}`)
    const patchedStrRef = patchChartPartXml(strRefXml, { title: 'new title' })
    expect(patchedStrRef).toContain('<c:v>new title</c:v>')
    const autoXml = chartSpace(`<c:title><c:overlay val="0"/></c:title>${barPlot('col')}`)
    const patchedAuto = patchChartPartXml(autoXml, { title: 'typed title' })
    expect(patchedAuto).toContain('<a:t>typed title</a:t>')
    expect(parseChartPartXml(patchedAuto, 'p')!.title).toBe('typed title')
    // Word auto titles already carry an empty c:tx/c:rich paragraph — the run
    // must land in it (CT_Title allows a single c:tx)
    const emptyTx =
      '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>' +
      '<a:endParaRPr lang="en-US"/></a:p></c:rich></c:tx></c:title>'
    const patchedEmptyTx = patchChartPartXml(chartSpace(`${emptyTx}${barPlot('col')}`), {
      title: 'kept single tx',
    })
    expect(patchedEmptyTx.match(/<c:tx>/g)).toHaveLength(1)
    // cache-less strRef tx: replaced wholesale by a rich body (CT_Tx choice)
    const bareStrRef = '<c:title><c:tx><c:strRef><c:f>Sheet1!$A$1</c:f></c:strRef></c:tx></c:title>'
    const patchedBare = patchChartPartXml(chartSpace(`${bareStrRef}${barPlot('col')}`), {
      title: 'from bare strRef',
    })
    expect(patchedBare.match(/<c:tx>/g)).toHaveLength(1)
    expect(patchedBare).not.toContain('<c:strRef>')
    expect(parseChartPartXml(patchedBare, 'p')!.title).toBe('from bare strRef')
    expect(/<a:r><a:t>kept single tx<\/a:t><\/a:r><a:endParaRPr/.test(patchedEmptyTx)).toBe(true)
  })

  it('strRef titles read the cached c:v text, not the placeholder', () => {
    const title =
      '<c:title><c:tx><c:strRef><c:f>Sheet1!$A$1</c:f>' +
      '<c:strCache><c:pt idx="0"><c:v>销售汇总</c:v></c:pt></c:strCache></c:strRef></c:tx></c:title>'
    const display = parseChartPartXml(
      chartSpace(`${title}<c:autoTitleDeleted val="0"/>${barPlot('col')}`),
      'p',
    )!
    expect(display.title).toBe('销售汇总')
  })
})
