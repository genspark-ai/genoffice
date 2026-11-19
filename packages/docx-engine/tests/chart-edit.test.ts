import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseChartPartXml, parseDocx, patchChartPartXml, saveDocx } from '../src/index'
import { CHART_PART_XML, buildChartDocx } from './helpers/build-docx'

describe('chart part parsing', () => {
  it('reads kind, title, categories and cached series values', () => {
    const display = parseChartPartXml(CHART_PART_XML, 'word/charts/chart1.xml')
    expect(display).toMatchObject({
      partPath: 'word/charts/chart1.xml',
      kind: 'bar',
      title: '销售统计',
      categories: ['一月', '二月', '三月'],
    })
    expect(display?.series).toEqual([
      { name: '华东', values: [120, 88.5, 96] },
      { name: '华南', values: [70, null, 110] },
    ])
  })

  it('surfaces the chart model on the protected block and keeps the part XML', async () => {
    const parsed = await parseDocx(await buildChartDocx())
    const block = parsed.blocks.find((b) => b.label === 'Chart')
    expect(block?.type).toBe('passthrough')
    expect(block?.chartDisplay?.title).toBe('销售统计')
    expect(block?.chartDisplay?.series[0].values).toEqual([120, 88.5, 96])
    expect(parsed.extras.chartParts['word/charts/chart1.xml']).toBe(CHART_PART_XML)
  })

  it('returns null for parts without cached series', () => {
    const empty =
      '<c:chartSpace xmlns:c="x"><c:chart><c:plotArea><c:scatterChart>' +
      '<c:ser><c:idx val="0"/></c:ser></c:scatterChart></c:plotArea></c:chart></c:chartSpace>'
    expect(parseChartPartXml(empty, 'word/charts/chart2.xml')).toBeNull()
  })
})

describe('patchChartPartXml', () => {
  it('patches title, series name, values and categories; keeps refs and structure', () => {
    const patched = patchChartPartXml(CHART_PART_XML, {
      title: '2026 营收 & 目标',
      categories: [null, '2月', null],
      series: [{ name: '华东大区', values: [null, 99, null] }, null],
    })
    // structure and data references stay byte-identical
    expect(patched).toContain('<c:f>Sheet1!$B$2:$B$4</c:f>')
    expect(patched).toContain('<c:formatCode>General</c:formatCode>')
    expect(patched).toContain('<c:barDir val="col"/>')
    // title collapses into the first run, second run blanked, entities escaped
    expect(patched).toContain('<a:t>2026 营收 &amp; 目标</a:t>')
    expect(patched).toContain('<a:t></a:t>')
    // targeted cache edits only
    expect(patched).toContain('<c:v>华东大区</c:v>')
    expect(patched).toContain('<c:v>华南</c:v>')
    expect(patched).toContain('<c:v>99</c:v>')
    expect(patched).toContain('<c:v>120</c:v>')
    expect(patched).toContain('<c:v>2月</c:v>')
    expect(patched).not.toContain('<c:v>88.5</c:v>')

    const reparsed = parseChartPartXml(patched, 'word/charts/chart1.xml')
    expect(reparsed?.title).toBe('2026 营收 & 目标')
    expect(reparsed?.categories).toEqual(['一月', '2月', '三月'])
    expect(reparsed?.series).toEqual([
      { name: '华东大区', values: [120, 99, 96] },
      { name: '华南', values: [70, null, 110] },
    ])
  })

  it('ignores edits it cannot anchor (gap points, missing title) and empty patches', () => {
    // idx 1 of series 2 has no cached pt: the edit is skipped, not invented
    const gap = patchChartPartXml(CHART_PART_XML, { series: [null, { values: [null, 42, null] }] })
    expect(gap).toBe(CHART_PART_XML)
    const noTitle = CHART_PART_XML.replace(/<c:title>[\s\S]*?<\/c:title>/, '')
    expect(patchChartPartXml(noTitle, { title: '新标题' })).toBe(noTitle)
    expect(patchChartPartXml(CHART_PART_XML, {})).toBe(CHART_PART_XML)
  })
})

describe('chart save round-trip', () => {
  it('saves a patched chart part via partXml and leaves everything else untouched', async () => {
    const source = await buildChartDocx()
    const parsed = await parseDocx(source)
    const originalPart = parsed.extras.chartParts['word/charts/chart1.xml']
    const patchedPart = patchChartPartXml(originalPart, {
      title: '更新后的标题',
      series: [{ values: [200, null, null] }, null],
    })

    const visible = parsed.blocks.filter((b) => !b.hidden)
    const saved = await saveDocx(
      parsed,
      visible.map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! })),
      { partXml: { 'word/charts/chart1.xml': patchedPart } },
    )

    const reparsed = await parseDocx(saved)
    const block = reparsed.blocks.find((b) => b.label === 'Chart')
    expect(block?.chartDisplay?.title).toBe('更新后的标题')
    expect(block?.chartDisplay?.series[0].values).toEqual([200, 88.5, 96])
    // the body paragraph is byte-identical; only the chart part changed
    const [oldZip, newZip] = await Promise.all([JSZip.loadAsync(source), JSZip.loadAsync(saved)])
    expect(await newZip.file('word/document.xml')!.async('string')).toBe(
      await oldZip.file('word/document.xml')!.async('string'),
    )
    expect(await newZip.file('word/styles.xml')!.async('string')).toBe(
      await oldZip.file('word/styles.xml')!.async('string'),
    )
  })

  it('stays byte-stable when no chart edit is made', async () => {
    const source = await buildChartDocx()
    const parsed = await parseDocx(source)
    const visible = parsed.blocks.filter((b) => !b.hidden)
    const saved = await saveDocx(
      parsed,
      visible.map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! })),
    )
    expect(saved).toEqual(source)
  })
})
