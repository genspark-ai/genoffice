/**
 * Tests for new insert capabilities: chart / SmartArt / audio-video / 3D / hyperlink / header-footer /
 * field textbox / WordArt outline / arbitrary preset geometry shapes.
 * All verify the insert → save → reopen roundtrip.
 */
import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import {
  openPptx,
  savePptx,
  addElement,
  addChart,
  editChartElement,
  markChartEditable,
  buildChartSpaceXml,
  parseChartXml,
  addSmartArt,
  addMedia,
  addModel3d,
  model3dPartOf,
  solidPng,
  setElementLink,
  getElementLink,
  getSlideLinks,
  groupElements,
  applyHeaderFooter,
  readHeaderFooter,
  duplicateSlide,
  type ChartElement,
  type GroupElement,
  type PassthroughElement,
  type PictureElement,
  type TextElement,
} from '../src/index'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

const OFF = { x: 914400, y: 914400, cx: 4572000, cy: 2743200 }

describe('addChart', () => {
  it('bar chart round-trips categories/series and renders as chart element', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addChart(opened, 0, {
      kind: 'bar',
      title: 'Quarterly Sales',
      categories: ['Q1', 'Q2', 'Q3'],
      series: [
        { name: 'East', values: [10, 20, 30] },
        { name: 'South', values: [15, 12, 22] },
      ],
      offset: { ...OFF },
    })
    expect(r).not.toBeNull()
    const el = r!.slide.elements.find((e) => e.id === r!.elementId)
    expect(el?.type).toBe('chart')

    const reopened = await openPptx(await savePptx(opened))
    const el2 = reopened.deck.slides[0]!.elements.at(-1) as ChartElement
    expect(el2.type).toBe('chart')
    expect(el2.chart.kind).toBe('bar')
    expect(el2.chart.categories).toEqual(['Q1', 'Q2', 'Q3'])
    expect(el2.chart.series.length).toBe(2)
    expect(el2.chart.series[0]!.values).toEqual([10, 20, 30])
    expect(el2.chart.legendPos).toBe('b')
    // Title isn't in ChartModel; verify the part XML on disk
    const chartPart = [...reopened.archive.entries.keys()].find((p) =>
      /^ppt\/charts\/chart\d+\.xml$/.test(p),
    )!
    expect(reopened.archive.readText(chartPart)).toContain('Quarterly Sales')
  })

  it('pie and doughnut charts parse back with their kinds', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    addChart(opened, 0, {
      kind: 'pie',
      categories: ['A', 'B'],
      series: [{ name: 'S', values: [1, 2] }],
      offset: { ...OFF },
    })
    addChart(opened, 0, {
      kind: 'doughnut',
      categories: ['A', 'B'],
      series: [{ name: 'S', values: [3, 4] }],
      offset: { ...OFF },
    })
    const reopened = await openPptx(await savePptx(opened))
    const els = reopened.deck.slides[0]!.elements.slice(-2) as ChartElement[]
    // The engine models a doughnut chart as pie + holePct
    expect(els.map((e) => e.chart.kind)).toEqual(['pie', 'pie'])
    expect(els.map((e) => e.chart.holePct)).toEqual([0, 50])
  })

  it('two charts on one slide get distinct parts', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    addChart(opened, 0, {
      kind: 'line',
      categories: ['a'],
      series: [{ name: 's', values: [1] }],
      offset: { ...OFF },
    })
    addChart(opened, 0, {
      kind: 'bar',
      categories: ['a'],
      series: [{ name: 's', values: [1] }],
      offset: { ...OFF },
    })
    const chartParts = [...opened.archive.entries.keys()].filter((p) =>
      /^ppt\/charts\/chart\d+\.xml$/.test(p),
    )
    expect(new Set(chartParts).size).toBe(chartParts.length)
    expect(chartParts.length).toBeGreaterThanOrEqual(2)
  })

  it('scatter chart round-trips: numeric categories go into c:xVal, series values into c:yVal, dual value axes', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addChart(opened, 0, {
      kind: 'scatter',
      categories: ['0.5', '1.8', '3'],
      series: [{ name: 'Sample', values: [10, 20, 15] }],
      offset: { ...OFF },
    })
    expect(r).not.toBeNull()

    const reopened = await openPptx(await savePptx(opened))
    const el = reopened.deck.slides[0]!.elements.at(-1) as ChartElement
    expect(el.type).toBe('chart')
    expect(el.chart.kind).toBe('scatter')
    expect(el.chart.scatterStyle).toBe('lineMarker')
    expect(el.chart.series[0]!.name).toBe('Sample')
    expect(el.chart.series[0]!.xValues).toEqual([0.5, 1.8, 3])
    expect(el.chart.series[0]!.values).toEqual([10, 20, 15])
  })

  it('scatter with non-numeric categories falls back to 1..n x values', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    addChart(opened, 0, {
      kind: 'scatter',
      categories: ['a', 'b', 'c'],
      series: [{ name: 's', values: [1, 2, 3] }],
      offset: { ...OFF },
    })
    const reopened = await openPptx(await savePptx(opened))
    const el = reopened.deck.slides[0]!.elements.at(-1) as ChartElement
    expect(el.chart.series[0]!.xValues).toEqual([1, 2, 3])
  })

  it('radar chart round-trips: category vertices + radarStyle standard', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addChart(opened, 0, {
      kind: 'radar',
      title: 'Skill Radar',
      categories: ['Listen', 'Speak', 'Read', 'Write'],
      series: [
        { name: 'A', values: [3, 4, 5, 2] },
        { name: 'B', values: [2, 5, 3, 4] },
      ],
      offset: { ...OFF },
    })
    expect(r).not.toBeNull()

    const reopened = await openPptx(await savePptx(opened))
    const el = reopened.deck.slides[0]!.elements.at(-1) as ChartElement
    expect(el.type).toBe('chart')
    expect(el.chart.kind).toBe('radar')
    expect(el.chart.radarStyle).toBe('standard')
    expect(el.chart.categories).toEqual(['Listen', 'Speak', 'Read', 'Write'])
    expect(el.chart.series.map((s) => s.values)).toEqual([
      [3, 4, 5, 2],
      [2, 5, 3, 4],
    ])
    const chartPart = [...reopened.archive.entries.keys()]
      .filter((p) => /^ppt\/charts\/chart\d+\.xml$/.test(p))
      .at(-1)!
    expect(reopened.archive.readText(chartPart)).toContain('Skill Radar')
  })

  it('buildChartSpaceXml scatter/radar produce valid OOXML recognizable by parseChartXml', () => {
    const scatterXml = buildChartSpaceXml({
      kind: 'scatter',
      categories: ['1', '2'],
      series: [{ name: 's', values: [3, 4] }],
      offset: { ...OFF },
    })
    expect(scatterXml).toContain('<c:scatterChart>')
    expect(scatterXml).toContain('<c:xVal>')
    const sm = parseChartXml(scatterXml)!
    expect(sm.kind).toBe('scatter')
    expect(sm.series[0]!.xValues).toEqual([1, 2])

    const radarXml = buildChartSpaceXml({
      kind: 'radar',
      categories: ['a', 'b', 'c'],
      series: [{ name: 's', values: [1, 2, 3] }],
      offset: { ...OFF },
    })
    expect(radarXml).toContain('<c:radarChart>')
    const rm = parseChartXml(radarXml)!
    expect(rm.kind).toBe('radar')
    expect(rm.categories).toEqual(['a', 'b', 'c'])
  })
})

describe('editChartElement style edits and preservation', () => {
  it('style toggles take effect; later type-only edits keep legend/gridlines/data labels/colors', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addChart(opened, 0, {
      kind: 'bar',
      categories: ['Q1', 'Q2'],
      series: [
        { name: 'A', values: [1, 2] },
        { name: 'B', values: [3, 4] },
      ],
      offset: { ...OFF },
    })!
    expect(
      editChartElement(opened, 0, r.elementId, {
        legendPos: 'r',
        dataLabels: true,
        gridlines: true,
        valAxisTitle: 'Amount (100M)',
        gapWidthPct: 80,
        colorScheme: ['#112233', '#445566'],
      }),
    ).toBe(true)

    let reopened = await openPptx(await savePptx(opened))
    let el = reopened.deck.slides[0]!.elements.at(-1) as ChartElement
    expect(el.chart.legendPos).toBe('r')
    expect(el.chart.dataLabels).toBe(true)
    expect(el.chart.valAxis?.gridColor).toBeTruthy()
    expect(el.chart.valAxis?.title).toBe('Amount (100M)')
    expect(el.chart.gapWidthPct).toBe(80)
    expect(el.chart.series.map((s) => s.color)).toEqual(['#112233', '#445566'])

    // Type-only change: style and colors carried over after rebuild
    expect(editChartElement(reopened, 0, el.id, { kind: 'line' })).toBe(true)
    reopened = await openPptx(await savePptx(reopened))
    el = reopened.deck.slides[0]!.elements.at(-1) as ChartElement
    expect(el.chart.kind).toBe('line')
    expect(el.chart.legendPos).toBe('r')
    expect(el.chart.dataLabels).toBe(true)
    expect(el.chart.valAxis?.gridColor).toBeTruthy()
    expect(el.chart.valAxis?.title).toBe('Amount (100M)')
    expect(el.chart.series.map((s) => s.color)).toEqual(['#112233', '#445566'])
  })

  it('percentStacked and horizontal bar direction are preserved through rebuilds (edits do not degrade)', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addChart(opened, 0, {
      kind: 'barPercentStacked',
      categories: ['Q1', 'Q2'],
      series: [
        { name: 'A', values: [1, 2] },
        { name: 'B', values: [3, 4] },
      ],
      offset: { ...OFF },
      barDir: 'bar',
    })!
    let reopened = await openPptx(await savePptx(opened))
    let el = reopened.deck.slides[0]!.elements.at(-1) as ChartElement
    expect(el.chart.grouping).toBe('percentStacked')
    expect(el.chart.barDir).toBe('bar')

    // Style-only change: type subdivision and direction do not degrade
    expect(editChartElement(reopened, 0, el.id, { gridlines: true })).toBe(true)
    reopened = await openPptx(await savePptx(reopened))
    el = reopened.deck.slides[0]!.elements.at(-1) as ChartElement
    expect(el.chart.grouping).toBe('percentStacked')
    expect(el.chart.barDir).toBe('bar')
    void r
  })

  it('barPercentStacked/horizontal bar XML: grouping/overlap/barDir/swapped axis positions', () => {
    const xml = buildChartSpaceXml({
      kind: 'barPercentStacked',
      categories: ['a'],
      series: [{ name: 's', values: [1] }],
      offset: { ...OFF },
      barDir: 'bar',
    })
    expect(xml).toContain('<c:barDir val="bar"/>')
    expect(xml).toContain('<c:grouping val="percentStacked"/>')
    expect(xml).toContain('<c:overlap val="100"/>')
    // Horizontal bar chart: category axis on the left, value axis at the bottom
    expect(xml).toMatch(/<c:catAx>.*?<c:axPos val="l"\/>/)
    expect(xml).toMatch(/<c:valAx>.*?<c:axPos val="b"\/>/)
  })

  it('chart title parse round-trip; style-only edits keep the title; switchRowCol swaps rows/columns', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addChart(opened, 0, {
      kind: 'bar',
      title: 'Quarterly Results',
      categories: ['Q1', 'Q2'],
      series: [
        { name: 'A', values: [1, 2] },
        { name: 'B', values: [3, 4] },
      ],
      offset: { ...OFF },
    })!
    let reopened = await openPptx(await savePptx(opened))
    let el = reopened.deck.slides[0]!.elements.at(-1) as ChartElement
    expect(el.chart.title).toBe('Quarterly Results')

    // Style-only change: title kept
    editChartElement(reopened, 0, el.id, { gridlines: true })
    reopened = await openPptx(await savePptx(reopened))
    el = reopened.deck.slides[0]!.elements.at(-1) as ChartElement
    expect(el.chart.title).toBe('Quarterly Results')

    // Row/column switch: categories ↔ series
    editChartElement(reopened, 0, el.id, { switchRowCol: true })
    reopened = await openPptx(await savePptx(reopened))
    el = reopened.deck.slides[0]!.elements.at(-1) as ChartElement
    expect(el.chart.categories).toEqual(['A', 'B'])
    expect(el.chart.series.map((s) => s.name)).toEqual(['Q1', 'Q2'])
    expect(el.chart.series[0]!.values).toEqual([1, 3])
    expect(el.chart.series[1]!.values).toEqual([2, 4])
    void r
  })

  it('markChartEditable: a foreign chart becomes editable after marking, saves round-trip', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addChart(opened, 0, {
      kind: 'bar',
      categories: ['a'],
      series: [{ name: 's', values: [1] }],
      offset: { ...OFF },
    })!
    // Simulate a foreign chart: strip the descr marker
    let reopened = await openPptx(await savePptx(opened))
    let el = reopened.deck.slides[0]!.elements.at(-1) as ChartElement
    el.anchor.originalXml = el.anchor.originalXml.replace(' descr="aislides-chart"', '')
    delete el.descr
    expect(editChartElement(reopened, 0, el.id, { gridlines: true })).toBe(false)

    expect(markChartEditable(reopened.deck.slides[0]!, el.id)).toBe(true)
    expect(editChartElement(reopened, 0, el.id, { gridlines: true })).toBe(true)

    reopened = await openPptx(await savePptx(reopened))
    el = reopened.deck.slides[0]!.elements.at(-1) as ChartElement
    expect(el.descr).toBe('aislides-chart')
    expect(el.chart.valAxis?.gridColor).toBeTruthy()
    void r
  })

  it('legendPos none turns the legend off; passing an empty axis title clears it', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addChart(opened, 0, {
      kind: 'bar',
      categories: ['a'],
      series: [{ name: 's', values: [1] }],
      offset: { ...OFF },
    })!
    editChartElement(opened, 0, r.elementId, { valAxisTitle: 'Title' })
    let reopened = await openPptx(await savePptx(opened))
    let el = reopened.deck.slides[0]!.elements.at(-1) as ChartElement
    expect(el.chart.valAxis?.title).toBe('Title')

    editChartElement(reopened, 0, el.id, { legendPos: 'none', valAxisTitle: '' })
    reopened = await openPptx(await savePptx(reopened))
    el = reopened.deck.slides[0]!.elements.at(-1) as ChartElement
    expect(el.chart.legendPos).toBeUndefined()
    expect(el.chart.valAxis?.title).toBeUndefined()
  })
})

describe('addSmartArt (shape-group simplification)', () => {
  it('process layout becomes an editable group with one shape per item', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addSmartArt(opened, 0, {
      layout: 'process',
      items: ['Research', 'Design', 'Develop', 'Launch'],
      offset: { ...OFF },
    })
    expect(r).not.toBeNull()

    const reopened = await openPptx(await savePptx(opened))
    const el = reopened.deck.slides[0]!.elements.at(-1) as GroupElement
    expect(el.type).toBe('group')
    expect(el.children.length).toBe(4)
    const texts = el.children
      .map((c) => (c as TextElement).text?.paragraphs[0]?.runs[0]?.text)
      .filter(Boolean)
    expect(texts).toEqual(['Research', 'Design', 'Develop', 'Launch'])
  })

  it('hierarchy layout adds connector bars between levels', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addSmartArt(opened, 0, {
      layout: 'hierarchy',
      items: ['CEO', 'R&D', 'Marketing'],
      offset: { ...OFF },
    })
    const el = r!.slide.elements.find((e) => e.id === r!.elementId) as GroupElement
    // 1 top-level + 2×(connector+node) = 5
    expect(el.children.length).toBe(5)
  })

  it('pyramid layout stacks a triangle apex over widening trapezoids', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addSmartArt(opened, 0, {
      layout: 'pyramid',
      items: ['Vision', 'Strategy', 'Execution'],
      offset: { ...OFF },
    })
    expect(r).not.toBeNull()
    const xml = opened.deck.slides[0]!.originalXml
    expect(xml).toContain('prst="triangle"')
    expect((xml.match(/prst="trapezoid"/g) ?? []).length).toBe(2)

    const reopened = await openPptx(await savePptx(opened))
    const el = reopened.deck.slides[0]!.elements.at(-1) as GroupElement
    expect(el.type).toBe('group')
    expect(el.children.length).toBe(3)
  })

  it('matrix layout lays items on a 2-column grid', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addSmartArt(opened, 0, {
      layout: 'matrix',
      items: ['S', 'W', 'O', 'T'],
      offset: { ...OFF },
    })
    const el = r!.slide.elements.find((e) => e.id === r!.elementId) as GroupElement
    expect(el.children.length).toBe(4)
  })

  it('venn layout writes translucent overlapping circles', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addSmartArt(opened, 0, {
      layout: 'venn',
      items: ['A', 'B', 'C'],
      offset: { ...OFF },
    })
    expect(r).not.toBeNull()
    const xml = opened.deck.slides[0]!.originalXml
    expect((xml.match(/<a:alpha val="55000"\/>/g) ?? []).length).toBe(3)
    expect((xml.match(/prst="ellipse"/g) ?? []).length).toBe(3)
  })
})

describe('addMedia / addModel3d', () => {
  it('video embeds media part with dual rels and parses back as picture.media', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const bytes = new Uint8Array([0, 0, 0, 24, 0x66, 0x74, 0x79, 0x70]) // fake mp4 header
    const r = addMedia(opened, 0, { kind: 'video', bytes, ext: 'mp4', offset: { ...OFF } })
    expect(r).not.toBeNull()

    const reopened = await openPptx(await savePptx(opened))
    const el = reopened.deck.slides[0]!.elements.at(-1) as PictureElement
    expect(el.type).toBe('picture')
    expect(el.media?.kind).toBe('video')
    expect(el.media?.target).toMatch(/^ppt\/media\/media\d+\.mp4$/)
    // Content_Types has an mp4 Default
    expect(reopened.archive.readText('[Content_Types].xml')).toContain('Extension="mp4"')
  })

  it('audio parses back with audio kind', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addMedia(opened, 0, {
      kind: 'audio',
      bytes: new Uint8Array([0x49, 0x44, 0x33]),
      ext: 'mp3',
      offset: { ...OFF },
    })
    expect(r).not.toBeNull()
    const reopened = await openPptx(await savePptx(opened))
    const el = reopened.deck.slides[0]!.elements.at(-1) as PictureElement
    expect(el.media?.kind).toBe('audio')
  })

  it('3d model embeds glb part and poster picture with descr marker', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const glb = new Uint8Array([0x67, 0x6c, 0x54, 0x46]) // "glTF"
    const r = addModel3d(opened, 0, { bytes: glb, ext: 'glb', offset: { ...OFF } })
    expect(r).not.toBeNull()

    const reopened = await openPptx(await savePptx(opened))
    const el = reopened.deck.slides[0]!.elements.at(-1) as PictureElement
    expect(el.type).toBe('picture')
    const part = model3dPartOf(el)
    expect(part).toMatch(/^ppt\/media\/media\d+\.glb$/)
    expect(reopened.archive.readBytes(part!)).toEqual(glb)
  })

  it('solidPng produces a decodable png signature', () => {
    const png = solidPng(4, 3, [10, 20, 30])
    expect([...png.slice(0, 8)]).toEqual([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
  })
})

describe('unsupported chart kinds', () => {
  // Classic-namespace chart type the model doesn't support (same bucket as stock/surface;
  // chartex parts like waterfall/funnel never even reach parseChartXml)
  const BUBBLE_CHART = `<?xml version="1.0"?><c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:layout/>
<c:bubbleChart><c:ser><c:idx val="0"/>
  <c:xVal><c:numRef><c:f>x</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:xVal>
  <c:yVal><c:numRef><c:f>y</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:yVal>
  <c:bubbleSize><c:numRef><c:f>s</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>9</c:v></c:pt></c:numCache></c:numRef></c:bubbleSize>
</c:ser></c:bubbleChart>
</c:plotArea></c:chart></c:chartSpace>`

  it('parseChartXml returns null for a bubble chart (no recognizable plot)', () => {
    expect(parseChartXml(BUBBLE_CHART)).toBeNull()
  })

  it('an unsupported chart degrades to a passthrough element and survives an unrelated edit + save byte-for-byte', async () => {
    // Build a deck with a chart, then swap its part for a bubble chart on disk
    const opened = await openPptx(fx('01_standard_business.pptx'))
    addChart(opened, 0, {
      kind: 'bar',
      categories: ['A', 'B'],
      series: [{ name: 'S', values: [1, 2] }],
      offset: { ...OFF },
    })
    const staged = await openPptx(await savePptx(opened))
    const chartPart = [...staged.archive.entries.keys()].find((p) =>
      /^ppt\/charts\/chart\d+\.xml$/.test(p),
    )!
    staged.archive.entries.set(chartPart, Buffer.from(BUBBLE_CHART, 'utf8'))
    const buf = await savePptx(staged)

    // Reopen: the graphicFrame no longer parses as a chart — it must become a passthrough chip, not vanish
    const reopened = await openPptx(buf)
    const pt = reopened.deck.slides[0]!.elements.find(
      (e) => e.type === 'passthrough',
    ) as PassthroughElement
    expect(pt).toBeDefined()
    expect(pt.kind).toBe('chart')
    const frameXmlBefore = pt.anchor.originalXml

    // Edit an unrelated element and save
    addElement(reopened.deck.slides[0]!, { kind: 'rect', offset: { ...OFF, y: 4000000 } })
    const saved = await openPptx(await savePptx(reopened))

    // The chart part and the graphicFrame bytes are untouched
    expect(saved.archive.readText(chartPart)).toBe(BUBBLE_CHART)
    const pt2 = saved.deck.slides[0]!.elements.find(
      (e) => e.type === 'passthrough',
    ) as PassthroughElement
    expect(pt2.kind).toBe('chart')
    expect(pt2.anchor.originalXml).toBe(frameXmlBefore)
  })
})

describe('setElementLink / getElementLink', () => {
  it('url link survives save → reopen', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, { kind: 'rect', offset: { ...OFF }, fillColor: '#4472C4' })
    const fresh = setElementLink(opened, 0, el.id, {
      kind: 'url',
      url: 'https://example.com/a?b=1',
    })
    expect(fresh).not.toBeNull()

    const reopened = await openPptx(await savePptx(opened))
    const last = reopened.deck.slides[0]!.elements.at(-1)!
    expect(getElementLink(reopened, 0, last.id)).toEqual({
      kind: 'url',
      url: 'https://example.com/a?b=1',
    })
  })

  it('slide jump link resolves to target slide index', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    duplicateSlide(opened, 0) // ensure at least 2 slides
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, { kind: 'roundRect', offset: { ...OFF } })
    setElementLink(opened, 0, el.id, { kind: 'slide', slideIndex: 1 })

    const reopened = await openPptx(await savePptx(opened))
    const last = reopened.deck.slides[0]!.elements.at(-1)!
    expect(getElementLink(reopened, 0, last.id)).toEqual({ kind: 'slide', slideIndex: 1 })
    // The action attribute exists
    expect(last.anchor.originalXml).toContain('ppaction://hlinksldjump')
  })

  it('clearing the link removes hlinkClick', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, { kind: 'rect', offset: { ...OFF } })
    const s1 = setElementLink(opened, 0, el.id, { kind: 'url', url: 'https://x.dev' })!
    const linked = s1.elements.at(-1)!
    const s2 = setElementLink(opened, 0, linked.id, null)!
    expect(getElementLink(opened, 0, s2.elements.at(-1)!.id)).toBeNull()
  })

  it('getSlideLinks collects every linked element on the slide', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    duplicateSlide(opened, 0)
    const slide = opened.deck.slides[0]!
    const a = addElement(slide, { kind: 'rect', offset: { ...OFF } })
    const s1 = setElementLink(opened, 0, a.id, { kind: 'url', url: 'https://example.com' })!
    const b = addElement(s1, { kind: 'roundRect', offset: { ...OFF, x: 5486400 } })
    const s2 = setElementLink(opened, 0, b.id, { kind: 'slide', slideIndex: 1 })!

    const links = getSlideLinks(opened, 0)
    expect(links).toHaveLength(2)
    const byId = new Map(links.map((l) => [l.elementId, l.target]))
    expect(byId.get(s2.elements.at(-2)!.id)).toEqual({ kind: 'url', url: 'https://example.com' })
    expect(byId.get(s2.elements.at(-1)!.id)).toEqual({ kind: 'slide', slideIndex: 1 })
  })

  it('getSlideLinks attributes a grouped child link to the child, not the group', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const a = addElement(slide, { kind: 'rect', offset: { ...OFF } })
    const s1 = setElementLink(opened, 0, a.id, { kind: 'url', url: 'https://grouped.dev' })!
    const b = addElement(s1, { kind: 'ellipse', offset: { ...OFF, x: 5486400 } })
    const linked = s1.elements.at(-2)!
    const grouped = groupElements(opened, 0, [linked.id, b.id])!
    expect(grouped).not.toBeNull()

    const links = getSlideLinks(opened, 0)
    expect(links).toHaveLength(1)
    expect(links[0]!.target).toEqual({ kind: 'url', url: 'https://grouped.dev' })
    expect(links[0]!.elementId).not.toBe(grouped.groupId)
    const group = grouped.slide.elements.find((e) => e.id === grouped.groupId) as GroupElement
    expect(group.children.some((c) => c.id === links[0]!.elementId)).toBe(true)
  })
})

describe('applyHeaderFooter', () => {
  it('writes footer + slide number + date placeholders into every slide', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const changed = applyHeaderFooter(opened, {
      footer: 'Confidential · Internal',
      slideNum: true,
      date: 'July 23, 2026',
    })
    expect(changed).toBe(true)

    const reopened = await openPptx(await savePptx(opened))
    for (const slide of reopened.deck.slides) {
      const state = readHeaderFooter(slide)
      expect(state.footer).toBe('Confidential · Internal')
      expect(state.slideNum).toBe(true)
      expect(state.date).toBe('July 23, 2026')
    }
  })

  it('re-applying replaces (not duplicates) placeholders and can clear them', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    applyHeaderFooter(opened, { footer: 'v1', slideNum: true })
    applyHeaderFooter(opened, { footer: 'v2', slideNum: false })
    const slide = opened.deck.slides[0]!
    const state = readHeaderFooter(slide)
    expect(state.footer).toBe('v2')
    expect(state.slideNum).toBe(false)
    const ftrCount = slide.elements.filter((e) => e.placeholder === 'ftr').length
    expect(ftrCount).toBe(1)
  })
})

describe('field runs & wordart & preset shapes', () => {
  it('slidenum field textbox round-trips as a:fld', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, {
      kind: 'textbox',
      offset: { ...OFF },
      paragraphs: [{ runs: [{ text: '1', field: 'slidenum' }] }],
    })
    expect(el.anchor.originalXml).toContain('type="slidenum"')

    const reopened = await openPptx(await savePptx(opened))
    const el2 = reopened.deck.slides[0]!.elements.at(-1) as TextElement
    const run = el2.text!.paragraphs[0]!.runs[0]!
    expect(run.field).toBe('slidenum')
  })

  it('wordart outline (rPr a:ln) round-trips', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    addElement(slide, {
      kind: 'textbox',
      offset: { ...OFF },
      paragraphs: [
        {
          runs: [
            {
              text: 'WordArt',
              bold: true,
              fontSize: 54,
              color: '#4472C4',
              outline: { color: '#FFFFFF', widthEmu: 19050 },
            },
          ],
        },
      ],
    })
    const reopened = await openPptx(await savePptx(opened))
    const el2 = reopened.deck.slides[0]!.elements.at(-1) as TextElement
    const run = el2.text!.paragraphs[0]!.runs[0]!
    expect(run.outline).toEqual({ color: '#FFFFFF', widthEmu: 19050 })
    expect(run.fontSize).toBe(54)
  })

  it('arbitrary preset geometry (star5 / chevron) round-trips', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    addElement(slide, { kind: 'star5', offset: { ...OFF }, fillColor: '#FFC000' })
    addElement(slide, { kind: 'chevron', offset: { ...OFF }, fillColor: '#70AD47' })
    const reopened = await openPptx(await savePptx(opened))
    const els = reopened.deck.slides[0]!.elements.slice(-2) as TextElement[]
    expect(els.map((e) => e.presetGeometry)).toEqual(['star5', 'chevron'])
  })

  it('shape stroke option writes spPr a:ln', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    addElement(slide, {
      kind: 'rect',
      offset: { ...OFF },
      fillColor: '#FFFFFF',
      stroke: { color: '#4472C4', widthEmu: 28575 },
    })
    const reopened = await openPptx(await savePptx(opened))
    const el2 = reopened.deck.slides[0]!.elements.at(-1) as TextElement
    expect(el2.stroke?.width).toBe(28575)
    expect(el2.stroke?.fill).toEqual({ type: 'solid', color: '#4472C4' })
  })
})

describe('per-data-point chart colors', () => {
  it('pointColors patch writes <c:dPt>, survives save/reopen and later style-only rebuilds', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addChart(opened, 0, {
      kind: 'doughnut',
      categories: ['蓝', '橙', '黄', '绿'],
      series: [{ name: 'S', values: [1, 2, 3, 4] }],
      offset: { ...OFF },
    })!
    // The issue's scenario: blue/orange/yellow/green per sector — impossible via the 5 schemes
    expect(
      editChartElement(opened, 0, r.elementId, {
        pointColors: { 0: { 0: '#4472C4', 1: '#ED7D31', 2: '#FFC000', 3: '#70AD47' } },
      }),
    ).toBe(true)

    let reopened = await openPptx(await savePptx(opened))
    let el = reopened.deck.slides[0]!.elements.at(-1) as ChartElement
    expect(el.chart.series[0]!.pointColors).toEqual(['#4472C4', '#ED7D31', '#FFC000', '#70AD47'])

    // A style-only rebuild must not drop the dPt colors
    expect(editChartElement(reopened, 0, el.id, { dataLabels: true })).toBe(true)
    reopened = await openPptx(await savePptx(reopened))
    el = reopened.deck.slides[0]!.elements.at(-1) as ChartElement
    expect(el.chart.series[0]!.pointColors).toEqual(['#4472C4', '#ED7D31', '#FFC000', '#70AD47'])
    expect(el.chart.dataLabels).toBe(true)
  })

  it('single-point override and null clear', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const r = addChart(opened, 0, {
      kind: 'bar',
      categories: ['Q1', 'Q2', 'Q3'],
      series: [{ name: 'A', values: [1, 2, 3] }],
      offset: { ...OFF },
    })!
    editChartElement(opened, 0, r.elementId, { pointColors: { 0: { 1: '#FF0000' } } })
    let reopened = await openPptx(await savePptx(opened))
    let el = reopened.deck.slides[0]!.elements.at(-1) as ChartElement
    expect(el.chart.series[0]!.pointColors?.[1]).toBe('#FF0000')
    expect(el.chart.series[0]!.pointColors?.[0]).toBeUndefined()

    editChartElement(reopened, 0, el.id, { pointColors: { 0: { 1: null } } })
    reopened = await openPptx(await savePptx(reopened))
    el = reopened.deck.slides[0]!.elements.at(-1) as ChartElement
    expect(el.chart.series[0]!.pointColors?.[1]).toBeUndefined()
  })
})
