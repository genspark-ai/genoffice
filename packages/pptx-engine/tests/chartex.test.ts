import { describe, it, expect } from 'vitest'
import { parseChartExXml } from '../src/chartex'

const FUNNEL = `<?xml version="1.0"?><cx:chartSpace xmlns:cx="cx" xmlns:a="a" xmlns:r="r">
<cx:chartData><cx:data id="0">
<cx:strDim type="cat"><cx:f>S!A</cx:f><cx:lvl ptCount="3"><cx:pt idx="0">C1</cx:pt><cx:pt idx="1">C2</cx:pt><cx:pt idx="2">C3</cx:pt></cx:lvl></cx:strDim>
<cx:numDim type="val"><cx:f>S!B</cx:f><cx:lvl ptCount="3"><cx:pt idx="0">50</cx:pt><cx:pt idx="1">100</cx:pt><cx:pt idx="2">200</cx:pt></cx:lvl></cx:numDim>
</cx:data></cx:chartData>
<cx:chart><cx:title/><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="funnel"><cx:dataId val="0"/></cx:series></cx:plotAreaRegion>
<cx:axis id="1"><cx:catScaling gapWidth="0.06"/></cx:axis></cx:plotArea></cx:chart></cx:chartSpace>`

const SUNBURST = `<?xml version="1.0"?><cx:chartSpace xmlns:cx="cx" xmlns:a="a" xmlns:r="r">
<cx:chartData><cx:data id="0">
<cx:strDim type="cat"><cx:f>S!A</cx:f>
<cx:lvl ptCount="3"><cx:pt idx="0">Leaf1</cx:pt><cx:pt idx="1">Leaf2</cx:pt><cx:pt idx="2"/></cx:lvl>
<cx:lvl ptCount="3"><cx:pt idx="0">Stem1</cx:pt><cx:pt idx="1">Stem1</cx:pt><cx:pt idx="2">Leaf3</cx:pt></cx:lvl>
<cx:lvl ptCount="3"><cx:pt idx="0">Root</cx:pt><cx:pt idx="1">Root</cx:pt><cx:pt idx="2">Root</cx:pt></cx:lvl>
</cx:strDim>
<cx:numDim type="size"><cx:f>S!B</cx:f><cx:lvl ptCount="3"><cx:pt idx="0">10</cx:pt><cx:pt idx="1">20</cx:pt><cx:pt idx="2">30</cx:pt></cx:lvl></cx:numDim>
</cx:data></cx:chartData>
<cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="sunburst">
<cx:dataPt idx="2"><cx:spPr><a:solidFill><a:srgbClr val="B0F0FF"><a:alpha val="0"/></a:srgbClr></a:solidFill></cx:spPr></cx:dataPt>
<cx:dataId val="0"/></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>`

describe('parseChartExXml', () => {
  it('funnel: categories + values + gap fraction + placeholder title', () => {
    const m = parseChartExXml(FUNNEL)!
    expect(m.kind).toBe('funnel')
    expect(m.categories).toEqual(['C1', 'C2', 'C3'])
    expect(m.series[0]!.values).toEqual([50, 100, 200])
    expect(m.gapWidthPct).toBeCloseTo(6)
    expect(m.title).toBe('Chart Title')
  })

  it('sunburst: leaf-first levels, sizes, transparent per-point color', () => {
    const m = parseChartExXml(SUNBURST)!
    expect(m.kind).toBe('sunburst')
    expect(m.sunburst!.levels).toEqual([
      ['Leaf1', 'Leaf2', ''],
      ['Stem1', 'Stem1', 'Leaf3'],
      ['Root', 'Root', 'Root'],
    ])
    expect(m.sunburst!.sizes).toEqual([10, 20, 30])
    expect(m.sunburst!.pointColors![2]).toBe('#B0F0FF00')
    expect(m.title).toBeUndefined()
  })

  it('explicit rich title text wins over the placeholder (incl. xml:space form)', () => {
    const withTitle = FUNNEL.replace(
      '<cx:title/>',
      '<cx:title><cx:tx><cx:rich><a:p><a:r><a:t xml:space="preserve">My </a:t></a:r><a:r><a:t>Funnel</a:t></a:r></a:p></cx:rich></cx:tx></cx:title>',
    )
    expect(parseChartExXml(withTitle)!.title).toBe('My Funnel')
  })

  it('unsupported layoutId returns null', () => {
    expect(parseChartExXml(FUNNEL.replace('layoutId="funnel"', 'layoutId="waterfall"'))).toBeNull()
  })
})
