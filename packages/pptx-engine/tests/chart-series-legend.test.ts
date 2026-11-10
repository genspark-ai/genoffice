import { describe, it, expect } from 'vitest'
import { parseChartXml } from '../src/chart'

const cats = (n: number) =>
  `<c:cat><c:strRef><c:f>x</c:f><c:strCache><c:ptCount val="${n}"/>` +
  Array.from({ length: n }, (_, i) => `<c:pt idx="${i}"><c:v>C${i}</c:v></c:pt>`).join('') +
  `</c:strCache></c:strRef></c:cat>`
const vals = (vs: number[]) =>
  `<c:val><c:numRef><c:f>y</c:f><c:numCache><c:ptCount val="${vs.length}"/>` +
  vs.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('') +
  `</c:numCache></c:numRef></c:val>`
const ser = (idx: number, name: string, vs: number[], extra = '') =>
  `<c:ser><c:idx val="${idx}"/><c:order val="${idx}"/><c:tx><c:strRef><c:f>n</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${name}</c:v></c:pt></c:strCache></c:strRef></c:tx>${extra}${cats(vs.length)}${vals(vs)}</c:ser>`
const space = (plotArea: string, chartExtra = '') =>
  `<?xml version="1.0"?><c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart>${chartExtra}<c:plotArea><c:layout/>${plotArea}</c:plotArea></c:chart></c:chartSpace>`

describe('chart parsing: pie of pie, legend order, spacer series, label precedence', () => {
  it('pie-of-pie parses as a pie with the split parameters', () => {
    const m = parseChartXml(
      space(
        `<c:ofPieChart><c:ofPieType val="pie"/><c:varyColors val="1"/>${ser(0, 'S', [0, 33, 27, 3])}` +
          `<c:gapWidth val="100"/><c:splitType val="pos"/><c:splitPos val="3"/><c:secondPieSize val="75"/></c:ofPieChart>`,
      ),
    )!
    expect(m.kind).toBe('pie')
    expect(m.ofPie).toEqual({ splitPos: 3, secondPieSize: 75, gapWidth: 100 })
  })

  it('stacked bar + line combo: side legend lists the stack top-first and c:legendEntry idx counts display positions', () => {
    const m = parseChartXml(
      space(
        `<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/>${ser(0, 'CCA', [1, 2])}${ser(2, 'Empty', [3, 4])}<c:axId val="1"/><c:axId val="2"/></c:barChart>` +
          `<c:lineChart><c:grouping val="standard"/>${ser(1, 'Growth', [5, 6])}<c:axId val="1"/><c:axId val="2"/></c:lineChart>` +
          `<c:catAx><c:axId val="1"/></c:catAx><c:valAx><c:axId val="2"/></c:valAx>`,
        '',
      ).replace(
        '</c:plotArea>',
        '</c:plotArea><c:legend><c:legendPos val="r"/><c:legendEntry><c:idx val="0"/><c:delete val="1"/></c:legendEntry></c:legend>',
      ),
    )!
    // display order [Empty, CCA, Growth]; position 0 (Empty) deleted
    expect(m.series.map((s) => s.name)).toEqual(['CCA', 'Empty', 'Growth'])
    expect(m.legendOrder).toEqual([0, 2])
  })

  it('clustered bars keep document order and an untouched legend has no explicit order', () => {
    const m = parseChartXml(
      space(
        `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, 'A', [1])}${ser(1, 'B', [2])}</c:barChart>`,
      ).replace('</c:plotArea>', '</c:plotArea><c:legend><c:legendPos val="r"/></c:legend>'),
    )!
    expect(m.legendOrder).toBeUndefined()
  })

  it('horizontal bars list the legend bottom-up; deletions count in that order', () => {
    const m = parseChartXml(
      space(
        `<c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/>${ser(0, 'S1', [1])}${ser(1, 'S2', [2])}${ser(2, 'S3', [3])}</c:barChart>`,
      ).replace(
        '</c:plotArea>',
        '</c:plotArea><c:legend><c:legendPos val="b"/><c:legendEntry><c:idx val="1"/><c:delete val="1"/></c:legendEntry></c:legend>',
      ),
    )!
    expect(m.legendOrder).toEqual([2, 0])
  })

  it('a bar series with <a:noFill/> is an invisible spacer', () => {
    const m = parseChartXml(
      space(
        `<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/>${ser(0, 'base', [10, 20], '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>')}${ser(1, 'delta', [4, 2], '<c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr>')}</c:barChart>`,
      ),
    )!
    expect(m.series[0]!.noFill).toBe(true)
    expect(m.series[1]!.noFill).toBeUndefined()
  })

  it('title size falls back to the rich paragraph defRPr when the run carries none', () => {
    const m = parseChartXml(
      space(
        `<c:barChart><c:barDir val="col"/>${ser(0, 'A', [1])}</c:barChart>`,
        '<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1200" b="1"/></a:pPr><a:r><a:rPr lang="en-IN"/><a:t>Cumulative</a:t></a:r></a:p></c:rich></c:tx></c:title>',
      ),
    )!
    expect(m.title).toBe('Cumulative')
    expect(m.titlePt).toBe(12)
  })

  it('series-level data labels win over the plot-level block; value+percent labels are flagged', () => {
    const m = parseChartXml(
      space(
        `<c:pieChart><c:varyColors val="1"/>${ser(0, 'S', [1, 3], '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/></c:dLbls>')}` +
          `<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/></c:dLbls></c:pieChart>`,
      ),
    )!
    expect(m.dataLabels).toBe(true)
    expect(m.dataLabelsValPct).toBe(true)
    expect(m.dataLabelsPct).toBeUndefined()
  })
})
