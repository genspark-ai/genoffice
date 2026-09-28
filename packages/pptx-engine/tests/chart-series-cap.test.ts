import { describe, expect, it } from 'vitest'
import { parseChartXml } from '../src/chart'

const ser = (i: number, points: number): string => {
  const vals = Array.from({ length: points }, (_, p) => `<c:pt idx="${p}"><c:v>${p}</c:v></c:pt>`)
  return (
    `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:v>S${i}</c:v></c:tx>` +
    `<c:val><c:numRef><c:numCache><c:ptCount val="${points}"/>${vals.join('')}</c:numCache></c:numRef></c:val></c:ser>`
  )
}

const chartXml = (series: string, plot = 'c:lineChart'): string =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
  `<c:chart><c:plotArea><${plot}><c:grouping val="standard"/>${series}</${plot}>` +
  '<c:catAx><c:axId val="1"/><c:axId val="2"/></c:catAx>' +
  '<c:valAx><c:axId val="2"/></c:valAx>' +
  '</c:plotArea></c:chart></c:chartSpace>'

describe('chart series count cap', () => {
  it('parses a normal chart unchanged', () => {
    const model = parseChartXml(chartXml(ser(0, 3) + ser(1, 3) + ser(2, 4)))!
    expect(model.series).toHaveLength(3)
    expect(model.categories).toHaveLength(4)
  })

  it('caps an oversized series array instead of spreading it', () => {
    let body = ''
    for (let i = 0; i < 400; i++) body += ser(i, 2)
    const model = parseChartXml(chartXml(body))!
    expect(model.series).toHaveLength(256)
    expect(model.categories).toHaveLength(2)
  })

  it('caps per plot, not per chart', () => {
    let first = ''
    let second = ''
    for (let i = 0; i < 200; i++) first += ser(i, 2)
    for (let i = 0; i < 200; i++) second += ser(i, 2)
    const model = parseChartXml(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
        '<c:chart><c:plotArea><c:barChart><c:grouping val="clustered"/>' +
        first +
        '</c:barChart><c:lineChart><c:grouping val="standard"/>' +
        second +
        '</c:lineChart>' +
        '<c:catAx><c:axId val="1"/><c:axId val="2"/></c:catAx><c:valAx><c:axId val="2"/></c:valAx>' +
        '</c:plotArea></c:chart></c:chartSpace>',
    )!
    expect(model.series).toHaveLength(256)
    expect(model.series.filter((s) => s.plotKind === 'line')).toHaveLength(56)
  })
})
