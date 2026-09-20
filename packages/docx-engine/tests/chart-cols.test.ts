/** Series workbook columns stay valid past Z, and multi-series parts keep clean refs. */
import { describe, expect, it } from 'vitest'
import { buildChartPartXml, colLetter, parseChartPartXml } from '../src/index'

describe('colLetter', () => {
  it('walks B through Z then continues with AA, AB', () => {
    expect(colLetter(0)).toBe('B')
    expect(colLetter(1)).toBe('C')
    expect(colLetter(24)).toBe('Z')
    expect(colLetter(25)).toBe('AA')
    expect(colLetter(26)).toBe('AB')
    expect(colLetter(51)).toBe('BA')
  })
})

describe('buildChartPartXml with 26+ series', () => {
  const series = Array.from({ length: 27 }, (_, i) => ({ name: `S${i}`, values: [i, i + 1] }))

  it('emits AA-style refs instead of bracket punctuation', () => {
    const xml = buildChartPartXml({
      kind: 'bar' as const,
      title: 'Wide',
      categories: ['Q1', 'Q2'],
      series,
    })
    expect(xml).toContain('Sheet1!$AA$1')
    expect(xml).toContain('Sheet1!$AA$2:$AA$3')
    expect(xml).not.toMatch(/Sheet1!\$[[\\]/)
  })

  it('round-trips every series name through our own parser', () => {
    const display = parseChartPartXml(
      buildChartPartXml({
        kind: 'bar' as const,
        title: 'Wide',
        categories: ['Q1', 'Q2'],
        series,
      }),
      'word/charts/chart1.xml',
    )!
    expect(display.series.map((s) => s.name)).toEqual(series.map((s) => s.name))
  })
})
