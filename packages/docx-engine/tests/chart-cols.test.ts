/** Series workbook columns stay valid past Z, and multi-series parts keep clean refs. */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  buildChartPartXml,
  buildChartWorkbookXlsxBase64,
  colLetter,
  parseChartPartXml,
  patchChartWorkbookXlsxBase64,
} from '../src/index'

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

  it('uses base-26 columns in generated and patched embedded workbooks', async () => {
    const workbook = await buildChartWorkbookXlsxBase64(['Q1', 'Q2'], series)
    const zip = await JSZip.loadAsync(Buffer.from(workbook, 'base64'))
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
    expect(sheet).toContain('<c r="AA1"')
    expect(sheet).toContain('<c r="AB1"')
    expect(sheet).not.toMatch(/r="[\\[\]]/)

    const workbookZip = await JSZip.loadAsync(Buffer.from(workbook, 'base64'))
    const workbookSheet = await workbookZip.file('xl/worksheets/sheet1.xml')!.async('text')
    workbookZip.file(
      'xl/worksheets/sheet1.xml',
      workbookSheet.replace(/(<worksheet\b[^>]*>)/, '$1<dimension ref="A1:B3"/>'),
    )
    const withDimension = await workbookZip.generateAsync({ type: 'base64' })
    const patched = await patchChartWorkbookXlsxBase64(withDimension, ['Q1', 'Q2'], series)
    const patchedZip = await JSZip.loadAsync(Buffer.from(patched!, 'base64'))
    const patchedSheet = await patchedZip.file('xl/worksheets/sheet1.xml')!.async('text')
    expect(patchedSheet).toContain('dimension ref="A1:AB3"')
    expect(patchedSheet).not.toMatch(/r="[\\[\]]/)
  })
})
