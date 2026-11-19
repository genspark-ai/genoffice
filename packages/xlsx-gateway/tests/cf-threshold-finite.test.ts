import { describe, expect, it } from 'vitest'
import { applyCfRules, CfEditError, type DxfSink } from '../src/gateway/xlsx-cf'

const SHEET = '<worksheet><sheetData/></worksheet>'

const dxfs: DxfSink = { internDxf: () => 0 }

function colorScale(stopValue: unknown): string {
  return applyCfRules(
    SHEET,
    [
      {
        ranges: [{ startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 }],
        stopIfTrue: false,
        rule: {
          type: 'colorScale',
          config: [
            { index: 0, value: { type: 'min' } },
            { index: 1, value: { type: 'num', value: stopValue } },
          ],
        },
      },
    ],
    dxfs,
  )
}

describe('CF threshold finiteness', () => {
  it('rejects a non-numeric threshold instead of writing val="NaN"', () => {
    expect(() => colorScale('abc')).toThrow(CfEditError)
    expect(() => colorScale('abc')).toThrow(/finite/)
    expect(() => colorScale(Number.NaN)).toThrow(CfEditError)
    expect(() => colorScale(Number.POSITIVE_INFINITY)).toThrow(CfEditError)
  })

  it('never emits val="NaN" or val="Infinity" for any rejected threshold', () => {
    for (const bad of ['abc', Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY]) {
      let xml = ''
      try {
        xml = colorScale(bad)
      } catch {
        // throwing is the expected outcome; nothing was written
      }
      expect(xml).not.toContain('val="NaN"')
      expect(xml).not.toContain('val="Infinity"')
      expect(xml).not.toContain('val="-Infinity"')
    }
  })

  it('still writes a finite threshold', () => {
    const xml = colorScale(75)
    expect(xml).toContain('<cfvo type="num" val="75"/>')
    expect(xml).toContain('<cfvo type="min"/>')
  })

  it('still writes a formula threshold verbatim', () => {
    const xml = applyCfRules(
      SHEET,
      [
        {
          ranges: [{ startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 }],
          stopIfTrue: false,
          rule: {
            type: 'colorScale',
            config: [
              { index: 0, value: { type: 'formula', value: '=A1*2' } },
              { index: 1, value: { type: 'max' } },
            ],
          },
        },
      ],
      dxfs,
    )
    expect(xml).toContain('<cfvo type="formula" val="=A1*2"/>')
  })
})
