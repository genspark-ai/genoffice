import { describe, expect, it } from 'vitest'

import {
  dataBarNeedsLayout,
  dataBarNeedsValues,
  defaultThreshold,
  emulateBarExtents,
  evaluateThresholdFormula,
  isSelfContainedFormula,
  layoutDataBar,
  middleAxisBounds,
  percentileInc,
  resolveBarBound,
  THRESHOLD_RANGE_CELL_CAP,
  type ThresholdReader,
} from '../src/renderer/cf-thresholds'

/// Sheet "Data" (the rule's own sheet): B2:B9 = 10..80 step 10, T5 = 42,
/// C1 = "text". Sheet "Other": A1 = 7. Table "Sales" with column "Amount"
/// over Data!D2:D4 = 1, 2, 3.
function reader(log: string[] = []): ThresholdReader {
  const cell = (sheet: string, row: number, column: number): number | undefined => {
    if (sheet === 'other') return row === 0 && column === 0 ? 7 : undefined
    if (sheet !== 'data') return undefined
    if (column === 1 && row >= 1 && row <= 8) return row * 10
    if (column === 19 && row === 4) return 42
    if (column === 2 && row === 0) return Number.NaN
    if (column === 3 && row >= 1 && row <= 3) return row
    return undefined
  }
  return {
    async readValues(sheetName, range) {
      const sheet = (sheetName ?? 'Data').toLowerCase()
      log.push(`${sheet}:${range.startRow}-${range.endRow}/${range.startColumn}-${range.endColumn}`)
      if (sheet !== 'data' && sheet !== 'other') return null
      const values: number[] = []
      for (let row = range.startRow; row <= range.endRow; row += 1) {
        for (let column = range.startColumn; column <= range.endColumn; column += 1) {
          const value = cell(sheet, row, column)
          if (value !== undefined) values.push(value)
        }
      }
      return values
    },
    definedName(name) {
      if (name.toLowerCase() === 'threshold') return '$T$5*2'
      if (name.toLowerCase() === 'nested') return 'Threshold+1'
      return null
    },
    tableColumn(table, column) {
      if (table !== 'Sales' || column !== 'Amount') return null
      return {
        sheetName: 'Data',
        range: { startRow: 1, endRow: 3, startColumn: 3, endColumn: 3 },
      }
    },
  }
}

describe('evaluateThresholdFormula', () => {
  it('reads a bare absolute cell reference on the rule sheet', async () => {
    // NewStyleConditionalFormattings.xlsx T2:T17 icon set threshold.
    expect(await evaluateThresholdFormula('$T$5', reader())).toBe(42)
    expect(await evaluateThresholdFormula('=$T$5', reader())).toBe(42)
  })

  it('reads sheet-qualified references, quoted or bare', async () => {
    expect(await evaluateThresholdFormula('Other!$A$1', reader())).toBe(7)
    expect(await evaluateThresholdFormula("'Other'!$A$1+1", reader())).toBe(8)
  })

  it('folds aggregates over absolute ranges', async () => {
    expect(await evaluateThresholdFormula('AVERAGE($B$2:$B$9)', reader())).toBe(45)
    expect(await evaluateThresholdFormula('SUM($B$2:$B$9)', reader())).toBe(360)
    expect(await evaluateThresholdFormula('MIN($B$2:$B$9)', reader())).toBe(10)
    expect(await evaluateThresholdFormula('MAX($B$2:$B$9)', reader())).toBe(80)
    expect(await evaluateThresholdFormula('MEDIAN($B$2:$B$9)', reader())).toBe(45)
    expect(await evaluateThresholdFormula('COUNT($B$2:$B$9)', reader())).toBe(8)
    expect(await evaluateThresholdFormula('PERCENTILE($B$2:$B$9,0.5)', reader())).toBe(45)
    expect(await evaluateThresholdFormula('PERCENTILE.INC($B$2:$B$9, 0.25)', reader())).toBe(27.5)
    expect(await evaluateThresholdFormula('QUARTILE($B$2:$B$9,3)', reader())).toBe(62.5)
    expect(await evaluateThresholdFormula('average(Other!$A$1:$A$1)', reader())).toBe(7)
  })

  it('folds MIN / MAX over a range at the cell cap without spreading', async () => {
    // 512k populated cells exceed the engine's call-argument limit; a
    // spread would throw and leave the formula stop in place.
    const values = Array.from({ length: THRESHOLD_RANGE_CELL_CAP }, (_, i) => i - 1000)
    const wide: ThresholdReader = {
      ...reader(),
      readValues: async () => values,
    }
    expect(await evaluateThresholdFormula('MIN($A$1:$A$512000)', wide)).toBe(-1000)
    expect(await evaluateThresholdFormula('MAX($A$1:$A$512000)', wide)).toBe(
      THRESHOLD_RANGE_CELL_CAP - 1001,
    )
  })

  it('composes aggregates, references and arithmetic', async () => {
    expect(await evaluateThresholdFormula('AVERAGE($B$2:$B$9)*1.5', reader())).toBe(67.5)
    expect(await evaluateThresholdFormula('MAX($B$2:$B$9)-$T$5', reader())).toBe(38)
    expect(await evaluateThresholdFormula('($T$5-2)/4', reader())).toBe(10)
    expect(await evaluateThresholdFormula('2*Other!$A$1+3', reader())).toBe(17)
  })

  it('treats relative references as 0, inside aggregates too', async () => {
    // databar.xlsx H3:H6 max threshold 3*A1+2 -> 2.
    expect(await evaluateThresholdFormula('3*A1+2', reader())).toBe(2)
    expect(await evaluateThresholdFormula('AVERAGE(B2:B9)', reader())).toBe(0)
    expect(await evaluateThresholdFormula('AVERAGE($B$2:B9)+5', reader())).toBe(5)
    expect(await evaluateThresholdFormula('Other!A1', reader())).toBe(0)
  })

  it('resolves defined names, nested up to a small depth', async () => {
    expect(await evaluateThresholdFormula('Threshold', reader())).toBe(84)
    expect(await evaluateThresholdFormula('Nested/5', reader())).toBe(17)
  })

  it('sums a table column through a structured reference', async () => {
    expect(await evaluateThresholdFormula('SUM(Sales[Amount])', reader())).toBe(6)
    expect(await evaluateThresholdFormula('AVERAGE(Sales[Amount])*2', reader())).toBe(4)
  })

  it('reads a blank cell as 0 and a text cell as unevaluable', async () => {
    expect(await evaluateThresholdFormula('$Z$9+1', reader())).toBe(1)
    expect(await evaluateThresholdFormula('$C$1+1', reader())).toBeNull()
  })

  it('returns null for shapes it cannot fold', async () => {
    expect(await evaluateThresholdFormula('TODAY()-30', reader())).toBeNull()
    expect(await evaluateThresholdFormula('Unknown+1', reader())).toBeNull()
    expect(await evaluateThresholdFormula('"12"', reader())).toBeNull()
    expect(await evaluateThresholdFormula('AVERAGE($B$2:$B$9,$T$5)', reader())).toBeNull()
    expect(await evaluateThresholdFormula('AVERAGE($Z$1:$Z$9)', reader())).toBeNull()
    expect(await evaluateThresholdFormula('AVERAGE(Missing!$A$1:$A$9)', reader())).toBeNull()
    expect(await evaluateThresholdFormula('', reader())).toBeNull()
  })

  it('folds a numeric literal without touching the sheet', async () => {
    const log: string[] = []
    expect(await evaluateThresholdFormula('0', reader(log))).toBe(0)
    expect(await evaluateThresholdFormula('2.5', reader(log))).toBe(2.5)
    expect(log).toEqual([])
  })
})

describe('isSelfContainedFormula', () => {
  it('accepts literal / function-only formulas', () => {
    expect(isSelfContainedFormula('TODAY()-30')).toBe(true)
    expect(isSelfContainedFormula('=DATE(2024,1,1)+LOG10(100)')).toBe(true)
    expect(isSelfContainedFormula('1E3*2')).toBe(true)
    expect(isSelfContainedFormula('LEN("A1")')).toBe(true)
  })

  it('rejects anything that reads the workbook', () => {
    expect(isSelfContainedFormula('$T$5')).toBe(false)
    expect(isSelfContainedFormula('A1+1')).toBe(false)
    expect(isSelfContainedFormula('Other!A1')).toBe(false)
    expect(isSelfContainedFormula('MyName*2')).toBe(false)
    expect(isSelfContainedFormula('SUM(Sales[Amount])')).toBe(false)
    expect(isSelfContainedFormula('TRUE')).toBe(false)
  })
})

describe('defaultThreshold', () => {
  it('gives the slot the dialog default', () => {
    expect(defaultThreshold('dataBar', 0, 2)).toEqual({ kind: 'min' })
    expect(defaultThreshold('dataBar', 1, 2)).toEqual({ kind: 'max' })
    expect(defaultThreshold('colorScale', 1, 3)).toEqual({ kind: 'percentile', value: '50' })
    expect(defaultThreshold('colorScale', 2, 3)).toEqual({ kind: 'max' })
    expect(defaultThreshold('iconSet', 1, 3)).toEqual({ kind: 'percent', value: '33' })
    expect(defaultThreshold('iconSet', 2, 3)).toEqual({ kind: 'percent', value: '67' })
    expect(defaultThreshold('iconSet', 3, 4)).toEqual({ kind: 'percent', value: '75' })
  })
})

describe('percentileInc', () => {
  it('interpolates like PERCENTILE.INC', () => {
    expect(percentileInc([1, 2, 3, 4], 0.5)).toBe(2.5)
    expect(percentileInc([1, 2, 3, 4], 0)).toBe(1)
    expect(percentileInc([1, 2, 3, 4], 1)).toBe(4)
    expect(percentileInc([5], 0.3)).toBe(5)
    expect(percentileInc([], 0.5)).toBeNull()
    expect(percentileInc([1, 2], 1.5)).toBeNull()
  })
})

describe('resolveBarBound', () => {
  const values = [1, 2, 3, 4, Number.NaN]
  it('maps every cfvo kind onto the data', () => {
    expect(resolveBarBound({ kind: 'num', value: '7' }, null, 'min')).toBe(7)
    expect(resolveBarBound({ kind: 'min' }, values, 'min')).toBe(1)
    expect(resolveBarBound({ kind: 'max' }, values, 'max')).toBe(4)
    expect(resolveBarBound({ kind: 'autoMin' }, values, 'min')).toBe(0)
    expect(resolveBarBound({ kind: 'autoMax' }, values, 'max')).toBe(4)
    expect(resolveBarBound({ kind: 'autoMax' }, [-3, -1], 'max')).toBe(0)
    expect(resolveBarBound({ kind: 'percent', value: '50' }, values, 'min')).toBe(2.5)
    expect(resolveBarBound({ kind: 'percentile', value: '50' }, values, 'min')).toBe(2.5)
  })

  it('is null without data or for unresolved kinds', () => {
    expect(resolveBarBound({ kind: 'min' }, null, 'min')).toBeNull()
    expect(resolveBarBound({ kind: 'min' }, [Number.NaN], 'min')).toBeNull()
    expect(resolveBarBound({ kind: 'formula', value: 'TODAY()' }, values, 'min')).toBeNull()
    expect(resolveBarBound({ kind: 'num', value: 'abc' }, values, 'min')).toBeNull()
  })
})

describe('emulateBarExtents', () => {
  it('widens the bounds so Univer 0..100% lands on Excel minLength..maxLength', () => {
    // Legacy 10/90 over values 1000..5000.
    const bounds = emulateBarExtents(1000, 5000, 10, 90)!
    expect(bounds).not.toBeNull()
    const univer = (value: number) => (value - bounds.min) / (bounds.max - bounds.min)
    expect(univer(1000)).toBeCloseTo(0.1, 12)
    expect(univer(5000)).toBeCloseTo(0.9, 12)
    expect(univer(3000)).toBeCloseTo(0.5, 12)
    // Explicit 2006 attributes.
    const wide = emulateBarExtents(100, 200, 20, 80)!
    expect((100 - wide.min) / (wide.max - wide.min)).toBeCloseTo(0.2, 12)
    expect((200 - wide.min) / (wide.max - wide.min)).toBeCloseTo(0.8, 12)
  })

  it('is null for x14 defaults and when the lowered bound would cross zero', () => {
    expect(emulateBarExtents(1, 4, 0, 100)).toBeNull()
    // 0..100 with 10/90 needs a lower bound of -12.5: Univer would draw a
    // two-sided axis layout, so keep its native bar.
    expect(emulateBarExtents(0, 100, 10, 90)).toBeNull()
    expect(emulateBarExtents(100, 1000, 10, 90)).toBeNull()
    // Boundary: min exactly one eighth of the span is representable.
    expect(emulateBarExtents(125, 1125, 10, 90)).toEqual({ min: 0, max: 1250 })
  })

  it('rejects degenerate inputs', () => {
    expect(emulateBarExtents(5, 5, 10, 90)).toBeNull()
    expect(emulateBarExtents(1, 4, 90, 10)).toBeNull()
    expect(emulateBarExtents(1, 4, -5, 90)).toBeNull()
  })
})

describe('middleAxisBounds', () => {
  it('centres the axis with one shared scale for two-signed data', () => {
    expect(middleAxisBounds(-2, 4)).toEqual({ min: -4, max: 4 })
    expect(middleAxisBounds(-10, 3)).toEqual({ min: -10, max: 10 })
  })

  it('leaves one-signed data alone (Excel draws it from the cell edge)', () => {
    // databar.xlsx B10:B13: axisPosition="middle" over 1..4 renders as
    // plain left-anchored bars in Excel.
    expect(middleAxisBounds(0, 4)).toBeNull()
    expect(middleAxisBounds(-4, 0)).toBeNull()
    expect(middleAxisBounds(1, 4)).toBeNull()
  })
})

describe('layoutDataBar', () => {
  const minMax = [{ kind: 'min' }, { kind: 'max' }]

  it('leaves a plain x14-style bar untouched without reading values', () => {
    const input = { cfvos: minMax, minLength: 0, maxLength: 100 }
    expect(dataBarNeedsLayout(input)).toBe(false)
    expect(dataBarNeedsValues(input)).toBe(false)
    expect(layoutDataBar(input, null)).toBeNull()
  })

  it('leaves a bar from an older sidecar (no extents on the wire) untouched', () => {
    expect(dataBarNeedsLayout({ cfvos: minMax })).toBe(false)
  })

  it('resolves x14 autoMin/autoMax to zero-anchored numbers', () => {
    const input = {
      cfvos: [{ kind: 'autoMin' }, { kind: 'autoMax' }],
      minLength: 0,
      maxLength: 100,
    }
    expect(dataBarNeedsValues(input)).toBe(true)
    expect(layoutDataBar(input, [1, 2, 3, 4])).toEqual([
      { kind: 'num', value: '0' },
      { kind: 'num', value: '4' },
    ])
    // One side numeric: only the auto side needs data.
    expect(
      layoutDataBar(
        {
          cfvos: [{ kind: 'autoMin' }, { kind: 'num', value: '10' }],
          minLength: 0,
          maxLength: 100,
        },
        [3, 5],
      ),
    ).toEqual([
      { kind: 'num', value: '0' },
      { kind: 'num', value: '10' },
    ])
  })

  it('keeps the file cfvos when the values are unavailable', () => {
    const input = {
      cfvos: [{ kind: 'autoMin' }, { kind: 'autoMax' }],
      minLength: 0,
      maxLength: 100,
    }
    expect(layoutDataBar(input, null)).toBeNull()
    expect(layoutDataBar({ cfvos: minMax, minLength: 10, maxLength: 90 }, null)).toBeNull()
  })

  it('emulates legacy 10/90 extents when representable', () => {
    const cfvos = layoutDataBar(
      { cfvos: minMax, minLength: 10, maxLength: 90 },
      [1000, 3000, 5000],
    )!
    expect(cfvos.map((cfvo) => cfvo.kind)).toEqual(['num', 'num'])
    const [low, high] = cfvos.map((cfvo) => Number(cfvo.value)) as [number, number]
    expect((1000 - low) / (high - low)).toBeCloseTo(0.1, 12)
    expect((5000 - low) / (high - low)).toBeCloseTo(0.9, 12)
    // percent 0 / percent 100 resolve to the data extremes first.
    const percent = layoutDataBar(
      {
        cfvos: [
          { kind: 'percent', value: '0' },
          { kind: 'percent', value: '100' },
        ],
        minLength: 10,
        maxLength: 90,
      },
      [1000, 5000],
    )!
    expect(percent.map((cfvo) => Number(cfvo.value))).toEqual([500, 5500])
  })

  it('keeps the native bar when the legacy extents would need a negative bound', () => {
    // num 0 / max over 0..100: lowering min to -12.5 would flip Univer into
    // its two-sided layout.
    expect(
      layoutDataBar(
        { cfvos: [{ kind: 'num', value: '0' }, { kind: 'max' }], minLength: 10, maxLength: 90 },
        [0, 50, 100],
      ),
    ).toBeNull()
  })

  it('centres the axis for a two-signed middle bar and leaves one-signed alone', () => {
    const middle = { cfvos: [{ kind: 'autoMin' }, { kind: 'autoMax' }], axisPosition: 'middle' }
    expect(layoutDataBar({ ...middle, minLength: 0, maxLength: 100 }, [-2, 1, 4])).toEqual([
      { kind: 'num', value: '-4' },
      { kind: 'num', value: '4' },
    ])
    // databar.xlsx B10:B13: Excel keeps left-anchored bars for 1..4.
    expect(layoutDataBar({ ...middle, minLength: 0, maxLength: 100 }, [1, 2, 3, 4])).toEqual([
      { kind: 'num', value: '0' },
      { kind: 'num', value: '4' },
    ])
    expect(
      layoutDataBar(
        { cfvos: minMax, axisPosition: 'middle', minLength: 0, maxLength: 100 },
        [-3, 6],
      ),
    ).toEqual([
      { kind: 'num', value: '-6' },
      { kind: 'num', value: '6' },
    ])
  })
})
