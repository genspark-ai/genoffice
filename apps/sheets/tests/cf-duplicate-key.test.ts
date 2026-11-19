import { BooleanNumber, CellValueType } from '@univerjs/core'
import { getCellValue } from '@univerjs/preset-sheets-conditional-formatting'
import { describe, expect, it } from 'vitest'

import { cfDisplayKey, wrapCalculateUnit, wrapCfViewModel } from '../src/renderer/cf-duplicate-key'

describe('cfDisplayKey', () => {
  it('keys a number and its text form identically', () => {
    expect(cfDisplayKey({ v: 1981233, t: CellValueType.NUMBER })).toBe('1981233')
    expect(cfDisplayKey({ v: '1981233', t: CellValueType.STRING })).toBe('1981233')
  })

  it('folds text case like Excel', () => {
    expect(cfDisplayKey({ v: 'Short' })).toBe(cfDisplayKey({ v: 'SHORT' }))
  })

  it('keeps leading zeros and whitespace distinct', () => {
    expect(cfDisplayKey({ v: '007' })).not.toBe(cfDisplayKey({ v: 7 }))
    expect(cfDisplayKey({ v: 'a ' })).not.toBe(cfDisplayKey({ v: 'a' }))
  })

  it('hides double noise the way General display does', () => {
    expect(cfDisplayKey({ v: 0.1 + 0.2 })).toBe(cfDisplayKey({ v: 0.3 }))
    expect(cfDisplayKey({ v: 1.5 })).not.toBe(cfDisplayKey({ v: 2 }))
  })

  it('treats booleans as their display text', () => {
    expect(cfDisplayKey({ v: BooleanNumber.TRUE, t: CellValueType.BOOLEAN })).toBe('TRUE')
    expect(cfDisplayKey({ v: 'true' })).toBe('TRUE')
  })

  it('reads rich text bodies and ignores blanks', () => {
    expect(cfDisplayKey({ p: { body: { dataStream: 'Abc\r\n' } } } as never)).toBe('ABC')
    expect(cfDisplayKey(null)).toBeNull()
    expect(cfDisplayKey({})).toBeNull()
    expect(cfDisplayKey({ v: '  ' })).toBeNull()
  })
})

type Unit = NonNullable<Parameters<typeof wrapCalculateUnit>[0]>

function fakeUnit(subType: string, cells: Record<string, unknown>): Unit {
  return {
    _rule: { rule: { type: 'highlightCell', subType } },
    _context: { getCellValue: (row, col) => (cells[`${row}_${col}`] ?? {}) as never },
  }
}

describe('wrapCalculateUnit', () => {
  const cells = {
    '0_0': { v: 1981233, t: CellValueType.NUMBER },
    '1_0': { v: '1981233', t: CellValueType.STRING },
    '2_0': { v: 'x' },
  }

  it('hands Univer the display key for duplicate/unique rules', () => {
    const unit = wrapCalculateUnit(fakeUnit('duplicateValues', cells))!
    const read = (row: number) => getCellValue(unit._context!.getCellValue(row, 0))
    expect(read(0)).toBe(read(1))
    expect(read(2)).toBe('X')
  })

  it('leaves other highlight sub-types on the raw cell', () => {
    const unit = wrapCalculateUnit(fakeUnit('number', cells))!
    expect(unit._context!.getCellValue(0, 0)).toBe(cells['0_0'])
  })

  it('follows a later sub-type change on the same unit', () => {
    const unit = wrapCalculateUnit(fakeUnit('number', cells))!
    unit._rule = { rule: { type: 'highlightCell', subType: 'uniqueValues' } }
    expect(getCellValue(unit._context!.getCellValue(0, 0))).toBe('1981233')
  })

  it('maps blanks to an empty cell so Univer skips them', () => {
    const unit = wrapCalculateUnit(fakeUnit('duplicateValues', { '0_0': { v: ' ' } }))!
    expect(getCellValue(unit._context!.getCellValue(0, 0))).toBeNull()
  })
})

describe('wrapCfViewModel', () => {
  it('wraps highlight units only and restores on dispose', () => {
    const created: Unit[] = []
    const original = function (
      this: unknown,
      _u: string,
      _s: string,
      rule: { rule: { type: string } },
    ) {
      const unit = fakeUnit(rule.rule.type === 'highlightCell' ? 'duplicateValues' : '', {
        '0_0': { v: 42 },
      })
      created.push(unit)
      return unit
    }
    const viewModel = { _createRuleCalculateUnitInstance: original }
    const handle = wrapCfViewModel(viewModel)
    const highlight = viewModel._createRuleCalculateUnitInstance('u', 's', {
      rule: { type: 'highlightCell' },
    })!
    const bar = viewModel._createRuleCalculateUnitInstance('u', 's', { rule: { type: 'dataBar' } })!
    expect(getCellValue(highlight._context!.getCellValue(0, 0))).toBe('42')
    expect(getCellValue(bar._context!.getCellValue(0, 0))).toBe(42)
    handle.dispose()
    expect(viewModel._createRuleCalculateUnitInstance).toBe(original)
    expect(created).toHaveLength(2)
  })
})
