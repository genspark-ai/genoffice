import { describe, expect, it } from 'vitest'
import {
  anchorOtherFormulaTrees,
  type OtherFormulaTreeLike,
} from '../src/renderer/other-formula-anchor'

const tree = (over: Partial<OtherFormulaTreeLike>): OtherFormulaTreeLike => ({
  isVirtual: false,
  formulaId: 'cf-1',
  unitId: 'u',
  subUnitId: 's',
  row: -1,
  column: -1,
  ...over,
})

const data = {
  u: {
    s: {
      'cf-1': {
        ranges: [
          { startRow: 0, startColumn: 0 },
          { startRow: 5, startColumn: 2 },
        ],
      },
    },
  },
}

describe('anchorOtherFormulaTrees', () => {
  it('moves the first tree of an other formula to its first range top-left', () => {
    const list = [tree({})]
    anchorOtherFormulaTrees(list, data, 0)
    expect(list[0]).toMatchObject({ row: 0, column: 0 })
  })

  it('leaves virtual, sheet-cell, already-anchored and unknown trees alone', () => {
    const list = [
      tree({ row: 3, column: 3 }),
      tree({ isVirtual: true }),
      tree({ formulaId: null }),
      tree({ formulaId: 'missing' }),
      tree({ unitId: 'other' }),
    ]
    anchorOtherFormulaTrees(list, data, 1)
    expect(list.slice(1).every((t) => t.row === -1 && t.column === -1)).toBe(true)
    anchorOtherFormulaTrees(list, data, 0)
    expect(list[0]).toMatchObject({ row: 3, column: 3 })
  })

  it('only touches trees appended after `from`', () => {
    const list = [tree({}), tree({})]
    anchorOtherFormulaTrees(list, data, 1)
    expect(list.map((t) => t.row)).toEqual([-1, 0])
  })
})
