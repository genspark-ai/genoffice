import { CellValueType } from '@univerjs/core'
import { AUTO_FILL_APPLY_TYPE } from '@univerjs/sheets'
import { describe, expect, it } from 'vitest'

import { invertedFillType } from '../src/renderer/ctrl-drag-fill'

const one = { rows: [0], cols: [0] }

describe('invertedFillType', () => {
  it('turns a lone number into a series, which the default drag would only copy', () => {
    expect(invertedFillType(one, () => ({ v: 1008 }))).toBe(AUTO_FILL_APPLY_TYPE.SERIES)
    expect(invertedFillType(one, () => ({ v: '7', t: CellValueType.NUMBER }))).toBe(
      AUTO_FILL_APPLY_TYPE.SERIES,
    )
  })

  it('copies everything the default drag would extend', () => {
    expect(invertedFillType({ rows: [0, 1], cols: [0] }, () => ({ v: 1 }))).toBe(
      AUTO_FILL_APPLY_TYPE.COPY,
    )
    expect(invertedFillType(one, () => ({ v: 'Item 1' }))).toBe(AUTO_FILL_APPLY_TYPE.COPY)
    expect(invertedFillType(one, () => null)).toBe(AUTO_FILL_APPLY_TYPE.COPY)
    expect(invertedFillType(one, () => ({ v: 3, f: '=A1+2' }))).toBe(AUTO_FILL_APPLY_TYPE.COPY)
    expect(invertedFillType(one, () => ({ v: 3, si: 'shared-1' }))).toBe(AUTO_FILL_APPLY_TYPE.COPY)
  })
})
