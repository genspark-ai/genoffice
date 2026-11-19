import { describe, expect, it } from 'vitest'
import { CellValueType } from '@univerjs/core'
import { cachedFormulaCellData } from '../src/renderer/univer-sync'

describe('cachedFormulaCellData', () => {
  it('types string caches so Univer does not coerce numeric-looking text', () => {
    expect(cachedFormulaCellData('=SWITCH(A1,"出勤","1")', '1')).toEqual({
      f: '=SWITCH(A1,"出勤","1")',
      v: '1',
      t: CellValueType.STRING,
    })
  })

  it('keeps numeric and boolean caches untyped', () => {
    expect(cachedFormulaCellData('=A1*2', 2)).toEqual({ f: '=A1*2', v: 2 })
    expect(cachedFormulaCellData('=A1>1', true)).toEqual({ f: '=A1>1', v: true })
  })

  it('drops missing caches entirely', () => {
    expect(cachedFormulaCellData('=A1', null)).toEqual({ f: '=A1' })
    expect(cachedFormulaCellData('=A1', undefined)).toEqual({ f: '=A1' })
  })
})
