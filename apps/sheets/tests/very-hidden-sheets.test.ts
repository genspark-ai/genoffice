import { describe, expect, it } from 'vitest'

import type { WorkbookFile } from '../src/shared/desktop-api'
import { isVeryHiddenSheet, vetoesUnhide } from '../src/renderer/very-hidden-sheets'

const file = {
  sheets: [
    { id: 'a', hidden: false },
    { id: 'b', hidden: true },
    { id: 'c', hidden: true, veryHidden: true },
  ],
} as unknown as WorkbookFile

describe('veryHidden sheet guard', () => {
  it('flags only the veryHidden sheet', () => {
    expect(['a', 'b', 'c'].map((id) => isVeryHiddenSheet(file, id))).toEqual([false, false, true])
    expect(isVeryHiddenSheet(undefined, 'c')).toBe(false)
  })

  it('vetoes the show command for it and nothing else', () => {
    const show = (subUnitId: string) => ({
      id: 'sheet.command.set-worksheet-show',
      params: { unitId: 'u', subUnitId },
    })
    expect(vetoesUnhide(file, show('c'))).toBe(true)
    expect(vetoesUnhide(file, show('b'))).toBe(false)
    expect(
      vetoesUnhide(file, { id: 'sheet.command.set-worksheet-hide', params: { subUnitId: 'c' } }),
    ).toBe(false)
    expect(vetoesUnhide(file, { id: 'sheet.command.set-worksheet-show' })).toBe(false)
  })
})
