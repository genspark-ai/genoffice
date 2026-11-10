import { describe, expect, it } from 'vitest'

import { handleApplyFormula, type DataToolsContext } from '../src/renderer/data-tools-actions'

function makeHarness() {
  const written: { f: string }[] = []
  const messages: string[] = []
  const worksheet = {
    getRange: () => ({
      setValue: (value: { f: string }) => {
        written.push(value)
      },
    }),
  }
  const workbook = {
    getActiveSheet: () => worksheet,
    getActiveRange: () => ({ getRow: () => 0, getColumn: () => 0 }),
  }
  const ctx = {
    univerRef: { current: { univerAPI: { getActiveWorkbook: () => workbook } } },
    setMessage: (message: string) => {
      messages.push(message)
    },
  } as unknown as DataToolsContext
  return { ctx, written, messages }
}

describe('handleApplyFormula paren balance', () => {
  it('counts parens outside string literals and quoted sheet names', () => {
    // A paren inside "..." or a quoted sheet name is text. Counting it made the
    // opens/closes compare unequal, so these valid formulas were rejected with
    // appUnbalancedParens and nothing was written to the cell.
    const accepted = [
      '=SUM(A1:A5)',
      '=IF(A1>0,"ok","no")',
      '=IF(A1=")",1,0)',
      '=LEN("a)b")',
      '=MID(A1,2,1)&"("&B1',
      '=LEN("a""(b")',
      "='Q1 (copy)'!A1",
    ]
    for (const formula of accepted) {
      const { ctx, written } = makeHarness()
      expect(handleApplyFormula(ctx, formula), formula).toBeNull()
      expect(written, formula).toEqual([{ f: formula }])
    }
  })

  it('still rejects a formula that is genuinely unbalanced', () => {
    for (const formula of ['=SUM(A1,A2', '=LEN("a)b"', '=IF(A1>0,"ok"']) {
      const { ctx, written } = makeHarness()
      expect(handleApplyFormula(ctx, formula), formula).not.toBeNull()
      expect(written, formula).toEqual([])
    }
  })
})
