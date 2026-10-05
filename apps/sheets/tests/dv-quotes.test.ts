import { describe, expect, it } from 'vitest'

import { applyDvRules } from '@genoffice/xlsx-gateway/gateway/xlsx-dv'
import { toUniverDvRule } from '../src/renderer/univer-sync'

const SHEET =
  '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>'

describe('dv quotes', () => {
  it('doubles embedded quotes in list literals', () => {
    const xml = applyDvRules(SHEET, [
      {
        ranges: [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
        rule: { type: 'list', formula1: 'a"b,c' },
      },
    ])
    expect(xml).toContain('<formula1>"a""b,c"</formula1>')
  })

  it('un-doubles on read, so a save/reopen cycle is stable', () => {
    // the reader must undo the writer, or every cycle grows the item:
    // a"b -> a""b -> a""""b
    const rule = toUniverDvRule(
      {
        ruleType: 'list',
        formulas: ['"a""b,c"'],
        ranges: [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
      } as never,
      'uid-1',
    )
    expect((rule as { formula1: string }).formula1).toBe('a"b,c')
  })
})
