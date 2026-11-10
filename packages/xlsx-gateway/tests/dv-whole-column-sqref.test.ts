import { describe, expect, it } from 'vitest'
import { applyDvRules, DvEditError } from '../src/gateway/xlsx-dv'
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from '../src/shared/grid-bounds'

const SHEET = '<worksheet><sheetData/></worksheet>'

function withExisting(sqref: string): string {
  return (
    `${SHEET}<dataValidations count="1">` +
    `<dataValidation type="list" sqref="${sqref}"><formula1>"old"</formula1></dataValidation>` +
    '</dataValidations>'
  )
}

const NEW_RULE = { type: 'list', formula1: '"new"' }

describe('xlsx-dv whole-column / whole-row sqref matching', () => {
  it('replaces an existing A:A rule when appending a full-column rule', () => {
    const xml = applyDvRules(
      withExisting('A:A'),
      [
        {
          ranges: [{ startRow: 0, endRow: MAX_GRID_ROWS - 1, startColumn: 0, endColumn: 0 }],
          rule: NEW_RULE,
        },
      ],
      { append: true },
    )
    expect(xml.match(/<dataValidation\b/g)).toHaveLength(1)
    expect(xml).not.toContain('"old"')
    expect(xml).toContain(`sqref="A1:A${MAX_GRID_ROWS}"`)
  })

  it('replaces an existing 1:1 rule when appending a full-row rule', () => {
    const xml = applyDvRules(
      withExisting('1:1'),
      [
        {
          ranges: [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: MAX_GRID_COLUMNS - 1 }],
          rule: NEW_RULE,
        },
      ],
      { append: true },
    )
    expect(xml.match(/<dataValidation\b/g)).toHaveLength(1)
    expect(xml).toContain('sqref="A1:XFD1"')
  })

  it('keeps an existing A:A rule when the new rule covers only part of the column', () => {
    const xml = applyDvRules(
      withExisting('A:A'),
      [{ ranges: [{ startRow: 0, endRow: 9, startColumn: 0, endColumn: 0 }], rule: NEW_RULE }],
      { append: true },
    )
    expect(xml.match(/<dataValidation\b/g)).toHaveLength(2)
  })

  it('rejects ranges outside the worksheet grid', () => {
    for (const range of [
      { startRow: 0, endRow: MAX_GRID_ROWS, startColumn: 0, endColumn: 0 },
      { startRow: 0, endRow: 0, startColumn: 0, endColumn: MAX_GRID_COLUMNS },
      { startRow: 0, endRow: 0, startColumn: -1, endColumn: 0 },
      { startRow: 5, endRow: 2, startColumn: 0, endColumn: 0 },
    ]) {
      expect(() => applyDvRules(SHEET, [{ ranges: [range], rule: NEW_RULE }])).toThrow(DvEditError)
    }
  })
})
