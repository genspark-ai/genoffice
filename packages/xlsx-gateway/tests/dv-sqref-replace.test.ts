import { describe, expect, it } from 'vitest'
import { applyDvRules } from '../src/gateway/xlsx-dv'

const SHEET = '<worksheet><sheetData/></worksheet>'

function appendOver(remove: {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}) {
  const existing =
    '<dataValidations count="1">' +
    '<dataValidation type="list" sqref="A1 $&amp; B1"><formula1>"x"</formula1></dataValidation>' +
    '</dataValidations>'
  return applyDvRules(
    `${SHEET}${existing}`,
    [
      {
        ranges: [remove],
        rule: { type: 'list', formula1: '"new"' },
      },
    ],
    { append: true },
  )
}

describe('xlsx-dv sqref rewrite', () => {
  it('inserts a surviving `$&` area literally instead of expanding it to the whole match', () => {
    const xml = appendOver({ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 })

    // The surviving areas are written back verbatim: "$&" and "B1".
    expect(xml).toContain('sqref="$&amp; B1"')
    // The whole-match expansion must NOT appear — that is the corruption.
    expect(xml).not.toContain('sqref="A1 $&amp; B1" B1"')
    const tag = /<dataValidation\b[^>]*>/.exec(xml)?.[0] ?? ''
    expect(tag).toContain('sqref="$&amp; B1"')
    expect(tag.endsWith('>')).toBe(true)
  })

  it('keeps a surviving area containing $1 literal', () => {
    const existing =
      '<dataValidations count="1">' +
      '<dataValidation type="list" sqref="A1 $1"><formula1>"x"</formula1></dataValidation>' +
      '</dataValidations>'
    const xml = applyDvRules(
      `${SHEET}${existing}`,
      [
        {
          ranges: [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
          rule: { type: 'list', formula1: '"new"' },
        },
      ],
      { append: true },
    )
    expect(xml).toContain('sqref="$1"')
  })
})
