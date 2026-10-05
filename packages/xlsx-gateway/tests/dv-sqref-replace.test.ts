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
    // A pre-existing rule whose sqref decodes to three areas: A1, a literal
    // "$&", and B1. Appending a rule over A1 removes A1 and rewrites the
    // remaining sqref. A string replacement expands "$&" to the entire match
    // (`sqref="A1 $&amp; B1"`), producing a duplicated, unquoted fragment
    // after the attribute list and a worksheet Excel must repair.
    const xml = appendOver({ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 })

    // The surviving areas are written back verbatim: "$&" and "B1".
    expect(xml).toContain('sqref="$&amp; B1"')
    // The whole-match expansion must NOT appear — that is the corruption.
    expect(xml).not.toContain('sqref="A1 $&amp; B1" B1"')
    // The element stays well-formed: the start tag ends at the first '>',
    // so a duplicated trailing area would leak past it.
    const tag = /<dataValidation\b[^>]*>/.exec(xml)?.[0] ?? ''
    expect(tag).toContain('sqref="$&amp; B1"')
    expect(tag.endsWith('>')).toBe(true)
  })

  it('keeps a surviving area containing $1 literal', () => {
    // $1 is a capture-group reference; with no capturing group it expands to
    // the empty string in a string replacement, silently dropping the area.
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
