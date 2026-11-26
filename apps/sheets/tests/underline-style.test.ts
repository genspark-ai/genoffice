/**
 * The underline variant survives the file→Univer load (#2912): a saved
 * `<u val="double">` used to come back as a single underline because every
 * mapping on the read side carried the on/off boolean only. The style load,
 * the rich-document base font, the run mapping and the journal's own run
 * extraction all keep the variant now.
 */
import { BooleanNumber } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import {
  cellFontTextStyle,
  fontTextStyleOf,
  toRichTextDocument,
  toUniverStyle,
} from '../src/renderer/univer-sync'
import { createEditJournal, recordSetRangeValues, toSaveEdits } from '../src/renderer/edit-journal'

const DOUBLE = { s: BooleanNumber.TRUE, t: 10 }

/** the sidecar wire's required booleans; the tests vary only the font bits */
const base = {
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  wrapText: false,
  diagonalUp: false,
  diagonalDown: false,
} as const

describe('cell style load keeps the underline variant', () => {
  it('maps double and the accounting forms to Univer’s decoration variant', () => {
    expect(toUniverStyle({ ...base, underline: true, underlineStyle: 'double' }).ul).toEqual(DOUBLE)
    expect(
      toUniverStyle({ ...base, underline: true, underlineStyle: 'doubleAccounting' }).ul,
    ).toEqual(DOUBLE)
  })

  it('leaves single without a variant and off without an underline', () => {
    expect(toUniverStyle({ ...base, underline: true }).ul).toEqual({ s: BooleanNumber.TRUE })
    expect(
      toUniverStyle({ ...base, underline: false, underlineStyle: 'double' }).ul,
    ).toBeUndefined()
    expect(toUniverStyle({ ...base, bold: true }).ul).toBeUndefined()
  })

  it('the rich-document base font carries the variant too', () => {
    expect(cellFontTextStyle({ ...base, underline: true, underlineStyle: 'double' }).ul).toEqual(
      DOUBLE,
    )
    expect(cellFontTextStyle({ ...base, underline: true }).ul).toEqual({ s: BooleanNumber.TRUE })
  })

  it('fontTextStyleOf keeps whatever variant the style already carries', () => {
    expect(fontTextStyleOf({ ul: DOUBLE }).ul).toEqual(DOUBLE)
    expect(fontTextStyleOf({ ul: { s: BooleanNumber.TRUE } }).ul).toEqual({
      s: BooleanNumber.TRUE,
    })
    expect(fontTextStyleOf(null).ul).toBeUndefined()
  })
})

describe('rich runs keep the underline variant', () => {
  it('maps a double-underlined run onto Univer’s decoration variant', () => {
    const p = toRichTextDocument('ab', [
      {
        text: 'a',
        bold: false,
        italic: false,
        underline: true,
        underlineStyle: 'double',
        strikethrough: false,
      },
      { text: 'b', bold: false, italic: false, underline: true, strikethrough: false },
    ])
    const ts = p?.body?.textRuns?.map((t) => t.ts)
    expect(ts?.[0]?.ul).toEqual(DOUBLE)
    expect(ts?.[1]?.ul).toEqual({ s: BooleanNumber.TRUE })
  })

  it('the journal extracts the variant from a Univer document and saves it back', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', {
      0: {
        0: {
          p: {
            body: {
              dataStream: 'ab\r\n',
              textRuns: [
                { st: 0, ed: 1, ts: { ul: DOUBLE } },
                { st: 1, ed: 2, ts: { ul: { s: BooleanNumber.TRUE } } },
              ],
            },
          },
        },
      },
    })
    const edits = toSaveEdits(journal)
    expect(edits[0]?.rich).toEqual([
      {
        text: 'a',
        bold: false,
        italic: false,
        underline: true,
        underlineStyle: 'double',
        strikethrough: false,
      },
      { text: 'b', bold: false, italic: false, underline: true, strikethrough: false },
    ])
  })
})
