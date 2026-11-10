import { describe, expect, it } from 'vitest'

import { redactionStatesSchema } from '../src/shared/desktop-api'
import {
  parseRedactionPart,
  serializeRedactionPart,
} from '@genoffice/xlsx-gateway/gateway/xlsx-redaction'

/**
 * The boundary the fill crosses on its way into the file: the renderer's mark
 * → the zod schema on the save payload → the part written to disk.
 *
 * It exists because `.strict()` on the save schema rejects an undeclared key
 * outright, so a field added to the mark and forgotten here does not degrade —
 * it refuses every save, and the mark silently never reaches the file.
 */
const MARK = { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0, label: 'client phone' }

describe('the save boundary keeps the displaced fill', () => {
  it('accepts a mark carrying previousFill', () => {
    const parsed = redactionStatesSchema.parse([
      { sheetName: 'Sheet1', marks: [{ ...MARK, previousFill: '#FFEE00' }] },
    ])
    expect(parsed[0]?.marks[0]?.previousFill).toBe('#FFEE00')
  })

  it('accepts a mark whose cell had no fill', () => {
    const parsed = redactionStatesSchema.parse([
      { sheetName: 'Sheet1', marks: [{ ...MARK, previousFill: null }] },
    ])
    expect(parsed[0]?.marks[0]?.previousFill).toBeNull()
  })

  it('accepts a mark from an older part that has no field at all', () => {
    const parsed = redactionStatesSchema.parse([{ sheetName: 'Sheet1', marks: [MARK] }])
    expect(parsed[0]?.marks[0]).not.toHaveProperty('previousFill')
  })

  it('carries the fill all the way into the written part', () => {
    const parsed = redactionStatesSchema.parse([
      { sheetName: 'Sheet1', marks: [{ ...MARK, previousFill: '#FFEE00' }] },
    ])
    const written = parseRedactionPart(serializeRedactionPart(parsed))
    expect(written[0]?.marks[0]?.previousFill).toBe('#FFEE00')
  })

  it('records a null fill into the part rather than dropping the field', () => {
    // Dropping it would make clearing unable to tell "there was nothing" from
    // "this part predates the tint", and only the first should clear the cell.
    const parsed = redactionStatesSchema.parse([
      { sheetName: 'Sheet1', marks: [{ ...MARK, previousFill: null }] },
    ])
    const part = serializeRedactionPart(parsed)
    expect(part).toContain('previousFill')
    expect(parseRedactionPart(part)[0]?.marks[0]?.previousFill).toBeNull()
  })

  it('refuses a fill that is not a colour', () => {
    // It is written back onto a cell as a literal, so anything else is either
    // a bug or an attempt to inject style through the save payload.
    expect(() =>
      redactionStatesSchema.parse([
        { sheetName: 'Sheet1', marks: [{ ...MARK, previousFill: 'red; background:url(x)' }] },
      ]),
    ).toThrow()
  })
})
