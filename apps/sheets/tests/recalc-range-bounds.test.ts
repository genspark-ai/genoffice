import { describe, expect, it } from 'vitest'
import { workbookRecalcRequestSchema } from '../src/shared/desktop-api'

/** Base request every case overrides — valid on its own */
const request = (over: Record<string, unknown>): Record<string, unknown> => ({
  sessionId: '00000000-0000-4000-8000-000000000000',
  edits: [],
  reads: [],
  ...over,
})

const read = (range: Record<string, number>): Record<string, unknown> => ({
  sheetId: 's1',
  range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0, ...range },
})

describe('workbookRecalcRequestSchema range bounds', () => {
  it('accepts the last real cell of an xlsx sheet', () => {
    const parsed = workbookRecalcRequestSchema.parse(
      request({
        reads: [
          read({ startRow: 1_048_575, endRow: 1_048_575, startColumn: 16_383, endColumn: 16_383 }),
        ],
      }),
    )
    expect(parsed.reads[0]!.range.startRow).toBe(1_048_575)
  })

  it('rejects a row past the 1,048,576-row sheet limit (usize→i32 truncation in the engine)', () => {
    expect(() =>
      workbookRecalcRequestSchema.parse(
        request({ reads: [read({ startRow: 1_048_576, endRow: 1_048_576 })] }),
      ),
    ).toThrow(/startRow/)
    expect(() =>
      workbookRecalcRequestSchema.parse(request({ reads: [read({ endRow: 2_000_000 })] })),
    ).toThrow(/endRow/)
  })

  it('rejects a column past the 16,384-column sheet limit', () => {
    expect(() =>
      workbookRecalcRequestSchema.parse(request({ reads: [read({ startColumn: 16_384 })] })),
    ).toThrow(/startColumn/)
    expect(() =>
      workbookRecalcRequestSchema.parse(request({ reads: [read({ endColumn: 99_999 })] })),
    ).toThrow(/endColumn/)
  })
})
