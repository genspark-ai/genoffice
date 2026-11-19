/**
 * A per-range password travels the real save path: the renderer puts the hash it
 * computed inside protectedRangeStates[].ranges[], the main process parses the
 * request with the strict workbookSaveRequestSchema, and the parsed ranges reach
 * the worksheet XML with the password attributes. The strict schema is the step
 * a writer-only test skips — and the step that rejected `password` outright.
 */
import { describe, expect, it } from 'vitest'
import { applyProtectedRanges } from '@genoffice/xlsx-gateway/gateway/xlsx-protection'
import {
  hashSheetPassword,
  legacySheetPasswordHash,
} from '@genoffice/xlsx-gateway/gateway/xlsx-protection-hash'
import { workbookSaveRequestSchema } from '../src/shared/desktop-api'

type RangeInput = Record<string, unknown>

/** The renderer's save request with nothing but an allow-edit-range change on one sheet */
function requestWithRanges(ranges: readonly RangeInput[]) {
  return {
    sessionId: '5d4f6f7a-1c2b-4e3d-9a8f-0b1c2d3e4f5a',
    mode: 'save',
    edits: [],
    structuralOps: [],
    chartEdits: [],
    visualEdits: [],
    visualAdditions: [],
    tableAdditions: [],
    pivotAdditions: [],
    sheetOps: [],
    sheetOrder: [],
    filterStates: [],
    hyperlinkEdits: [],
    cfStates: [],
    dvStates: [],
    pageSetupStates: [],
    noteStates: [],
    formulaValues: [],
    pivotCacheRefreshPaths: [],
    pivotRefreshUpdates: [],
    sheetProtections: [],
    sparklineAdditions: [],
    definedNamesState: null,
    protectedRangeStates: [{ sheetId: 'sheet-1', ranges }],
  }
}

/** What sheets-main hands the gateway: the parsed ranges, keyed by sheet name */
function worksheetXmlAfter(request: ReturnType<typeof requestWithRanges>): string {
  const parsed = workbookSaveRequestSchema.parse(request)
  const states = parsed.protectedRangeStates.map((state) => ({
    sheetName: `Sheet ${state.sheetId}`,
    ranges: state.ranges,
  }))
  expect(states).toHaveLength(1)
  return applyProtectedRanges('<sheetData/>', states[0]!.ranges)
}

describe('allow-edit range password through the save request schema', () => {
  it('a spi hash from hashSheetPassword passes the schema and lands in the worksheet', async () => {
    const password = await hashSheetPassword('open sesame', 1_000)
    expect('algorithmName' in password).toBe(true)
    const xml = worksheetXmlAfter(requestWithRanges([{ name: 'Totals', sqref: 'B2:D9', password }]))
    if (!('algorithmName' in password)) throw new Error('expected a spi hash')
    expect(xml).toContain(`<protectedRange sqref="B2:D9" name="Totals" algorithmName="SHA-512"`)
    expect(xml).toContain(`hashValue="${password.hashValue}"`)
    expect(xml).toContain(`saltValue="${password.saltValue}"`)
    expect(xml).toContain('spinCount="1000"')
  })

  it('a legacy hash passes the schema and lands as the password attribute', () => {
    const legacy = legacySheetPasswordHash('open sesame')
    const xml = worksheetXmlAfter(
      requestWithRanges([{ name: 'Totals', sqref: 'B2:D9', password: { legacy } }]),
    )
    expect(xml).toContain(`<protectedRange sqref="B2:D9" name="Totals" password="${legacy}"/>`)
  })

  it('a range without a password still saves, and mixed lists keep each range its own', () => {
    const xml = worksheetXmlAfter(
      requestWithRanges([
        { name: 'Open', sqref: 'A1' },
        { name: 'Locked', sqref: 'C1', password: { legacy: '8492' } },
      ]),
    )
    expect(xml).toContain('<protectedRange sqref="A1" name="Open"/>')
    expect(xml).toContain('<protectedRange sqref="C1" name="Locked" password="8492"/>')
  })

  it('the schema stays strict: a plaintext password or an unknown key is rejected', () => {
    // the renderer must hash before sending — a raw string is not a SheetPasswordHash
    expect(() =>
      workbookSaveRequestSchema.parse(
        requestWithRanges([{ name: 'R', sqref: 'A1', password: 'open sesame' }]),
      ),
    ).toThrow()
    expect(() =>
      workbookSaveRequestSchema.parse(requestWithRanges([{ name: 'R', sqref: 'A1', pin: '1234' }])),
    ).toThrow()
    // a hash with a stray key is not accepted either
    expect(() =>
      workbookSaveRequestSchema.parse(
        requestWithRanges([{ name: 'R', sqref: 'A1', password: { legacy: '8492', extra: 1 } }]),
      ),
    ).toThrow()
  })
})
