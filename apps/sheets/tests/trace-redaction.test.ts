import { describe, expect, it } from 'vitest'
import { executeWorkbookTool, type SheetsSkillDeps } from '../src/renderer/ai/tools'
import type { ActiveSheetInfo } from '../src/renderer/ai/tools'
import { buildRedactionIndex } from '../src/renderer/ai/redact'

/**
 * The trace tools print cell values, and a cell the reader withheld keeps its
 * real value on the sheet — that is the point of the feature, not a side
 * effect. So every value they print has to go through the redaction index, and
 * `trace_dependents` has the harder half: a `=B2` beside a withheld B2 shows
 * B2's value without the trace ever naming B2.
 */

const SECRET = '13800138000'

const info: ActiveSheetInfo = {
  mode: 'demo',
  sheetId: 'sheet-1',
  sheetName: 'Sheet1',
  revision: 0,
  knownAddresses: ['B2', 'C2'],
  sheets: [
    { id: 'sheet-1', name: 'Sheet1' },
    { id: 'sheet-2', name: 'Summary' },
  ],
  selection: 'B2:C2',
}

const deps = (overrides: Partial<SheetsSkillDeps> = {}): SheetsSkillDeps =>
  ({
    getActiveSheetInfo: () => info,
    readCells: () => ({}),
    readFormats: () => ({}),
    readSheetFeatures: () => '',
    findCells: () => ({ matches: [], truncated: false, incompleteSheets: [] }),
    ...overrides,
  }) as unknown as SheetsSkillDeps

/** B2 withheld — the phone number */
const indexWithB2 = buildRedactionIndex(
  [
    {
      sheetName: 'Sheet1',
      marks: [{ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1, label: 'client phone' }],
    },
  ],
  info.sheets,
)

const call = (name: string, input: Record<string, unknown>) => ({ id: 't1', name, input }) as never

const run = (d: SheetsSkillDeps, name: string, input: Record<string, unknown>) => {
  const out = executeWorkbookTool(call(name, input), d)
  if (out instanceof Promise) throw new Error('expected a sync tool')
  return out
}

describe('trace_precedents with a withheld cell', () => {
  const tracing = deps({
    redactions: () => indexWithB2,
    tracePrecedents: () => ({
      formula: '=B2*2',
      value: 24690,
      refs: [
        {
          label: 'B2',
          cellCount: 1,
          hasError: false,
          samples: [{ address: 'B2', value: SECRET }],
        },
      ],
      usesNames: false,
    }),
  })

  it('prints the marker where a withheld precedent would be', () => {
    const out = run(tracing, 'trace_precedents', { address: 'C2' })
    expect(out.output).not.toContain(SECRET)
    expect(out.output).toContain('{{client phone}}')
  })

  it('withholds the traced cell’s own value, which is its precedents’ value', () => {
    // `C2 = 24690` when `B2` is the withheld number says as much as printing
    // B2 would
    const out = run(tracing, 'trace_precedents', { address: 'C2' })
    expect(out.output).not.toContain('24690')
    expect(out.output).toContain('withheld from the model')
  })

  it('leaves an untraced cell alone when nothing on the way is withheld', () => {
    const plain = deps({
      tracePrecedents: () => ({
        formula: '=B2*2',
        value: 24690,
        refs: [
          {
            label: 'B2',
            cellCount: 1,
            hasError: false,
            samples: [{ address: 'B2', value: 12345 }],
          },
        ],
        usesNames: false,
      }),
    })
    const out = run(plain, 'trace_precedents', { address: 'C2' })
    expect(out.output).toContain('24690')
    expect(out.output).toContain('12345')
  })

  it('says the value of a plain cell being traced is withheld, not its value', () => {
    const plainCell = deps({
      redactions: () => indexWithB2,
      tracePrecedents: () => ({ value: SECRET, refs: [], usesNames: false }),
    })
    const out = run(plainCell, 'trace_precedents', { address: 'B2' })
    expect(out.output).not.toContain(SECRET)
    expect(out.output).toContain('{{client phone}}')
  })
})

describe('trace_dependents with a withheld cell', () => {
  it('withholds a dependent whose value came from the traced cell', () => {
    const out = run(
      deps({
        redactions: () => indexWithB2,
        traceDependents: () => ({
          dependents: [
            {
              sheetName: 'Sheet1',
              address: 'C2',
              formula: '=B2',
              value: SECRET,
            },
          ],
          truncated: false,
          incompleteSheets: [],
        }),
      }),
      'trace_dependents',
      { address: 'B2' },
    )
    expect(out.output).not.toContain(SECRET)
    expect(out.output).toContain('withheld from the model')
  })

  it('prints a dependent that reads nothing withheld as usual', () => {
    const out = run(
      deps({
        traceDependents: () => ({
          dependents: [{ sheetName: 'Sheet1', address: 'C2', formula: '=A2', value: 42 }],
          truncated: false,
          incompleteSheets: [],
        }),
      }),
      'trace_dependents',
      { address: 'B2' },
    )
    expect(out.output).toContain('42')
  })
})
