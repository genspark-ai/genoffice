/**
 * The write guard: a model batch may not land on a cell the reader withheld.
 *
 * A mark is a label over a real value, so two failures are possible and only
 * one is guarded here — the model **writing over** the value. (Leaking it is
 * the read path's job, covered in redact-read.test.ts.) The overwrite is the
 * quiet one: the file still opens, it just no longer says what the reader
 * wrote, and nothing on screen mentions it.
 *
 * So the tests below pin three things: the ops that must be refused, the ops
 * that must stay allowed (an over-broad guard makes the feature unusable), and
 * the message — which has to name the cell and the label, or the model cannot
 * tell the user what it hit.
 */
import { describe, expect, it, vi } from 'vitest'

import { buildRedactionIndex, NO_REDACTIONS, type RedactionIndex } from '../src/renderer/ai/redact'
import { redactGuardForOps } from '../src/renderer/ai/redact-guard'
import { executeWorkbookTool, type SheetsSkillDeps } from '../src/renderer/ai/tools'

const SHEETS = [
  { id: 'sh1', name: 'Customers' },
  { id: 'sh2', name: 'Orders' },
]

/// B2 on Customers — the phone number the reader does not want written over.
const INDEX: RedactionIndex = buildRedactionIndex(
  [
    {
      sheetName: 'Customers',
      marks: [{ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1, label: 'client phone' }],
    },
  ],
  SHEETS,
)

const sheetName = (sheetId: string): string | undefined =>
  SHEETS.find((sheet) => sheet.id === sheetId)?.name

function guard(ops: unknown[], index: RedactionIndex = INDEX) {
  return redactGuardForOps(ops, index, sheetName)
}

describe('a write that reaches a withheld cell is refused', () => {
  it('names the cell and the label, so the model can tell the user what it hit', () => {
    const refusal = guard([{ op: 'set_cell', sheetId: 'sh1', address: 'B2', value: 'x' }])
    expect(refusal).not.toBeNull()
    expect(refusal?.label).toBe('client phone')
    expect(refusal?.where).toBe('Customers!B2')
    // The label reaches the model in the same {{…}} form every other app uses.
    expect(refusal?.reason).toContain('{{client phone}}')
    expect(refusal?.reason).toContain('Customers!B2')
  })

  it('refuses every op that can change a cell', () => {
    const ops: Record<string, unknown>[] = [
      { op: 'set_cell', sheetId: 'sh1', address: 'B2', value: 'x' },
      { op: 'set_formula', sheetId: 'sh1', address: 'B2', formula: '=1+1' },
      { op: 'clear_cell', sheetId: 'sh1', address: 'B2' },
      {
        op: 'set_range',
        sheetId: 'sh1',
        start: 'A1',
        values: [
          ['a', 'b'],
          ['c', 'd'],
        ],
      },
      { op: 'clear_range', sheetId: 'sh1', range: 'B1:B5' },
      { op: 'fill_range', sheetId: 'sh1', source: 'A1', target: 'B1:B5' },
      { op: 'copy_range', sheetId: 'sh1', source: 'A1', target: 'B2' },
      { op: 'convert_to_values', sheetId: 'sh1', range: 'B2' },
      { op: 'sort_range', sheetId: 'sh1', range: 'A1:B5', byColumn: 'A', order: 'asc' },
      { op: 'find_replace', sheetId: 'sh1', range: 'B1:B5', find: 'a', replace: 'b' },
      { op: 'merge_cells', sheetId: 'sh1', range: 'B1:B5' },
      { op: 'add_table', sheetId: 'sh1', range: 'A1:B5' },
    ]
    for (const op of ops) {
      expect(guard([op]), `${op.op} should be refused`).not.toBeNull()
    }
  })

  it('refuses a batch when any op in it reaches the mark, and says which', () => {
    // propose_operations is all-or-nothing, so a refusal must name the op that
    // caused it rather than the first one in the batch.
    const refusal = guard([
      { op: 'set_cell', sheetId: 'sh1', address: 'A1', value: 'fine' },
      { op: 'sort_range', sheetId: 'sh1', range: 'A1:B5', byColumn: 'A', order: 'asc' },
    ])
    expect(refusal?.reason).toContain('sort_range')
  })

  it('refuses a delete that would take the withheld cell with it', () => {
    // The mark does not travel with the cell, so whatever shifts into its
    // place would be read as the withheld one — and the real value is gone.
    expect(guard([{ op: 'delete_rows', sheetId: 'sh1', row: 2, count: 1 }])).not.toBeNull()
    expect(guard([{ op: 'delete_cols', sheetId: 'sh1', column: 'B', count: 1 }])).not.toBeNull()
  })

  it('refuses dropping or duplicating a sheet that carries marks', () => {
    // Duplicating copies the withheld values into a sheet with no marks at
    // all, which publishes them.
    expect(guard([{ op: 'delete_sheet', sheetId: 'sh1' }])).not.toBeNull()
    expect(guard([{ op: 'duplicate_sheet', sheetId: 'sh1' }])).not.toBeNull()
  })

  it('refuses an insert that would slide a marked cell out from under its mark', () => {
    expect(guard([{ op: 'insert_rows', sheetId: 'sh1', row: 2, count: 1 }])).not.toBeNull()
    expect(guard([{ op: 'insert_rows', sheetId: 'sh1', row: 1, count: 3 }])).not.toBeNull()
    expect(guard([{ op: 'insert_cols', sheetId: 'sh1', column: 'B', count: 1 }])).not.toBeNull()
  })
})

describe('what stays allowed', () => {
  it('leaves a workbook that withholds nothing completely alone', () => {
    const ops = [
      { op: 'set_cell', sheetId: 'sh1', address: 'B2', value: 'x' },
      { op: 'sort_range', sheetId: 'sh1', range: 'A1:B5', byColumn: 'A', order: 'asc' },
      { op: 'delete_sheet', sheetId: 'sh1' },
    ]
    for (const op of ops) {
      expect(guard([op], NO_REDACTIONS), `${op.op} should be allowed`).toBeNull()
    }
  })

  it('allows a write beside the mark, and on another sheet', () => {
    expect(guard([{ op: 'set_cell', sheetId: 'sh1', address: 'B3', value: 'x' }])).toBeNull()
    expect(guard([{ op: 'set_cell', sheetId: 'sh1', address: 'C2', value: 'x' }])).toBeNull()
    expect(guard([{ op: 'set_cell', sheetId: 'sh2', address: 'B2', value: 'x' }])).toBeNull()
  })

  it('allows formatting, notes and other metadata that leave the value alone', () => {
    // Refusing these would make the feature unusable for ordinary work: none
    // of them changes what the cell says.
    const ops: Record<string, unknown>[] = [
      { op: 'format_range', sheetId: 'sh1', range: 'B2', format: { bold: true } },
      { op: 'set_note', sheetId: 'sh1', address: 'B2', text: 'checked' },
      { op: 'set_hyperlink', sheetId: 'sh1', address: 'B2', url: 'https://example.com' },
      { op: 'set_filter', sheetId: 'sh1', range: 'A1:B5' },
      {
        op: 'set_data_validation',
        sheetId: 'sh1',
        range: 'B2',
        rule: { kind: 'list', values: ['a'] },
      },
    ]
    for (const op of ops) {
      expect(guard([op]), `${op.op} should be allowed`).toBeNull()
    }
  })

  it('allows a one-cell merge, which destroys nothing', () => {
    expect(guard([{ op: 'merge_cells', sheetId: 'sh1', range: 'B2:B2' }])).toBeNull()
  })

  it('allows an insert that only touches rows below the mark', () => {
    // The rule is positional, not "inserts are forbidden": an insert at or
    // above a mark slides the withheld value out from under it, while one
    // below leaves it exactly where it is.
    const lower = buildRedactionIndex(
      [
        {
          sheetName: 'Customers',
          marks: [{ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1, label: 'x' }],
        },
      ],
      SHEETS,
    )
    expect(guard([{ op: 'insert_rows', sheetId: 'sh1', row: 9, count: 1 }], lower)).toBeNull()
    expect(guard([{ op: 'insert_cols', sheetId: 'sh1', column: 'D', count: 1 }], lower)).toBeNull()
    expect(guard([{ op: 'insert_rows', sheetId: 'sh1', row: 2, count: 1 }], lower)).not.toBeNull()
  })
})

describe('propose_operations refuses the batch before anything is applied', () => {
  const info = {
    mode: 'demo' as const,
    sheetId: 'sh1',
    sheetName: 'Customers',
    revision: 0,
    knownAddresses: [],
    sheets: SHEETS,
  }
  const plan = {
    transactionId: 'agent-1',
    baseRevision: 0,
    cellChanges: [],
    sheetRenames: [],
    structuralChanges: [],
    formatChanges: [],
    warnings: [],
  }
  function deps(overrides: Partial<SheetsSkillDeps> = {}): SheetsSkillDeps {
    return {
      getActiveSheetInfo: () => info,
      readCells: () => ({}),
      readFormats: () => ({}),
      readSheetFeatures: () => 'Feature state of sheet Customers',
      findCells: () => ({ matches: [], truncated: false, incompleteSheets: [] }),
      selectRange: () => ({ ok: true, sheetName: 'Customers' }),
      tracePrecedents: () => ({ refs: [] }),
      traceDependents: () => ({ dependents: [], truncated: false, incompleteSheets: [] }),
      proposeOperations: () => ({ ok: true, plan }),
      redactions: () => INDEX,
      ...overrides,
    }
  }

  it('never reaches the plan when a withheld cell is in range', () => {
    const propose = vi.fn(() => ({ ok: true, plan }))
    const result = executeWorkbookTool(
      {
        id: 'call-1',
        name: 'propose_operations',
        input: {
          summary: 'update the phone column',
          operations: [{ op: 'set_cell', sheetId: 'sh1', address: 'B2', value: 'call them' }],
        },
      },
      deps({ proposeOperations: propose as never }),
    )
    expect(propose).not.toHaveBeenCalled()
    expect((result as { isError?: boolean }).isError).toBe(true)
    expect((result as { output: string }).output).toContain('{{client phone}}')
    expect((result as { output: string }).output).toContain('all-or-nothing')
  })

  it('applies normally when the batch stays off the withheld cell', () => {
    const result = executeWorkbookTool(
      {
        id: 'call-2',
        name: 'propose_operations',
        input: {
          summary: 'update the name column',
          operations: [{ op: 'set_cell', sheetId: 'sh1', address: 'A1', value: 'Acme' }],
        },
      },
      deps(),
    )
    expect((result as { isError?: boolean }).isError).toBeFalsy()
  })

  it('treats a session whose part could not be read as fully withheld', () => {
    // The load side installs the withhold-everything index on an unreadable
    // part; without it the guard would wave every write through on a workbook
    // whose marks are unknown.
    const propose = vi.fn(() => ({ ok: true, plan }))
    const result = executeWorkbookTool(
      {
        id: 'call-3',
        name: 'propose_operations',
        input: {
          summary: 'anything',
          operations: [{ op: 'set_cell', sheetId: 'sh1', address: 'A1', value: 'x' }],
        },
      },
      deps({
        redactions: () => ({
          isEmpty: false,
          sheetIds: [],
          labelAt: () => 'private',
          labelsFor: () => [],
          marksFor: () => [
            { startRow: 0, endRow: 1_048_575, startColumn: 0, endColumn: 16_383, label: 'private' },
          ],
        }),
        proposeOperations: propose as never,
      }),
    )
    expect(propose).not.toHaveBeenCalled()
    expect((result as { output: string }).output).toContain('{{private}}')
  })
})
