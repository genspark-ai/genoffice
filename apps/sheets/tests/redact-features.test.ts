import { describe, expect, it } from 'vitest'

import { buildRedactionIndex } from '../src/renderer/ai/redact'
import { readSheetFeatures, type WorkbookReadContext } from '../src/renderer/ai/workbook-readers'

const NOTE_TEXT = '客户电话 13800138000'
const DV_LITERAL = '13800138000'

type RangeStub = {
  getA1Notation: () => string
  getRow: () => number
  getLastRow: () => number
  getColumn: () => number
  getLastColumn: () => number
}

function range(
  a1: string,
  startRow: number,
  endRow: number,
  startColumn: number,
  endColumn: number,
): RangeStub {
  return {
    getA1Notation: () => a1,
    getRow: () => startRow,
    getLastRow: () => endRow,
    getColumn: () => startColumn,
    getLastColumn: () => endColumn,
  }
}

/// `read_sheet_features` reports workbook structures, not cells — but two of
/// those structures carry cell-adjacent content: a note is anchored to a cell
/// and holds free text, and a data-validation rule can store a literal value.
/// Both are the same leak as an unprojected cell, reached through a side door.
function ctxWith(options: { withheld: boolean }): WorkbookReadContext {
  const worksheet = {
    getSheetId: () => 'sh1',
    getSheetName: () => 'Customers',
    getFreeze: () => ({ xSplit: 0, ySplit: 0 }),
    getDataValidations: () => [
      {
        getRanges: () => [range('B2:B10', 1, 9, 1, 1)],
        rule: {
          type: 'cellIs',
          operator: 'equal',
          formula1: DV_LITERAL,
          formula2: undefined,
        },
      },
      {
        getRanges: () => [range('D2:D10', 1, 9, 3, 3)],
        rule: { type: 'list', operator: undefined, formula1: 'a,b,c', formula2: undefined },
      },
    ],
    getNotes: () => [
      { row: 1, col: 1, note: NOTE_TEXT },
      { row: 2, col: 3, note: 'public note' },
    ],
  }
  return {
    univerRef: {
      current: {
        univerAPI: {
          getActiveWorkbook: () => ({
            getActiveSheet: () => worksheet,
            getSheetBySheetId: () => worksheet,
            getDefinedNames: () => [],
          }),
        },
      } as never,
    },
    lazyWorkbookRef: { current: null },
    adapterRef: { current: { getSnapshot: () => ({ revision: 0, sheets: [] }) } } as never,
    redactionIndexRef: {
      current: options.withheld
        ? buildRedactionIndex(
            [
              {
                sheetName: 'Customers',
                marks: [
                  {
                    startRow: 1,
                    endRow: 9,
                    startColumn: 1,
                    endColumn: 1,
                    label: '客户电话',
                  },
                ],
              },
            ],
            [{ id: 'sh1', name: 'Customers' }],
          )
        : buildRedactionIndex([], [{ id: 'sh1', name: 'Customers' }]),
    },
  } as unknown as WorkbookReadContext
}

describe('read_sheet_features does not leak a withheld value through a side door', () => {
  it('replaces a note anchored on a withheld cell with the placeholder', () => {
    const output = readSheetFeatures(ctxWith({ withheld: true }))
    expect(output).not.toContain(NOTE_TEXT)
    expect(output).toContain('- B2: {{客户电话}}')
  })

  it('leaves a note on a cell that is not withheld alone', () => {
    // Hiding every note would gut the tool; only the withheld one changes.
    const output = readSheetFeatures(ctxWith({ withheld: true }))
    expect(output).toContain('- D3: public note')
  })

  it('drops the formula of a validation rule covering a withheld cell', () => {
    const output = readSheetFeatures(ctxWith({ withheld: true }))
    expect(output).not.toContain(DV_LITERAL)
    // The rule itself is still reported: the model has to know not to break it.
    expect(output).toContain('B2:B10: cellIs equal (withheld)')
  })

  it('leaves a validation rule on an untouched column alone', () => {
    const output = readSheetFeatures(ctxWith({ withheld: true }))
    expect(output).toContain('D2:D10: list a,b,c')
  })

  it('reports everything verbatim when nothing is withheld', () => {
    const output = readSheetFeatures(ctxWith({ withheld: false }))
    expect(output).toContain(NOTE_TEXT)
    expect(output).toContain(`B2:B10: cellIs equal ${DV_LITERAL}`)
  })
})
