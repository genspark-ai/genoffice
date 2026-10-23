import { describe, expect, it } from 'vitest'

import {
  addMark,
  clearMark,
  columnLabel,
  hasPendingRedactionChange,
  indexFor,
  rangeLabelOf,
  redactIntentFor,
} from '../src/renderer/redact-actions'
import { NO_REDACTIONS } from '../src/renderer/ai/redact'
import {
  parseRedactionPart,
  serializeRedactionPart,
  type SheetRedactionState,
} from '@genoffice/xlsx-gateway/gateway/xlsx-redaction'
import type { SelectionRequest } from '../src/renderer/redact-menu'
import { createWorkbookSkill } from '../src/renderer/ai/workbook-skill'

const SHEETS = [
  { id: 'sh1', name: 'Customers' },
  { id: 'sh2', name: 'Orders' },
]

function selection(overrides: Partial<SelectionRequest> = {}): SelectionRequest {
  const startRow = overrides.startRow ?? 1
  const startColumn = overrides.startColumn ?? 1
  return {
    sheetId: 'sh1',
    startRow,
    endRow: overrides.endRow ?? startRow,
    startColumn,
    endColumn: overrides.endColumn ?? startColumn,
    isSingleCell:
      (overrides.endRow ?? startRow) === startRow &&
      (overrides.endColumn ?? startColumn) === startColumn,
    ...overrides,
  }
}

const nameOf = (id: string) => SHEETS.find((sheet) => sheet.id === id)?.name

describe('column labels', () => {
  it('matches the addresses a spreadsheet uses', () => {
    expect(columnLabel(0)).toBe('A')
    expect(columnLabel(25)).toBe('Z')
    expect(columnLabel(26)).toBe('AA')
    expect(columnLabel(27)).toBe('AB')
    expect(columnLabel(701)).toBe('ZZ')
    expect(columnLabel(702)).toBe('AAA')
  })
})

describe('the range a right-click names', () => {
  it('reads as one cell for a single cell and as a span otherwise', () => {
    expect(rangeLabelOf(selection())).toBe('B2')
    expect(
      rangeLabelOf(
        selection({ startRow: 1, endRow: 499, startColumn: 1, endColumn: 1, isSingleCell: false }),
      ),
    ).toBe('B2:B500')
  })
})

describe('what a right-click means', () => {
  const MARKED: SheetRedactionState[] = [
    {
      sheetName: 'Customers',
      marks: [{ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1, label: 'client phone' }],
    },
  ]

  it('asks for a label on a cell that is not withheld', () => {
    // C4, deliberately clear of the mark on B2.
    const intent = redactIntentFor(
      selection({ startRow: 3, startColumn: 2 }),
      indexFor(MARKED, SHEETS),
      nameOf,
    )
    expect(intent.kind).toBe('ask')
    if (intent.kind !== 'ask') return
    expect(intent.dialog.sheetName).toBe('Customers')
    expect(intent.dialog.rangeLabel).toBe('C4')
  })

  it('clears the mark when the selection is already inside one', () => {
    const intent = redactIntentFor(selection(), indexFor(MARKED, SHEETS), nameOf)
    expect(intent.kind).toBe('clear')
    if (intent.kind !== 'clear') return
    expect(intent.mark).toEqual({ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 })
  })

  it('clears the mark when the selection is one cell of a marked range', () => {
    // A whole column marked, then one cell of it clicked: the reader means
    // "stop hiding this", not "hide this one cell too".
    const wide: SheetRedactionState[] = [
      {
        sheetName: 'Customers',
        marks: [{ startRow: 1, endRow: 9, startColumn: 1, endColumn: 1, label: 'client phone' }],
      },
    ]
    const intent = redactIntentFor(selection(), indexFor(wide, SHEETS), nameOf)
    expect(intent.kind).toBe('clear')
    if (intent.kind !== 'clear') return
    // The cleared area is the whole mark, not just the clicked cell.
    expect(intent.mark).toEqual({ startRow: 1, endRow: 9, startColumn: 1, endColumn: 1 })
  })

  it('asks rather than clearing when the selection only partly overlaps', () => {
    // Overlapping is not covering: the new selection is bigger than the mark,
    // so it is a new thing to withhold, not a request to undo the old one.
    const intent = redactIntentFor(
      selection({ startRow: 0, endRow: 1, startColumn: 1, endColumn: 1, isSingleCell: false }),
      indexFor(MARKED, SHEETS),
      nameOf,
    )
    expect(intent.kind).toBe('ask')
  })

  it('does nothing for a sheet this workbook does not have', () => {
    const intent = redactIntentFor(selection({ sheetId: 'gone' }), indexFor(MARKED, SHEETS), nameOf)
    expect(intent.kind).toBe('ignore')
  })

  it('does nothing at all on a workbook that withholds nothing', () => {
    // The shared empty index answers isEmpty before any of this runs, so the
    // cost is a single property read rather than a scan per right-click.
    expect(NO_REDACTIONS.isEmpty).toBe(true)
    const intent = redactIntentFor(selection(), NO_REDACTIONS, nameOf)
    expect(intent.kind).toBe('ask')
  })
})

describe('adding and clearing marks', () => {
  const MARK = { startRow: 1, endRow: 1, startColumn: 1, endColumn: 1, label: 'client phone' }

  it('adds a mark without touching the others', () => {
    const before: SheetRedactionState[] = [
      { sheetName: 'Orders', marks: [{ ...MARK, label: 'order id' }] },
    ]
    expect(addMark(before, 'Customers', MARK)).toEqual([
      ...before,
      { sheetName: 'Customers', marks: [MARK] },
    ])
  })

  it('replaces a mark over the same cells rather than stacking a second', () => {
    const before: SheetRedactionState[] = [{ sheetName: 'Customers', marks: [MARK] }]
    const renamed = { ...MARK, label: 'mobile' }
    const after = addMark(before, 'Customers', renamed)
    expect(after).toEqual([{ sheetName: 'Customers', marks: [renamed] }])
  })

  it('clears only the named mark', () => {
    const other = { ...MARK, startRow: 5, label: 'address' }
    const before: SheetRedactionState[] = [{ sheetName: 'Customers', marks: [MARK, other] }]
    expect(clearMark(before, 'Customers', MARK)).toEqual([
      { sheetName: 'Customers', marks: [other] },
    ])
  })

  it('drops a sheet once its last mark is cleared', () => {
    const before: SheetRedactionState[] = [{ sheetName: 'Customers', marks: [MARK] }]
    expect(clearMark(before, 'Customers', MARK)).toEqual([])
  })

  it('leaves other sheets alone when clearing', () => {
    const before: SheetRedactionState[] = [
      { sheetName: 'Customers', marks: [MARK] },
      { sheetName: 'Orders', marks: [{ ...MARK, label: 'x' }] },
    ]
    expect(clearMark(before, 'Customers', MARK)).toEqual([before[1]])
  })
})

describe('the index follows the state the save writes', () => {
  it('withholds a cell the moment the mark is added', () => {
    // An index that lagged behind the state would withhold nothing right after
    // the reader asked for it, and the next save would then record marks the
    // model was never shown.
    const index = indexFor(
      addMark([], 'Customers', {
        startRow: 1,
        endRow: 1,
        startColumn: 1,
        endColumn: 1,
        label: 'client phone',
      }),
      SHEETS,
    )
    expect(index.labelAt('sh1', 1, 1)).toBe('client phone')
  })

  it('stops withholding once the mark is cleared', () => {
    const marked = addMark([], 'Customers', {
      startRow: 1,
      endRow: 1,
      startColumn: 1,
      endColumn: 1,
      label: 'client phone',
    })
    const index = indexFor(clearMark(marked, 'Customers', marked[0]!.marks[0]!), SHEETS)
    expect(index.labelAt('sh1', 1, 1)).toBeNull()
  })
})

describe('a newly withheld cell counts as a pending change', () => {
  /**
   * The save ticks ask "is the edit journal empty?" and skip the write when it
   * is. A mark is not in the journal, so without this a workbook whose only
   * edit is a withheld cell never autosaves, never gets a crash-recovery copy,
   * and loses the mark when the tab closes — silently, because the reader did
   * see the mark drawn on the grid.
   */
  const LOADED: SheetRedactionState[] = [
    {
      sheetName: 'Customers',
      marks: [{ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1, label: 'client phone' }],
    },
  ]
  const MARK = { startRow: 4, endRow: 4, startColumn: 1, endColumn: 1, label: 'mobile' }

  it('sees a mark the reader just added', () => {
    expect(hasPendingRedactionChange(addMark(LOADED, 'Customers', MARK), LOADED)).toBe(true)
  })

  it('sees a mark the reader just cleared', () => {
    expect(
      hasPendingRedactionChange(clearMark(LOADED, 'Customers', LOADED[0]!.marks[0]!), LOADED),
    ).toBe(true)
  })

  it('sees a mark that moved by one row', () => {
    const moved: SheetRedactionState[] = [
      {
        sheetName: 'Customers',
        marks: [{ ...LOADED[0]!.marks[0]!, startRow: 2 }],
      },
    ]
    expect(hasPendingRedactionChange(moved, LOADED)).toBe(true)
  })

  it('sees a mark whose label was edited', () => {
    const renamed: SheetRedactionState[] = [
      { sheetName: 'Customers', marks: [{ ...LOADED[0]!.marks[0]!, label: 'mobile' }] },
    ]
    expect(hasPendingRedactionChange(renamed, LOADED)).toBe(true)
  })

  it('sees nothing when the marks are exactly what the file had', () => {
    expect(hasPendingRedactionChange(LOADED, LOADED)).toBe(false)
    expect(hasPendingRedactionChange([], [])).toBe(false)
  })

  it('treats a re-mark of the same cell as a change', () => {
    // Same area, different label: the file on disk still carries the old one.
    const reMarked = addMark(LOADED, 'Customers', {
      ...MARK,
      startRow: 1,
      endRow: 1,
      label: 'mobile',
    })
    expect(hasPendingRedactionChange(reMarked, LOADED)).toBe(true)
  })
})

describe('the prompt tells the model which placeholders exist', () => {
  /**
   * The base prompt is imported raw and fixed at build time, so the section is
   * a getter. That is the whole risk: a model told about a placeholder that is
   * not there, or not told about one that is, misreads the sheet — and the
   * second failure is silent, because a cell reading {{client phone}} with no
   * explanation looks like the author's own wording.
   */
  function skillWith(marks: Record<string, string[]>) {
    return createWorkbookSkill({
      redactions: () => ({
        isEmpty: Object.keys(marks).length === 0,
        sheetIds: Object.keys(marks),
        labelAt: () => null,
        labelsFor: (sheetId: string) => marks[sheetId] ?? [],
        marksFor: () => [],
      }),
      getActiveSheetInfo: () =>
        ({
          mode: 'lazy',
          sheetId: 'sh1',
          sheetName: 'Customers',
          knownAddresses: [],
          sheets: [
            { id: 'sh1', name: 'Customers' },
            { id: 'sh2', name: 'Orders' },
          ],
        }) as never,
    } as never)
  }

  it('names every label in play, across sheets', () => {
    const prompt = skillWith({ sh1: ['client phone'], sh2: ['order number'] }).systemPrompt
    expect(prompt).toContain('- {{client phone}}')
    expect(prompt).toContain('- {{order number}}')
  })

  it('warns about the two ways a model could recover a withheld value', () => {
    const prompt = skillWith({ sh1: ['client phone'] }).systemPrompt
    // Adding the hidden numbers back into a total.
    expect(prompt).toContain('left out of every statistic')
    // Reading an empty search as "not in the workbook".
    expect(prompt).toContain('absent from `find_cells` results')
  })

  it('says nothing about placeholders in a workbook that has none', () => {
    const prompt = skillWith({}).systemPrompt
    expect(prompt).not.toContain('Private placeholders')
    expect(prompt).not.toContain('{{')
  })

  it('reflects a mark added after the skill was built', () => {
    // The getter is what makes this true; a string captured at build time
    // would still describe the workbook as it was when the panel opened.
    const marks: Record<string, string[]> = {}
    const skill = skillWith(marks)
    expect(skill.systemPrompt).not.toContain('Private placeholders')
    marks.sh1 = ['client phone']
    expect(skill.systemPrompt).toContain('{{client phone}}')
  })
})

describe('a mark carries the fill it displaced', () => {
  /**
   * The tint is the only thing the mark ever does to the reader's own file, so
   * this is the promise that makes it acceptable: nothing is lost. Clearing has
   * to put the original colour back, not blank the cell.
   */
  const MARK = { startRow: 1, endRow: 1, startColumn: 1, endColumn: 1, label: 'client phone' }

  it('stores a previous fill and gives it back on clear', () => {
    const marked = addMark([], 'Customers', { ...MARK, previousFill: '#FFEE00' })
    expect(marked[0]?.marks[0]?.previousFill).toBe('#FFEE00')
    const cleared = clearMark(marked, 'Customers', marked[0]!.marks[0]!)
    // Nothing survives, so the reader's yellow is back where it was.
    expect(cleared).toEqual([])
  })

  it('records "no fill" as null rather than dropping the field', () => {
    // Dropping it would make clearing unable to tell "there was nothing" from
    // "this part predates the tint", and the first has to clear the cell.
    const marked = addMark([], 'Customers', { ...MARK, previousFill: null })
    expect(marked[0]?.marks[0]).toHaveProperty('previousFill', null)
  })

  it('leaves a mark without the field alone', () => {
    const marked = addMark([], 'Customers', MARK)
    expect(marked[0]?.marks[0]).not.toHaveProperty('previousFill')
  })

  it('keeps the displaced fill when a marked cell is re-marked', () => {
    // Re-marking replaces the mark, but the cell still carries the tint — so
    // the value to restore is the reader's original, not the tint.
    const once = addMark([], 'Customers', { ...MARK, previousFill: '#FFEE00' })
    const twice = addMark(once, 'Customers', { ...MARK, previousFill: '#FFEE00', label: 'mobile' })
    expect(twice[0]?.marks).toHaveLength(1)
    expect(twice[0]?.marks[0]?.previousFill).toBe('#FFEE00')
  })

  it('survives the carrier round trip with the fill intact', () => {
    const marked = addMark([], 'Customers', { ...MARK, previousFill: '#FFEE00' })
    expect(parseRedactionPart(serializeRedactionPart(marked))).toEqual(marked)
  })

  it('survives the round trip when there was no fill to restore', () => {
    const marked = addMark([], 'Customers', { ...MARK, previousFill: null })
    expect(parseRedactionPart(serializeRedactionPart(marked))).toEqual(marked)
  })
})
