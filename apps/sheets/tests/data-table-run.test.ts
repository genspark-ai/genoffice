import { beforeEach, describe, expect, it } from 'vitest'

import { applyDataTable } from '../src/renderer/data-table'
import { setManualCalculation } from '../src/renderer/calc-options'
import { setModuleLang } from '../src/renderer/i18n/locale'
import type { UniverRuntime } from '../src/renderer/univer-state'

beforeEach(() => {
  // Stable error-message assertions; zh is the authored source, en the mirror.
  setModuleLang('en')
})

type Scalar = number | string | boolean | null
interface ICellDataLike {
  readonly v?: Scalar
}

/**
 * A Univer runtime fake shaped exactly around what applyDataTable touches:
 * a cell map, a designated formula cell computed from the map on read, and a
 * formula service whose calculationEnd listeners fire on the microtask after
 * each write (synchronously for Calculate Now, which is manual mode).
 * `extraFormulaCells` marks additional coordinates that report a formula —
 * the input-cell validation tests need an input that "is" a formula while
 * the region's formula cell also reports one.
 */
function makeRuntime(
  formulaCell: { row: number; column: number },
  compute: () => Scalar,
  extraFormulaCells: readonly { row: number; column: number }[] = [],
) {
  const cells = new Map<string, Scalar>()
  const key = (row: number, column: number) => `${row}:${column}`
  const writes: string[] = []
  const events: string[] = []
  let calcEnds = 0
  let forcedCalcs = 0
  const listeners = new Set<() => void>()
  const fire = (forced: boolean): void => {
    calcEnds += 1
    if (forced) forcedCalcs += 1
    for (const listener of [...listeners]) listener()
  }
  const pendCalc = (): void => {
    queueMicrotask(() => fire(false))
  }
  const isFormula = (row: number, column: number): boolean =>
    (row === formulaCell.row && column === formulaCell.column) ||
    extraFormulaCells.some((cell) => cell.row === row && cell.column === column)
  const range = (row: number, column: number, numRows = 1, numColumns = 1) => ({
    getFormula: () => (isFormula(row, column) ? '=FAKE()' : null),
    getValue: () =>
      row === formulaCell.row && column === formulaCell.column
        ? compute()
        : (cells.get(key(row, column)) ?? null),
    setValue: (value: number | string | boolean | ICellDataLike) => {
      writes.push(`set ${key(row, column)} ${JSON.stringify(value)}`)
      if (typeof value === 'object' && value !== null) {
        if (value.v === null || value.v === undefined) cells.delete(key(row, column))
        else cells.set(key(row, column), value.v)
      } else cells.set(key(row, column), value)
      pendCalc()
    },
    setValues: (matrix: readonly (readonly ICellDataLike[])[]) => {
      writes.push(`setValues ${row}:${column} ${numRows}x${numColumns} ${JSON.stringify(matrix)}`)
      matrix.forEach((rowValues, rowIndex) => {
        rowValues.forEach((data, columnIndex) => {
          const at = key(row + rowIndex, column + columnIndex)
          if (data.v === null || data.v === undefined) cells.delete(at)
          else cells.set(at, data.v)
        })
      })
      pendCalc()
    },
  })
  const undoService = {
    __tempBatchingUndoRedo: () => {
      events.push('batch-begin')
      return { dispose: () => events.push('batch-dispose') }
    },
  }
  const commandService = { beforeCommandExecuted: () => undefined }
  const worksheet = { getRange: range }
  const workbook = { getId: () => 'wb', getActiveSheet: () => worksheet }
  const runtime = {
    univer: {
      __getInjector: () => ({
        get: (id: unknown) =>
          String(id) === 'univer.undo-redo.service' ? undoService : commandService,
      }),
    },
    univerAPI: {
      getActiveWorkbook: () => workbook,
      getFormula: () => ({
        calculationEnd: (listener: () => void) => {
          listeners.add(listener)
          return { dispose: () => listeners.delete(listener) }
        },
        executeCalculation: () => fire(true),
      }),
    },
  } as unknown as UniverRuntime
  const seed = (row: number, column: number, value: Scalar): void => {
    cells.set(`${row}:${column}`, value)
  }
  const read = (row: number, column: number): Scalar => cells.get(key(row, column)) ?? null
  return {
    runtime,
    seed,
    read,
    writes,
    events,
    calcEnds: () => calcEnds,
    forcedCalcs: () => forcedCalcs,
  }
}

const progress = () => {
  const calls: { done: number; total: number }[] = []
  return { calls, onProgress: (done: number, total: number) => void calls.push({ done, total }) }
}

describe('applyDataTable: one-variable column table', () => {
  it('writes one result per value, restores the input cell, batches one undo item', async () => {
    // Region B4:C7 — formula C4 = colInput*2, values B5:B7 = 1..3, input D1
    const fake = makeRuntime({ row: 3, column: 2 }, () => {
      const input = fake.read(0, 3)
      return typeof input === 'number' ? input * 2 : '#ERR'
    })
    fake.seed(0, 3, 10)
    fake.seed(4, 1, 1)
    fake.seed(5, 1, 2)
    fake.seed(6, 1, 3)
    const { calls, onProgress } = progress()

    const outcome = await applyDataTable(fake.runtime, {
      tableRange: 'B4:C7',
      rowInputCell: null,
      columnInputCell: 'D1',
      onProgress,
    })

    expect(outcome).toEqual({ kind: 'column', cells: 3 })
    expect(fake.read(4, 2)).toBe(2)
    expect(fake.read(5, 2)).toBe(4)
    expect(fake.read(6, 2)).toBe(6)
    // Excel behavior: the input cell keeps its original contents.
    expect(fake.read(0, 3)).toBe(10)
    expect(calls).toEqual([
      { done: 1, total: 3 },
      { done: 2, total: 3 },
      { done: 3, total: 3 },
    ])
    // One journaled block write for the results, then the input restore.
    expect(fake.writes.filter((write) => write.startsWith('setValues'))).toEqual([
      `setValues 4:2 3x1 ${JSON.stringify([[{ v: 2 }], [{ v: 4 }], [{ v: 6 }]])}`,
    ])
    expect(fake.events[0]).toBe('batch-begin')
    expect(fake.events[fake.events.length - 1]).toBe('batch-dispose')
    expect(fake.events.filter((event) => event === 'batch-begin')).toHaveLength(1)
  })

  it('restores an originally empty input cell to empty', async () => {
    const fake = makeRuntime({ row: 3, column: 2 }, () => {
      const input = fake.read(0, 3)
      return typeof input === 'number' ? input + 1 : '#ERR'
    })
    fake.seed(4, 1, 7)
    await applyDataTable(fake.runtime, {
      tableRange: 'B4:C5',
      rowInputCell: null,
      columnInputCell: 'D1',
    })
    expect(fake.read(0, 3)).toBeNull()
    expect(fake.read(4, 2)).toBe(8)
  })

  it('works in manual calculation mode via forced Calculate Now', async () => {
    const fake = makeRuntime({ row: 3, column: 2 }, () => {
      const input = fake.read(0, 3)
      return typeof input === 'number' ? input * 3 : '#ERR'
    })
    setManualCalculation(fake.runtime, true)
    fake.seed(0, 3, 0)
    fake.seed(4, 1, 2)
    fake.seed(5, 1, 4)
    await applyDataTable(fake.runtime, {
      tableRange: 'B4:C6',
      rowInputCell: null,
      columnInputCell: 'D1',
    })
    expect(fake.read(4, 2)).toBe(6)
    expect(fake.read(5, 2)).toBe(12)
    // Each substitution was forced through executeCalculation.
    expect(fake.forcedCalcs()).toBeGreaterThanOrEqual(3)
  })
})

describe('applyDataTable: one-variable row table', () => {
  it('lays results out across the formula row', async () => {
    // Region B4:D5 — formula B5 = rowInput+100, values C4:D4 = 1,2, input F9
    const fake = makeRuntime({ row: 4, column: 1 }, () => {
      const input = fake.read(8, 5)
      return typeof input === 'number' ? input + 100 : '#ERR'
    })
    fake.seed(8, 5, 0)
    fake.seed(3, 2, 1)
    fake.seed(3, 3, 2)
    const outcome = await applyDataTable(fake.runtime, {
      tableRange: 'B4:D5',
      rowInputCell: 'F9',
      columnInputCell: null,
    })
    expect(outcome).toEqual({ kind: 'row', cells: 2 })
    expect(fake.read(4, 2)).toBe(101)
    expect(fake.read(4, 3)).toBe(102)
    expect(fake.read(8, 5)).toBe(0)
  })
})

describe('applyDataTable: two-variable grid', () => {
  it('substitutes top-row values into the row input and left-column values into the column input', async () => {
    // Region B4:D6 — corner B4 = rowInput*10 + colInput; C4:D4 = 1,2; B5:B6 = 10,20
    const fake = makeRuntime({ row: 3, column: 1 }, () => {
      const rowInput = fake.read(0, 5)
      const columnInput = fake.read(1, 5)
      return typeof rowInput === 'number' && typeof columnInput === 'number'
        ? rowInput * 10 + columnInput
        : '#ERR'
    })
    fake.seed(0, 5, 0)
    fake.seed(1, 5, 0)
    fake.seed(3, 2, 1)
    fake.seed(3, 3, 2)
    fake.seed(4, 1, 10)
    fake.seed(5, 1, 20)
    const { calls, onProgress } = progress()

    const outcome = await applyDataTable(fake.runtime, {
      tableRange: 'B4:D6',
      rowInputCell: 'F1',
      columnInputCell: 'F2',
      onProgress,
    })

    expect(outcome).toEqual({ kind: 'grid', cells: 4 })
    // Results C5:D6 — rows follow the left column (10, 20), columns the top row (1, 2).
    expect(fake.read(4, 2)).toBe(20)
    expect(fake.read(4, 3)).toBe(30)
    expect(fake.read(5, 2)).toBe(30)
    expect(fake.read(5, 3)).toBe(40)
    expect(fake.read(0, 5)).toBe(0)
    expect(fake.read(1, 5)).toBe(0)
    expect(calls[calls.length - 1]).toEqual({ done: 4, total: 4 })
    expect(fake.writes.filter((write) => write.startsWith('setValues'))).toEqual([
      `setValues 4:2 2x2 ${JSON.stringify([
        [{ v: 20 }, { v: 30 }],
        [{ v: 30 }, { v: 40 }],
      ])}`,
    ])
  })
})

describe('applyDataTable: value handling', () => {
  it('writes formula error text as the cell value', async () => {
    const fake = makeRuntime({ row: 3, column: 2 }, () => {
      const input = fake.read(0, 3)
      return input === 0 ? '#DIV/0!' : '#ERR'
    })
    fake.seed(0, 3, 1)
    fake.seed(4, 1, 0)
    await applyDataTable(fake.runtime, {
      tableRange: 'B4:C5',
      rowInputCell: null,
      columnInputCell: 'D1',
    })
    expect(fake.read(4, 2)).toBe('#DIV/0!')
  })
})

describe('applyDataTable: validation and failure paths', () => {
  it('rejects a table region whose formula cell holds no formula, without writing', async () => {
    // Formula cell B4 is empty: compute() only answers for C4-shaped reads.
    const fake = makeRuntime({ row: 9, column: 9 }, () => 0)
    fake.seed(4, 1, 1)
    await expect(
      applyDataTable(fake.runtime, {
        tableRange: 'B4:C5',
        rowInputCell: null,
        columnInputCell: 'D1',
      }),
    ).rejects.toThrow('must contain a formula')
    expect(fake.writes).toHaveLength(0)
    expect(fake.events).toHaveLength(0)
  })

  it('rejects non-numeric substitution values', async () => {
    const fake = makeRuntime({ row: 3, column: 2 }, () => 1)
    fake.seed(0, 3, 0)
    fake.seed(4, 1, 'text')
    await expect(
      applyDataTable(fake.runtime, {
        tableRange: 'B4:C5',
        rowInputCell: null,
        columnInputCell: 'D1',
      }),
    ).rejects.toThrow('must be numbers')
  })

  it('rejects input cells holding formulas or text', async () => {
    const formulaFake = makeRuntime({ row: 3, column: 2 }, () => 1, [{ row: 0, column: 3 }])
    formulaFake.seed(4, 1, 1)
    // Input cell D1 reports a formula (the extra formula cell), so the run
    // would destroy it — Excel rejects this too.
    await expect(
      applyDataTable(formulaFake.runtime, {
        tableRange: 'B4:C5',
        rowInputCell: null,
        columnInputCell: 'D1',
      }),
    ).rejects.toThrow('plain number')

    const textFake = makeRuntime({ row: 3, column: 2 }, () => 1)
    textFake.seed(0, 3, 'abc')
    textFake.seed(4, 1, 1)
    await expect(
      applyDataTable(textFake.runtime, {
        tableRange: 'B4:C5',
        rowInputCell: null,
        columnInputCell: 'D1',
      }),
    ).rejects.toThrow('plain number')
  })

  it('rejects an input cell inside the table region and same row/column inputs', async () => {
    const fake = makeRuntime({ row: 3, column: 1 }, () => 1)
    fake.seed(3, 2, 1)
    fake.seed(4, 1, 2)
    fake.seed(0, 5, 0)
    fake.seed(1, 5, 0)
    await expect(
      applyDataTable(fake.runtime, {
        tableRange: 'B4:D6',
        rowInputCell: 'C4',
        columnInputCell: 'F2',
      }),
    ).rejects.toThrow('cannot be inside the table range')
    await expect(
      applyDataTable(fake.runtime, {
        tableRange: 'B4:D6',
        rowInputCell: 'F1',
        columnInputCell: 'F1',
      }),
    ).rejects.toThrow('must differ')
  })

  it('rejects a region larger than the cell cap before touching the sheet', async () => {
    const fake = makeRuntime({ row: 0, column: 0 }, () => 1)
    await expect(
      applyDataTable(fake.runtime, {
        tableRange: 'A1:CW102',
        rowInputCell: 'ZZ1',
        columnInputCell: 'ZZ2',
      }),
    ).rejects.toThrow('limit')
    expect(fake.writes).toHaveLength(0)
  })

  it('rejects a malformed range or input reference', async () => {
    const fake = makeRuntime({ row: 3, column: 2 }, () => 1)
    await expect(
      applyDataTable(fake.runtime, {
        tableRange: 'not a range',
        rowInputCell: null,
        columnInputCell: 'D1',
      }),
    ).rejects.toThrow('A1-style')
    await expect(
      applyDataTable(fake.runtime, {
        tableRange: 'B4:C5',
        rowInputCell: 'D1:D2',
        columnInputCell: null,
      }),
    ).rejects.toThrow('A1-style')
  })

  it('settles the undo batch and restores inputs when a read fails mid-run', async () => {
    let reads = 0
    const fake = makeRuntime({ row: 3, column: 2 }, () => {
      reads += 1
      if (reads > 2) throw new Error('read blew up')
      return 1
    })
    fake.seed(0, 3, 42)
    fake.seed(4, 1, 1)
    fake.seed(5, 1, 2)
    fake.seed(6, 1, 3)
    await expect(
      applyDataTable(fake.runtime, {
        tableRange: 'B4:C7',
        rowInputCell: null,
        columnInputCell: 'D1',
      }),
    ).rejects.toThrow('read blew up')
    // The third substitution never landed results, and D1 keeps its 42.
    expect(fake.read(0, 3)).toBe(42)
    expect(fake.read(4, 2)).toBeNull()
    expect(fake.events.filter((event) => event === 'batch-dispose')).toHaveLength(1)
  })
})
