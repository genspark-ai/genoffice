import { CellValueType, type IStyleData } from '@univerjs/core'
import { describe, expect, it } from 'vitest'
import {
  applyOperation,
  buildPasteGrid,
  isBlankSourceCell,
  linkFormula,
  needsOwnLayer,
  parseClipboardText,
  pasteSpecialAvailability,
  textCell,
  transformCell,
  transposeGrid,
  type PasteSpecialOptions,
  type SourceCell,
  type SourceGridCell,
} from '../src/renderer/paste-special'

const num = (v: number) => ({ v, t: CellValueType.NUMBER })
const str = (v: string) => ({ v, t: CellValueType.STRING })
const base: PasteSpecialOptions = {
  type: 'all',
  operation: 'none',
  skipBlanks: false,
  transpose: false,
  link: false,
}
const noRelocate = (f: string) => f

describe('transposeGrid', () => {
  it('swaps rows and columns, padding ragged rows', () => {
    expect(
      transposeGrid([
        [1, 2, 3],
        [4, 5],
      ]),
    ).toEqual([
      [1, 4],
      [2, 5],
      [3, undefined],
    ])
    expect(transposeGrid([])).toEqual([])
  })
})

describe('applyOperation', () => {
  it('combines numbers and treats a blank target as 0', () => {
    expect(applyOperation(num(10), num(4), 'add')?.v).toBe(14)
    expect(applyOperation(num(10), num(4), 'subtract')?.v).toBe(6)
    expect(applyOperation(num(10), num(4), 'multiply')?.v).toBe(40)
    expect(applyOperation(num(10), num(4), 'divide')?.v).toBe(2.5)
    expect(applyOperation(null, num(4), 'add')?.v).toBe(4)
    expect(applyOperation(null, num(4), 'subtract')?.v).toBe(-4)
  })
  it('leaves text or boolean operands and blank sources alone', () => {
    expect(applyOperation(str('x'), num(4), 'add')).toBeNull()
    expect(applyOperation(num(1), str('4'), 'add')).toBeNull()
    expect(applyOperation(num(1), { v: 1, t: CellValueType.BOOLEAN }, 'add')).toBeNull()
    expect(applyOperation(num(1), null, 'multiply')).toBeNull()
    expect(applyOperation(num(1), num(2), 'none')).toBeNull()
  })
  it('divides by zero into #DIV/0! and clears any target formula', () => {
    const r = applyOperation({ v: 8, f: '=A1', t: CellValueType.NUMBER }, num(0), 'divide')
    expect(r).toEqual({ v: '#DIV/0!', t: CellValueType.STRING, f: null, si: null })
    expect(applyOperation({ v: 8, f: '=A1', t: CellValueType.NUMBER }, num(2), 'add')).toEqual({
      v: 10,
      t: CellValueType.NUMBER,
      f: null,
      si: null,
    })
  })
})

describe('linkFormula', () => {
  it('omits the sheet on the same sheet and quotes names that need it', () => {
    expect(linkFormula(null, 0, 0)).toBe('=A1')
    expect(linkFormula('Data', 4, 27)).toBe('=Data!AB5')
    expect(linkFormula('Q1 data', 0, 1)).toBe("='Q1 data'!B1")
    expect(linkFormula("Bob's", 0, 0)).toBe("='Bob''s'!A1")
    expect(linkFormula('2024', 0, 0)).toBe("='2024'!A1")
  })
})

describe('transformCell', () => {
  const styledStyle: IStyleData = {
    bl: 1,
    n: { pattern: '0.00' },
    bd: { t: { s: 1, cl: { rgb: '#000000' } } },
  }
  const styled: NonNullable<SourceCell> = { v: 3, t: CellValueType.NUMBER, s: styledStyle }
  const input = (source: SourceGridCell['cell'], target?: ReturnType<typeof num>) => ({
    source,
    target,
    rowDelta: 1,
    columnDelta: 0,
  })
  it('splits the Excel paste types', () => {
    expect(transformCell(input(styled), { ...base, type: 'values' }, noRelocate)).toEqual({
      v: 3,
      t: CellValueType.NUMBER,
      f: null,
      si: null,
      p: null,
    })
    expect(transformCell(input(styled), { ...base, type: 'formats' }, noRelocate)).toEqual({
      s: styled.s,
    })
    expect(transformCell(input(styled), { ...base, type: 'values-numfmt' }, noRelocate)).toEqual({
      v: 3,
      t: CellValueType.NUMBER,
      f: null,
      si: null,
      p: null,
      s: { n: { pattern: '0.00' } },
    })
    const noBorders = transformCell(
      input(styled),
      { ...base, type: 'all-except-borders' },
      noRelocate,
    )
    expect(noBorders?.s).toEqual({ bl: 1, n: { pattern: '0.00' } })
    expect(transformCell(input(styled), { ...base, type: 'col-widths' }, noRelocate)).toBeNull()
  })
  it('relocates formulas by the paste offset and keeps values for Formulas paste', () => {
    const relocate = (f: string, dr: number, dc: number) => `${f}@${dr},${dc}`
    expect(
      transformCell(input({ v: 5, f: '=A1' }), { ...base, type: 'formulas' }, relocate),
    ).toEqual({ f: '=A1@1,0', si: null, v: null, p: null })
    expect(transformCell(input(num(5)), { ...base, type: 'formulas' }, relocate)).toMatchObject({
      v: 5,
      f: null,
    })
    expect(
      transformCell(
        input({ v: 5, f: '=A1' }),
        { ...base, type: 'formulas', transpose: true },
        relocate,
      ),
    ).toMatchObject({ v: 5, f: null })
  })
  it('skips blanks and applies the operation on value pastes only', () => {
    expect(transformCell(input(null), { ...base, skipBlanks: true }, noRelocate)).toBeNull()
    expect(transformCell(input(null), base, noRelocate)).toMatchObject({ v: null })
    expect(
      transformCell(
        input(num(2), num(5)),
        { ...base, type: 'values', operation: 'multiply' },
        noRelocate,
      ),
    ).toMatchObject({ v: 10 })
    expect(
      transformCell(input(styled, num(5)), { ...base, type: 'all', operation: 'add' }, noRelocate),
    ).toEqual({ v: 8, t: CellValueType.NUMBER, f: null, si: null, s: styled.s })
    expect(
      transformCell(
        input(num(2), num(5)),
        { ...base, type: 'formats', operation: 'add' },
        noRelocate,
      ),
    ).toBeNull()
  })
})

describe('buildPasteGrid', () => {
  const cell = (v: number | null, row: number, column: number): SourceGridCell => ({
    cell: v === null ? null : num(v),
    row,
    column,
  })
  const source = [
    [cell(1, 0, 0), cell(2, 0, 1), cell(3, 0, 2)],
    [cell(4, 1, 0), cell(null, 1, 1), cell(6, 1, 2)],
  ]
  const target = (rows: number, columns: number) => ({
    startRow: 10,
    startColumn: 5,
    rows,
    columns,
    cellAt: () => num(100),
  })
  const values = (grid: ReturnType<typeof buildPasteGrid>) =>
    grid.cells.map((line) => line.map((c) => (c ? (c.f ?? c.v) : null)))

  it('transposes before writing', () => {
    const grid = buildPasteGrid(
      source,
      { ...base, transpose: true },
      target(1, 1),
      noRelocate,
      null,
    )
    expect([grid.rows, grid.columns]).toEqual([3, 2])
    expect(values(grid)).toEqual([
      [1, 4],
      [2, null],
      [3, 6],
    ])
  })
  it('tiles when the selection is an exact multiple and honours skip blanks', () => {
    const grid = buildPasteGrid(
      source,
      { ...base, skipBlanks: true, operation: 'add' },
      target(4, 3),
      noRelocate,
      null,
    )
    expect([grid.rows, grid.columns]).toEqual([4, 3])
    expect(values(grid)).toEqual([
      [101, 102, 103],
      [104, null, 106],
      [101, 102, 103],
      [104, null, 106],
    ])
  })
  it('ignores a selection that is not a multiple', () => {
    const grid = buildPasteGrid(source, base, target(3, 2), noRelocate, null)
    expect([grid.rows, grid.columns]).toEqual([2, 3])
  })
  it('writes link formulas to the source cells, never transposed', () => {
    const grid = buildPasteGrid(
      source,
      { ...base, link: true, transpose: true },
      target(1, 1),
      noRelocate,
      'Data',
    )
    expect(values(grid)).toEqual([
      ['=Data!A1', '=Data!B1', '=Data!C1'],
      ['=Data!A2', '=Data!B2', '=Data!C2'],
    ])
    const same = buildPasteGrid(source, { ...base, link: true }, target(1, 1), noRelocate, null)
    expect(values(same)[0]![0]).toBe('=A1')
  })
  it('hands the formula relocator the offset from source to target', () => {
    const seen: [number, number][] = []
    const relocate = (f: string, dr: number, dc: number) => {
      seen.push([dr, dc])
      return f
    }
    buildPasteGrid(
      [[{ cell: { v: 1, f: '=B2' }, row: 2, column: 3 }]],
      { ...base, type: 'formulas' },
      target(1, 1),
      relocate,
      null,
    )
    expect(seen).toEqual([[8, 2]])
  })
})

describe('clipboard text source', () => {
  it('parses TSV with quoted fields and types the cells', () => {
    expect(parseClipboardText('a\tb\r\n"x\ty"\t"say ""hi"""\n')).toEqual([
      ['a', 'b'],
      ['x\ty', 'say "hi"'],
    ])
    expect(textCell('12.5')).toEqual({ v: 12.5, t: CellValueType.NUMBER })
    expect(textCell('TRUE')).toEqual({ v: 1, t: CellValueType.BOOLEAN })
    expect(textCell('abc')).toEqual({ v: 'abc', t: CellValueType.STRING })
    expect(textCell('')).toBeNull()
    expect(isBlankSourceCell(textCell(''))).toBe(true)
  })
})

describe('availability and routing', () => {
  it('delegates plain variants to Univer and keeps the switches for our layer', () => {
    expect(needsOwnLayer(base)).toBe(false)
    expect(needsOwnLayer({ ...base, type: 'values' })).toBe(false)
    expect(needsOwnLayer({ ...base, type: 'values-numfmt' })).toBe(true)
    expect(needsOwnLayer({ ...base, transpose: true })).toBe(true)
    expect(needsOwnLayer({ ...base, operation: 'add' })).toBe(true)
    expect(needsOwnLayer({ ...base, link: true })).toBe(true)
    expect(needsOwnLayer({ ...base, type: 'col-widths', transpose: true, operation: 'add' })).toBe(
      false,
    )
  })
  it('narrows the dialog for cut, outside text and empty clipboards', () => {
    const internal = {
      kind: 'internal' as const,
      unitId: 'u',
      subUnitId: 's',
      rows: [0],
      cols: [0],
    }
    const copy = pasteSpecialAvailability({ ...internal, cut: false })
    expect(copy.types.has('formulas-numfmt')).toBe(true)
    expect(copy.types.has('comments')).toBe(false)
    expect(copy.link).toBe(true)
    const cut = pasteSpecialAvailability({ ...internal, cut: true })
    expect([...cut.types]).toEqual(['all'])
    expect(cut.transpose).toBe(false)
    const text = pasteSpecialAvailability({ kind: 'text', rows: [['1']] })
    expect([...text.types].sort()).toEqual(['all', 'values'])
    expect(text.transpose).toBe(true)
    expect(text.operation).toBe(false)
    expect(pasteSpecialAvailability({ kind: 'none' }).types.size).toBe(0)
  })
})
