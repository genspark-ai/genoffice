/**
 * Skill-template context variables in the sheets panel. SELECTION follows the
 * run's frozen scope — the same snapshot get_workbook_context reports — as a
 * sheet-qualified address plus the values while the range is small, capped
 * whole at 500 chars; a dropped or missing scope omits the variable exactly
 * like the composer's scope chip. The values come from the redaction-aware
 * readCells the model's tools use, stubbed here with plain cells (a withheld
 * cell simply arrives as its placeholder value).
 */
import { describe, expect, it, vi } from 'vitest'
import { sheetSkillVars, type ReadSelectionValues } from '../src/renderer/ai/skill-vars'

const NOW = new Date(2026, 9, 29)

describe('sheetSkillVars', () => {
  it('builds the standard set with the frozen scope as SELECTION', () => {
    const readValues = vi.fn<ReadSelectionValues>(() => ({
      B2: { value: 'Quarter' },
      C2: { value: 1200 },
      B3: { value: 'Q2' },
      C3: { value: 1400 },
    }))
    const vars = sheetSkillVars({
      lang: 'en',
      app: 'Sheets',
      fileName: 'Q3.xlsx',
      scope: { a1: 'B2:C3', sheetId: 's1' },
      sheetNameOf: () => 'Sheet1',
      readValues,
      now: NOW,
    })
    expect(readValues).toHaveBeenCalledWith(['B2', 'C2', 'B3', 'C3'], 's1')
    expect(vars.APP).toBe('Sheets')
    expect(vars.LANGUAGE).toBe('English')
    expect(vars.FILE_NAME).toBe('Q3.xlsx')
    expect(vars.TODAY).toBe(NOW.toLocaleDateString('en'))
    expect(vars.SELECTION).toBe('Sheet1!B2:C3: B2=Quarter, C2=1200; B3=Q2, C3=1400')
  })

  it('omits FILE_NAME and SELECTION when the app has no file and the run has no scope', () => {
    for (const scope of [null, undefined]) {
      const vars = sheetSkillVars({
        lang: 'en',
        app: 'Sheets',
        scope,
        sheetNameOf: () => 'Sheet1',
        readValues: () => ({}),
      })
      expect('FILE_NAME' in vars).toBe(false)
      expect('SELECTION' in vars).toBe(false)
    }
  })

  it('qualifies the address with the raw sheetId when the sheet name is unknown', () => {
    const vars = sheetSkillVars({
      lang: 'en',
      app: 'Sheets',
      scope: { a1: 'B2:C3', sheetId: 'sheet-7' },
      sheetNameOf: () => undefined,
      readValues: () => ({ B2: { value: 1 } }),
    })
    expect(vars.SELECTION).toBe('sheet-7!B2:C3: B2=1')
  })

  it('skips empty cells and puts a formula behind its computed value, like read_range', () => {
    const vars = sheetSkillVars({
      lang: 'en',
      app: 'Sheets',
      scope: { a1: 'B2:C3', sheetId: 's1' },
      sheetNameOf: () => 'Sheet1',
      readValues: () => ({
        B2: { value: null },
        C2: { value: '' },
        B3: { value: 42, formula: '=SUM(B2:B2)' },
        C3: { value: null, formula: '=C2*2' },
      }),
    })
    // the whole empty row drops out; a formula with no computed value stands alone
    expect(vars.SELECTION).toBe('Sheet1!B2:C3: B3=42 (=SUM(B2:B2)), C3==C2*2')
  })

  it('reads values at exactly the small-range cap but not past it', () => {
    // a full 8×8 grid so the row-major grouping is visible in the output
    const cells: Record<string, { value: number }> = {}
    let n = 0
    for (let row = 1; row <= 8; row += 1) {
      for (const column of ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']) {
        cells[`${column}${row}`] = { value: n }
        n += 1
      }
    }
    const small = vi.fn<ReadSelectionValues>(() => cells)
    const vars = sheetSkillVars({
      lang: 'en',
      app: 'Sheets',
      scope: { a1: 'A1:H8', sheetId: 's1' }, // 8×8 = 64 cells
      sheetNameOf: () => 'Sheet1',
      readValues: small,
    })
    expect(small).toHaveBeenCalledTimes(1)
    expect(vars.SELECTION).toContain('A1=0, B1=1')
    expect(vars.SELECTION).toContain('H1=7; A2=8')

    const big = vi.fn<ReadSelectionValues>(() => cells)
    const bigVars = sheetSkillVars({
      lang: 'en',
      app: 'Sheets',
      scope: { a1: 'A1:I8', sheetId: 's1' }, // 9×8 = 72 cells
      sheetNameOf: () => 'Sheet1',
      readValues: big,
    })
    expect(big).not.toHaveBeenCalled()
    expect(bigVars.SELECTION).toBe('Sheet1!A1:I8')
  })

  it('keeps the address alone when every cell is empty', () => {
    const vars = sheetSkillVars({
      lang: 'en',
      app: 'Sheets',
      scope: { a1: 'B2:C3', sheetId: 's1' },
      sheetNameOf: () => 'Sheet1',
      readValues: () => ({}),
    })
    expect(vars.SELECTION).toBe('Sheet1!B2:C3')
  })

  it('caps SELECTION at 500 chars, address included', () => {
    const long = 'x'.repeat(600)
    const vars = sheetSkillVars({
      lang: 'en',
      app: 'Sheets',
      scope: { a1: 'B2:D2', sheetId: 's1' },
      sheetNameOf: () => 'Sheet1',
      readValues: () => ({ B2: { value: long } }),
    })
    expect(vars.SELECTION!.length).toBe(500)
    expect(vars.SELECTION!.endsWith('…')).toBe(true)
  })

  it('survives a reader that throws mid-teardown: the address still names the target', () => {
    const vars = sheetSkillVars({
      lang: 'en',
      app: 'Sheets',
      scope: { a1: 'B2:C3', sheetId: 's1' },
      sheetNameOf: () => 'Sheet1',
      readValues: () => {
        throw new Error('workbook disposing')
      },
    })
    expect(vars.SELECTION).toBe('Sheet1!B2:C3')
  })
})
