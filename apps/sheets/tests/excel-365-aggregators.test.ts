/**
 * GROUPBY / PIVOTBY: the executors, and the lexer rewrite that lets a bare
 * function name (GROUPBY's third argument, the shape Excel writes) travel
 * into the calculation as a string. The data set follows Excel's own
 * documentation examples. One shared workbook: the engine's settle timers
 * outlive an early dispose, so the cases run in sequence on one instance.
 */
import {
  ICommandService,
  IUniverInstanceService,
  LocaleType,
  LogLevel,
  Univer,
  UniverInstanceType,
} from '@univerjs/core'
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula'
import { describe, expect, it } from 'vitest'

import { installExcel365Aggregators } from '../src/renderer/excel-365-aggregators'
import type { UniverRuntime } from '../src/renderer/univer-state'

const settle = () => new Promise((resolve) => setTimeout(resolve, 1200))

describe('GROUPBY and PIVOTBY', () => {
  it('groups, pivots, totals, sorts and filters — and the bare function name travels', async () => {
    const univer = new Univer({ logLevel: LogLevel.ERROR, locale: LocaleType.EN_US, locales: {} })
    univer.registerPlugin(UniverFormulaEnginePlugin)
    univer.registerPlugin(UniverSheetsPlugin)
    univer.registerPlugin(UniverSheetsFormulaPlugin)
    const injector = univer.__getInjector()
    const runtime = { univer, univerAPI: null as never } as unknown as UniverRuntime
    installExcel365Aggregators(runtime)
    univer.createUnit(UniverInstanceType.UNIVER_SHEET, {
      id: 'wb1',
      sheetOrder: ['s1'],
      name: 'wb',
      styles: {},
      sheets: {
        s1: {
          id: 's1',
          name: 'S1',
          rowCount: 40,
          columnCount: 8,
          // Excel's documentation table: Product | Category | Sales
          cellData: {
            0: {
              0: { v: 'Product', t: 1 },
              1: { v: 'Category', t: 1 },
              2: { v: 'Sales', t: 1 },
            },
            1: { 0: { v: 'Apple', t: 1 }, 1: { v: 'Fruit', t: 1 }, 2: { v: 10 } },
            2: { 0: { v: 'Banana', t: 1 }, 1: { v: 'Fruit', t: 1 }, 2: { v: 20 } },
            3: { 0: { v: 'Carrot', t: 1 }, 1: { v: 'Veg', t: 1 }, 2: { v: 30 } },
            4: { 0: { v: 'Durian', t: 1 }, 1: { v: 'Veg', t: 1 }, 2: { v: 5 } },
            5: { 0: { v: 'Ambrosia', t: 1 }, 1: { v: 'Fruit', t: 1 }, 2: { v: 7 } },
          },
        },
      },
    } as never)
    const commandService = injector.get(ICommandService)
    let cursor = 10
    /** set one formula below the data, settle, return the raw value */
    const calc = async (formula: string) => {
      const at = cursor++
      await commandService.executeCommand('sheet.command.set-range-values', {
        unitId: 'wb1',
        subUnitId: 's1',
        range: { startRow: at, endRow: at + 1, startColumn: 0, endColumn: 1 },
        // the value keys are absolute row numbers, not range offsets
        value: { [at]: { 0: { f: formula } } },
      })
      await settle()
      const workbook = injector.get(IUniverInstanceService).getUnit('wb1') as unknown as {
        getSheetBySheetId(id: string): {
          getCellRaw(row: number, column: number): { v?: unknown } | undefined
        }
      }
      return workbook.getSheetBySheetId('s1').getCellRaw(at, 0)?.v
    }

    // the bare SUM in the function position is the shape Excel writes; the
    // lexer rewrite carries it into the executor as a string. Automatic
    // headers (values open with text over a number) plus the grand total
    // Excel shows by default: 30 + 30 + 60
    expect(await calc('=SUM(GROUPBY(B1:B4,C1:C4,SUM))')).toBe(120)
    // auto headers shown: the detected header row, then the groups
    expect(await calc('=INDEX(GROUPBY(B1:B4,C1:C4,SUM),1,1)')).toBe('Category')
    expect(await calc('=INDEX(GROUPBY(B1:B4,C1:C4,SUM),2,1)')).toBe('Fruit')
    // numeric values argue no headers, groups sort ascending
    // by the row field, and the grand total is there without asking
    expect(await calc('=INDEX(GROUPBY(B2:B4,C2:C4,SUM),1,1)')).toBe('Fruit')
    expect(await calc('=INDEX(GROUPBY(B2:B4,C2:C4,SUM),1,2)')).toBe(30)
    expect(await calc('=INDEX(GROUPBY(B2:B4,C2:C4,SUM),2,1)')).toBe('Veg')
    expect(await calc('=INDEX(GROUPBY(B2:B4,C2:C4,SUM),2,2)')).toBe(30)
    expect(await calc('=INDEX(GROUPBY(B2:B4,C2:C4,SUM),3,1)')).toBe('Total')
    expect(await calc('=INDEX(GROUPBY(B2:B4,C2:C4,SUM),3,2)')).toBe(60)
    // the default sort: source order starts at Apple, ascending starts at
    // Ambrosia
    expect(await calc('=INDEX(GROUPBY(A2:A6,C2:C6,SUM),1,1)')).toBe('Ambrosia')
    // field_headers 1: headers exist but stay hidden
    expect(await calc('=INDEX(GROUPBY(B1:B4,C1:C4,SUM,1),1,1)')).toBe('Fruit')
    expect(await calc('=SUM(GROUPBY(B1:B4,C1:C4,SUM,1))')).toBe(120)
    // field_headers 2: no headers in the data, positional labels generated
    expect(await calc('=INDEX(GROUPBY(B2:B4,C2:C4,SUM,2),1,1)')).toBe('Field 1')
    expect(await calc('=INDEX(GROUPBY(B2:B4,C2:C4,SUM,2),1,2)')).toBe('Value 1')
    expect(await calc('=INDEX(GROUPBY(B2:B4,C2:C4,SUM,2),2,1)')).toBe('Fruit')
    // field_headers 3: headers exist and show
    expect(await calc('=INDEX(GROUPBY(B1:B4,C1:C4,SUM,3),1,1)')).toBe('Category')
    expect(await calc('=SUM(GROUPBY(B1:B4,C1:C4,SUM,3))')).toBe(120)
    // total_depth 0 stays the way to ask for no grand total
    expect(await calc('=SUM(GROUPBY(B1:B4,C1:C4,SUM,3,0))')).toBe(60)
    // a nested SUM(...) keeps its own semantics (30+30 groups + 60 Total)
    expect(await calc('=SUM(INDEX(GROUPBY(B1:B4,C1:C4,SUM,0,1),0,2))')).toBe(120)
    // headers off, grand total on: one Fruit group of 30 with a Total row
    expect(await calc('=INDEX(GROUPBY(B2:B3,C2:C3,SUM,0,1,-1),1,2)')).toBe(30)
    expect(await calc('=INDEX(GROUPBY(B2:B3,C2:C3,SUM,0,1,-1),2,1)')).toBe('Total')
    // the filter array drops Banana before grouping: 10 + 30
    expect(await calc('=SUM(GROUPBY(B1:B4,C1:C4,SUM,3,0,0,{TRUE;TRUE;FALSE;TRUE}))')).toBe(40)
    // PIVOTBY: the column keys always ride the first row (they are the pivot's
    // column axis, not a header), blank corner without field names, and the
    // grand total column and row Excel shows by default
    expect(await calc('=INDEX(PIVOTBY(A2:A4,B2:B4,C2:C4,SUM,0),1,2)')).toBe('Fruit')
    expect(await calc('=INDEX(PIVOTBY(A2:A4,B2:B4,C2:C4,SUM,0),1,3)')).toBe('Veg')
    expect(await calc('=INDEX(PIVOTBY(A2:A4,B2:B4,C2:C4,SUM,0),1,4)')).toBe('Total')
    // Apple under Fruit = 10; Carrot under Veg = 30 (rows shifted by the keys row)
    expect(await calc('=INDEX(PIVOTBY(A2:A4,B2:B4,C2:C4,SUM,0),2,2)')).toBe(10)
    expect(await calc('=INDEX(PIVOTBY(A2:A4,B2:B4,C2:C4,SUM,0),4,3)')).toBe(30)
    // an empty intersection is blank and does not take the workbook down
    expect(String(await calc('=INDEX(PIVOTBY(A2:A4,B2:B4,C2:C4,SUM,0),2,3)'))).toBe('')
    // col_sort_order -1 puts the column keys in descending order
    expect(await calc('=INDEX(PIVOTBY(A2:A4,B2:B4,C2:C4,SUM,0,1,1,1,-1),1,2)')).toBe('Veg')
    // PIVOTBY's automatic headers: detected and shown — the row field's name
    // rides the keys row's corner
    expect(await calc('=INDEX(PIVOTBY(A1:A4,B1:B4,C1:C4,SUM),1,1)')).toBe('Product')
    // a GROUPBY total row over two row fields: the label cells past the first
    // are blanks a reader can touch without crashing
    expect(String(await calc('=INDEX(GROUPBY(A2:B4,C2:C4,SUM),4,2)'))).toBe('')
    // the workbook still calculates plainly afterwards
    expect(await calc('=SUM(C2:C4)')).toBe(60)
    univer.dispose()
  }, 60_000)
})
