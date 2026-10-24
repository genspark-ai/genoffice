import { describe, expect, it } from 'vitest'

import { NO_REDACTIONS, buildRedactionIndex } from '../src/renderer/ai/redact'
import { WITHHELD_UNKNOWN } from '../src/renderer/ai/redact-load'
import { modelTextOf, type ModelViewWorkbook } from '../src/renderer/ai/redact-view'
import type { SheetRedactionState } from '@genoffice/xlsx-gateway/gateway/xlsx-redaction'

const PHONE = '13800138000'
const LABEL = '客户电话'

const SHEETS = [
  { id: 'sh1', name: 'Customers' },
  { id: 'sh2', name: 'Orders' },
]

/** B2 holds the phone number; A1 and C1 are ordinary cells around it. */
const WORKBOOK: ModelViewWorkbook = {
  sheets: [
    {
      id: 'sh1',
      name: 'Customers',
      cells: {
        A1: { value: 'Acme Ltd' },
        B2: { value: PHONE },
        C1: { value: 42 },
      },
    },
    { id: 'sh2', name: 'Orders', cells: { A1: { value: 'widget' } } },
  ],
}

function indexWith(
  overrides: SheetRedactionState[] = [
    {
      sheetName: 'Customers',
      marks: [{ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1, label: LABEL }],
    },
  ],
) {
  return buildRedactionIndex(overrides, SHEETS)
}

describe('the workbook as the model may read it', () => {
  it('never carries the value of a withheld cell', () => {
    // The one assertion that matters. If this function can return raw text on
    // any path, the whole feature is a comment: the secret crosses the wire and
    // nothing else in the app notices.
    const out = modelTextOf(WORKBOOK, indexWith())
    expect(out).not.toContain(PHONE)
    expect(JSON.stringify(out)).not.toContain('13800138')
  })

  it('shows the label the reader chose, in place of the value', () => {
    const out = modelTextOf(WORKBOOK, indexWith())
    expect(out).toContain('B2: {{客户电话}}')
  })

  it('leaves the cells around a mark untouched', () => {
    // Over-redaction is its own bug: a workbook where everything reads as
    // withheld teaches the model nothing and hides the reader's own data from
    // the one consumer they wanted to use.
    const out = modelTextOf(WORKBOOK, indexWith())
    expect(out).toContain('A1: Acme Ltd')
    expect(out).toContain('C1: 42')
    expect(out).toContain('widget')
  })

  it('prints nothing withheld when nothing is withheld', () => {
    const out = modelTextOf(WORKBOOK, NO_REDACTIONS)
    expect(out).toContain('B2: 13800138000')
    expect(out).not.toContain('{{')
  })

  it('withholds every cell when the marks could not be read', () => {
    // WITHHELD_UNKNOWN is what an unreadable redaction part resolves to. Reading
    // it as "nothing is withheld" would hand over precisely the values the part
    // was written to hide.
    const out = modelTextOf(WORKBOOK, WITHHELD_UNKNOWN)
    expect(out).not.toContain(PHONE)
    expect(out).not.toContain('Acme Ltd')
    expect(out).toContain('{{private}}')
  })

  it('withholds a cell whose address cannot be placed on the grid', () => {
    const workbook: ModelViewWorkbook = {
      sheets: [{ id: 'sh1', name: 'Customers', cells: { 'not-an-address': { value: PHONE } } }],
    }
    const out = modelTextOf(workbook, indexWith())
    expect(out).not.toContain(PHONE)
    expect(out).toContain('{{private}}')
  })

  it('withholds a sheet whose cells cannot be looked up at all', () => {
    const workbook = {
      sheets: [{ id: '', name: 'Customers', cells: { A1: { value: 'Acme Ltd' } } }],
    } as unknown as ModelViewWorkbook
    const out = modelTextOf(workbook, indexWith())
    expect(out).not.toContain('Acme Ltd')
    expect(out).toContain('{{private}}')
  })

  it('withholds a cell whose value has no shape it can read', () => {
    // A cell value is a scalar. Anything else is a shape this function does not
    // recognise, and `String(value)` on it would serialise whatever it holds.
    const workbook = {
      sheets: [{ id: 'sh1', name: 'Customers', cells: { A1: { value: { secret: PHONE } } } }],
    } as unknown as ModelViewWorkbook
    const out = modelTextOf(workbook, NO_REDACTIONS)
    expect(out).not.toContain(PHONE)
    expect(out).not.toContain('secret')
  })

  it('folds a value that carries line breaks onto one line', () => {
    // One cell is one line here; a raw newline in a value tears the line
    // structure apart and the model reads the rest as another cell.
    const workbook: ModelViewWorkbook = {
      sheets: [{ id: 'sh1', name: 'Customers', cells: { A1: { value: 'one\ntwo\tthree' } } }],
    }
    expect(modelTextOf(workbook, NO_REDACTIONS)).toContain('A1: one two three')
  })

  it('lists cells in reading order', () => {
    const workbook: ModelViewWorkbook = {
      sheets: [
        {
          id: 'sh1',
          name: 'Customers',
          cells: { C1: { value: 'c' }, A2: { value: 'a2' }, B1: { value: 'b' } },
        },
      ],
    }
    expect(modelTextOf(workbook, NO_REDACTIONS)).toBe('## Customers (id=sh1)\nB1: b\nC1: c\nA2: a2')
  })

  it('answers with something rather than throwing on a workbook it does not recognise', () => {
    // A throw here is a leak with extra steps: the caller's fallback is its own
    // unredacted text.
    const junk = [null, undefined, {}, { sheets: 'not an array' }, { sheets: [null, 7] }]
    for (const workbook of junk) {
      expect(() => modelTextOf(workbook as unknown as ModelViewWorkbook, indexWith())).not.toThrow()
      expect(modelTextOf(workbook as unknown as ModelViewWorkbook, indexWith())).not.toContain(
        PHONE,
      )
    }
  })

  it('keeps an empty sheet visible', () => {
    const workbook: ModelViewWorkbook = { sheets: [{ id: 'sh1', name: 'Customers', cells: {} }] }
    expect(modelTextOf(workbook, indexWith())).toBe('## Customers (id=sh1)')
  })
})
