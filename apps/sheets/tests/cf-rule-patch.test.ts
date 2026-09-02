import { describe, expect, it } from 'vitest'
import { BorderStyleTypes } from '@univerjs/core'
import { FConditionalFormattingBuilder } from '@univerjs/sheets-conditional-formatting/facade'
import {
  buildConditionalRule,
  patchBuiltHighlightRule,
  type BuiltHighlightRule,
} from '../src/renderer/univer-sync'
import type { WorkbookCellStyle } from '../src/shared/desktop-api'

type CfRule = Parameters<typeof buildConditionalRule>[2]

const worksheet = {
  newConditionalFormattingRule: () => new FConditionalFormattingBuilder(),
} as unknown as Parameters<typeof buildConditionalRule>[0]

function cfRule(overrides: Partial<CfRule>): CfRule {
  return {
    ranges: [{ startRow: 0, startColumn: 0, endRow: 3, endColumn: 0 }],
    ruleType: 'cellIs',
    formulas: [],
    priority: 1,
    percent: false,
    bottom: false,
    cfvos: [],
    colors: [],
    iconReverse: false,
    showValue: true,
    ...overrides,
  }
}

function dxf(overrides: Partial<WorkbookCellStyle>): WorkbookCellStyle {
  return {
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    wrapText: false,
    diagonalUp: false,
    diagonalDown: false,
    ...overrides,
  }
}

describe('patchBuiltHighlightRule', () => {
  it('re-targets the placeholder text condition to containsErrors', () => {
    const target: BuiltHighlightRule = {
      operator: 'containsText',
      value: '',
      style: { bg: { rgb: '#FF0000' } },
    }
    patchBuiltHighlightRule(target, 'containsErrors', dxf({ fillColor: '#FF0000' }))
    expect(target.operator).toBe('containsErrors')
    expect('value' in target).toBe(false)
    expect(target.style).toMatchObject({ bg: { rgb: '#FF0000' } })
  })

  it('re-targets to notContainsErrors', () => {
    const target: BuiltHighlightRule = { operator: 'containsText', value: '' }
    patchBuiltHighlightRule(target, 'notContainsErrors', undefined)
    expect(target.operator).toBe('notContainsErrors')
    expect('value' in target).toBe(false)
  })

  it('leaves other rule types untouched', () => {
    const target: BuiltHighlightRule = { operator: 'containsText', value: 'apple' }
    patchBuiltHighlightRule(target, 'containsText', undefined)
    expect(target).toEqual({ operator: 'containsText', value: 'apple' })
  })

  it('injects dxf borders onto the built style', () => {
    const target: BuiltHighlightRule = {}
    patchBuiltHighlightRule(
      target,
      'cellIs',
      dxf({
        fontColor: '#008000',
        borderTop: { style: 'thin', color: '#008000' },
        borderBottom: { style: 'medium' },
      }),
    )
    expect(target.style?.bd).toEqual({
      t: { s: BorderStyleTypes.THIN, cl: { rgb: '#008000' } },
      b: { s: BorderStyleTypes.MEDIUM, cl: { rgb: '#000000' } },
    })
  })

  it('merges borders and number format without dropping builder style keys', () => {
    const target: BuiltHighlightRule = { style: { bg: { rgb: '#FFFF00' } } }
    patchBuiltHighlightRule(
      target,
      'expression',
      dxf({ numberFormat: '0.00%', borderLeft: { style: 'dashed', color: '#0000FF' } }),
    )
    expect(target.style).toMatchObject({
      bg: { rgb: '#FFFF00' },
      n: { pattern: '0.00%' },
      bd: { l: { s: BorderStyleTypes.DASHED, cl: { rgb: '#0000FF' } } },
    })
  })

  it('does not create a style for a border-free dxf', () => {
    const target: BuiltHighlightRule = {}
    patchBuiltHighlightRule(target, 'cellIs', dxf({ fillColor: '#FF0000' }))
    expect(target.style).toBeUndefined()
  })
})

describe('buildConditionalRule', () => {
  it('builds notContainsErrors as a text rule with the error operator and dxf fill', () => {
    const built = buildConditionalRule(
      worksheet,
      [dxf({ fillColor: '#FF0000' })],
      cfRule({ ruleType: 'notContainsErrors', dxfIndex: 0, formulas: ['NOT(ISERROR(A1))'] }),
    )
    expect(built?.rule).toMatchObject({
      type: 'highlightCell',
      subType: 'text',
      operator: 'notContainsErrors',
    })
    expect((built?.rule as { value?: string }).value).toBeUndefined()
    expect((built?.rule as BuiltHighlightRule).style?.bg).toBeTruthy()
    expect(built?.ranges).toEqual([{ startRow: 0, startColumn: 0, endRow: 3, endColumn: 0 }])
  })

  it('builds containsErrors', () => {
    const built = buildConditionalRule(
      worksheet,
      [],
      cfRule({ ruleType: 'containsErrors', formulas: ['ISERROR(A1)'] }),
    )
    expect(built?.rule).toMatchObject({ subType: 'text', operator: 'containsErrors' })
  })

  it('builds containsText as a formula rule so matching is case-insensitive (prod_048)', () => {
    // The rule text is "Short" but the cell holds "SHORT": Excel matches
    // (SEARCH semantics), Univer's native text condition (indexOf) does not.
    // The formula is synthesized against the sorted anchor, never taken from
    // the file (whose formula is anchored to its first sqref cell instead).
    const built = buildConditionalRule(
      worksheet,
      [dxf({ fillColor: '#FF0000' })],
      cfRule({
        ruleType: 'containsText',
        text: 'Short',
        dxfIndex: 0,
        formulas: ['NOT(ISERROR(SEARCH("Short",D3)))'],
      }),
    )
    expect(built?.rule).toMatchObject({
      subType: 'formula',
      value: '=NOT(ISERROR(SEARCH("Short",A1)))',
    })
  })

  it('synthesizes the formula when the file rule carries none', () => {
    const built = buildConditionalRule(
      worksheet,
      [],
      cfRule({ ruleType: 'endsWith', text: 'x"y', formulas: [] }),
    )
    expect(built?.rule).toMatchObject({
      subType: 'formula',
      value: '=RIGHT(A1,LEN("x""y"))="x""y"',
    })
  })

  it('keeps the native text condition for over-budget ranges', () => {
    const built = buildConditionalRule(
      worksheet,
      [],
      cfRule({
        ruleType: 'containsText',
        text: 'Short',
        formulas: ['NOT(ISERROR(SEARCH("Short",A1)))'],
        ranges: [{ startRow: 0, startColumn: 0, endRow: 99_999, endColumn: 4 }],
      }),
    )
    expect(built?.rule).toMatchObject({ subType: 'text', operator: 'containsText' })
  })

  it('carries cfRule/@stopIfTrue onto the built rule', () => {
    const built = buildConditionalRule(
      worksheet,
      [],
      cfRule({ ruleType: 'containsText', text: 'x', formulas: [], stopIfTrue: true }),
    )
    expect((built as { stopIfTrue?: boolean }).stopIfTrue).toBe(true)
    const without = buildConditionalRule(
      worksheet,
      [],
      cfRule({ ruleType: 'containsText', text: 'x', formulas: [] }),
    )
    expect((without as { stopIfTrue?: boolean }).stopIfTrue).toBe(false)
  })

  it('carries a dxf border onto the built cellIs rule style', () => {
    const built = buildConditionalRule(
      worksheet,
      [dxf({ bold: true, borderTop: { style: 'thin', color: '#92D050' } })],
      cfRule({ ruleType: 'cellIs', operator: 'equal', dxfIndex: 0, formulas: ['1'] }),
    )
    expect(built?.rule).toMatchObject({ subType: 'number', operator: 'equal' })
    const style = (built?.rule as BuiltHighlightRule).style
    expect(style?.bd).toEqual({ t: { s: BorderStyleTypes.THIN, cl: { rgb: '#92D050' } } })
    expect(style?.bl).toBe(1)
  })
})
