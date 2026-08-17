import { describe, expect, it } from 'vitest'

import {
  MAX_EXPANDED_CELL_OPS,
  expandToPrimitiveOps,
  structuralOpLabel,
  workbookCommandBatchSchema,
  workbookOperationSchema,
} from '../src/domain/workbook-dsl'
import { columnIndex, columnLabel, parseRange, rangeCellCount } from '../src/domain/cell-address'

describe('cell-address helpers', () => {
  it('round-trips column labels', () => {
    for (const [label, index] of [
      ['A', 0],
      ['Z', 25],
      ['AA', 26],
      ['AZ', 51],
      ['BA', 52],
    ] as const) {
      expect(columnIndex(label)).toBe(index)
      expect(columnLabel(index)).toBe(label)
    }
  })

  it('normalizes reversed ranges and accepts single cells', () => {
    expect(parseRange('C3:A1')).toEqual({ startRow: 0, startColumn: 0, endRow: 2, endColumn: 2 })
    expect(parseRange('B2')).toEqual({ startRow: 1, startColumn: 1, endRow: 1, endColumn: 1 })
    expect(rangeCellCount(parseRange('A1:B5'))).toBe(10)
  })
})

describe('expandToPrimitiveOps', () => {
  it('expands set_range row-major from the start cell', () => {
    const ops = expandToPrimitiveOps([
      {
        op: 'set_range',
        sheetId: 's',
        start: 'B2',
        values: [
          [1, '=A1'],
          [null, 'text'],
        ],
      },
    ])
    expect(ops).toEqual([
      { op: 'set_cell', sheetId: 's', address: 'B2', value: 1 },
      { op: 'set_formula', sheetId: 's', address: 'C2', formula: '=A1' },
      { op: 'set_cell', sheetId: 's', address: 'B3', value: null },
      { op: 'set_cell', sheetId: 's', address: 'C3', value: 'text' },
    ])
  })

  it('rejects jagged set_range values with an instructive error', () => {
    expect(() =>
      expandToPrimitiveOps([
        {
          op: 'set_range',
          sheetId: 's',
          start: 'A1',
          values: [['Q', 'A', 'note'], ['Q only']],
        },
      ]),
    ).toThrow(/rectangular.*row 1 has 3 cell\(s\) but row 2 has 1/)
  })

  it('accepts range as an alias for start when its size matches values', () => {
    const ops = expandToPrimitiveOps([
      {
        op: 'set_range',
        sheetId: 's',
        range: 'B2:C3',
        values: [
          [1, 2],
          [3, 4],
        ],
      },
    ])
    expect(ops.map((op) => (op as { address: string }).address)).toEqual(['B2', 'C2', 'B3', 'C3'])
  })

  it('rejects a set_range range whose size does not match the values grid', () => {
    expect(() =>
      expandToPrimitiveOps([
        {
          op: 'set_range',
          sheetId: 's',
          range: 'A1:D1',
          values: [['only', 'two']],
        },
      ]),
    ).toThrow(/spans 1×4 cells but values is 1 row\(s\) × 2 cell\(s\)/)
  })

  it('size-checks range even when start is also given, and rejects disagreeing targets', () => {
    // consistent start + range: allowed, and the size check still applies
    expect(() =>
      expandToPrimitiveOps([
        {
          op: 'set_range',
          sheetId: 's',
          start: 'B2',
          range: 'B2:C3',
          values: [
            [1, 2],
            [3, 4],
          ],
        },
      ]),
    ).not.toThrow()
    expect(() =>
      expandToPrimitiveOps([
        {
          op: 'set_range',
          sheetId: 's',
          start: 'B2',
          range: 'B2:D4',
          values: [
            [1, 2],
            [3, 4],
          ],
        },
      ]),
    ).toThrow(/spans 3×3 cells but values is 2 row\(s\) × 2 cell\(s\)/)
    expect(() =>
      expandToPrimitiveOps([
        {
          op: 'set_range',
          sheetId: 's',
          start: 'A1',
          range: 'B2:C3',
          values: [
            [1, 2],
            [3, 4],
          ],
        },
      ]),
    ).toThrow(/disagree on the top-left cell/)
  })

  it('rejects a set_range with neither start nor range', () => {
    expect(() =>
      expandToPrimitiveOps([
        {
          op: 'set_range',
          sheetId: 's',
          values: [[1]],
        },
      ]),
    ).toThrow(/needs "start"/)
  })

  it('expands clear_range into clear_cell primitives', () => {
    const ops = expandToPrimitiveOps([{ op: 'clear_range', sheetId: 's', range: 'A1:B1' }])
    expect(ops).toEqual([
      { op: 'clear_cell', sheetId: 's', address: 'A1' },
      { op: 'clear_cell', sheetId: 's', address: 'B1' },
    ])
  })

  it('passes structural and rename operations through untouched', () => {
    const structural = { op: 'insert_rows' as const, sheetId: 's', row: 3, count: 2 }
    const rename = { op: 'rename_sheet' as const, sheetId: 's', name: 'Data' }
    expect(expandToPrimitiveOps([structural, rename])).toEqual([structural, rename])
  })

  it('rejects batches that expand past the cell-op ceiling', () => {
    expect(() =>
      expandToPrimitiveOps([
        { op: 'clear_range', sheetId: 's', range: `A1:C${MAX_EXPANDED_CELL_OPS}` },
      ]),
    ).toThrow(/2000/)
  })
})

describe('workbookCommandBatchSchema mixing rule', () => {
  const base = {
    dslVersion: 1 as const,
    transactionId: 'tx',
    baseRevision: 0,
    summary: 'mix check',
  }

  it('rejects structural + cell-content in one batch', () => {
    expect(() =>
      workbookCommandBatchSchema.parse({
        ...base,
        operations: [
          { op: 'add_sheet', name: 'New' },
          { op: 'set_cell', sheetId: 's', address: 'A1', value: 1 },
        ],
      }),
    ).toThrow(/separate batches/)
  })

  it('allows rename_sheet alongside either kind', () => {
    expect(() =>
      workbookCommandBatchSchema.parse({
        ...base,
        operations: [
          { op: 'rename_sheet', sheetId: 's', name: 'Data' },
          { op: 'insert_rows', sheetId: 's', row: 1, count: 1 },
        ],
      }),
    ).not.toThrow()
    expect(() =>
      workbookCommandBatchSchema.parse({
        ...base,
        operations: [
          { op: 'rename_sheet', sheetId: 's', name: 'Data' },
          { op: 'set_range', sheetId: 's', start: 'A1', values: [[1]] },
        ],
      }),
    ).not.toThrow()
  })
})

describe('operation schemas', () => {
  it('caps structural counts and validates column labels', () => {
    expect(() =>
      workbookOperationSchema.parse({ op: 'insert_rows', sheetId: 's', row: 0, count: 1 }),
    ).toThrow()
    expect(() =>
      workbookOperationSchema.parse({ op: 'delete_rows', sheetId: 's', row: 1, count: 501 }),
    ).toThrow()
    expect(() =>
      workbookOperationSchema.parse({ op: 'insert_cols', sheetId: 's', column: 'c', count: 1 }),
    ).toThrow()
    expect(() => workbookOperationSchema.parse({ op: 'add_sheet', name: 'bad[name]' })).toThrow()
  })
})

describe('extended operations', () => {
  it('parses the new visual, data-tool, and sheet-management ops', () => {
    const ops = [
      { op: 'add_chart', sheetId: 's', chartType: 'pie', dataRange: 'A1:B10', title: 'Share' },
      { op: 'add_shape', sheetId: 's', shapeType: 'textbox', anchorCell: 'E2', text: 'Note' },
      { op: 'edit_shape', visualId: 'added-shape-abc-1', text: 'New text', fillColor: '#DDEBF7' },
      { op: 'edit_shape', visualId: 'added-shape-abc-1', anchorCell: 'H4' },
      { op: 'add_image', sheetId: 's', path: '~/logo.png', anchorCell: 'B2' },
      {
        op: 'add_image',
        sheetId: 's',
        path: 'https://cdn.example.com/gen/img-1',
        anchorCell: 'B9',
      },
      { op: 'add_table', sheetId: 's', range: 'A1:D10' },
      {
        op: 'add_table',
        sheetId: 's',
        range: 'A1:D10',
        name: 'Sales_Q3',
        style: 'TableStyleLight9',
        bandedRows: false,
      },
      {
        op: 'add_pivot',
        sheetId: 's',
        sourceRange: 'A1:D100',
        targetCell: 'F1',
        rowFields: 'Region',
        values: [
          { field: 'Amount', agg: 'sum' },
          { field: 'Quantity', agg: 'count' },
        ],
      },
      {
        op: 'add_pivot',
        sheetId: 's',
        sourceRange: 'A1:D100',
        targetCell: 'F1',
        targetSheetId: 's2',
        rowFields: 'Region',
        columnField: 'Quarter',
        values: [{ field: 'Amount', agg: 'average' }],
        name: 'Summary',
      },
      { op: 'set_rows_hidden', sheetId: 's', row: 3, count: 2, hidden: true },
      { op: 'set_cols_hidden', sheetId: 's', column: 'C', hidden: false },
      { op: 'set_hyperlink', sheetId: 's', address: 'B2', target: 'https://example.com' },
      { op: 'set_hyperlink', sheetId: 's', address: 'B3', target: null },
      { op: 'protect_sheet', sheetId: 's', protected: true },
      { op: 'set_filter', sheetId: 's', range: 'A1:D100' },
      { op: 'clear_filter', sheetId: 's' },
      { op: 'set_filter_criteria', sheetId: 's', column: 'B', values: ['East', 'South'] },
      { op: 'set_filter_criteria', sheetId: 's', column: 'B', values: null },
      {
        op: 'add_conditional_format',
        sheetId: 's',
        range: 'B2:B50',
        rule: {
          kind: 'number',
          operator: 'greaterThan',
          value: 100,
          format: { fillColor: '#C6EFCE' },
        },
      },
      { op: 'clear_conditional_formats', sheetId: 's' },
      {
        op: 'set_data_validation',
        sheetId: 's',
        range: 'C2:C50',
        validation: { kind: 'list', values: ['Yes', 'No'] },
      },
      { op: 'set_data_validation', sheetId: 's', range: 'C2:C50', validation: null },
      { op: 'add_defined_name', name: 'SalesArea', ref: 'Sheet1!$A$1:$B$10' },
      { op: 'delete_defined_name', name: 'SalesArea' },
      {
        op: 'set_page_setup',
        sheetId: 's',
        orientation: 'landscape',
        paperSize: 9,
        printArea: 'A1:H40',
      },
      { op: 'set_freeze', sheetId: 's', rows: 1, columns: 0 },
      { op: 'set_freeze', sheetId: 's', rows: 0, columns: 0 },
      { op: 'set_note', sheetId: 's', address: 'B2', text: 'Tax inclusive' },
      { op: 'set_note', sheetId: 's', address: 'B2', text: null },
      { op: 'refresh_pivot', sheetId: 's' },
      { op: 'set_page_setup', sheetId: 's', fitToWidth: 1, printArea: null },
      { op: 'duplicate_sheet', sheetId: 's', name: 'Copy' },
      { op: 'set_sheet_hidden', sheetId: 's', hidden: true },
      { op: 'move_sheet', sheetId: 's', position: 1 },
    ]
    for (const op of ops) {
      expect(() => workbookOperationSchema.parse(op), JSON.stringify(op)).not.toThrow()
    }
  })

  it('accepts the extended format_range fields', () => {
    expect(() =>
      workbookOperationSchema.parse({
        op: 'format_range',
        sheetId: 's',
        range: 'A1:C1',
        format: {
          strikethrough: true,
          fontFamily: '微软雅黑',
          fontSize: 14,
          verticalAlign: 'center',
          wrapText: true,
          textRotation: 45,
          indent: 2,
        },
      }),
    ).not.toThrow()
    expect(() =>
      workbookOperationSchema.parse({
        op: 'format_range',
        sheetId: 's',
        range: 'A1',
        format: { textRotation: 'vertical' },
      }),
    ).not.toThrow()
    expect(() =>
      workbookOperationSchema.parse({
        op: 'format_range',
        sheetId: 's',
        range: 'A1',
        format: { textRotation: 120 },
      }),
    ).toThrow()
    expect(() =>
      workbookOperationSchema.parse({
        op: 'format_range',
        sheetId: 's',
        range: 'A1',
        format: { fontSize: 0 },
      }),
    ).toThrow()
  })

  it('rejects invalid extended ops at the schema layer', () => {
    expect(() =>
      workbookOperationSchema.parse({
        op: 'add_chart',
        sheetId: 's',
        chartType: 'donut',
        dataRange: 'A1:B10',
      }),
    ).toThrow()
    expect(() =>
      workbookOperationSchema.parse({
        op: 'add_shape',
        sheetId: 's',
        shapeType: 'star',
        anchorCell: 'A1',
      }),
    ).toThrow()
    expect(() =>
      workbookOperationSchema.parse({ op: 'add_defined_name', name: '1bad', ref: 'A1' }),
    ).toThrow()
    expect(() =>
      workbookOperationSchema.parse({
        op: 'set_data_validation',
        sheetId: 's',
        range: 'A1',
        validation: { kind: 'dateBetween', start: '2026/01/01', end: '2026-12-31' },
      }),
    ).toThrow()
  })

  it('requires at least one property on edit_shape', () => {
    expect(() => expandToPrimitiveOps([{ op: 'edit_shape', visualId: 'added-shape-x-1' }])).toThrow(
      /at least one/,
    )
    expect(() =>
      expandToPrimitiveOps([{ op: 'edit_shape', visualId: 'added-shape-x-1', text: 'ok' }]),
    ).not.toThrow()
  })

  it('validates set_page_setup at expansion', () => {
    expect(() => expandToPrimitiveOps([{ op: 'set_page_setup', sheetId: 's' }])).toThrow(
      /at least one/,
    )
    expect(() =>
      expandToPrimitiveOps([
        {
          op: 'set_page_setup',
          sheetId: 's',
          scale: 80,
          fitToWidth: 1,
        },
      ]),
    ).toThrow(/mutually exclusive/)
    expect(() =>
      expandToPrimitiveOps([
        {
          op: 'set_page_setup',
          sheetId: 's',
          orientation: 'landscape',
        },
      ]),
    ).not.toThrow()
  })

  it('rejects pivot cross-field violations at expansion', () => {
    expect(() =>
      expandToPrimitiveOps([
        {
          op: 'add_pivot',
          sheetId: 's',
          sourceRange: 'A1:D100',
          targetCell: 'F1',
          rowFields: 'Region',
          columnField: 'Quarter',
          values: [
            { field: 'Amount', agg: 'sum' },
            { field: 'Quantity', agg: 'sum' },
          ],
        },
      ]),
    ).toThrow(/exactly one values entry/)
    expect(() =>
      expandToPrimitiveOps([
        {
          op: 'add_pivot',
          sheetId: 's',
          sourceRange: 'A1:D100',
          targetCell: 'F1',
          rowFields: 'Region',
          values: [{ field: 'Region', agg: 'count' }],
        },
      ]),
    ).toThrow(/cannot also be/)
  })

  it('validates cross-field rules at expansion', () => {
    expect(() =>
      expandToPrimitiveOps([
        {
          op: 'add_conditional_format',
          sheetId: 's',
          range: 'A1:A5',
          rule: { kind: 'number', operator: 'between', value: 1, format: { bold: true } },
        },
      ]),
    ).toThrow(/value2/)
    expect(() =>
      expandToPrimitiveOps([
        {
          op: 'set_data_validation',
          sheetId: 's',
          range: 'A1',
          validation: { kind: 'numberBetween', min: 10, max: 1 },
        },
      ]),
    ).toThrow(/min/)
    expect(() =>
      expandToPrimitiveOps([
        {
          op: 'set_data_validation',
          sheetId: 's',
          range: 'A1',
          validation: { kind: 'dateBetween', start: '2026-12-31', end: '2026-01-01' },
        },
      ]),
    ).toThrow(/start/)
  })

  it('classifies new layout ops as content-compatible and sheet ops as structural', () => {
    const base = { dslVersion: 1 as const, transactionId: 'tx', baseRevision: 0, summary: 'mix' }
    expect(() =>
      workbookCommandBatchSchema.parse({
        ...base,
        operations: [
          { op: 'set_cell', sheetId: 's', address: 'A1', value: 1 },
          { op: 'add_chart', sheetId: 's', chartType: 'column', dataRange: 'A1:B5' },
          { op: 'set_hyperlink', sheetId: 's', address: 'A1', target: 'https://example.com' },
        ],
      }),
    ).not.toThrow()
    expect(() =>
      workbookCommandBatchSchema.parse({
        ...base,
        operations: [
          { op: 'duplicate_sheet', sheetId: 's' },
          { op: 'set_cell', sheetId: 's', address: 'A1', value: 1 },
        ],
      }),
    ).toThrow(/separate batches/)
  })
})

describe('structuralOpLabel', () => {
  it('describes each structural operation', () => {
    expect(structuralOpLabel({ op: 'insert_rows', sheetId: 's', row: 5, count: 2 })).toBe(
      'Insert 2 rows before row 5',
    )
    expect(structuralOpLabel({ op: 'delete_rows', sheetId: 's', row: 2, count: 3 })).toBe(
      'Delete rows 2–4',
    )
    expect(structuralOpLabel({ op: 'insert_cols', sheetId: 's', column: 'C', count: 1 })).toBe(
      'Insert 1 column before column C',
    )
    expect(structuralOpLabel({ op: 'delete_cols', sheetId: 's', column: 'B', count: 2 })).toBe(
      'Delete columns B–C',
    )
    expect(structuralOpLabel({ op: 'add_sheet', name: 'Summary' })).toBe('Add sheet "Summary"')
  })
})

describe('find_replace expansion', () => {
  const reader =
    (values: Record<string, string | number>) =>
    (address: string): { value: string | number | null; formula?: string } => ({
      value: values[address] ?? null,
    })

  it('replaces substrings case-insensitively by default, skipping non-text', () => {
    const ops = expandToPrimitiveOps(
      [
        {
          op: 'find_replace',
          sheetId: 's',
          range: 'A1:A4',
          find: 'apple',
          replace: 'pear',
        },
      ],
      reader({ A1: 'Apple pie', A2: 'APPLE apple', A3: 42, A4: 'banana' }),
    )
    expect(ops).toEqual([
      { op: 'set_cell', sheetId: 's', address: 'A1', value: 'pear pie' },
      { op: 'set_cell', sheetId: 's', address: 'A2', value: 'pear pear' },
    ])
  })

  it('honors matchCase and wholeCell', () => {
    const caseOps = expandToPrimitiveOps(
      [
        {
          op: 'find_replace',
          sheetId: 's',
          range: 'A1:A2',
          find: 'Apple',
          replace: 'Pear',
          matchCase: true,
        },
      ],
      reader({ A1: 'Apple', A2: 'apple' }),
    )
    expect(caseOps).toEqual([{ op: 'set_cell', sheetId: 's', address: 'A1', value: 'Pear' }])
    const wholeOps = expandToPrimitiveOps(
      [
        {
          op: 'find_replace',
          sheetId: 's',
          range: 'A1:A2',
          find: 'apple',
          replace: 'pear',
          wholeCell: true,
        },
      ],
      reader({ A1: 'apple', A2: 'apple pie' }),
    )
    expect(wholeOps).toEqual([{ op: 'set_cell', sheetId: 's', address: 'A1', value: 'pear' }])
  })

  it('keeps a literal $ in the replacement and skips formula cells', () => {
    const ops = expandToPrimitiveOps(
      [
        {
          op: 'find_replace',
          sheetId: 's',
          range: 'A1:A2',
          find: 'x',
          replace: '$&',
        },
      ],
      (address) => (address === 'A1' ? { value: 'x marks' } : { value: 'x', formula: '=X1' }),
    )
    expect(ops).toEqual([{ op: 'set_cell', sheetId: 's', address: 'A1', value: '$& marks' }])
  })

  it('requires a reader and rejects oversized ranges', () => {
    expect(() =>
      expandToPrimitiveOps([
        {
          op: 'find_replace',
          sheetId: 's',
          range: 'A1:B2',
          find: 'a',
          replace: 'b',
        },
      ]),
    ).toThrow(/needs the current cell contents/)
    expect(() =>
      expandToPrimitiveOps(
        [
          {
            op: 'find_replace',
            sheetId: 's',
            range: 'A1:Z1000',
            find: 'a',
            replace: 'b',
          },
        ],
        reader({}),
      ),
    ).toThrow(/more than/)
  })
})
