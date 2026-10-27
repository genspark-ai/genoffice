/**
 * The eight "show values as" modes that reference a base field — Difference
 * From, % Difference From, % Of, % of Parent Total, Running Total In, %
 * Running Total In, Rank ascending/descending — across the bake path
 * (buildPivotLayout), the refresh path (recomputePivotData), and the OOXML
 * writer/parser pair. Excel semantics (cross-checked against LibreOffice's
 * pivot reference-value implementation): the base item's own cells are blank
 * for difference/percentDiff ("100%" for % Of), totals spanning the base
 * field are blank, divide-by-zero is blank (this engine's #DIV/0!
 * convention), running totals restart per outer member / other-axis group,
 * ranks share the smaller rank on ties (RANK.EQ).
 */
import { describe, expect, it } from 'vitest'

import { recomputePivotData } from '../src/domain/pivot-engine'
import { buildPivotLayout } from '../src/domain/pivot-layout'
import { parsePivotDefinition } from '../src/gateway/xlsx-pivot'
import { applyPivotAdditions, type PivotAddition } from '../src/gateway/xlsx-pivot-add'
import type { MutablePackage } from '../src/gateway/xlsx-drawing-add'

type Cell = string | number

const GRID: readonly (readonly Cell[])[] = [
  ['Region', 'Product', 'Amount'],
  ['East', 'A', 10],
  ['East', 'B', 20],
  ['West', 'A', 30],
  ['West', 'B', 40],
  ['East', 'A', 5],
  ['West', 'B', 1],
]
// Region × Product: East [A=15, B=20, total 35], West [A=30, B=41, total 71],
// grand 106; product totals A=45, B=61; rows-only East 35, West 71.

// ─── Bake: Difference From / % Difference From / % Of ───────────────────────

describe('difference family (bake)', () => {
  it('difference from a base item of the column field', () => {
    const layout = buildPivotLayout(GRID, {
      rowFields: ['Region'],
      columnField: 'Product',
      values: [
        {
          field: 'Amount',
          agg: 'sum',
          showDataAs: 'difference',
          baseField: 'Product',
          baseItem: 'A',
        },
      ],
    })
    // matrix rows: [label, A, B, Total]; the base item column is blank
    expect(layout.matrix[1]).toEqual(['East', null, 5, null])
    expect(layout.matrix[2]).toEqual(['West', null, 11, null])
    // The grand-total row fixes the column (base) coordinate, so its non-base
    // columns keep the calculation (61 − 45).
    expect(layout.matrix[3]).toEqual(['Grand Total', null, 16, null])
  })

  it('% Of divides by the base item and shows 100% on the base item itself', () => {
    const layout = buildPivotLayout(GRID, {
      rowFields: ['Region'],
      columnField: 'Product',
      values: [
        {
          field: 'Amount',
          agg: 'sum',
          showDataAs: 'percentOf',
          baseField: 'Product',
          baseItem: 'A',
        },
      ],
    })
    expect(layout.matrix[1]).toEqual(['East', 1, 20 / 15, null])
    expect(layout.matrix[2]).toEqual(['West', 1, 41 / 30, null])
  })

  it('% difference blanks the base item cell and totals', () => {
    const layout = buildPivotLayout(GRID, {
      rowFields: ['Region'],
      columnField: 'Product',
      values: [
        {
          field: 'Amount',
          agg: 'sum',
          showDataAs: 'percentDiff',
          baseField: 'Product',
          baseItem: 'A',
        },
      ],
    })
    expect(layout.matrix[1]).toEqual(['East', null, 5 / 15, null])
    expect(layout.matrix[2]).toEqual(['West', null, 11 / 30, null])
  })

  it('base field on the row axis blanks the base item rows, keeps the column coordinate', () => {
    const layout = buildPivotLayout(GRID, {
      rowFields: ['Region'],
      columnField: 'Product',
      values: [
        {
          field: 'Amount',
          agg: 'sum',
          showDataAs: 'difference',
          baseField: 'Region',
          baseItem: 'East',
        },
      ],
    })
    expect(layout.matrix[1]).toEqual(['East', null, null, null])
    expect(layout.matrix[2]).toEqual(['West', 15, 21, 36])
    expect(layout.matrix[3]).toEqual(['Grand Total', null, null, null])
  })

  it('multi-level rows blank the base item member and its subtotal', () => {
    const layout = buildPivotLayout(GRID, {
      rowFields: ['Region', 'Product'],
      values: [
        {
          field: 'Amount',
          agg: 'sum',
          showDataAs: 'difference',
          baseField: 'Region',
          baseItem: 'East',
        },
      ],
    })
    // rows: East/A, East/B, East Subtotal, West/A, West/B, West Subtotal, Grand
    expect(layout.matrix.slice(1, 8).map((row) => row[2])).toEqual([
      null,
      null,
      null,
      30 - 15,
      41 - 20,
      71 - 35,
      null,
    ])
  })

  it('multi-level rows with an inner base field compare at the same inner level', () => {
    const layout = buildPivotLayout(GRID, {
      rowFields: ['Region', 'Product'],
      values: [
        {
          field: 'Amount',
          agg: 'sum',
          showDataAs: 'difference',
          baseField: 'Product',
          baseItem: 'A',
        },
      ],
    })
    expect(layout.matrix.slice(1, 8).map((row) => row[2])).toEqual([
      null, // the base item's own cell
      5,
      null, // subtotal spans the base level
      null,
      11,
      null,
      null,
    ])
  })

  it('a zero base item divides to blank (Excel #DIV/0!)', () => {
    const grid: readonly (readonly Cell[])[] = [
      ['Region', 'Product', 'Amount'],
      ['East', 'Z', 5],
      ['East', 'Z', -5],
      ['East', 'A', 15],
      ['East', 'B', 7],
    ]
    const layout = buildPivotLayout(grid, {
      rowFields: ['Region'],
      columnField: 'Product',
      values: [
        {
          field: 'Amount',
          agg: 'sum',
          showDataAs: 'percentOf',
          baseField: 'Product',
          baseItem: 'Z',
        },
      ],
    })
    expect(layout.matrix[1]).toEqual(['East', null, null, null, null])
  })
})

// ─── Bake: % of Parent Total ────────────────────────────────────────────────

describe('% of Parent Total (bake)', () => {
  it('divides by the base field parent level subtotal, 100% on its own subtotals', () => {
    const layout = buildPivotLayout(GRID, {
      rowFields: ['Region', 'Product'],
      values: [
        { field: 'Amount', agg: 'sum', showDataAs: 'percentOfParent', baseField: 'Product' },
      ],
    })
    expect(layout.matrix.slice(1, 8).map((row) => row[2])).toEqual([
      15 / 35,
      20 / 35,
      1,
      30 / 71,
      41 / 71,
      1,
      1,
    ])
  })

  it('an outermost base field divides by the grand total', () => {
    const layout = buildPivotLayout(GRID, {
      rowFields: ['Region', 'Product'],
      values: [{ field: 'Amount', agg: 'sum', showDataAs: 'percentOfParent', baseField: 'Region' }],
    })
    expect(layout.matrix.slice(1, 8).map((row) => row[2])).toEqual([
      15 / 106,
      20 / 106,
      35 / 106,
      30 / 106,
      41 / 106,
      71 / 106,
      1,
    ])
  })

  it('keeps the other axis fixed: percent of the column within the row-parent', () => {
    const layout = buildPivotLayout(GRID, {
      rowFields: ['Region'],
      columnField: 'Product',
      values: [{ field: 'Amount', agg: 'sum', showDataAs: 'percentOfParent', baseField: 'Region' }],
    })
    expect(layout.matrix[1]).toEqual(['East', 15 / 45, 20 / 61, 35 / 106])
    expect(layout.matrix[2]).toEqual(['West', 30 / 45, 41 / 61, 71 / 106])
  })
})

// ─── Bake: Running Total In / % Running Total In ────────────────────────────

describe('running totals (bake)', () => {
  it('accumulates in base-field item order and blanks spanning totals', () => {
    const layout = buildPivotLayout(GRID, {
      rowFields: ['Region'],
      values: [{ field: 'Amount', agg: 'sum', showDataAs: 'runTotal', baseField: 'Region' }],
    })
    expect(layout.matrix.slice(1).map((row) => row[1])).toEqual([35, 106, null])
  })

  it('restarts for each outer member of the base axis', () => {
    const layout = buildPivotLayout(GRID, {
      rowFields: ['Region', 'Product'],
      values: [{ field: 'Amount', agg: 'sum', showDataAs: 'runTotal', baseField: 'Product' }],
    })
    // East A 15 → East B 35; restarts at West: A 30 → B 71; subtotals/grand blank.
    expect(layout.matrix.slice(1, 8).map((row) => row[2])).toEqual([
      15,
      35,
      null,
      30,
      71,
      null,
      null,
    ])
  })

  it('restarts for each other-axis group (accumulation runs along columns)', () => {
    const layout = buildPivotLayout(GRID, {
      rowFields: ['Region'],
      columnField: 'Product',
      values: [{ field: 'Amount', agg: 'sum', showDataAs: 'runTotal', baseField: 'Product' }],
    })
    expect(layout.matrix[1]).toEqual(['East', 15, 35, null])
    expect(layout.matrix[2]).toEqual(['West', 30, 71, null])
    // The grand-total row fixes the column (base) coordinate, so it keeps
    // accumulating down the rows of each column.
    expect(layout.matrix[3]).toEqual(['Grand Total', 45, 106, null])
  })

  it('% running total divides by the group total', () => {
    const layout = buildPivotLayout(GRID, {
      rowFields: ['Region'],
      columnField: 'Product',
      values: [
        { field: 'Amount', agg: 'sum', showDataAs: 'percentOfRunningTotal', baseField: 'Product' },
      ],
    })
    expect(layout.matrix[1]).toEqual(['East', 15 / 35, 1, null])
    expect(layout.matrix[2]).toEqual(['West', 30 / 71, 1, null])
  })
})

// ─── Bake: Rank ─────────────────────────────────────────────────────────────

describe('rank (bake)', () => {
  const TIE: readonly (readonly Cell[])[] = [
    ['Region', 'Product', 'Amount'],
    ['East', 'A', 10],
    ['East', 'B', 10],
    ['East', 'C', 20],
    ['West', 'A', 30],
    ['West', 'B', 5],
    ['West', 'C', 5],
  ]

  it('ascending ranks with ties sharing the smaller rank (RANK.EQ)', () => {
    const layout = buildPivotLayout(TIE, {
      rowFields: ['Region'],
      columnField: 'Product',
      values: [{ field: 'Amount', agg: 'sum', showDataAs: 'rankAscending', baseField: 'Product' }],
    })
    expect(layout.matrix[1]).toEqual(['East', 1, 1, 3, null])
    expect(layout.matrix[2]).toEqual(['West', 3, 1, 1, null])
    // Grand-total row: column totals 40/15/25 rank 3/1/2 ascending.
    expect(layout.matrix[3]).toEqual(['Grand Total', 3, 1, 2, null])
  })

  it('descending ranks put the largest first, ties sharing the rank', () => {
    const layout = buildPivotLayout(TIE, {
      rowFields: ['Region'],
      columnField: 'Product',
      values: [{ field: 'Amount', agg: 'sum', showDataAs: 'rankDescending', baseField: 'Product' }],
    })
    expect(layout.matrix[1]).toEqual(['East', 2, 2, 1, null])
    expect(layout.matrix[2]).toEqual(['West', 1, 2, 2, null])
  })

  it('ranks can run along the row axis per column group', () => {
    const layout = buildPivotLayout(GRID, {
      rowFields: ['Region'],
      columnField: 'Product',
      values: [{ field: 'Amount', agg: 'sum', showDataAs: 'rankDescending', baseField: 'Region' }],
    })
    // The grand-total column fixes the row (base) coordinate, so it ranks
    // the row totals too.
    expect(layout.matrix[1]).toEqual(['East', 2, 2, 2])
    expect(layout.matrix[2]).toEqual(['West', 1, 1, 1])
  })
})

// ─── Validation ─────────────────────────────────────────────────────────────

describe('base field/item validation (bake)', () => {
  it('rejects a base field outside the axes', () => {
    expect(() =>
      buildPivotLayout(GRID, {
        rowFields: ['Region'],
        values: [{ field: 'Amount', agg: 'sum', showDataAs: 'runTotal', baseField: 'Product' }],
      }),
    ).toThrow(/must sit on the pivot rows or columns/)
  })

  it('rejects a base item outside the base field members', () => {
    expect(() =>
      buildPivotLayout(GRID, {
        rowFields: ['Region'],
        values: [
          {
            field: 'Amount',
            agg: 'sum',
            showDataAs: 'difference',
            baseField: 'Region',
            baseItem: 'North',
          },
        ],
      }),
    ).toThrow(/Base item "North" is not a member/)
  })

  it('rejects a base-field mode without a base field', () => {
    expect(() =>
      buildPivotLayout(GRID, {
        rowFields: ['Region'],
        values: [{ field: 'Amount', agg: 'sum', showDataAs: 'runTotal' }],
      }),
    ).toThrow(/needs a base field/)
  })
})

// ─── OOXML write + parse round trip ─────────────────────────────────────────

const WORKBOOK = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`

const WORKBOOK_RELS = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`

function inMemoryPackage(entries: Record<string, string>): MutablePackage {
  const files = new Map(Object.entries(entries))
  return {
    paths: async () => [...files.keys()],
    has: async (path) => files.has(path),
    readText: async (path) => {
      const content = files.get(path)
      if (content === undefined) throw new Error(`missing part: ${path}`)
      return content
    },
    write: (path, content) => void files.set(path, content),
    add: (path, content) => void files.set(path, content),
    addBinary: (path, bytes) => void files.set(path, new TextDecoder().decode(bytes)),
    remove: (path) => void files.delete(path),
  }
}

function worksheetOf(grid: readonly (readonly Cell[])[]): string {
  const rows = grid
    .map((row, r) => {
      const cells = row
        .map((value, c) =>
          typeof value === 'number'
            ? `<c r="${String.fromCharCode(65 + c)}${r + 1}"><v>${value}</v></c>`
            : `<c r="${String.fromCharCode(65 + c)}${r + 1}" t="inlineStr"><is><t>${value}</t></is></c>`,
        )
        .join('')
      return `<row r="${r + 1}">${cells}</row>`
    })
    .join('')
  return `<?xml version="1.0" encoding="UTF-8"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows}</sheetData></worksheet>`
}

type BakeOp = Parameters<typeof buildPivotLayout>[1]

async function writePivotPackage(
  grid: readonly (readonly Cell[])[],
  op: BakeOp,
): Promise<{ table: string; cache: string }> {
  const layout = buildPivotLayout(grid, op)
  const addition: PivotAddition = {
    worksheetPath: 'xl/worksheets/sheet2.xml',
    sourceSheetName: 'Data',
    sourceArea: {
      startRow: 0,
      startColumn: 0,
      endRow: grid.length - 1,
      endColumn: grid[0]!.length - 1,
    },
    location: { startRow: 0, startColumn: 5, endRow: 4, endColumn: 7 },
    name: 'Pivot1',
    ...layout.definition,
  } as PivotAddition
  const pkg = inMemoryPackage({
    '[Content_Types].xml': CONTENT_TYPES,
    'xl/workbook.xml': WORKBOOK,
    'xl/_rels/workbook.xml.rels': WORKBOOK_RELS,
    'xl/worksheets/sheet1.xml': worksheetOf(grid),
  })
  await applyPivotAdditions(pkg, [addition], WORKBOOK, new Set())
  return {
    table: await pkg.readText('xl/pivotTables/pivotTable1.xml'),
    cache: await pkg.readText('xl/pivotCache/pivotCacheDefinition1.xml'),
  }
}

describe('OOXML round trip', () => {
  it('writes difference with baseField/baseItem and parses it back', async () => {
    const { table, cache } = await writePivotPackage(GRID, {
      rowFields: ['Region'],
      values: [
        {
          field: 'Amount',
          agg: 'sum',
          showDataAs: 'difference',
          baseField: 'Region',
          baseItem: 'West',
        },
      ],
    })
    expect(table).toContain('showDataAs="difference" baseField="0" baseItem="1"')
    const definition = parsePivotDefinition(table, cache)
    expect(definition.unsupported).toEqual([])
    expect(definition.dataFields[0]).toMatchObject({
      showDataAs: 'difference',
      baseField: 0,
      baseItem: 1,
    })
  })

  it('spells % Of as "percent" in the file and keeps General for difference', async () => {
    const percent = await writePivotPackage(GRID, {
      rowFields: ['Region'],
      values: [
        {
          field: 'Amount',
          agg: 'sum',
          showDataAs: 'percentOf',
          baseField: 'Region',
          baseItem: 'West',
        },
      ],
    })
    expect(percent.table).toContain('showDataAs="percent" baseField="0" baseItem="1"')
    // percent modes default to the 0.00% numFmt (id 10); difference stays General.
    expect(percent.table).toMatch(/numFmtId="10"/)
    const definition = parsePivotDefinition(percent.table, percent.cache)
    expect(definition.dataFields[0]).toMatchObject({
      showDataAs: 'percentOf',
      baseField: 0,
      baseItem: 1,
    })
    const difference = await writePivotPackage(GRID, {
      rowFields: ['Region'],
      values: [
        {
          field: 'Amount',
          agg: 'sum',
          showDataAs: 'difference',
          baseField: 'Region',
          baseItem: 'West',
        },
      ],
    })
    expect(difference.table).not.toMatch(/numFmtId="10"/)
  })

  it('writes runTotal with baseField only (baseItem stays 0)', async () => {
    const { table, cache } = await writePivotPackage(GRID, {
      rowFields: ['Region'],
      values: [{ field: 'Amount', agg: 'sum', showDataAs: 'runTotal', baseField: 'Region' }],
    })
    expect(table).toContain('showDataAs="runTotal" baseField="0" baseItem="0"')
    expect(parsePivotDefinition(table, cache).dataFields[0]).toMatchObject({
      showDataAs: 'runTotal',
      baseField: 0,
    })
  })

  it('fails closed on (previous)/(next) base items and off-axis base fields', async () => {
    const { table, cache } = await writePivotPackage(GRID, {
      rowFields: ['Region'],
      values: [
        {
          field: 'Amount',
          agg: 'sum',
          showDataAs: 'difference',
          baseField: 'Region',
          baseItem: 'West',
        },
      ],
    })
    const previous = parsePivotDefinition(
      table.replace('baseItem="1"', 'baseItem="1048828"'),
      cache,
    )
    expect(previous.unsupported.join()).toMatch(/previous/)
    const next = parsePivotDefinition(table.replace('baseItem="1"', 'baseItem="1048829"'), cache)
    expect(next.unsupported.join()).toMatch(/next/)
    const offAxis = parsePivotDefinition(table.replace('baseField="0"', 'baseField="2"'), cache)
    expect(offAxis.unsupported.join()).toMatch(/base field/)
    const offRange = parsePivotDefinition(table.replace('baseField="0"', 'baseField="9"'), cache)
    expect(offRange.unsupported.join()).toMatch(/base field/)
  })
})

// ─── Bake vs recompute: the two switches stay in sync ───────────────────────

describe('recompute matches the baked grid', () => {
  const cases: readonly { label: string; op: BakeOp; rowLevels: number }[] = [
    {
      label: 'difference (row base, East)',
      op: {
        rowFields: ['Region'],
        columnField: 'Product',
        values: [
          {
            field: 'Amount',
            agg: 'sum',
            showDataAs: 'difference',
            baseField: 'Region',
            baseItem: 'East',
          },
        ],
      },
      rowLevels: 1,
    },
    {
      label: 'difference (column base, A)',
      op: {
        rowFields: ['Region'],
        columnField: 'Product',
        values: [
          {
            field: 'Amount',
            agg: 'sum',
            showDataAs: 'difference',
            baseField: 'Product',
            baseItem: 'A',
          },
        ],
      },
      rowLevels: 1,
    },
    {
      label: 'percentDiff (column base, A)',
      op: {
        rowFields: ['Region'],
        columnField: 'Product',
        values: [
          {
            field: 'Amount',
            agg: 'sum',
            showDataAs: 'percentDiff',
            baseField: 'Product',
            baseItem: 'A',
          },
        ],
      },
      rowLevels: 1,
    },
    {
      label: 'percentOf (column base, B)',
      op: {
        rowFields: ['Region'],
        columnField: 'Product',
        values: [
          {
            field: 'Amount',
            agg: 'sum',
            showDataAs: 'percentOf',
            baseField: 'Product',
            baseItem: 'B',
          },
        ],
      },
      rowLevels: 1,
    },
    {
      label: 'percentOfParent (row base)',
      op: {
        rowFields: ['Region'],
        columnField: 'Product',
        values: [
          { field: 'Amount', agg: 'sum', showDataAs: 'percentOfParent', baseField: 'Region' },
        ],
      },
      rowLevels: 1,
    },
    {
      label: 'runTotal (row base)',
      op: {
        rowFields: ['Region'],
        columnField: 'Product',
        values: [{ field: 'Amount', agg: 'sum', showDataAs: 'runTotal', baseField: 'Region' }],
      },
      rowLevels: 1,
    },
    {
      label: 'runTotal (column base)',
      op: {
        rowFields: ['Region'],
        columnField: 'Product',
        values: [{ field: 'Amount', agg: 'sum', showDataAs: 'runTotal', baseField: 'Product' }],
      },
      rowLevels: 1,
    },
    {
      label: 'percentOfRunningTotal (column base)',
      op: {
        rowFields: ['Region'],
        columnField: 'Product',
        values: [
          {
            field: 'Amount',
            agg: 'sum',
            showDataAs: 'percentOfRunningTotal',
            baseField: 'Product',
          },
        ],
      },
      rowLevels: 1,
    },
    {
      label: 'rankAscending (column base)',
      op: {
        rowFields: ['Region'],
        columnField: 'Product',
        values: [
          { field: 'Amount', agg: 'sum', showDataAs: 'rankAscending', baseField: 'Product' },
        ],
      },
      rowLevels: 1,
    },
    {
      label: 'rankDescending (row base)',
      op: {
        rowFields: ['Region'],
        columnField: 'Product',
        values: [
          { field: 'Amount', agg: 'sum', showDataAs: 'rankDescending', baseField: 'Region' },
        ],
      },
      rowLevels: 1,
    },
    {
      label: 'runTotal resets per outer member (multi-level rows)',
      op: {
        rowFields: ['Region', 'Product'],
        values: [{ field: 'Amount', agg: 'sum', showDataAs: 'runTotal', baseField: 'Product' }],
      },
      rowLevels: 2,
    },
    {
      label: 'percentOfParent at inner level (multi-level rows)',
      op: {
        rowFields: ['Region', 'Product'],
        values: [
          { field: 'Amount', agg: 'sum', showDataAs: 'percentOfParent', baseField: 'Product' },
        ],
      },
      rowLevels: 2,
    },
    {
      label: 'difference at outer level (multi-level rows)',
      op: {
        rowFields: ['Region', 'Product'],
        values: [
          {
            field: 'Amount',
            agg: 'sum',
            showDataAs: 'difference',
            baseField: 'Product',
            baseItem: 'A',
          },
        ],
      },
      rowLevels: 2,
    },
  ]

  for (const { label, op, rowLevels } of cases) {
    it(`${label} recomputes identically`, async () => {
      const layout = buildPivotLayout(GRID, op)
      const { table, cache } = await writePivotPackage(GRID, op)
      const definition = parsePivotDefinition(table, cache)
      expect(definition.unsupported).toEqual([])
      const { data } = recomputePivotData(definition, GRID)
      // The baked matrix drops its header row(s) and the row-label columns.
      const columnLevels = Array.isArray(op.columnField)
        ? op.columnField.length
        : op.columnField
          ? 1
          : 0
      const headerRows = Math.max(1, columnLevels)
      const expected = layout.matrix.slice(headerRows).map((row) => row.slice(rowLevels))
      expect(data).toEqual(expected)
    })
  }

  it('rank ties share the smaller rank through the whole write/recompute cycle', async () => {
    // East products {A:10, B:10, C:20} tie — RANK.EQ gives 1/1/3 through the
    // engine, not a strict 1/2/3 ordering.
    const op: Parameters<typeof buildPivotLayout>[1] = {
      rowFields: ['Region'],
      columnField: 'Product',
      values: [
        {
          field: 'Amount',
          agg: 'sum' as const,
          showDataAs: 'rankAscending' as const,
          baseField: 'Product',
        },
      ],
    }
    const TIE: readonly (readonly Cell[])[] = [
      ['Region', 'Product', 'Amount'],
      ['East', 'A', 10],
      ['East', 'B', 10],
      ['East', 'C', 20],
    ]
    const { table, cache } = await writePivotPackage(TIE, op)
    const { data } = recomputePivotData(parsePivotDefinition(table, cache), TIE)
    // The grand-total row fixes the column (base) coordinate, so the column
    // totals (also 10/10/20) repeat the tie ranks.
    expect(data).toEqual([
      [1, 1, 3, null],
      [1, 1, 3, null],
    ])
  })
})

// ─── Refresh engine directly (hand-written XML) ─────────────────────────────

describe('recompute engine (parsed XML)', () => {
  const PIVOT_XML = (mode: string, baseField = 0, baseItem = 0): string =>
    '<pivotTableDefinition name="PivotTable1" cacheId="1">' +
    '<location ref="E3:H7" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>' +
    '<pivotFields count="3">' +
    '<pivotField axis="axisRow" showAll="0"><items count="3">' +
    '<item x="0"/><item x="1"/><item t="default"/></items></pivotField>' +
    '<pivotField axis="axisCol" showAll="0"><items count="3">' +
    '<item x="0"/><item x="1"/><item t="default"/></items></pivotField>' +
    '<pivotField dataField="1" showAll="0"/>' +
    '</pivotFields>' +
    '<rowFields count="1"><field x="0"/></rowFields>' +
    '<rowItems count="3"><i><x/></i><i><x v="1"/></i><i t="grand"><x/></i></rowItems>' +
    '<colFields count="1"><field x="1"/></colFields>' +
    '<colItems count="3"><i><x/></i><i><x v="1"/></i><i t="grand"><x/></i></colItems>' +
    `<dataFields count="1"><dataField name="Sum of Amount" fld="2" showDataAs="${mode}" baseField="${baseField}" baseItem="${baseItem}"/></dataFields>` +
    '</pivotTableDefinition>'

  const CACHE_XML =
    '<pivotCacheDefinition>' +
    '<cacheSource type="worksheet"><worksheetSource ref="A1:C7" sheet="Data"/></cacheSource>' +
    '<cacheFields count="3">' +
    '<cacheField name="Region"><sharedItems count="2"><s v="East"/><s v="West"/></sharedItems></cacheField>' +
    '<cacheField name="Product"><sharedItems count="2"><s v="A"/><s v="B"/></sharedItems></cacheField>' +
    '<cacheField name="Amount"><sharedItems containsString="0" containsNumber="1"/></cacheField>' +
    '</cacheFields>' +
    '</pivotCacheDefinition>'

  it('difference from a row-axis base item blanks its own rows', () => {
    const { data } = recomputePivotData(
      parsePivotDefinition(PIVOT_XML('difference', 0, 0), CACHE_XML),
      GRID,
    )
    expect(data).toEqual([
      [null, null, null],
      [15, 21, 36],
      [null, null, null],
    ])
  })

  it('difference from a column-axis base item blanks the base column', () => {
    const { data } = recomputePivotData(
      parsePivotDefinition(PIVOT_XML('difference', 1, 0), CACHE_XML),
      GRID,
    )
    expect(data).toEqual([
      [null, 5, null],
      [null, 11, null],
      [null, 16, null],
    ])
  })

  it('running total restarts per column group (base field on rows)', () => {
    const { data } = recomputePivotData(
      parsePivotDefinition(PIVOT_XML('runTotal', 0), CACHE_XML),
      GRID,
    )
    // Accumulation restarts per column; the Total column still accumulates the
    // row totals (it fixes the row, i.e. base, coordinate).
    expect(data).toEqual([
      [15, 20, 35],
      [45, 61, 106],
      [null, null, null],
    ])
  })

  it('rank among the row members per column', () => {
    const { data } = recomputePivotData(
      parsePivotDefinition(PIVOT_XML('rankAscending', 0), CACHE_XML),
      GRID,
    )
    expect(data).toEqual([
      [1, 1, 1],
      [2, 2, 2],
      [null, null, null],
    ])
  })

  it('"percent" parses to the percentOf mode', () => {
    const definition = parsePivotDefinition(PIVOT_XML('percent', 1, 1), CACHE_XML)
    expect(definition.unsupported).toEqual([])
    expect(definition.dataFields[0]).toMatchObject({
      showDataAs: 'percentOf',
      baseField: 1,
      baseItem: 1,
    })
  })
})
