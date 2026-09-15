/**
 * Tests for sheets-skill (W7 deliverable).
 */

import { describe, expect, it, beforeEach, vi } from 'vitest'
import { ReactUIAdapter } from '@genoffice/agent-runtime'
import {
  createSheetsSkillExtension,
  createGetWorkbookContextTool,
  createReadRangeTool,
  createAggregateRangeTool,
  createFindCellsTool,
  createCreateSheetTool,
  type SheetsEditor,
  type SheetsRange,
  type CellValue,
} from '../src/extensions/sheets-skill'

class MockSheetsEditor implements SheetsEditor {
  private sheets: Map<string, CellValue[][]>
  active: string
  createdSheets: string[] = []

  constructor() {
    this.sheets = new Map()
    this.sheets.set('Sales', [
      [{ raw: 'Region' }, { raw: 'Q1' }, { raw: 'Q2' }, { raw: 'Q3' }],
      [{ raw: 'North' }, { raw: 100 }, { raw: 120 }, { raw: 140 }],
      [{ raw: 'South' }, { raw: 80 }, { raw: 90 }, { raw: 95 }],
      [{ raw: 'East' }, { raw: 150 }, { raw: 160 }, { raw: 170 }],
      [{ raw: 'West' }, { raw: 70 }, { raw: 75 }, { raw: 80 }],
    ])
    this.active = 'Sales'
  }

  getWorkbookSummary() {
    let totalCells = 0
    let totalFormulas = 0
    for (const rows of this.sheets.values()) {
      for (const row of rows) {
        for (const c of row) {
          totalCells++
          if (typeof c.raw === 'string' && c.raw.startsWith('=')) totalFormulas++
        }
      }
    }
    return {
      sheetNames: Array.from(this.sheets.keys()),
      activeSheet: this.active,
      totalCells,
      totalFormulas,
    }
  }

  readRange(range: SheetsRange): CellValue[][] {
    const rows = this.sheets.get(range.sheet) ?? []
    return rows.slice(range.startRow, range.endRow + 1).map((row) =>
      row.slice(range.startCol, range.endCol + 1)
    )
  }

  aggregateRange(range: SheetsRange, op: 'sum' | 'avg' | 'count' | 'min' | 'max'): number | null {
    const data = this.readRange(range).flat()
    if (op === 'count') return data.length
    const nums = data.map((c) => Number(c.raw)).filter((n) => Number.isFinite(n))
    if (nums.length === 0) return null
    if (op === 'sum') return nums.reduce((a, b) => a + b, 0)
    if (op === 'avg') return nums.reduce((a, b) => a + b, 0) / nums.length
    if (op === 'min') return Math.min(...nums)
    if (op === 'max') return Math.max(...nums)
    return null
  }

  findCells(sheet: string, query: string, maxResults = 20): string[] {
    const rows = this.sheets.get(sheet) ?? []
    const results: string[] = []
    const q = query.toLowerCase()
    for (let r = 0; r < rows.length && results.length < maxResults; r++) {
      for (let c = 0; c < rows[r]!.length; c++) {
        const raw = String(rows[r]![c]!.raw ?? '')
        if (raw.toLowerCase().includes(q)) {
          results.push(`${sheet}!R${r + 1}C${c + 1}`)
        }
      }
    }
    return results
  }

  getSheetFeatures() {
    return { mergedRanges: [], frozenPanes: { row: 1, col: 1 } }
  }

  createNewDocument(name: string) {
    this.sheets.set(name, [[{ raw: '' }]])
    this.active = name
    this.createdSheets.push(name)
  }
}

describe('sheets-skill: get_workbook_context', () => {
  let adapter: ReactUIAdapter
  let editor: MockSheetsEditor

  beforeEach(() => {
    adapter = new ReactUIAdapter()
    editor = new MockSheetsEditor()
    adapter.setEditorInstance(editor)
  })

  it('returns sheet names, active sheet, cell and formula counts', async () => {
    const tool = createGetWorkbookContextTool({ uiAdapter: adapter })
    const result = await tool.execute('id', {}, undefined, undefined, {} as never)
    const text = (result.content[0] as { type: 'text'; text: string }).text
    expect(text).toContain('1 sheet(s)')
    expect(text).toContain('[Sales]')
    expect(text).toContain('active="Sales"')
    expect(text).toContain('20 cells')
    expect(result.details.sheetCount).toBe(1)
  })
})

describe('sheets-skill: read_range', () => {
  let adapter: ReactUIAdapter
  let editor: MockSheetsEditor

  beforeEach(() => {
    adapter = new ReactUIAdapter()
    editor = new MockSheetsEditor()
    adapter.setEditorInstance(editor)
  })

  it('returns TSV for a small range', async () => {
    const tool = createReadRangeTool({ uiAdapter: adapter })
    const result = await tool.execute('id', { sheet: 'Sales', startRow: 0, startCol: 0, endRow: 1, endCol: 3 }, undefined, undefined, {} as never)
    const text = (result.content[0] as { type: 'text'; text: string }).text
    expect(text).toContain('Region\tQ1\tQ2\tQ3')
    expect(text).toContain('North\t100\t120\t140')
  })

  it('rejects oversized ranges', async () => {
    const tool = createReadRangeTool({ uiAdapter: adapter })
    const result = await tool.execute('id', { sheet: 'Sales', startRow: 0, startCol: 0, endRow: 100, endCol: 100 }, undefined, undefined, {} as never)
    expect((result.content[0] as { type: 'text'; text: string }).text).toContain('too large')
    expect(result.details.truncated).toBe(true)
  })
})

describe('sheets-skill: aggregate_range', () => {
  let adapter: ReactUIAdapter
  let editor: MockSheetsEditor

  beforeEach(() => {
    adapter = new ReactUIAdapter()
    editor = new MockSheetsEditor()
    adapter.setEditorInstance(editor)
  })

  it.each([
    // 12 numeric cells: 100+120+140+80+90+95+150+160+170+70+75+80 = 1330
    ['sum', 1330],
    ['count', 12],
    ['avg', 1330 / 12],
    ['min', 70],
    ['max', 170],
  ])('computes %s correctly', async (op, expected) => {
    const tool = createAggregateRangeTool({ uiAdapter: adapter })
    const result = await tool.execute('id', {
      sheet: 'Sales', startRow: 1, startCol: 1, endRow: 4, endCol: 3,
      op: op as 'sum',
    }, undefined, undefined, {} as never)
    expect((result.details.result as number)).toBeCloseTo(expected as number, 1)
  })
})

describe('sheets-skill: find_cells', () => {
  let adapter: ReactUIAdapter
  let editor: MockSheetsEditor

  beforeEach(() => {
    adapter = new ReactUIAdapter()
    editor = new MockSheetsEditor()
    adapter.setEditorInstance(editor)
  })

  it('returns A1-style refs for matches', async () => {
    const tool = createFindCellsTool({ uiAdapter: adapter })
    const result = await tool.execute('id', { sheet: 'Sales', query: 'north' }, undefined, undefined, {} as never)
    const text = (result.content[0] as { type: 'text'; text: string }).text
    expect(text).toContain('Sales!R2C1')
    expect(result.details.count).toBe(1)
  })

  it('reports no matches', async () => {
    const tool = createFindCellsTool({ uiAdapter: adapter })
    const result = await tool.execute('id', { sheet: 'Sales', query: 'zzz' }, undefined, undefined, {} as never)
    expect((result.content[0] as { type: 'text'; text: string }).text).toContain('No matches')
  })
})

describe('sheets-skill: create_document', () => {
  it('delegates to host createNewDocument', async () => {
    const adapter = new ReactUIAdapter()
    const editor = new MockSheetsEditor()
    adapter.setEditorInstance(editor)
    const tool = createCreateSheetTool({ uiAdapter: adapter })
    const result = await tool.execute('id', { sheetName: 'NewSheet' }, undefined, undefined, {} as never)
    expect(editor.createdSheets).toContain('NewSheet')
    expect(result.details.created).toBe(true)
  })

  it('returns helpful error when host missing createNewDocument', async () => {
    const adapter = new ReactUIAdapter()
    const editor: SheetsEditor = {
      getWorkbookSummary: () => ({ sheetNames: [], activeSheet: '', totalCells: 0, totalFormulas: 0 }),
      readRange: () => [],
      aggregateRange: () => null,
      findCells: () => [],
      getSheetFeatures: () => ({ mergedRanges: [], frozenPanes: null }),
    }
    adapter.setEditorInstance(editor)
    const tool = createCreateSheetTool({ uiAdapter: adapter })
    const result = await tool.execute('id', {}, undefined, undefined, {} as never)
    expect((result.content[0] as { type: 'text'; text: string }).text).toContain('createNewDocument')
  })
})

describe('sheets-skill extension factory', () => {
  it('registers all 5 sheets tools by default', () => {
    const adapter = new ReactUIAdapter()
    adapter.setEditorInstance(new MockSheetsEditor())
    const registered: string[] = []
    const pi = { registerTool: (t: { name: string }) => registered.push(t.name), on: () => {} } as never
    createSheetsSkillExtension({ uiAdapter: adapter })(pi)
    expect(registered).toEqual([
      'get_workbook_context', 'read_range', 'aggregate_range', 'find_cells', 'create_document',
    ])
  })

  it('enabledTools restricts the registered set', () => {
    const adapter = new ReactUIAdapter()
    adapter.setEditorInstance(new MockSheetsEditor())
    const registered: string[] = []
    const pi = { registerTool: (t: { name: string }) => registered.push(t.name), on: () => {} } as never
    createSheetsSkillExtension({ uiAdapter: adapter, enabledTools: ['read_range'] })(pi)
    expect(registered).toEqual(['read_range'])
  })

  it('loads into a real pi session', async () => {
    const { createOfficeSession } = await import('@genoffice/agent-runtime')
    const adapter = new ReactUIAdapter()
    adapter.setEditorInstance(new MockSheetsEditor())
    const { session, dispose } = await createOfficeSession({
      cwd: process.cwd(),
      extensionFactories: [createSheetsSkillExtension({ uiAdapter: adapter })],
    })
    const names = session.getAllTools().map((t) => t.name)
    expect(names).toContain('read_range')
    expect(names).toContain('aggregate_range')
    expect(names).toContain('find_cells')
    dispose()
  }, 30_000)
})
