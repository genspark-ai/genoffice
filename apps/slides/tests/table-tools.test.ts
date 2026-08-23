/**
 * AI table tools (add_table / edit_table_cell / edit_table_structure):
 * tool-level input validation + IPC call orchestration (window.slidesApi stubbed; engine paths covered by pptx-engine tests).
 */
import { describe, it, expect, vi, beforeEach } from 'vitest'
import { createSlidesSkill, type DeckAccess } from '../src/renderer/ai/slides-skill'
import type { RenderSlide } from '@genoffice/pptx-render'
import type { AgentToolCall } from '../src/shared/ipc'

const SLIDE = {
  widthPx: 1280,
  heightPx: 720,
  nodes: [],
  background: null,
} as unknown as RenderSlide

function makeAccess() {
  const applied: Array<{ idx: number }> = []
  const access: DeckAccess = {
    getSlides: () => [SLIDE],
    getCurrent: () => 0,
    getSelectedIds: () => [],
    applySlide: (idx) => {
      applied.push({ idx })
    },
    applyDeck: () => {},
    fitWidthPx: 1280,
    landGeneratedPages: async () => ({ ok: true, pages: 1 }),
  }
  return { access, applied }
}

const call = (name: string, input: Record<string, unknown>): AgentToolCall => ({
  id: 't',
  name,
  input,
})

beforeEach(() => {
  ;(window as any).slidesApi = {
    addTable: vi.fn(async (_op: any) => ({ slide: SLIDE, sourceId: 'tbl_1' })),
    editTableCell: vi.fn(async () => SLIDE),
    tableStructure: vi.fn(async () => ({ slide: SLIDE, sourceId: 'tbl_1' })),
  }
})

describe('add_table', () => {
  it('inserts table and fills cells row by row', async () => {
    const { access, applied } = makeAccess()
    const skill = createSlidesSkill(access)
    const r = await skill.executeTool!(
      call('add_table', {
        slideIndex: 0,
        rows: 2,
        cols: 2,
        cells: [['Name', 'Qty'], ['Apple']],
      }),
    )
    expect(r.isError).toBeUndefined()
    expect(r.mutated).toBe(true)
    const api = (window as any).slidesApi
    expect(api.addTable).toHaveBeenCalledOnce()
    expect(api.addTable.mock.calls[0][0]).toMatchObject({
      slideIndex: 0,
      rows: 2,
      cols: 2,
      fitWidthPx: 1280,
    })
    // 3 non-empty cells -> 3 editTableCell calls with correct coordinates
    expect(api.editTableCell).toHaveBeenCalledTimes(3)
    expect(api.editTableCell.mock.calls.map((c: any[]) => [c[0].row, c[0].col])).toEqual([
      [0, 0],
      [0, 1],
      [1, 0],
    ])
    expect(api.editTableCell.mock.calls[0][0].paragraphs).toEqual([{ runs: [{ text: 'Name' }] }])
    expect(applied.length).toBe(1)
    expect(r.output).toContain('tbl_1')
  })

  it('rejects invalid dimensions and out-of-range slideIndex', async () => {
    const { access } = makeAccess()
    const skill = createSlidesSkill(access)
    expect(
      (await skill.executeTool!(call('add_table', { slideIndex: 0, rows: 0, cols: 3 }))).isError,
    ).toBe(true)
    expect(
      (await skill.executeTool!(call('add_table', { slideIndex: 0, rows: 3, cols: 99 }))).isError,
    ).toBe(true)
    expect(
      (await skill.executeTool!(call('add_table', { slideIndex: 5, rows: 2, cols: 2 }))).isError,
    ).toBe(true)
    expect((window as any).slidesApi.addTable).not.toHaveBeenCalled()
  })
})

describe('edit_table_cell', () => {
  it('replaces one cell text via IPC', async () => {
    const { access, applied } = makeAccess()
    const skill = createSlidesSkill(access)
    const r = await skill.executeTool!(
      call('edit_table_cell', {
        slideIndex: 0,
        sourceId: 'tbl_1',
        row: 1,
        col: 0,
        paragraphs: [{ text: 'Banana', bold: true }],
      }),
    )
    expect(r.isError).toBeUndefined()
    const op = (window as any).slidesApi.editTableCell.mock.calls[0][0]
    expect(op).toMatchObject({ slideIndex: 0, sourceId: 'tbl_1', row: 1, col: 0 })
    expect(op.paragraphs[0].runs[0]).toMatchObject({ text: 'Banana', bold: true })
    expect(applied.length).toBe(1)
  })

  it('fails when the table/cell is missing', async () => {
    ;(window as any).slidesApi.editTableCell = vi.fn(async () => null)
    const { access } = makeAccess()
    const skill = createSlidesSkill(access)
    const r = await skill.executeTool!(
      call('edit_table_cell', {
        slideIndex: 0,
        sourceId: 'nope',
        row: 9,
        col: 9,
        paragraphs: [{ text: 'x' }],
      }),
    )
    expect(r.isError).toBe(true)
  })
})

describe('edit_table_structure', () => {
  it('passes kind/index/before through to IPC', async () => {
    const { access, applied } = makeAccess()
    const skill = createSlidesSkill(access)
    const r = await skill.executeTool!(
      call('edit_table_structure', {
        slideIndex: 0,
        sourceId: 'tbl_1',
        kind: 'insert-row',
        index: 1,
        before: true,
      }),
    )
    expect(r.isError).toBeUndefined()
    expect((window as any).slidesApi.tableStructure.mock.calls[0][0]).toMatchObject({
      slideIndex: 0,
      sourceId: 'tbl_1',
      kind: 'insert-row',
      index: 1,
      before: true,
    })
    expect(applied.length).toBe(1)
  })

  it('rejects unknown kind without calling IPC', async () => {
    const { access } = makeAccess()
    const skill = createSlidesSkill(access)
    const r = await skill.executeTool!(
      call('edit_table_structure', {
        slideIndex: 0,
        sourceId: 'tbl_1',
        kind: 'merge-cells',
        index: 0,
      }),
    )
    expect(r.isError).toBe(true)
    expect((window as any).slidesApi.tableStructure).not.toHaveBeenCalled()
  })
})
