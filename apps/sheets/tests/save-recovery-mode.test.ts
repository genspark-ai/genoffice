/**
 * A dirty workbook gets a crash-recovery copy through the
 * same save pipeline, but recovery mode must never touch the opened file, prompt,
 * or clear the journal — the whole point is that unsaved work survives a crash.
 */
import JSZip from 'jszip'
import { describe, expect, it, vi, beforeEach } from 'vitest'
import { applyCellEditsToXlsx } from '../src/gateway/xlsx-gateway'
import { handleSave, type SaveContext } from '../src/renderer/save-actions'
import { createEditJournal, recordSetRangeValues } from '../src/renderer/edit-journal'
import { buildEditFixture } from './fixture-builder'

const saveWorkbookEdits = vi.fn()
const writeWorkbookRecovery = vi.fn()

beforeEach(() => {
  saveWorkbookEdits.mockReset().mockResolvedValue({ canceled: true })
  writeWorkbookRecovery.mockReset().mockResolvedValue({ ok: true })
  ;(globalThis as unknown as { window: unknown }).window = {
    desktopApi: { saveWorkbookEdits, writeWorkbookRecovery },
  }
})

function ctxWith(opts: { dirty: boolean; needsSaveAs?: boolean; restoredFromRecovery?: boolean }): {
  ctx: SaveContext
  messages: string[]
  journal: ReturnType<typeof createEditJournal>
  overlay: Map<string, Map<string, { v?: string | number | boolean | null }>>
} {
  const journal = createEditJournal()
  if (opts.dirty) recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { v: 'edited' } } })
  const messages: string[] = []
  const overlay = new Map<string, Map<string, { v?: string | number | boolean | null }>>()
  return {
    messages,
    journal,
    overlay,
    ctx: {
      univerRef: { current: null },
      stashViewRestore: () => {},
      lazyWorkbookRef: {
        current: {
          editJournal: journal,
          recalc: {
            timer: null,
            generation: 0,
            failed: false,
            formulaCells: new Map(),
            overlay,
          },
          file: {
            sessionId: '11111111-1111-4111-8111-111111111111',
            needsSaveAs: !!opts.needsSaveAs,
            restoredFromRecovery: !!opts.restoredFromRecovery,
          },
        },
      } as never,
      setMessage: (m: string) => messages.push(m),
      openLazyWorkbook: () => {},
    },
  }
}

describe('handleSave recovery mode', () => {
  it('writes a recovery copy and never the opened file', async () => {
    const { ctx, messages } = ctxWith({ dirty: true })
    await handleSave(ctx, 'recovery')
    expect(writeWorkbookRecovery).toHaveBeenCalledTimes(1)
    expect(saveWorkbookEdits).not.toHaveBeenCalled()
    // silent: no status messages for a background copy
    expect(messages).toEqual([])
  })

  it('sends the same payload shape the save pipeline gets', async () => {
    const { ctx } = ctxWith({ dirty: true })
    await handleSave(ctx, 'recovery')
    const payload = writeWorkbookRecovery.mock.calls[0]![0] as Record<string, unknown>
    expect(payload.sessionId).toBe('11111111-1111-4111-8111-111111111111')
    // 'recovery' is a renderer-side mode; the main process still sees a normal save request
    expect(payload.mode).toBe('save')
    expect(Array.isArray(payload.edits)).toBe(true)
    expect((payload.edits as unknown[]).length).toBe(1)
    expect(payload).toHaveProperty('structuralOps')
    expect(payload).toHaveProperty('sparklineAdditions')
  })

  it('does nothing when there is nothing pending', async () => {
    const { ctx, messages } = ctxWith({ dirty: false })
    await handleSave(ctx, 'recovery')
    expect(writeWorkbookRecovery).not.toHaveBeenCalled()
    expect(messages).toEqual([])
  })

  it('a failed copy is swallowed (best-effort, never surfaces)', async () => {
    writeWorkbookRecovery.mockRejectedValue(new Error('disk full'))
    const { ctx, messages } = ctxWith({ dirty: true })
    await expect(handleSave(ctx, 'recovery')).resolves.toBeUndefined()
    expect(messages).toEqual([])
  })
})

describe('handleSave formula cache overlay', () => {
  it('never pairs an immediately saved replacement formula with the previous cached value', async () => {
    const { ctx, journal, overlay } = ctxWith({ dirty: false })
    recordSetRangeValues(journal, 'sheet-1', { 2: { 1: { f: '=2+2' } } })
    overlay.set(
      'sheet-1',
      new Map([
        // B3 used to be SUM(C1:C2), whose cached value was 5.
        ['2:1', { v: 5 }],
        ['0:2', { v: 7 }],
      ]),
    )
    let savedWorksheet = ''
    saveWorkbookEdits.mockImplementationOnce(async (payload) => {
      const request = payload as {
        edits: {
          row: number
          column: number
          writeValue: boolean
          value: string | number | boolean | null
          formula?: string
        }[]
        formulaValues: {
          sheetId: string
          row: number
          column: number
          value: string | number | boolean | null
        }[]
      }
      const mutation = await applyCellEditsToXlsx(
        await buildEditFixture(),
        request.edits.map((edit) => ({
          sheetName: 'Data',
          row: edit.row,
          column: edit.column,
          writeValue: edit.writeValue,
          cell: {
            value: edit.value,
            ...(edit.formula === undefined ? {} : { formula: edit.formula }),
          },
        })),
        [],
        [],
        undefined,
        [],
        [],
        [],
        [],
        [],
        null,
        [],
        [],
        [
          {
            sheetName: 'Data',
            cells: request.formulaValues.map(({ row, column, value }) => ({
              row,
              column,
              value,
            })),
          },
        ],
      )
      const zip = await JSZip.loadAsync(mutation.buffer)
      savedWorksheet = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
      return { canceled: true }
    })

    await handleSave(ctx, 'save')

    const payload = saveWorkbookEdits.mock.calls[0]![0] as {
      edits: { row: number; column: number; formula?: string }[]
      formulaValues: { row: number; column: number; value: unknown }[]
    }
    expect(payload.edits).toContainEqual(
      expect.objectContaining({ row: 2, column: 1, formula: '=2+2' }),
    )
    expect(payload.formulaValues).toEqual([{ sheetId: 'sheet-1', row: 0, column: 2, value: 7 }])
    const savedB3 = /<c r="B3"[^>]*>[\s\S]*?<\/c>/.exec(savedWorksheet)?.[0]
    expect(savedB3).toContain('<f>2+2</f>')
    expect(savedB3).not.toContain('<v>')
  })
})

describe('handleSave restored-recovery write-back', () => {
  it('a clean restored session still saves: the workbook bytes are the change', async () => {
    const { ctx } = ctxWith({ dirty: false, restoredFromRecovery: true })
    await handleSave(ctx, 'save')
    expect(saveWorkbookEdits).toHaveBeenCalledTimes(1)
    const payload = saveWorkbookEdits.mock.calls[0]![0] as Record<string, unknown>
    expect(payload.mode).toBe('save')
    expect(payload.restoreWriteBack).toBe(true)
  })

  it('a clean ordinary session still refuses a pointless save', async () => {
    const { ctx, messages } = ctxWith({ dirty: false })
    await handleSave(ctx, 'save')
    expect(saveWorkbookEdits).not.toHaveBeenCalled()
    expect(messages.length).toBe(1)
  })

  it('never flags recovery-mode writes as a restore write-back', async () => {
    const { ctx } = ctxWith({ dirty: true, restoredFromRecovery: true })
    await handleSave(ctx, 'recovery')
    expect(writeWorkbookRecovery).toHaveBeenCalledTimes(1)
    const payload = writeWorkbookRecovery.mock.calls[0]![0] as Record<string, unknown>
    expect(payload.restoreWriteBack).toBeUndefined()
  })
})
