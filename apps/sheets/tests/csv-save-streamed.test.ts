/**
 * Saving a CSV session that streams (too large for a full load): the
 * "keep this format?" question must come BEFORE the full-load guard —
 * a streamed session never reaches preloadComplete, so guarding first
 * dead-ended plain Save forever with the Save-As-.xlsx escape hatch
 * unreachable.
 */
import { beforeEach, describe, expect, it, vi } from 'vitest'
import { handleSave, type SaveContext } from '../src/renderer/save-actions'
import { createEditJournal, recordSetRangeValues } from '../src/renderer/edit-journal'

const saveWorkbookEdits = vi.fn()
const writeWorkbookRecovery = vi.fn()
const confirmCsvSave = vi.fn()

beforeEach(() => {
  saveWorkbookEdits.mockReset().mockResolvedValue({ canceled: true })
  writeWorkbookRecovery.mockReset().mockResolvedValue({ ok: true })
  confirmCsvSave.mockReset()
  ;(globalThis as unknown as { window: unknown }).window = {
    desktopApi: { saveWorkbookEdits, writeWorkbookRecovery, confirmCsvSave },
  }
})

// confirmedCsvSaves is remembered per path across handleSave calls, so every
// test uses its own csv path to stay independent.
let serial = 0

function csvCtx(opts: { preloadComplete: boolean }): { ctx: SaveContext; messages: string[] } {
  const journal = createEditJournal()
  recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { v: 'edited' } } })
  const messages: string[] = []
  serial += 1
  return {
    messages,
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
            overlay: new Map(),
          },
          flags: { preloadComplete: opts.preloadComplete },
          file: {
            sessionId: '11111111-1111-4111-8111-111111111111',
            needsSaveAs: false,
            restoredFromRecovery: false,
            csvPath: `/tmp/streamed-${serial}.csv`,
          },
        },
      } as never,
      setMessage: (message: string) => messages.push(message),
      openLazyWorkbook: () => {},
    },
  }
}

describe('handleSave on a streamed CSV session', () => {
  it('asks the format question and routes the xlsx choice to Save As', async () => {
    confirmCsvSave.mockResolvedValue('xlsx')
    const { ctx } = csvCtx({ preloadComplete: false })
    await handleSave(ctx, 'save')
    expect(confirmCsvSave).toHaveBeenCalledTimes(1)
    expect(saveWorkbookEdits).toHaveBeenCalledTimes(1)
    const payload = saveWorkbookEdits.mock.calls[0]![0] as Record<string, unknown>
    expect(payload.mode).toBe('save-as')
  })

  it('keeps CSV unavailable while streaming, without remembering the choice', async () => {
    confirmCsvSave.mockResolvedValue('csv')
    const { ctx, messages } = csvCtx({ preloadComplete: false })
    await handleSave(ctx, 'save')
    expect(saveWorkbookEdits).not.toHaveBeenCalled()
    expect(messages.length).toBe(1)
    // Not remembered: the next Save must offer the choice (and its escape
    // hatch) again instead of silently repeating the message.
    await handleSave(ctx, 'save')
    expect(confirmCsvSave).toHaveBeenCalledTimes(2)
  })

  it('cancel aborts without saving', async () => {
    confirmCsvSave.mockResolvedValue('cancel')
    const { ctx, messages } = csvCtx({ preloadComplete: false })
    await handleSave(ctx, 'save')
    expect(saveWorkbookEdits).not.toHaveBeenCalled()
    expect(messages.length).toBe(1)
  })

  it('remembers the CSV choice once the workbook is fully loaded', async () => {
    confirmCsvSave.mockResolvedValue('csv')
    const { ctx } = csvCtx({ preloadComplete: true })
    await handleSave(ctx, 'save')
    await handleSave(ctx, 'save')
    expect(confirmCsvSave).toHaveBeenCalledTimes(1)
  })
})
