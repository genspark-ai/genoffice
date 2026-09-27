import { afterEach, describe, expect, it, vi } from 'vitest'
import type { CopyElementsOp, DeleteElementsOp } from '../src/shared/ipc'
import type { ActionCtx } from '../src/renderer/action-context'
import {
  copySelected,
  cutSelected,
  deleteSelected,
  pasteParagraphs,
} from '../src/renderer/clipboard-actions'
import { renderSelectionToPngBase64 } from '../src/renderer/selection-image'

vi.mock('../src/renderer/selection-image', () => ({ renderSelectionToPngBase64: vi.fn() }))
vi.mock('../src/renderer/export-render', () => ({ renderSlidesToPngBase64: vi.fn() }))
vi.mock('../src/renderer/i18n/locale', () => ({ t: (key: string) => key }))

function setup() {
  const api = {
    copyElements: vi.fn(async (_op: CopyElementsOp) => 2),
    copyElementsImage: vi.fn(async () => true),
  }
  vi.stubGlobal('window', { slidesApi: api })
  const ctx = {
    current: 0,
    selectedIds: ['a', 'b'],
    slide: { nodes: [], widthPx: 1280, heightPx: 720 },
    images: new Map(),
    setHasClipboard: vi.fn(),
    setStatus: vi.fn(),
  } as unknown as ActionCtx
  return { api, ctx }
}

afterEach(() => {
  vi.unstubAllGlobals()
  vi.resetAllMocks()
})

describe('copy selected slide elements', () => {
  it('copies editable data immediately, then attaches the PNG to the same clipboard token', async () => {
    const { api, ctx } = setup()
    let complete!: (png: string) => void
    vi.mocked(renderSelectionToPngBase64).mockReturnValue(
      new Promise((r) => {
        complete = r
      }),
    )
    const copying = copySelected(ctx)
    await vi.waitFor(() => expect(renderSelectionToPngBase64).toHaveBeenCalled())
    const op = api.copyElements.mock.calls[0]![0]
    expect(op).toMatchObject({
      slideIndex: 0,
      sourceIds: ['a', 'b'],
      clipboardToken: expect.any(String),
    })
    expect(ctx.setHasClipboard).toHaveBeenCalledWith(true)
    expect(api.copyElementsImage).not.toHaveBeenCalled()
    complete('png-data')
    await copying
    expect(api.copyElementsImage).toHaveBeenCalledWith(op.clipboardToken, 'png-data')
  })

  it('keeps editable copying available when rendering fails', async () => {
    const { api, ctx } = setup()
    vi.mocked(renderSelectionToPngBase64).mockRejectedValue(new Error('image failed'))
    await expect(copySelected(ctx)).resolves.toBeUndefined()
    expect(ctx.setHasClipboard).toHaveBeenCalledWith(true)
    expect(api.copyElementsImage).not.toHaveBeenCalled()
  })

  it('does not render or modify the clipboard with no copied elements', async () => {
    const { api, ctx } = setup()
    api.copyElements.mockResolvedValue(0)
    await copySelected(ctx)
    expect(renderSelectionToPngBase64).not.toHaveBeenCalled()
    expect(api.copyElementsImage).not.toHaveBeenCalled()
    expect(ctx.setHasClipboard).not.toHaveBeenCalled()
  })
})

describe('delete selected slide elements', () => {
  function deleteSetup(selectedIds: string[]) {
    const api = {
      copyElements: vi.fn(async () => selectedIds.length),
      deleteElements: vi.fn(async (_op: DeleteElementsOp) => ({ nodes: [] })),
    }
    vi.stubGlobal('window', { slidesApi: api })
    const ctx = {
      current: 3,
      selectedIds,
      applySlide: vi.fn(),
      setSelectedIds: vi.fn(),
      setHasClipboard: vi.fn(),
      setStatus: vi.fn(),
    } as unknown as ActionCtx
    return { api, ctx }
  }

  it('deletes the whole selection in one call and applies the slide once', async () => {
    const { api, ctx } = deleteSetup(['a', 'b', 'c'])
    await deleteSelected(ctx)
    expect(api.deleteElements).toHaveBeenCalledTimes(1)
    expect(api.deleteElements).toHaveBeenCalledWith({ slideIndex: 3, sourceIds: ['a', 'b', 'c'] })
    expect(ctx.applySlide).toHaveBeenCalledTimes(1)
    expect(ctx.applySlide).toHaveBeenCalledWith(3, { nodes: [] })
    expect(ctx.setSelectedIds).toHaveBeenCalledWith([])
  })

  it('clears the selection without an IPC call when nothing is selected', async () => {
    const { api, ctx } = deleteSetup([])
    await deleteSelected(ctx)
    expect(api.deleteElements).not.toHaveBeenCalled()
    expect(ctx.setSelectedIds).toHaveBeenCalledWith([])
  })

  it('leaves the page alone when the main process refuses the batch', async () => {
    const { api, ctx } = deleteSetup(['a'])
    api.deleteElements.mockResolvedValue(null as never)
    await deleteSelected(ctx)
    expect(ctx.applySlide).not.toHaveBeenCalled()
    expect(ctx.setSelectedIds).toHaveBeenCalledWith([])
  })

  it('cut copies, deletes as one batch, then reports the cut count', async () => {
    const { api, ctx } = deleteSetup(['a', 'b'])
    await cutSelected(ctx)
    expect(api.copyElements).toHaveBeenCalledWith({
      slideIndex: 3,
      sourceIds: ['a', 'b'],
      cut: true,
    })
    expect(api.deleteElements).toHaveBeenCalledTimes(1)
    expect(ctx.setStatus).toHaveBeenCalledWith('appStatusCut')
  })
})

describe('pasteParagraphs', () => {
  it('maps lines to single-run paragraphs unchanged', () => {
    expect(pasteParagraphs('a\nb\r\nc')).toEqual({
      paragraphs: [{ runs: [{ text: 'a' }] }, { runs: [{ text: 'b' }] }, { runs: [{ text: 'c' }] }],
      truncated: false,
    })
  })

  it('keeps a large but ordinary paste intact', () => {
    const out = pasteParagraphs(`${'x'.repeat(200)}\n`.repeat(4000))
    expect(out.truncated).toBe(false)
    expect(out.paragraphs).toHaveLength(4001)
  })

  it('caps hostile clipboard text and reports the trim', () => {
    const start = Date.now()
    const out = pasteParagraphs(`${'x'.repeat(100000)}\n`.repeat(5000))
    expect(Date.now() - start).toBeLessThan(5000)
    expect(out.truncated).toBe(true)
    expect(out.paragraphs.length).toBeLessThanOrEqual(50_000)
    const chars = out.paragraphs.flatMap((p) => p.runs.map((r) => r.text)).join('').length
    expect(chars).toBeLessThanOrEqual(1_000_000)
    expect(pasteParagraphs('a\n'.repeat(60_000)).truncated).toBe(true)
  })
})
