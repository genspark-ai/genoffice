import { afterEach, describe, expect, it, vi } from 'vitest'
import type { CopyElementsOp } from '../src/shared/ipc'
import type { ActionCtx } from '../src/renderer/action-context'
import { copySelected } from '../src/renderer/clipboard-actions'
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
