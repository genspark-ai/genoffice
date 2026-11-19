import { afterEach, describe, expect, it, vi } from 'vitest'
import type { SavePictureOp, SavePictureResult } from '../src/shared/ipc'
import type { ActionCtx } from '../src/renderer/action-context'
import { canSaveAsPicture, saveSelectionAsPicture } from '../src/renderer/picture-edit-actions'
import { renderSelectionToPngBase64 } from '../src/renderer/selection-image'

vi.mock('../src/renderer/selection-image', () => ({ renderSelectionToPngBase64: vi.fn() }))
vi.mock('../src/renderer/i18n/locale', () => ({
  t: (key: string, params?: Record<string, unknown>) =>
    params ? `${key}:${Object.values(params).join(',')}` : key,
}))

function setup(over: Partial<Record<keyof ActionCtx, unknown>> = {}) {
  const api = {
    savePicture: vi.fn(async (_op: SavePictureOp): Promise<SavePictureResult> => ({
      ok: true,
      path: '/out/x.png',
    })),
  }
  vi.stubGlobal('window', { slidesApi: api })
  const ctx = {
    current: 0,
    selectedIds: ['a', 'b'],
    enteredGroupId: null,
    slide: {
      nodes: [
        { type: 'shape', sourceId: 'a' },
        { type: 'picture', sourceId: 'b', name: 'Logo: v2/final' },
      ],
    },
    images: new Map(),
    setStatus: vi.fn(),
    ...over,
  } as unknown as ActionCtx
  vi.mocked(renderSelectionToPngBase64).mockResolvedValue('png-data')
  return { api, ctx }
}

afterEach(() => {
  vi.unstubAllGlobals()
  vi.resetAllMocks()
})

describe('saveSelectionAsPicture', () => {
  it('renders the whole selection to one PNG and saves it as Picture.png', async () => {
    const { api, ctx } = setup()
    await saveSelectionAsPicture(ctx)
    expect(renderSelectionToPngBase64).toHaveBeenCalledWith(ctx.slide, ['a', 'b'], ctx.images)
    expect(api.savePicture).toHaveBeenCalledWith({ pngBase64: 'png-data', defaultName: 'Picture' })
    expect(ctx.setStatus).toHaveBeenCalledWith('appStatusPictureSaved:/out/x.png')
  })

  it('names a single element after its own name with path characters removed', async () => {
    const { api, ctx } = setup()
    await saveSelectionAsPicture(ctx, ['b'])
    expect(renderSelectionToPngBase64).toHaveBeenCalledWith(ctx.slide, ['b'], ctx.images)
    expect(api.savePicture.mock.calls[0]![0].defaultName).toBe('Logo v2 final')
  })

  it('cancelling the dialog leaves the status alone', async () => {
    const { api, ctx } = setup()
    api.savePicture.mockResolvedValue({ ok: false })
    await saveSelectionAsPicture(ctx)
    expect(ctx.setStatus).not.toHaveBeenCalled()
  })

  it('reports write and render failures in the status bar', async () => {
    const { api, ctx } = setup()
    api.savePicture.mockResolvedValue({ ok: false, error: 'EACCES' })
    await saveSelectionAsPicture(ctx)
    expect(ctx.setStatus).toHaveBeenCalledWith('appStatusPictureSaveFailed:EACCES')
    vi.mocked(renderSelectionToPngBase64).mockRejectedValue(new Error('no bounds'))
    await saveSelectionAsPicture(ctx)
    expect(ctx.setStatus).toHaveBeenLastCalledWith('appStatusPictureSaveFailed:Error: no bounds')
    expect(api.savePicture).toHaveBeenCalledTimes(1)
  })

  it('is unavailable for an empty selection or inside an entered group', async () => {
    const { api, ctx } = setup({ enteredGroupId: 'g1' })
    expect(canSaveAsPicture(ctx, ['a'])).toBe(false)
    expect(canSaveAsPicture({ ...ctx, enteredGroupId: null } as ActionCtx, [])).toBe(false)
    expect(canSaveAsPicture({ ...ctx, enteredGroupId: null } as ActionCtx, ['a'])).toBe(true)
    await saveSelectionAsPicture(ctx)
    expect(renderSelectionToPngBase64).not.toHaveBeenCalled()
    expect(api.savePicture).not.toHaveBeenCalled()
  })
})
