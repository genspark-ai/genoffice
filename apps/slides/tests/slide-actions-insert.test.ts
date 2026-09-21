import { beforeEach, describe, expect, it, vi } from 'vitest'
import { addSlide, addSlideAt, addSlideWithLayout } from '../src/renderer/slide-actions'
import type { ActionCtx } from '../src/renderer/action-context'

// export-render pulls node-canvas, which the jsdom test environment cannot load
vi.mock('../src/renderer/export-render', () => ({ renderSlidesToPngBase64: vi.fn() }))

const api = {
  addBlankSlide: vi.fn(async () => ({ slides: [{}, {}, {}], index: 2 })),
  addSlideWithLayout: vi.fn(async () => ({ slides: [{}, {}, {}], index: 2 })),
}

function makeCtx(): ActionCtx {
  return {
    slide: {},
    slides: [{}, {}],
    current: 1,
    selectedSlides: [0, 1],
    setSlides: vi.fn(),
    setCurrent: vi.fn(),
    setSelectedSlides: vi.fn(),
    setSelectedIds: vi.fn(),
    setEditing: vi.fn(),
    setDirty: vi.fn(),
  } as unknown as ActionCtx
}

describe('inserting a slide resets the rail multi-selection', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    ;(window as unknown as { slidesApi: typeof api }).slidesApi = api
  })

  it.each([
    ['addSlide', (ctx: ActionCtx) => addSlide(ctx)],
    ['addSlideAt', (ctx: ActionCtx) => addSlideAt(ctx, 2)],
    [
      'addSlideWithLayout',
      (ctx: ActionCtx) => addSlideWithLayout(ctx, 'ppt/slideLayouts/slideLayout1.xml'),
    ],
  ])('%s selects only the new slide', async (_name, run) => {
    const ctx = makeCtx()
    await run(ctx)
    expect(ctx.setCurrent).toHaveBeenCalledWith(2)
    expect(ctx.setSelectedSlides).toHaveBeenCalledWith([2])
  })
})
