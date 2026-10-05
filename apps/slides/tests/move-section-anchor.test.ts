import { beforeEach, describe, expect, it, vi } from 'vitest'
import { moveSectionDir } from '../src/renderer/slide-actions'
import type { ActionCtx } from '../src/renderer/action-context'
import type { RenderSlide } from '@genoffice/pptx-render'

// export-render pulls node-canvas, which the jsdom test environment cannot load
vi.mock('../src/renderer/export-render', () => ({ renderSlidesToPngBase64: vi.fn() }))

/** partPath is the stable slide identity across a reorder. */
const slide = (partPath: string): RenderSlide => ({ partPath, nodes: [] }) as unknown as RenderSlide

// section B (slides 2-3) moves up, landing after slide 0
const before = [slide('s1'), slide('s2'), slide('s3'), slide('s4')]
const after = [slide('s1'), slide('s3'), slide('s4'), slide('s2')]

type SectionInfo = { id: string; name: string; slideIndices: number[] }
type MoveSectionResult = { slides: RenderSlide[]; sections: SectionInfo[] }

const api = {
  moveSection: vi.fn(async (): Promise<MoveSectionResult | null> => ({
    slides: after,
    sections: [{ id: 'B', name: 'Body', slideIndices: [1, 2] }],
  })),
}

function makeCtx(current: number): ActionCtx {
  return {
    slide: before[current],
    slides: before,
    current,
    selectedSlides: [current],
    setSlides: vi.fn(),
    setSections: vi.fn(),
    setCurrent: vi.fn(),
    setSelectedSlides: vi.fn(),
    setSelectedIds: vi.fn(),
    setEditing: vi.fn(),
    setDirty: vi.fn(),
    setStatus: vi.fn(),
  } as unknown as ActionCtx
}

describe('moveSectionDir keeps current on the same slide', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    ;(window as unknown as { slidesApi: typeof api }).slidesApi = api
  })

  it('re-anchors current onto the moved slide', async () => {
    // slide 3 ('s4') ends up at index 2 after B moves above it
    const ctx = makeCtx(3)
    await moveSectionDir(ctx, 'B', 'up')
    expect(ctx.setSlides).toHaveBeenCalledWith(after)
    expect(ctx.setCurrent).toHaveBeenCalledWith(2)
  })

  it('re-anchors for a move down as well', async () => {
    const ctx = makeCtx(0)
    await moveSectionDir(ctx, 'B', 'down')
    // 's1' is untouched by the reorder, so the anchor is unchanged
    expect(ctx.setCurrent).toHaveBeenCalledWith(0)
  })

  it('leaves current alone when the slide has no stable part path', async () => {
    const ctx = makeCtx(3)
    ;(ctx.slides as RenderSlide[])[3] = { nodes: [] } as unknown as RenderSlide
    await moveSectionDir(ctx, 'B', 'up')
    expect(ctx.setSlides).toHaveBeenCalledWith(after)
    expect(ctx.setCurrent).not.toHaveBeenCalled()
  })

  it('does not touch the deck when the move is refused', async () => {
    api.moveSection.mockResolvedValueOnce(null)
    const ctx = makeCtx(3)
    await moveSectionDir(ctx, 'B', 'up')
    expect(ctx.setSlides).not.toHaveBeenCalled()
    expect(ctx.setCurrent).not.toHaveBeenCalled()
  })
})
