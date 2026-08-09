/** AI picture tools: crop_image / set_picture_opacity / replace_image dispatch and guards. */
import { describe, it, expect, vi, beforeEach } from 'vitest'
import { createSlidesSkill, type DeckAccess } from '../src/renderer/ai/slides-skill'
import type { RenderSlide, PlacedBox } from '@genoffice/pptx-render'
import type { AgentToolCall } from '../src/shared/ipc'

const box = (x: number, y: number, w: number, h: number): PlacedBox => ({
  x,
  y,
  w,
  h,
  rotationDeg: 0,
  flipH: false,
  flipV: false,
  centerX: x + w / 2,
  centerY: y + h / 2,
})

const deck = {
  widthPx: 1280,
  heightPx: 720,
  nodes: [
    { id: 'pic1', sourceId: 'pic1', type: 'picture', box: box(100, 100, 300, 200) },
    {
      id: 'sh1',
      sourceId: 'sh1',
      type: 'shape',
      box: box(500, 100, 200, 100),
      fill: { kind: 'none' },
    },
    {
      id: 'g1',
      sourceId: 'g1',
      type: 'group',
      box: box(600, 400, 200, 200),
      children: [{ id: 'pic2', sourceId: 'pic2', type: 'picture', box: box(0, 0, 100, 100) }],
    },
  ],
} as unknown as RenderSlide

function mkAccess(): DeckAccess {
  return {
    getSlides: () => [deck],
    getCurrent: () => 0,
    getSelectedIds: () => [],
    applySlide: () => {},
    applyDeck: () => {},
    fitWidthPx: 1280,
    retryBackoffMs: 0,
  } as unknown as DeckAccess
}

const call = (name: string, input: Record<string, unknown>): AgentToolCall => ({
  id: 't',
  name,
  input: { slideIndex: 0, sourceId: 'pic1', ...input },
})

beforeEach(() => {
  ;(window as unknown as { slidesApi: unknown }).slidesApi = {
    editPictureSrcRect: vi.fn(async () => deck),
    editPictureOpacity: vi.fn(async () => deck),
    replacePictureUrl: vi.fn(async () => deck),
  }
})
const api = () =>
  (window as unknown as { slidesApi: Record<string, ReturnType<typeof vi.fn>> }).slidesApi

describe('crop_image', () => {
  it('applies a clamped srcRect and reports mutation', async () => {
    const r = await createSlidesSkill(mkAccess()).executeTool!(
      call('crop_image', { l: 0.1, t: 0, r: 0.25, b: 0 }),
    )
    expect(r.mutated).toBe(true)
    expect(api().editPictureSrcRect).toHaveBeenCalledWith({
      slideIndex: 0,
      sourceId: 'pic1',
      srcRect: { l: 0.1, t: 0, r: 0.25, b: 0 },
    })
  })

  it('all-zero fractions remove the crop (srcRect null)', async () => {
    await createSlidesSkill(mkAccess()).executeTool!(call('crop_image', { l: 0, t: 0, r: 0, b: 0 }))
    expect(api().editPictureSrcRect).toHaveBeenCalledWith(
      expect.objectContaining({ srcRect: null }),
    )
  })

  it('refuses a crop that removes the whole image', async () => {
    const r = await createSlidesSkill(mkAccess()).executeTool!(
      call('crop_image', { l: 0.6, t: 0, r: 0.6, b: 0 }),
    )
    expect(r.isError).toBe(true)
    expect(api().editPictureSrcRect).not.toHaveBeenCalled()
  })

  it('refuses non-picture targets and grouped pictures', async () => {
    const skill = createSlidesSkill(mkAccess())
    const r1 = await skill.executeTool!(
      call('crop_image', { sourceId: 'sh1', l: 0, t: 0, r: 0, b: 0 }),
    )
    expect(r1.isError).toBe(true)
    expect(r1.output).toContain('not a picture')
    const r2 = await skill.executeTool!(
      call('crop_image', { sourceId: 'pic2', l: 0, t: 0, r: 0, b: 0 }),
    )
    expect(r2.isError).toBe(true)
    expect(r2.output).toContain('group')
    expect(api().editPictureSrcRect).not.toHaveBeenCalled()
  })
})

describe('set_picture_opacity', () => {
  it('applies a valid opacity', async () => {
    const r = await createSlidesSkill(mkAccess()).executeTool!(
      call('set_picture_opacity', { opacity: 0.35 }),
    )
    expect(r.mutated).toBe(true)
    expect(api().editPictureOpacity).toHaveBeenCalledWith({
      slideIndex: 0,
      sourceId: 'pic1',
      opacity: 0.35,
    })
  })

  it('rejects out-of-range opacity', async () => {
    const r = await createSlidesSkill(mkAccess()).executeTool!(
      call('set_picture_opacity', { opacity: 1.5 }),
    )
    expect(r.isError).toBe(true)
    expect(api().editPictureOpacity).not.toHaveBeenCalled()
  })
})

describe('replace_image', () => {
  it('swaps in place and passes keepCrop through as keepSrcRect', async () => {
    const r = await createSlidesSkill(mkAccess()).executeTool!(
      call('replace_image', { url: 'https://example.com/a.png', keepCrop: true }),
    )
    expect(r.mutated).toBe(true)
    expect(api().replacePictureUrl).toHaveBeenCalledWith({
      slideIndex: 0,
      sourceId: 'pic1',
      url: 'https://example.com/a.png',
      keepSrcRect: true,
    })
  })

  it('rejects non-http urls', async () => {
    const r = await createSlidesSkill(mkAccess()).executeTool!(
      call('replace_image', { url: 'file:///etc/passwd' }),
    )
    expect(r.isError).toBe(true)
    expect(api().replacePictureUrl).not.toHaveBeenCalled()
  })
})
