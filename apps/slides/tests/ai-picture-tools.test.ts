/** AI picture tool replace_image: dispatch and guards (crop/opacity moved to apply_ops setPictureSrcRect/setPictureOpacity). */
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

  it('rejects unknown url schemes', async () => {
    const r = await createSlidesSkill(mkAccess()).executeTool!(
      call('replace_image', { url: 'data:image/png;base64,AAAA' }),
    )
    expect(r.isError).toBe(true)
    expect(api().replacePictureUrl).not.toHaveBeenCalled()
  })

  it('forwards file:// urls — the main process resolves only the generated-image store', async () => {
    await createSlidesSkill(mkAccess()).executeTool!(
      call('replace_image', { url: 'file:///tmp/genoffice-ai-images/1234.png' }),
    )
    expect(api().replacePictureUrl).toHaveBeenCalledWith(
      expect.objectContaining({ url: 'file:///tmp/genoffice-ai-images/1234.png' }),
    )
  })
})
