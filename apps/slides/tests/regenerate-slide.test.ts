/** Skill-layer behavior of the regenerate_slide (redo one page in place) and delete_slide tools. */
import { describe, it, expect, vi, beforeEach } from 'vitest'
import { createSlidesSkill, type DeckAccess } from '../src/renderer/ai/slides-skill'
import type { RenderSlide } from '@genoffice/pptx-render'
import type { AgentToolCall } from '../src/shared/ipc'

const page = { widthPx: 1280, heightPx: 720, nodes: [] } as unknown as RenderSlide

function mkAccess(
  slides: RenderSlide[],
  overrides: Partial<DeckAccess> = {},
): DeckAccess & { applyDeck: ReturnType<typeof vi.fn> } {
  const applyDeck = vi.fn()
  return {
    getSlides: () => slides,
    getCurrent: () => 0,
    getSelectedIds: () => [],
    applySlide: () => {},
    applyDeck,
    fitWidthPx: 1280,
    isCloudPageGenEnabled: async () => true,
    ...overrides,
  } as unknown as DeckAccess & { applyDeck: ReturnType<typeof vi.fn> }
}

const call = (name: string, input: Record<string, unknown>): AgentToolCall => ({
  id: 't',
  name,
  input,
})

beforeEach(() => {
  ;(window as any).slidesApi = { deleteSlide: vi.fn(async () => [page, page]) }
})

const cloudOk = () => vi.fn(async () => ({ ok: true, marker: 'cloudpptx:/tmp/p.pptx' }))

describe('regenerate_slide', () => {
  it('brief → cloud generates marker → calls access.regenerateSlide to land it', async () => {
    const regenerateSlide = vi.fn(async () => ({ ok: true }))
    const generatePageCloud = cloudOk()
    const skill = createSlidesSkill(
      mkAccess([page, page], { regenerateSlide, generatePageCloud, retryBackoffMs: 0 }),
    )
    const r = await skill.executeTool!(
      call('regenerate_slide', {
        slideIndex: 1,
        brief: 'Redo as three-column cards, keep the title "NEW"',
      }),
    )
    expect(r.isError).toBeUndefined()
    expect(r.mutated).toBe(true)
    expect(generatePageCloud).toHaveBeenCalledOnce()
    expect(regenerateSlide).toHaveBeenCalledWith(1, 'cloudpptx:/tmp/p.pptx')
    expect(r.output).toContain('page 2')
  })

  it('slideIndex out of range → errors without invoking the pipeline', async () => {
    const regenerateSlide = vi.fn(async () => ({ ok: true }))
    const skill = createSlidesSkill(
      mkAccess([page], { regenerateSlide, generatePageCloud: cloudOk(), retryBackoffMs: 0 }),
    )
    const r = await skill.executeTool!(call('regenerate_slide', { slideIndex: 3, brief: 'x' }))
    expect(r.isError).toBe(true)
    expect(regenerateSlide).not.toHaveBeenCalled()
  })

  it('empty brief → errors', async () => {
    const skill = createSlidesSkill(
      mkAccess([page], {
        regenerateSlide: vi.fn(async () => ({ ok: true })),
        generatePageCloud: cloudOk(),
        retryBackoffMs: 0,
      }),
    )
    const r = await skill.executeTool!(call('regenerate_slide', { slideIndex: 0, brief: '' }))
    expect(r.isError).toBe(true)
  })

  it('cloud generation fails (after 1 retry) → error passed through', async () => {
    const generatePageCloud = vi.fn(async () => ({ ok: false, error: 'cloud timeout' }))
    const skill = createSlidesSkill(
      mkAccess([page], {
        regenerateSlide: vi.fn(async () => ({ ok: true })),
        generatePageCloud,
        retryBackoffMs: 0,
      }),
    )
    const r = await skill.executeTool!(call('regenerate_slide', { slideIndex: 0, brief: 'x' }))
    expect(r.isError).toBe(true)
    expect(r.output).toContain('cloud timeout')
    expect(generatePageCloud).toHaveBeenCalledTimes(2)
  })

  it('cloud unavailable → the local spec builder produces the marker and lands it', async () => {
    const regenerateSlide = vi.fn(async () => ({ ok: true }))
    const generatePageLocal = vi.fn(async () => ({ ok: true, marker: 'localpptx:/tmp/p.pptx' }))
    const skill = createSlidesSkill(
      mkAccess([page, page], {
        regenerateSlide,
        generatePageLocal,
        isCloudPageGenEnabled: async () => false,
        retryBackoffMs: 0,
      }),
    )
    const r = await skill.executeTool!(call('regenerate_slide', { slideIndex: 1, brief: 'redo' }))
    expect(r.isError).toBeUndefined()
    expect(generatePageLocal).toHaveBeenCalledOnce()
    expect(regenerateSlide).toHaveBeenCalledWith(1, 'localpptx:/tmp/p.pptx')
  })

  it('local generation fails twice → error passed through, page untouched', async () => {
    const generatePageLocal = vi.fn(async () => ({ ok: false, error: 'spec rejected' }))
    const regenerateSlide = vi.fn(async () => ({ ok: true }))
    const skill = createSlidesSkill(
      mkAccess([page], {
        regenerateSlide,
        generatePageLocal,
        isCloudPageGenEnabled: async () => false,
        retryBackoffMs: 0,
      }),
    )
    const r = await skill.executeTool!(call('regenerate_slide', { slideIndex: 0, brief: 'x' }))
    expect(r.isError).toBe(true)
    expect(r.output).toContain('spec rejected')
    expect(generatePageLocal).toHaveBeenCalledTimes(2)
    expect(regenerateSlide).not.toHaveBeenCalled()
  })

  it('local image failures show up in the redo report', async () => {
    const skill = createSlidesSkill(
      mkAccess([page], {
        regenerateSlide: vi.fn(async () => ({ ok: true })),
        generatePageLocal: vi.fn(async () => ({
          ok: true,
          marker: 'localpptx:/tmp/p.pptx',
          imageFailures: ['https://img.example/broken.jpg'],
        })),
        isCloudPageGenEnabled: async () => false,
        retryBackoffMs: 0,
      }),
    )
    const r = await skill.executeTool!(call('regenerate_slide', { slideIndex: 0, brief: 'x' }))
    expect(r.isError).toBeUndefined()
    expect(r.output).toContain('Missing images')
    expect(r.output).toContain('https://img.example/broken.jpg')
  })

  it('landing fails → error passed through with a retry hint', async () => {
    const skill = createSlidesSkill(
      mkAccess([page], {
        regenerateSlide: vi.fn(async () => ({ ok: false, error: 'conversion timeout' })),
        generatePageCloud: cloudOk(),
        retryBackoffMs: 0,
      }),
    )
    const r = await skill.executeTool!(call('regenerate_slide', { slideIndex: 0, brief: 'x' }))
    expect(r.isError).toBe(true)
    expect(r.output).toContain('conversion timeout')
  })

  it('htmlGenerated=true after success (native tools no longer blocked by the anti-handcrafting gate)', async () => {
    const skill = createSlidesSkill(
      mkAccess([page], {
        regenerateSlide: vi.fn(async () => ({ ok: true })),
        generatePageCloud: cloudOk(),
        retryBackoffMs: 0,
      }),
    )
    await skill.executeTool!(call('regenerate_slide', { slideIndex: 0, brief: 'x' }))
    ;(window as any).slidesApi.addElement = vi.fn(async () => ({ slide: page, sourceId: 'e1' }))
    const r = await skill.executeTool!(
      call('add_text_box', {
        slideIndex: 0,
        x: 1,
        y: 1,
        w: 10,
        h: 10,
        paragraphs: [{ text: 'x' }],
      }),
    )
    expect(r.isError).toBeUndefined()
  })
})

describe('delete_slide', () => {
  it('deletes the given slide and writes back via applyDeck', async () => {
    const access = mkAccess([page, page, page])
    const r = await createSlidesSkill(access).executeTool!(call('delete_slide', { slideIndex: 2 }))
    expect(r.isError).toBeUndefined()
    expect(r.mutated).toBe(true)
    expect((window as any).slidesApi.deleteSlide).toHaveBeenCalledWith(2)
    expect(access.applyDeck).toHaveBeenCalledOnce()
    expect(r.output).toContain('has 2 pages')
  })

  it('only one slide left → refused', async () => {
    const r = await createSlidesSkill(mkAccess([page])).executeTool!(
      call('delete_slide', { slideIndex: 0 }),
    )
    expect(r.isError).toBe(true)
    expect((window as any).slidesApi.deleteSlide).not.toHaveBeenCalled()
  })

  it('out of range → errors', async () => {
    const r = await createSlidesSkill(mkAccess([page, page])).executeTool!(
      call('delete_slide', { slideIndex: 5 }),
    )
    expect(r.isError).toBe(true)
  })
})
