import { beforeEach, describe, expect, it, vi } from 'vitest'
import type { RenderNode, RenderSlide } from '@genoffice/pptx-render'
import {
  firstPlaceholderId,
  nextPlaceholder,
  nextPlaceholderId,
} from '../src/renderer/placeholder-nav'
import type { ActionCtx, EditingState } from '../src/renderer/action-context'
import type { EditParagraph } from '../src/shared/ipc'

const text = (sourceId: string, placeholder?: string): RenderNode =>
  ({ type: 'text', sourceId, box: { x: 0, y: 0, w: 1, h: 1 }, placeholder }) as RenderNode
const shape = (sourceId: string, placeholder?: string): RenderNode =>
  ({ type: 'shape', sourceId, box: { x: 0, y: 0, w: 1, h: 1 }, placeholder }) as RenderNode
const connector = (sourceId: string): RenderNode =>
  ({ ...shape(sourceId, 'body'), line: { widthPx: 1 } }) as unknown as RenderNode
const picture = (sourceId: string): RenderNode =>
  ({ type: 'picture', sourceId, box: { x: 0, y: 0, w: 1, h: 1 } }) as unknown as RenderNode

const decoration = (sourceId: string, placeholder: string): RenderNode =>
  ({ ...text(sourceId, placeholder), decoration: true }) as RenderNode

const nodes = [
  text('title', 'title'),
  picture('pic'),
  text('free'),
  connector('conn'),
  shape('body', 'body'),
  text('foot', 'ftr'),
]

describe('placeholder order helpers', () => {
  it('follows spTree order and skips pictures, connectors and plain text boxes', () => {
    expect(firstPlaceholderId(nodes)).toBe('title')
    expect(nextPlaceholderId(nodes, 'title')).toBe('body')
  })

  it('moves from a non-placeholder to the next placeholder after it', () => {
    expect(nextPlaceholderId(nodes, 'free')).toBe('body')
  })

  it('returns null past the last placeholder or on a slide without any', () => {
    expect(nextPlaceholderId(nodes, 'body')).toBeNull()
    expect(firstPlaceholderId([picture('p'), text('t')])).toBeNull()
  })

  it('scans from the start for an unknown current id', () => {
    expect(nextPlaceholderId(nodes, 'nope')).toBe('title')
  })

  it('skips master/layout decorations even when they carry a title placeholder', () => {
    const withDeco = [decoration('m-title', 'title'), decoration('m-ftr', 'ftr'), ...nodes]
    expect(firstPlaceholderId(withDeco)).toBe('title')
    expect(nextPlaceholderId(withDeco, 'm-title')).toBe('title')
  })

  it('skips date, footer and slide-number placeholders on the slide itself', () => {
    const hf = [text('d', 'dt'), text('f', 'ftr'), text('n', 'sldNum'), text('b', 'body')]
    expect(firstPlaceholderId(hf)).toBe('b')
    expect(nextPlaceholderId(hf, 'b')).toBeNull()
    expect(nextPlaceholderId(nodes, 'body')).toBeNull()
  })
})

const paragraphs: EditParagraph[] = [{ runs: [{ text: 'hi' }] }] as unknown as EditParagraph[]
const newSlide = {
  nodes: [text('n-title', 'ctrTitle'), text('n-sub', 'subTitle')],
} as unknown as RenderSlide

const api = {
  editText: vi.fn(async () => ({ nodes }) as unknown as RenderSlide),
  deleteElement: vi.fn(async () => ({ nodes: [] }) as unknown as RenderSlide),
  addBlankSlide: vi.fn(async () => ({
    slides: [{ nodes }, newSlide] as unknown as RenderSlide[],
    index: 1,
  })),
}

function makeCtx(editing: EditingState): ActionCtx {
  return {
    editing,
    slide: { nodes },
    slides: [{ nodes }],
    current: 0,
    applySlide: vi.fn(),
    setSlides: vi.fn(),
    setCurrent: vi.fn(),
    setSelectedSlides: vi.fn(),
    setSelectedIds: vi.fn(),
    setEditing: vi.fn(),
    setDirty: vi.fn(),
  } as unknown as ActionCtx
}

describe('nextPlaceholder', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    ;(window as unknown as { slidesApi: typeof api }).slidesApi = api
  })

  it('commits changed text, then enters the next placeholder', async () => {
    const ctx = makeCtx({ sourceId: 'title' })
    await nextPlaceholder(ctx, paragraphs)
    expect(api.editText).toHaveBeenCalledWith({ slideIndex: 0, sourceId: 'title', paragraphs })
    expect(ctx.applySlide).toHaveBeenCalledWith(0, { nodes })
    expect(ctx.setEditing).toHaveBeenCalledWith({ sourceId: 'body' })
    expect(ctx.setSelectedIds).toHaveBeenCalledWith(['body'])
    expect(api.addBlankSlide).not.toHaveBeenCalled()
  })

  it('skips the commit when the text is unchanged', async () => {
    const ctx = makeCtx({ sourceId: 'title' })
    await nextPlaceholder(ctx, null)
    expect(api.editText).not.toHaveBeenCalled()
    expect(ctx.setEditing).toHaveBeenCalledWith({ sourceId: 'body' })
  })

  it('on the last placeholder inserts a same-layout slide and enters its first placeholder', async () => {
    const ctx = makeCtx({ sourceId: 'body' })
    await nextPlaceholder(ctx, null)
    expect(api.addBlankSlide).toHaveBeenCalledWith({
      sourceIndex: 0,
      fitWidthPx: expect.any(Number),
    })
    expect(ctx.setSlides).toHaveBeenCalledWith([{ nodes }, newSlide])
    expect(ctx.setCurrent).toHaveBeenCalledWith(1)
    expect(ctx.setSelectedSlides).toHaveBeenCalledWith([1])
    expect(ctx.setEditing).toHaveBeenCalledWith({ sourceId: 'n-title' })
    expect(ctx.setSelectedIds).toHaveBeenCalledWith(['n-title'])
    expect(ctx.setDirty).toHaveBeenCalledWith(true)
  })

  it('leaves edit mode with the shape selected when the new slide has no placeholder', async () => {
    api.addBlankSlide.mockResolvedValueOnce({
      slides: [{ nodes }, { nodes: [] }] as unknown as RenderSlide[],
      index: 1,
    })
    const ctx = makeCtx({ sourceId: 'body' })
    await nextPlaceholder(ctx, null)
    expect(ctx.setEditing).toHaveBeenCalledWith(null)
    expect(ctx.setSelectedIds).toHaveBeenCalledWith([])
  })

  it('drops an untouched click-to-type text box instead of committing it', async () => {
    const ctx = makeCtx({ sourceId: 'free', discardIfEmpty: true })
    await nextPlaceholder(ctx, null)
    expect(api.deleteElement).toHaveBeenCalledWith({ slideIndex: 0, sourceId: 'free' })
    expect(api.editText).not.toHaveBeenCalled()
    expect(ctx.setEditing).toHaveBeenCalledWith({ sourceId: 'body' })
  })

  it('commits a grouped child through its group and navigates from the group', async () => {
    const grouped = [
      { type: 'group', sourceId: 'g', children: [] } as unknown as RenderNode,
      ...nodes,
    ]
    const ctx = makeCtx({ sourceId: 'child', groupId: 'g' })
    ;(ctx as { slide: RenderSlide }).slide = { nodes: grouped } as unknown as RenderSlide
    await nextPlaceholder(ctx, paragraphs)
    expect(api.editText).toHaveBeenCalledWith({
      slideIndex: 0,
      sourceId: 'child',
      paragraphs,
      groupId: 'g',
    })
    expect(ctx.setEditing).toHaveBeenCalledWith({ sourceId: 'title' })
  })
})
