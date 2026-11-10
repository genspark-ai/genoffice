import { beforeEach, describe, expect, it, vi } from 'vitest'
import { buildCtxItems } from '../src/renderer/context-menu-items'
import * as slideActions from '../src/renderer/slide-actions'
import * as clipboardActions from '../src/renderer/clipboard-actions'
import { t } from '../src/renderer/i18n/locale'
import type { ActionCtx } from '../src/renderer/action-context'

vi.mock('../src/renderer/konva-adapter', () => ({ isEditableText: () => false }))
vi.mock('../src/renderer/clipboard-actions', () => ({ pasteSlideAfter: vi.fn() }))
vi.mock('../src/renderer/slide-actions', () => ({ addSlideAt: vi.fn(), addSectionAt: vi.fn() }))
vi.mock('../src/renderer/show-actions', () => ({ startSlideShow: vi.fn() }))
vi.mock('../src/renderer/arrange-actions', () => ({}))
vi.mock('../src/renderer/insert-actions', () => ({}))
vi.mock('../src/renderer/picture-edit-actions', () => ({}))
vi.mock('../src/renderer/table-actions', () => ({}))
vi.mock('../src/renderer/style-actions', () => ({}))

function makeGapCtx(pos: number, canPasteSlide = true) {
  return {
    ctxMenu: { kind: 'gap', x: 0, y: 0, pos },
    slides: [{}, {}, {}],
    sections: [],
    selectedIds: [],
    slide: null,
    current: 0,
    hasClipboard: false,
    canPasteSlide,
  } as unknown as ActionCtx
}

const find = (ctx: ActionCtx, key: Parameters<typeof t>[0]) =>
  buildCtxItems(ctx).find((i) => i?.label === t(key))

describe('thumbnail rail blank-space context menu', () => {
  beforeEach(() => vi.clearAllMocks())

  it('new slide lands at the insertion point', () => {
    find(makeGapCtx(2), 'appCtxNewSlide')?.onClick?.()
    expect(slideActions.addSlideAt).toHaveBeenCalledWith(expect.anything(), 2)
  })

  it('paste goes after the slide above the gap, including above the first slide', () => {
    find(makeGapCtx(0), 'appCtxPasteSlide')?.onClick?.()
    expect(clipboardActions.pasteSlideAfter).toHaveBeenCalledWith(expect.anything(), -1)
    find(makeGapCtx(3), 'appCtxPasteSlideKeepSource')?.onClick?.()
    expect(clipboardActions.pasteSlideAfter).toHaveBeenCalledWith(expect.anything(), 2, 'source')
  })

  it('paste is disabled without a slide on the clipboard; cut/copy always are', () => {
    const ctx = makeGapCtx(1, false)
    expect(find(ctx, 'appCtxPasteSlide')?.disabled).toBe(true)
    expect(find(ctx, 'appCtxCutSlide')?.disabled).toBe(true)
    expect(find(ctx, 'appCtxCopySlide')?.disabled).toBe(true)
  })

  it('a section can start on the slide below the gap but not below the last slide', () => {
    const above = find(makeGapCtx(1), 'ribbonAddSection')
    expect(above?.disabled).toBeFalsy()
    above?.onClick?.()
    expect(slideActions.addSectionAt).toHaveBeenCalledWith(expect.anything(), 1)
    expect(find(makeGapCtx(3), 'ribbonAddSection')?.disabled).toBe(true)
  })
})
