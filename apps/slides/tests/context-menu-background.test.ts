import { beforeEach, describe, expect, it, vi } from 'vitest'
import { buildCtxItems } from '../src/renderer/context-menu-items'
import * as styleActions from '../src/renderer/style-actions'
import { t } from '../src/renderer/i18n/locale'
import type { ActionCtx } from '../src/renderer/action-context'

vi.mock('../src/renderer/konva-adapter', () => ({ isEditableText: () => false }))
vi.mock('../src/renderer/clipboard-actions', () => ({}))
vi.mock('../src/renderer/slide-actions', () => ({}))
vi.mock('../src/renderer/show-actions', () => ({}))
vi.mock('../src/renderer/arrange-actions', () => ({}))
vi.mock('../src/renderer/insert-actions', () => ({}))
vi.mock('../src/renderer/picture-edit-actions', () => ({}))
vi.mock('../src/renderer/table-actions', () => ({}))
vi.mock('../src/renderer/style-actions', () => ({
  onBackground: vi.fn().mockResolvedValue(undefined),
}))

function makeThumbCtx(index: number) {
  return {
    ctxMenu: { kind: 'thumb', x: 0, y: 0, index },
    slides: [{}, {}, {}],
    sections: [],
    selectedIds: [],
    slide: null,
    current: 0,
    hasClipboard: false,
    canPasteSlide: false,
    setCurrent: vi.fn(),
    openBgFormat: vi.fn(),
  } as unknown as ActionCtx & {
    setCurrent: ReturnType<typeof vi.fn>
    openBgFormat: ReturnType<typeof vi.fn>
  }
}

describe('thumbnail context-menu background entries', () => {
  beforeEach(() => vi.clearAllMocks())

  it('change-background-picture opens the picker for the clicked slide', () => {
    const ctx = makeThumbCtx(2)
    const item = buildCtxItems(ctx).find((i) => i?.label === t('appCtxChangeBgImage'))
    expect(item).toBeTruthy()
    item?.onClick?.()
    expect(ctx.setCurrent).toHaveBeenCalledWith(2)
    expect(styleActions.onBackground).toHaveBeenCalledWith(ctx, {
      kind: 'image',
      mode: 'stretch',
      pick: true,
      slideIndex: 2,
    })
  })

  it('format-background selects the clicked slide then opens the pane', () => {
    const ctx = makeThumbCtx(1)
    const item = buildCtxItems(ctx).find((i) => i?.label === t('appCtxFormatBackground'))
    expect(item).toBeTruthy()
    item?.onClick?.()
    expect(ctx.setCurrent).toHaveBeenCalledWith(1)
    expect(ctx.openBgFormat).toHaveBeenCalled()
  })
})
