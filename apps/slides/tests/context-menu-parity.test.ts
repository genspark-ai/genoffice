import { beforeEach, describe, expect, it, vi } from 'vitest'
import { buildCtxItems } from '../src/renderer/context-menu-items'
import * as slideActions from '../src/renderer/slide-actions'
import * as arrangeActions from '../src/renderer/arrange-actions'
import * as pictureEditActions from '../src/renderer/picture-edit-actions'
import { defaultShapeStyle } from '../src/renderer/default-shape'
import { t } from '../src/renderer/i18n/locale'
import type { ActionCtx } from '../src/renderer/action-context'
import type { CtxItem } from '../src/renderer/components/ContextMenu'

vi.mock('../src/renderer/konva-adapter', () => ({
  isEditableText: (n: { type: string }) => n.type === 'shape' || n.type === 'text',
}))
vi.mock('../src/renderer/clipboard-actions', () => ({}))
vi.mock('../src/renderer/slide-actions', () => ({
  addSlide: vi.fn(),
  addSlideAt: vi.fn(),
  setSlideLayoutAt: vi.fn(),
}))
vi.mock('../src/renderer/show-actions', () => ({}))
vi.mock('../src/renderer/arrange-actions', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../src/renderer/arrange-actions')>()),
  reorderSelected: vi.fn(),
  groupSelected: vi.fn(),
  ungroupSelected: vi.fn(),
  regroupSelected: vi.fn(),
}))
vi.mock('../src/renderer/insert-actions', () => ({}))
vi.mock('../src/renderer/picture-edit-actions', () => ({
  replacePicture: vi.fn(),
  canSaveAsPicture: vi.fn(
    (ctx: { slide?: unknown; enteredGroupId?: string | null }, ids: string[]) =>
      !!ctx.slide && ids.length > 0 && !ctx.enteredGroupId,
  ),
  saveSelectionAsPicture: vi.fn(),
}))
vi.mock('../src/renderer/table-actions', () => ({}))
vi.mock('../src/renderer/style-actions', () => ({}))

const layouts = [
  { path: 'ppt/slideLayouts/slideLayout1.xml', name: 'Title Slide', layoutType: 'title' },
  { path: 'builtin:blank', name: 'Blank', layoutType: 'blank' },
]

const shape = {
  type: 'shape',
  sourceId: 's1',
  fill: { kind: 'solid', color: '#112233' },
  stroke: { color: '445566', widthPt: 2 },
}
const picture = { type: 'picture', sourceId: 'p1' }
const connector = { type: 'shape', sourceId: 'c1', line: { points: [] }, fill: { kind: 'none' } }

function makeCtx(over: Partial<Record<keyof ActionCtx, unknown>>) {
  return {
    slides: [{}, {}],
    sections: [],
    selectedIds: [],
    selectedSlides: [],
    ungroupedSets: [],
    slide: { partPath: 'ppt/slides/slide2.xml', nodes: [shape, picture, connector] },
    current: 1,
    hasClipboard: false,
    canPasteSlide: false,
    layouts,
    showRuler: false,
    showGrid: true,
    showGuides: false,
    toggleRuler: vi.fn(),
    toggleGrid: vi.fn(),
    toggleGuides: vi.fn(),
    setCurrent: vi.fn(),
    openBgFormat: vi.fn(),
    openFormat: vi.fn(),
    newComment: vi.fn(),
    startEdit: vi.fn(),
    openChangeShape: vi.fn(),
    findNodeCtx: () => null,
    ...over,
  } as unknown as ActionCtx
}

const labels = (items: Array<CtxItem | null>) => items.map((i) => i?.label ?? '|')
const find = (items: Array<CtxItem | null>, label: string) => items.find((i) => i?.label === label)

describe('canvas context menu (PowerPoint order)', () => {
  beforeEach(() => vi.clearAllMocks())

  it('lists cut/copy/paste, view toggles, new slide + layout, background and comment', () => {
    const ctx = makeCtx({ ctxMenu: { kind: 'canvas', x: 0, y: 0 } })
    const items = buildCtxItems(ctx)
    expect(labels(items)).toEqual([
      t('appCtxCut'),
      t('appCtxCopy'),
      t('appCtxPaste'),
      '|',
      t('ribbonRuler'),
      t('ribbonGridlines'),
      t('ribbonGuides'),
      '|',
      t('appCtxNewSlide'),
      t('ribbonLayout'),
      t('appCtxResetSlide'),
      '|',
      t('appCtxFormatBackground'),
      t('ribbonNewComment'),
    ])
    expect(find(items, t('appCtxCut'))?.disabled).toBe(true)
    expect(find(items, t('appCtxCopy'))?.disabled).toBe(true)
  })

  it('view toggles reflect the ribbon state and flip it', () => {
    const ctx = makeCtx({ ctxMenu: { kind: 'canvas', x: 0, y: 0 } })
    const items = buildCtxItems(ctx)
    expect(find(items, t('ribbonRuler'))?.checked).toBe(false)
    expect(find(items, t('ribbonGridlines'))?.checked).toBe(true)
    find(items, t('ribbonGuides'))?.onClick?.()
    expect(ctx.toggleGuides).toHaveBeenCalled()
  })

  it('Layout lists the deck layouts with localized built-in names and applies to the current slide', () => {
    const ctx = makeCtx({ ctxMenu: { kind: 'canvas', x: 0, y: 0 } })
    const layout = find(buildCtxItems(ctx), t('ribbonLayout'))!
    expect(layout.disabled).toBe(false)
    expect(labels(layout.sub!)).toEqual([t('ribbonLayoutTitleSlide'), t('ribbonLayoutBlank')])
    layout.sub![1]?.onClick?.()
    expect(slideActions.setSlideLayoutAt).toHaveBeenCalledWith(ctx, 1, 'builtin:blank')
    find(buildCtxItems(ctx), t('appCtxResetSlide'))?.onClick?.()
    expect(slideActions.setSlideLayoutAt).toHaveBeenCalledWith(ctx, 1)
  })

  it('Layout is disabled until the layout list has been fetched', () => {
    const ctx = makeCtx({ ctxMenu: { kind: 'canvas', x: 0, y: 0 }, layouts: null })
    expect(find(buildCtxItems(ctx), t('ribbonLayout'))?.disabled).toBe(true)
  })

  it('New Comment opens the comments composer', () => {
    const ctx = makeCtx({ ctxMenu: { kind: 'canvas', x: 0, y: 0 } })
    find(buildCtxItems(ctx), t('ribbonNewComment'))?.onClick?.()
    expect(ctx.newComment).toHaveBeenCalled()
  })
})

describe('thumbnail context menu (PowerPoint order)', () => {
  beforeEach(() => vi.clearAllMocks())

  it('orders clipboard, slide ops, section, layout, background, hide, comment', () => {
    const ctx = makeCtx({ ctxMenu: { kind: 'thumb', x: 0, y: 0, index: 0 } })
    expect(labels(buildCtxItems(ctx))).toEqual([
      t('appCtxCutSlide'),
      t('appCtxCopySlide'),
      t('appCtxPasteSlide'),
      t('appCtxPasteSlideKeepSource'),
      t('appCtxPasteSlideAsPicture'),
      '|',
      t('appCtxNewSlide'),
      t('appCtxDuplicateSlide'),
      t('appCtxDeleteSlide'),
      '|',
      t('appCtxAddSectionBefore'),
      '|',
      t('ribbonLayout'),
      t('appCtxResetSlide'),
      '|',
      t('appCtxChangeBgImage'),
      t('appCtxFormatBackground'),
      '|',
      t('appCtxHideSlide'),
      '|',
      t('ribbonNewComment'),
    ])
  })

  it('New Slide lands after the last selected slide with that slide as its source', () => {
    const ctx = makeCtx({
      ctxMenu: { kind: 'thumb', x: 0, y: 0, index: 0 },
      selectedSlides: [0, 1],
    })
    find(buildCtxItems(ctx), t('appCtxNewSlide'))?.onClick?.()
    expect(slideActions.addSlideAt).toHaveBeenCalledWith(ctx, 2)
  })

  it('layout and comment act on the clicked slide, not the current one', () => {
    const ctx = makeCtx({ ctxMenu: { kind: 'thumb', x: 0, y: 0, index: 0 } })
    find(buildCtxItems(ctx), t('ribbonLayout'))!.sub![0]?.onClick?.()
    expect(slideActions.setSlideLayoutAt).toHaveBeenCalledWith(
      ctx,
      0,
      'ppt/slideLayouts/slideLayout1.xml',
    )
    find(buildCtxItems(ctx), t('appCtxResetSlide'))?.onClick?.()
    expect(slideActions.setSlideLayoutAt).toHaveBeenCalledWith(ctx, 0)
    find(buildCtxItems(ctx), t('ribbonNewComment'))?.onClick?.()
    expect(ctx.setCurrent).toHaveBeenCalledWith(0)
    expect(ctx.newComment).toHaveBeenCalled()
  })
})

describe('shape context menu (PowerPoint order)', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    localStorage.clear()
  })

  const shapeCtx = () =>
    makeCtx({ ctxMenu: { kind: 'element', x: 0, y: 0, targetId: 's1' }, selectedIds: ['s1'] })

  it('follows PowerPoint up to Format, then keeps rotate/flip, align, comment and delete', () => {
    expect(labels(buildCtxItems(shapeCtx()))).toEqual([
      t('appCtxCut'),
      t('appCtxCopy'),
      t('appCtxPaste'),
      '|',
      t('appCtxEditText'),
      t('appCtxChangeShape'),
      t('appCtxEditPoints'),
      t('appCtxGroup'),
      t('appCtxBringToFront'),
      t('appCtxSendToBack'),
      t('appCtxSaveAsPicture'),
      t('appCtxHyperlink'),
      t('appCtxSetDefaultShape'),
      t('appCtxSizePosition'),
      t('paneFormatTitleTyped', { type: t('paneFormatShape') }),
      '|',
      t('appCtxRotateLeft90'),
      t('appCtxRotateRight90'),
      '|',
      t('appCtxFlipH'),
      t('appCtxFlipV'),
      '|',
      t('appCtxAlignLeft'),
      t('appCtxAlignCenterH'),
      t('appCtxAlignRight'),
      t('appCtxAlignTop'),
      t('appCtxAlignCenterV'),
      t('appCtxAlignBottom'),
      '|',
      t('ribbonNewComment'),
      '|',
      t('appCtxDelete'),
    ])
  })

  it('z-order lives in two flyouts', () => {
    const items = buildCtxItems(shapeCtx())
    const front = find(items, t('appCtxBringToFront'))!
    const back = find(items, t('appCtxSendToBack'))!
    expect(labels(front.sub!)).toEqual([t('appCtxBringToFront'), t('appCtxBringForward')])
    expect(labels(back.sub!)).toEqual([t('appCtxSendToBack'), t('appCtxSendBackward')])
    front.sub![1]?.onClick?.()
    expect(arrangeActions.reorderSelected).toHaveBeenCalledWith(expect.anything(), 's1', 'forward')
    back.sub![0]?.onClick?.()
    expect(arrangeActions.reorderSelected).toHaveBeenCalledWith(expect.anything(), 's1', 'back')
  })

  it('Group flyout enables Group for a multi-selection and Ungroup for a group', () => {
    const single = find(buildCtxItems(shapeCtx()), t('appCtxGroup'))!
    expect(single.disabled).toBe(true)
    expect(single.sub!.map((i) => i?.disabled)).toEqual([true, true, true])
    const multi = find(
      buildCtxItems(
        makeCtx({
          ctxMenu: { kind: 'element', x: 0, y: 0, targetId: 's1' },
          selectedIds: ['s1', 'p1'],
        }),
      ),
      t('appCtxGroup'),
    )!
    expect(multi.disabled).toBe(false)
    expect(multi.sub![0]?.disabled).toBe(false)
    multi.sub![0]?.onClick?.()
    expect(arrangeActions.groupSelected).toHaveBeenCalled()
    const grouped = find(
      buildCtxItems(
        makeCtx({
          ctxMenu: { kind: 'element', x: 0, y: 0, targetId: 'g1' },
          selectedIds: ['g1'],
          slide: { nodes: [{ type: 'group', sourceId: 'g1', children: [] }] },
        }),
      ),
      t('appCtxGroup'),
    )!
    expect(grouped.sub![1]?.disabled).toBe(false)
    grouped.sub![1]?.onClick?.()
    expect(arrangeActions.ungroupSelected).toHaveBeenCalled()
  })

  it('Regroup is offered only when a selected element remembers an ungrouped set', () => {
    const remembered = makeCtx({
      ctxMenu: { kind: 'element', x: 0, y: 0, targetId: 's1' },
      selectedIds: ['s1'],
      ungroupedSets: [{ slidePath: 'ppt/slides/slide2.xml', sourceIds: ['s1', 'p1'] }],
    })
    const group = find(buildCtxItems(remembered), t('appCtxGroup'))!
    expect(group.disabled).toBe(false)
    expect(group.sub!.map((i) => i?.label)).toEqual([
      t('appCtxGroup'),
      t('appCtxUngroup'),
      t('appCtxRegroup'),
    ])
    expect(group.sub![2]?.hint).toBeUndefined()
    expect(group.sub![2]?.disabled).toBe(false)
    group.sub![2]?.onClick?.()
    expect(arrangeActions.regroupSelected).toHaveBeenCalled()
    const otherSlide = find(
      buildCtxItems(
        makeCtx({
          ctxMenu: { kind: 'element', x: 0, y: 0, targetId: 's1' },
          selectedIds: ['s1'],
          ungroupedSets: [{ slidePath: 'ppt/slides/slide1.xml', sourceIds: ['s1', 'p1'] }],
        }),
      ),
      t('appCtxGroup'),
    )!
    expect(otherSlide.sub![2]?.disabled).toBe(true)
  })

  it('multi-selection hides single-only entries and disables z-order', () => {
    const items = buildCtxItems(
      makeCtx({
        ctxMenu: { kind: 'element', x: 0, y: 0, targetId: 's1' },
        selectedIds: ['s1', 'p1'],
      }),
    )
    expect(find(items, t('appCtxHyperlink'))).toBeUndefined()
    expect(find(items, t('appCtxSizePosition'))).toBeUndefined()
    expect(find(items, t('appCtxSetDefaultShape'))).toBeUndefined()
    expect(find(items, t('appCtxBringToFront'))?.disabled).toBe(true)
  })

  it('Save as Picture saves the whole selection, or the clicked element when nothing is selected', () => {
    const multi = makeCtx({
      ctxMenu: { kind: 'element', x: 0, y: 0, targetId: 's1' },
      selectedIds: ['s1', 'p1'],
    })
    const item = find(buildCtxItems(multi), t('appCtxSaveAsPicture'))!
    expect(item.disabled).toBe(false)
    item.onClick?.()
    expect(pictureEditActions.saveSelectionAsPicture).toHaveBeenCalledWith(multi, ['s1', 'p1'])
    const unselected = makeCtx({ ctxMenu: { kind: 'element', x: 0, y: 0, targetId: 'p1' } })
    find(buildCtxItems(unselected), t('appCtxSaveAsPicture'))?.onClick?.()
    expect(pictureEditActions.saveSelectionAsPicture).toHaveBeenLastCalledWith(unselected, ['p1'])
  })

  it('Save as Picture is disabled inside an entered group and absent from the table menu', () => {
    const entered = makeCtx({
      ctxMenu: { kind: 'element', x: 0, y: 0, targetId: 's1' },
      selectedIds: ['s1'],
      enteredGroupId: 'g1',
    })
    expect(find(buildCtxItems(entered), t('appCtxSaveAsPicture'))?.disabled).toBe(true)
    const table = makeCtx({
      ctxMenu: { kind: 'element', x: 0, y: 0, targetId: 't1', cell: { row: 0, col: 0 } },
      selectedIds: ['t1'],
      slide: { nodes: [{ type: 'table', sourceId: 't1', cells: [], rows: [], cols: [] }] },
    })
    expect(find(buildCtxItems(table), t('appCtxSaveAsPicture'))).toBeUndefined()
  })

  it('Size and Position opens the format pane on its size tab', () => {
    const ctx = shapeCtx()
    const items = buildCtxItems(ctx)
    find(items, t('appCtxSizePosition'))?.onClick?.()
    expect(ctx.openFormat).toHaveBeenCalledWith('size')
    find(items, t('paneFormatTitleTyped', { type: t('paneFormatShape') }))?.onClick?.()
    expect(ctx.openFormat).toHaveBeenLastCalledWith()
  })

  it('Set as Default Shape persists the fill and line new shapes start with', () => {
    expect(defaultShapeStyle()).toEqual({ fillColor: '#C43E1C' })
    find(buildCtxItems(shapeCtx()), t('appCtxSetDefaultShape'))?.onClick?.()
    expect(defaultShapeStyle()).toEqual({
      fillColor: '#112233',
      stroke: { color: '#445566', widthPt: 2 },
    })
  })

  it('connectors offer neither Change Shape, Set as Default Shape nor rotate/flip', () => {
    const items = buildCtxItems(
      makeCtx({ ctxMenu: { kind: 'element', x: 0, y: 0, targetId: 'c1' }, selectedIds: ['c1'] }),
    )
    expect(find(items, t('appCtxChangeShape'))).toBeUndefined()
    expect(find(items, t('appCtxSetDefaultShape'))).toBeUndefined()
    expect(find(items, t('appCtxRotateLeft90'))).toBeUndefined()
  })
})

describe('picture context menu', () => {
  beforeEach(() => vi.clearAllMocks())

  it('adds Change Picture before Crop and skips the shape-only entries', () => {
    const ctx = makeCtx({
      ctxMenu: { kind: 'element', x: 0, y: 0, targetId: 'p1' },
      selectedIds: ['p1'],
    })
    const items = buildCtxItems(ctx)
    const l = labels(items)
    expect(l.slice(l.indexOf(t('appCtxHyperlink')), l.indexOf('|', 4))).toEqual([
      t('appCtxHyperlink'),
      t('ribbonReplacePicture'),
      t('appCtxCropPicture'),
      t('appCtxRemoveBackground'),
      t('appCtxSizePosition'),
      t('paneFormatTitleTyped', { type: t('paneFormatPicture') }),
    ])
    expect(find(items, t('appCtxEditText'))).toBeUndefined()
    expect(find(items, t('appCtxSetDefaultShape'))).toBeUndefined()
    find(items, t('ribbonReplacePicture'))?.onClick?.()
    expect(pictureEditActions.replacePicture).toHaveBeenCalledWith(ctx)
  })
})
