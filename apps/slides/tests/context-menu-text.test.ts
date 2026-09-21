import { beforeEach, describe, expect, it, vi } from 'vitest'
import { buildCtxItems } from '../src/renderer/context-menu-items'
import * as styleActions from '../src/renderer/style-actions'
import * as insertActions from '../src/renderer/insert-actions'
import { restoreEditSelection } from '../src/renderer/TextEditOverlay'
import { t } from '../src/renderer/i18n/locale'
import type { ActionCtx } from '../src/renderer/action-context'

vi.mock('../src/renderer/konva-adapter', () => ({ isEditableText: () => false }))
vi.mock('../src/renderer/TextEditOverlay', () => ({ restoreEditSelection: vi.fn(() => true) }))
vi.mock('../src/renderer/clipboard-actions', () => ({}))
vi.mock('../src/renderer/slide-actions', () => ({}))
vi.mock('../src/renderer/show-actions', () => ({}))
vi.mock('../src/renderer/arrange-actions', () => ({}))
vi.mock('../src/renderer/insert-actions', () => ({ openLinkDialog: vi.fn() }))
vi.mock('../src/renderer/picture-edit-actions', () => ({}))
vi.mock('../src/renderer/table-actions', () => ({}))
vi.mock('../src/renderer/style-actions', () => ({
  onFormat: vi.fn(),
  onParagraphFormat: vi.fn(),
}))

const nativeClipboard = vi.fn()
const openFormat = vi.fn()

function makeTextCtx(collapsed: boolean, hasClipboard = true) {
  return {
    ctxMenu: { kind: 'text', x: 0, y: 0, collapsed },
    slides: [{}],
    sections: [],
    selectedIds: ['s1'],
    slide: null,
    current: 0,
    hasClipboard,
    canPasteSlide: false,
    editing: { sourceId: 's1' },
    editingCell: null,
    findNodeCtx: () => ({ node: { type: 'text' } }),
    openFormat,
  } as unknown as ActionCtx
}

const find = (ctx: ActionCtx, key: Parameters<typeof t>[0]) =>
  buildCtxItems(ctx).find((i) => i?.label === t(key))

describe('text edit context menu', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    ;(window as unknown as { slidesApi: unknown }).slidesApi = { nativeClipboard }
    document.execCommand = vi.fn()
  })

  it('offers only commands the editor already supports, in PowerPoint order', () => {
    const labels = buildCtxItems(makeTextCtx(false)).map((i) => i?.label ?? '|')
    expect(labels).toEqual([
      t('appCtxCut'),
      t('appCtxCopy'),
      t('appCtxPaste'),
      '|',
      t('ribbonBold'),
      t('ribbonItalic'),
      t('ribbonUnderline'),
      '|',
      t('ribbonBullets'),
      t('ribbonNumbering'),
      '|',
      t('appCtxHyperlink'),
      t('appCtxSelectAll'),
      '|',
      t('paneFormatTitleTyped', { type: t('paneFormatTextBox') }),
    ])
  })

  it('cut/copy need a ranged selection; paste needs clipboard content', () => {
    const caret = makeTextCtx(true, false)
    expect(find(caret, 'appCtxCut')?.disabled).toBe(true)
    expect(find(caret, 'appCtxCopy')?.disabled).toBe(true)
    expect(find(caret, 'appCtxPaste')?.disabled).toBe(true)
    const ranged = makeTextCtx(false)
    expect(find(ranged, 'appCtxCut')?.disabled).toBeFalsy()
    expect(find(ranged, 'appCtxCopy')?.disabled).toBeFalsy()
    expect(find(ranged, 'appCtxPaste')?.disabled).toBeFalsy()
  })

  it('clipboard commands go through the native clipboard so the edit session survives', () => {
    const ctx = makeTextCtx(false)
    find(ctx, 'appCtxCut')?.onClick?.()
    find(ctx, 'appCtxPaste')?.onClick?.()
    expect(nativeClipboard.mock.calls).toEqual([['cut'], ['paste']])
    expect(restoreEditSelection).toHaveBeenCalledTimes(2)
  })

  it('formatting restores the saved selection before dispatching the ribbon action', () => {
    const ctx = makeTextCtx(true)
    find(ctx, 'ribbonBold')?.onClick?.()
    expect(restoreEditSelection).toHaveBeenCalled()
    expect(styleActions.onFormat).toHaveBeenCalledWith('bold')
    find(ctx, 'ribbonNumbering')?.onClick?.()
    expect(styleActions.onParagraphFormat).toHaveBeenCalledWith(ctx, { bullet: 'number' })
    find(ctx, 'appCtxHyperlink')?.onClick?.()
    expect(insertActions.openLinkDialog).toHaveBeenCalledWith(ctx)
    find(ctx, 'appCtxSelectAll')?.onClick?.()
    expect(document.execCommand).toHaveBeenCalledWith('selectAll')
  })

  it('format entry names the edited element, also for a table cell edit', () => {
    const cell = {
      ...makeTextCtx(true),
      editing: null,
      editingCell: { sourceId: 'tbl', row: 0, col: 0 },
      findNodeCtx: () => ({ node: { type: 'table' } }),
    } as unknown as ActionCtx
    const item = buildCtxItems(cell).at(-1)
    expect(item?.label).toBe(t('paneFormatTitleTyped', { type: t('ribbonGroupTable') }))
    item?.onClick?.()
    expect(openFormat).toHaveBeenCalled()
  })

  it('a cell edit drops bullets/numbering (no cell-scoped path) but keeps selection commands', () => {
    const cell = {
      ...makeTextCtx(false),
      editing: null,
      editingCell: { sourceId: 'tbl', row: 0, col: 0 },
    } as unknown as ActionCtx
    const labels = buildCtxItems(cell).map((i) => i?.label ?? '|')
    expect(labels).not.toContain(t('ribbonBullets'))
    expect(labels).not.toContain(t('ribbonNumbering'))
    expect(labels.filter((l) => l === '|')).toHaveLength(3)
    expect(labels).toEqual(expect.arrayContaining([t('ribbonBold'), t('appCtxHyperlink')]))
  })
})
