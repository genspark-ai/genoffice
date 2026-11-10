import { describe, expect, it, vi } from 'vitest'
import { buildCtxItems } from '../src/renderer/context-menu-items'
import { t } from '../src/renderer/i18n/locale'
import type { ActionCtx } from '../src/renderer/action-context'
import type { CtxItem } from '../src/renderer/components/ContextMenu'

vi.mock('../src/renderer/konva-adapter', () => ({
  isEditableText: (n: { type: string }) => n.type === 'shape' || n.type === 'text',
}))
vi.mock('../src/renderer/clipboard-actions', () => ({}))
vi.mock('../src/renderer/slide-actions', () => ({}))
vi.mock('../src/renderer/show-actions', () => ({}))
vi.mock('../src/renderer/arrange-actions', () => ({ regroupCandidates: () => null }))
vi.mock('../src/renderer/insert-actions', () => ({}))
vi.mock('../src/renderer/picture-edit-actions', () => ({ canSaveAsPicture: () => false }))
vi.mock('../src/renderer/table-actions', () => ({}))
vi.mock('../src/renderer/style-actions', () => ({}))

const box = { x: 0, y: 0, w: 100, h: 50 }
const nodes = [
  {
    type: 'shape',
    sourceId: 'poly',
    box,
    fill: { kind: 'none' },
    polygonPoints: [0, 0, 1, 0, 1, 1],
  },
  {
    type: 'shape',
    sourceId: 'arc',
    box,
    fill: { kind: 'none' },
    fillPathData: 'M 0 0',
    strokePathData: 'M 0 0',
  },
  { type: 'text', sourceId: 'txt', box, fill: { kind: 'none' } },
  { type: 'shape', sourceId: 'line', box, fill: { kind: 'none' }, line: { points: [] } },
  { type: 'picture', sourceId: 'pic', box },
  { type: 'group', sourceId: 'grp', box, children: [] },
]

function ctxFor(id: string): ActionCtx {
  return {
    slides: [{}],
    sections: [],
    selectedIds: [id],
    selectedSlides: [],
    slide: { nodes },
    current: 0,
    hasClipboard: false,
    layouts: [],
    ctxMenu: { kind: 'element', x: 0, y: 0, targetId: id },
    setEditPointsTarget: vi.fn(),
    findNodeCtx: () => null,
  } as unknown as ActionCtx
}

const entry = (id: string): CtxItem | undefined =>
  buildCtxItems(ctxFor(id)).find((i) => i?.label === t('appCtxEditPoints')) ?? undefined

describe('Edit Points context-menu entry', () => {
  it('is enabled for single-path shapes and text boxes', () => {
    expect(entry('poly')?.disabled).toBe(false)
    expect(entry('txt')?.disabled).toBe(false)
  })

  it('is disabled for layered geometry and absent for connectors, pictures and groups', () => {
    expect(entry('arc')?.disabled).toBe(true)
    expect(entry('line')).toBeUndefined()
    expect(entry('pic')).toBeUndefined()
    expect(entry('grp')).toBeUndefined()
  })

  it('enters the mode on the clicked shape with no vertex selected', () => {
    const ctx = ctxFor('poly')
    buildCtxItems(ctx)
      .find((i) => i?.label === t('appCtxEditPoints'))
      ?.onClick?.()
    expect(ctx.setEditPointsTarget).toHaveBeenCalledWith({ sourceId: 'poly', vertex: null })
  })
})
