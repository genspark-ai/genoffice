import { beforeEach, describe, expect, it, vi } from 'vitest'
import { handleGlobalKeydown } from '../src/renderer/keyboard-actions'
import { createPreviewTracker, deleteSelectedVertex } from '../src/renderer/edit-points-actions'
import * as clipboardActions from '../src/renderer/clipboard-actions'
import type { ActionCtx } from '../src/renderer/action-context'

vi.mock('../src/renderer/clipboard-actions', () => ({ deleteSelected: vi.fn() }))
vi.mock('../src/renderer/slide-actions', () => ({}))
vi.mock('../src/renderer/file-actions', () => ({}))
vi.mock('../src/renderer/arrange-actions', () => ({}))
vi.mock('../src/renderer/show-actions', () => ({}))
vi.mock('../src/renderer/style-actions', () => ({}))

const square = {
  type: 'shape',
  sourceId: 's',
  box: { x: 0, y: 0, w: 100, h: 100 },
  fill: { kind: 'none' },
  polygonPoints: [0, 0, 100, 0, 100, 100, 0, 100],
}

function makeCtx(over: Record<string, unknown> = {}): ActionCtx {
  return {
    slideShow: false,
    presenter: false,
    editing: null,
    selectedIds: ['s'],
    slide: { nodes: [square] },
    slides: [{}],
    current: 0,
    selectedSlides: [0],
    masterItems: null,
    inkTool: 'select',
    viewMode: 'normal',
    brushMode: null,
    enteredGroupId: null,
    overlayOpen: false,
    findNodeCtx: () => null,
    setSelectedIds: vi.fn(),
    setEditPointsTarget: vi.fn(),
    commitEditPoints: vi.fn(),
    editPointsTarget: null,
    ...over,
  } as unknown as ActionCtx
}

const keydown = (key: string) => new KeyboardEvent('keydown', { key, cancelable: true })

describe('Edit Points keys', () => {
  beforeEach(() => vi.clearAllMocks())

  it('Escape leaves the mode before touching the selection', () => {
    const ctx = makeCtx({ editPointsTarget: { sourceId: 's', vertex: 1 } })
    const e = keydown('Escape')
    handleGlobalKeydown(ctx, e)
    expect(e.defaultPrevented).toBe(true)
    expect(ctx.setEditPointsTarget).toHaveBeenCalledWith(null)
    expect(ctx.setSelectedIds).not.toHaveBeenCalled()
  })

  it('Delete removes the selected vertex instead of the shape', () => {
    const ctx = makeCtx({ editPointsTarget: { sourceId: 's', vertex: 1 } })
    handleGlobalKeydown(ctx, keydown('Delete'))
    expect(clipboardActions.deleteSelected).not.toHaveBeenCalled()
    expect(ctx.commitEditPoints).toHaveBeenCalledWith(
      's',
      {
        w: 100,
        h: 100,
        cmds: [
          { op: 'M', pts: [0, 0] },
          { op: 'L', pts: [100, 100] },
          { op: 'L', pts: [0, 100] },
          { op: 'Z', pts: [] },
        ],
      },
      false,
    )
    expect(ctx.setEditPointsTarget).toHaveBeenCalledWith({ sourceId: 's', vertex: null })
  })

  it('Delete is a no-op while a handle drag is in progress', () => {
    const ctx = makeCtx({ editPointsTarget: { sourceId: 's', vertex: 1, dragging: true } })
    const e = keydown('Delete')
    handleGlobalKeydown(ctx, e)
    expect(e.defaultPrevented).toBe(true)
    expect(ctx.commitEditPoints).not.toHaveBeenCalled()
    expect(ctx.setEditPointsTarget).not.toHaveBeenCalled()
    expect(clipboardActions.deleteSelected).not.toHaveBeenCalled()
  })

  it('Delete with no vertex selected still deletes the shape', () => {
    const ctx = makeCtx({ editPointsTarget: { sourceId: 's', vertex: null } })
    handleGlobalKeydown(ctx, keydown('Delete'))
    expect(clipboardActions.deleteSelected).toHaveBeenCalled()
  })

  it('refuses to drop below three vertices', () => {
    const tri = { ...square, polygonPoints: [0, 0, 100, 0, 50, 100] }
    const ctx = makeCtx({
      slide: { nodes: [tri] },
      editPointsTarget: { sourceId: 's', vertex: 0 },
    })
    deleteSelectedVertex(ctx)
    expect(ctx.commitEditPoints).not.toHaveBeenCalled()
  })
})

describe('preview tracker', () => {
  const commit = { slideIndex: 2, sourceId: 's', path: { w: 1, h: 1, cmds: [] } }

  it('holds the last previewed path until a final commit closes the gesture', () => {
    const tr = createPreviewTracker()
    expect(tr.flush()).toBeNull()
    tr.note(commit, true)
    const later = { ...commit, path: { w: 2, h: 2, cmds: [] } }
    tr.note(later, true)
    expect(tr.flush()).toBe(later)
    expect(tr.flush()).toBeNull()
  })

  it('flushes to the slide the previews went to, whatever page is current now', () => {
    const tr = createPreviewTracker()
    tr.note(commit, true)
    const flushed = tr.flush()!
    expect(flushed.slideIndex).toBe(2)
    expect(flushed).toBe(commit)
  })

  it('a final commit leaves nothing to flush', () => {
    const tr = createPreviewTracker()
    tr.note(commit, true)
    tr.note(commit, false)
    expect(tr.flush()).toBeNull()
  })
})
